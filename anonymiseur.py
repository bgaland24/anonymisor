#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
anonymiseur.py — Anonymisation par lot .txt / .md / .docx / .pptx / .xlsx / .pdf

Objectif : remplacer des noms propres identifiants (organisations, personnes,
prestataires, produits) par des pseudonymes/jetons définis dans un fichier de
correspondances, supprimer les identifiants techniques par motif (emails, IP,
URLs, ports, chemins), et nettoyer les métadonnées / commentaires / suivi de
modifications des fichiers Office.

ZÉRO dépendance externe : bibliothèque standard Python 3.8+ uniquement.

DEUX MODES
==========
1) PRÉ-REMPLISSAGE (--scan)
   Scanne le dossier, repère des candidats (acronymes en MAJUSCULES, séquences
   capitalisées) et écrit un fichier de correspondances pré-rempli à compléter.
   Ne modifie AUCUN document.

2) ANONYMISATION (par défaut)
   Applique le fichier de correspondances + les motifs techniques, écrit les
   documents nettoyés dans le dossier de sortie, produit un rapport CSV de tous
   les remplacements, puis lance une passe de vérification qui signale tout
   terme interdit ayant survécu.

LIMITES HONNÊTES
================
- Un terme coupé entre deux "runs" Word (ex. "SEM"+"AE") peut échapper au
  remplacement : la passe de vérification le signalera. Ouvrir/réenregistrer le
  fichier dans Word avant traitement réduit ce risque.
- L'acceptation du suivi de modifications est best-effort (regex). Pour une
  garantie totale : accepter les modifications dans Word, ou convertir via
  LibreOffice, AVANT de passer le script.
- Les .doc / .ppt hérités (binaires) ne sont PAS traités : les convertir
  d'abord en .docx / .pptx (LibreOffice : soffice --convert-to docx fichier.doc).
- Les .pdf sont convertis en .docx via Microsoft Word (dépendance OPTIONNELLE
  pywin32 + Word installé) puis durcis comme un .docx. Un PDF SCANNÉ (image de
  texte) donnera une sortie « blanche » : son texte est dans les pixels, non
  anonymisable sans OCR.
- Ce script neutralise les noms ; il ne neutralise PAS le contexte métier
  (secteur, taille, budget, stack legacy) sauf si tu ajoutes ces expressions
  dans le fichier de correspondances. Le contexte seul peut rester identifiant.
"""

import argparse
import base64
import csv
import glob
import os
import re
import shutil
import sys
import types
import zipfile
from collections import Counter

TEXT_EXT = {".txt", ".md", ".markdown"}
OFFICE_EXT = {".docx", ".pptx", ".xlsx"}
PDF_EXT = {".pdf"}
LEGACY_EXT = {".doc", ".ppt", ".xls"}

# Fichiers qui n'ont pas pu être lus (verrouillés, OneDrive hors-ligne, corrompus).
# Rempli au fil du traitement, affiché en récapitulatif à la fin.
NON_LUS = []


def _recap_non_lus():
    """Affiche le récapitulatif des fichiers ignorés (s'il y en a)."""
    if not NON_LUS:
        return
    uniques = sorted(set(NON_LUS))
    print(f"\n⚠️  {len(uniques)} fichier(s) NON LU(s) — ignoré(s) et ABSENT(s) du "
          f"résultat :")
    for c in uniques:
        print(f"    - {c}")
    print("    -> Ferme-les dans Office (ou télécharge-les depuis OneDrive : clic "
          "droit → « Toujours conserver sur cet appareil »), puis relance.")

# --------------------------------------------------------------------------- #
#  Fichier de correspondances
# --------------------------------------------------------------------------- #
def charger_correspondances(chemin):
    """Lit un TSV 'original<TAB>remplacement'. Lignes vides / '#...' ignorées."""
    paires = []
    # utf-8-sig : tolère un éventuel BOM ajouté par Excel / le Bloc-notes /
    # PowerShell lors de l'édition manuelle du fichier de correspondances.
    with open(chemin, encoding="utf-8-sig") as f:
        for n, ligne in enumerate(f, 1):
            ligne = ligne.rstrip("\n").rstrip("\r")
            if not ligne.strip() or ligne.lstrip().startswith("#"):
                continue
            if "\t" in ligne:
                orig, repl = ligne.split("\t", 1)
            else:
                # Tolérance : pas de tabulation, mais un remplacement entre
                # crochets en fin de ligne (ex. « plants [PRODUIT2] »). On
                # coupe juste avant ce dernier jeton [..].
                m = re.match(r"^(.+?)\s+(\[[^\]]+\])\s*$", ligne)
                if not m:
                    print(f"  [avert] ligne {n} ignorée (ni tabulation, ni "
                          f"remplacement [..] en fin de ligne) : {ligne!r}")
                    continue
                orig, repl = m.group(1), m.group(2)
            orig, repl = orig.strip(), repl.strip()
            if orig == "" or repl == "":
                print(f"  [avert] ligne {n} ignorée (champ vide) : {ligne!r}")
                continue
            # Préfixe "re:" => le terme de gauche est une EXPRESSION RÉGULIÈRE
            # (ni échappée, ni encadrée par \b). Pratique pour les URLs, etc.
            est_regex = False
            if orig.startswith("re:"):
                est_regex = True
                orig = orig[3:].strip()
                if orig == "":
                    print(f"  [avert] ligne {n} ignorée (motif regex vide) : {ligne!r}")
                    continue
            paires.append((orig, repl, est_regex))
    # Plus long d'abord : "Marketo Engage" avant "Marketo", "ODK Cloud" avant "ODK"
    paires.sort(key=lambda p: len(p[0]), reverse=True)
    return paires


# Groupes de lettres équivalentes (base + variantes accentuées, deux casses)
# pour rendre les termes LITTÉRAUX insensibles aux accents : "Eric" attrape
# aussi "Éric", "Qualité" attrape "qualite", etc.
_GROUPES_ACCENTS = [
    "aàáâãäåAÀÁÂÃÄÅ", "eéèêëEÉÈÊË", "iíìîïIÍÌÎÏ",
    "oóòôõöOÓÒÔÕÖ", "uúùûüUÚÙÛÜ", "cçCÇ", "yÿYŸ", "nñNÑ",
]
_CHAR2CLASSE = {}
for _g in _GROUPES_ACCENTS:
    _classe = "[" + _g + "]"
    for _ch in _g:
        _CHAR2CLASSE[_ch] = _classe


def _motif_litteral(orig):
    """Patron regex accent-insensible pour un terme littéral (sans \\b)."""
    return "".join(_CHAR2CLASSE.get(ch, re.escape(ch)) for ch in orig)


def compiler_motifs_termes(paires, sensible_casse):
    """Compile un motif regex par terme (avec frontières de mot)."""
    flags = 0 if sensible_casse else re.IGNORECASE
    compiles = []
    for orig, repl, est_regex in paires:
        if est_regex:
            try:
                motif = re.compile(orig, flags)
            except re.error as e:
                print(f"  [avert] motif regex invalide ignoré : {orig!r} ({e})")
                continue
        else:
            # \b fonctionne avec les lettres accentuées (str unicode en Python 3) ;
            # _motif_litteral rend en plus le terme insensible aux accents.
            motif = re.compile(r"\b" + _motif_litteral(orig) + r"\b", flags)
        compiles.append((orig, motif, repl, est_regex))
    return compiles


# --------------------------------------------------------------------------- #
#  Motifs techniques (identifiants par forme)
# --------------------------------------------------------------------------- #
def construire_scrubbers(args):
    s = []
    if not args.no_email:
        s.append(("email", re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"),
                  "[EMAIL]"))
    if not args.no_url:
        s.append(("url", re.compile(r"https?://[^\s<>\"')]+"), "[URL]"))
    if not args.no_port:
        # Conservateur : seulement après localhost / IP, ou "port NNNN".
        # Placé AVANT le scrub IP pour capter "192.168.0.1:8080" -> "[IP]:[PORT]".
        s.append(("port", re.compile(r"\b(localhost|(?:\d{1,3}\.){3}\d{1,3}):\d{1,5}\b"),
                  r"\1:[PORT]"))
        # "port 432" -> "port [PORT]" ; "port https 9001" -> "port https [PORT]"
        # (un nom de protocole optionnel entre "port" et le numéro est conservé).
        s.append(("port_mot",
                  re.compile(r"\bport\s+((?:https?|ssh|ftps?|sftp|smtps?|imaps?|"
                             r"pop3?|rdp|tcp|udp|tls|ssl|dns|ldaps?)\s+)?\d{1,5}\b",
                             re.IGNORECASE),
                  r"port \1[PORT]"))
    if not args.no_mac:
        # Adresse MAC (6 octets hexa séparés par : ou -). Placé AVANT l'IPv6 car
        # le motif IPv6 capturerait sinon une MAC du type 00:1A:2B:3C:4D:5E.
        s.append(("mac",
                  re.compile(r"\b(?:[0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}\b"),
                  "[MAC]"))
    if not args.no_ip:
        s.append(("ipv4", re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b"), "[IP]"))
        s.append(("ipv6", re.compile(r"\b(?:[A-Fa-f0-9]{1,4}:){2,7}[A-Fa-f0-9]{1,4}\b"),
                  "[IPV6]"))
    if not args.no_path:
        s.append(("chemin_unc", re.compile(r"\\\\[^\s<>\"']+"), "[CHEMIN]"))
        s.append(("chemin_win", re.compile(r"\b[A-Za-z]:\\[^\s<>\"']+"), "[CHEMIN]"))
        s.append(("chemin_mnt", re.compile(r"/(?:mnt|home|srv|opt|var)/[^\s<>\"']+"),
                  "[CHEMIN]"))
    return s


# --------------------------------------------------------------------------- #
#  Application des remplacements sur une chaîne de texte
# --------------------------------------------------------------------------- #
def remplacer_texte(texte, termes, scrubbers, compteur):
    """Ordre d'application, dans cet ordre précis :
      1) scrubber EMAIL — un email est un bloc atomique : on le neutralise EN
         PREMIER, pour qu'aucune règle ne casse son domaine
         (« x@semae.fr » doit devenir « [EMAIL] », pas « x@[ORG-S].fr ») ;
      2) règles REGEX du fichier (re:) — ex. URLs internes ;
      3) autres scrubbers (url, ip, mac, port, chemin) ;
      4) termes EXACTS du fichier.
    """
    # 1) Email d'abord (bloc atomique)
    for nom, motif, repl in scrubbers:
        if nom != "email":
            continue
        texte, n = motif.subn(repl, texte)
        if n:
            compteur[f"<{nom}>"] += n
    # 2) Règles regex (re:)
    for orig, motif, repl, est_regex in termes:
        if not est_regex:
            continue
        texte, n = motif.subn(repl.replace("\\", "\\\\"), texte)
        if n:
            compteur[orig] += n
    # 3) Autres scrubbers techniques (url, ip, mac, port, chemin)
    for nom, motif, repl in scrubbers:
        if nom == "email":
            continue
        texte, n = motif.subn(repl, texte)
        if n:
            compteur[f"<{nom}>"] += n
    # 4) Termes exacts (littéraux)
    for orig, motif, repl, est_regex in termes:
        if est_regex:
            continue
        texte, n = motif.subn(repl.replace("\\", "\\\\"), texte)
        if n:
            compteur[orig] += n
    return texte


# --------------------------------------------------------------------------- #
#  Fichiers texte simples
# --------------------------------------------------------------------------- #
def traiter_texte(src, dst, termes, scrubbers):
    compteur = Counter()
    try:
        with open(src, encoding="utf-8", errors="replace") as f:
            contenu = f.read()
    except OSError as e:
        print(f"  [ignoré] {os.path.basename(src)} illisible (ouvert dans une appli "
              f"ou OneDrive hors-ligne ?) : {e.strerror or e}")
        NON_LUS.append(src)
        return None
    contenu = remplacer_texte(contenu, termes, scrubbers, compteur)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(contenu)
    return compteur


# --------------------------------------------------------------------------- #
#  Fichiers Office (zip de XML)
# --------------------------------------------------------------------------- #
def _remplacer_dans_noeuds_texte(xml, termes, scrubbers, compteur):
    """Remplace uniquement le texte entre '>' et '<' (jamais les balises)."""
    def repl(m):
        inner = m.group(1)
        if not inner.strip():
            return m.group(0)
        return ">" + remplacer_texte(inner, termes, scrubbers, compteur) + "<"
    return re.sub(r">([^<>]*)<", repl, xml)


def _accepter_suivi_modifs(xml):
    """Best-effort : accepte insertions, supprime suppressions et révisions."""
    # Supprime les blocs supprimés (texte effacé)
    xml = re.sub(r"<w:del\b[^>]*>.*?</w:del>", "", xml, flags=re.DOTALL)
    # Déballe les insertions (garde le texte inséré)
    xml = re.sub(r"<w:ins\b[^>]*>(.*?)</w:ins>", r"\1", xml, flags=re.DOTALL)
    # Supprime les marqueurs de changement de format
    for tag in ("rPr", "pPr", "sectPr", "tblPr", "tcPr", "trPr"):
        xml = re.sub(rf"<w:{tag}Change\b[^>]*>.*?</w:{tag}Change>", "", xml,
                     flags=re.DOTALL)
        xml = re.sub(rf"<w:{tag}Change\b[^>]*/>", "", xml)
    return xml


def _supprimer_marqueurs_commentaires(xml):
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        xml = re.sub(rf"<w:{tag}\b[^>]*/>", "", xml)
        xml = re.sub(rf"<w:{tag}\b[^>]*>.*?</w:{tag}>", "", xml, flags=re.DOTALL)
    return xml


# Cibles d'hyperliens externes : http(s), mailto, ftp (les Type="http://schemas"
# et les cibles internes styles.xml/media/... ne commencent pas par ces schémas
# DANS un attribut Target -> intacts).
_RE_TARGET_LIEN = re.compile(r'(Target=")((?:https?|mailto|ftp)[^"]*)(")',
                             re.IGNORECASE)
_RE_URL_ABSOLUE = re.compile(r"(?i)^(?:https?|mailto|ftp):")


def _nettoyer_liens_rels(xml, termes, scrubbers, compteur):
    """Offusque les cibles d'hyperliens externes d'un .rels (http(s), mailto…).

    La cible d'un lien Office est un ATTRIBUT (Target="..."), donc hors des
    nœuds texte traités ailleurs. Si le remplacement n'est plus une URL absolue
    (ex. « [LIEN OUTIL COLLABORATIF] »), Word la résoudrait comme un lien
    RELATIF — révélant le chemin OneDrive du document. On force alors une cible
    absolue neutre pour éviter cette fuite.
    """
    def repl(m):
        avant = m.group(2)
        apres = remplacer_texte(avant, termes, scrubbers, compteur)
        if apres != avant and not _RE_URL_ABSOLUE.match(apres):
            apres = "https://anonymise.invalid/"
        return m.group(1) + apres + m.group(3)
    return _RE_TARGET_LIEN.sub(repl, xml)


# Membres à supprimer entièrement (commentaires, miniatures, personnes,
# objets OLE embarqués, ET propriétés du document : on retire core/app/custom
# plutôt que de les vider — la partie n'existe plus, donc zéro champ résiduel.
# Word régénère un minimum à la prochaine sauvegarde.)
#
# Objets OLE (word/ppt/xl + /embeddings/) : ce sont des fichiers embarqués —
# souvent un Excel/Word ENTIER avec ses PROPRES métadonnées et données — que
# l'anonymiseur ne sait pas ouvrir récursivement. On les supprime donc en bloc.
# Leur aperçu (image) est traité par ailleurs ; la référence <o:OLEObject>
# restante est tolérée par Word (objet « non disponible »).
def _membre_a_supprimer(nom):
    n = nom.lower()
    motifs = (
        "word/comments", "word/people.xml", "word/commentsids.xml",
        "word/commentsextended.xml", "word/commentsextensible.xml",
        "ppt/comments/", "ppt/authors.xml", "ppt/cmauthors.xml",
        "xl/comments", "xl/threadedcomments/", "xl/persons/",
        "docprops/thumbnail",
        "docprops/core.xml", "docprops/app.xml", "docprops/custom.xml",
        "word/embeddings/", "ppt/embeddings/", "xl/embeddings/",
    )
    return any(n.startswith(p) or n == p for p in motifs) or \
        ("/comments" in n and n.endswith(".xml") and "ppt/" in n)


# --------------------------------------------------------------------------- #
#  Images : suppression (défaut) ou strip des métadonnées (--garder-images)
# --------------------------------------------------------------------------- #
_IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff",
            ".emf", ".wmf", ".svg", ".ico", ".webp"}
_RE_MEDIA = re.compile(r"(?:word|ppt|xl)/(?:[^/]+/)*media/", re.IGNORECASE)

# Placeholders 1×1 SANS métadonnée. Remplacer les octets d'origine (au lieu de
# retirer la partie) garde le document structurellement valide : les références
# r:embed restent résolues, on évite les images « cassées ». Le PNG transparent
# sert de substitut universel (Word décode d'après le contenu, pas l'extension).
_PH_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk"
    "+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")
_PH_GIF = base64.b64decode(
    "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7")
_PH_SVG = b'<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>'


def _est_media_image(nom):
    return bool(_RE_MEDIA.search(nom)) and \
        os.path.splitext(nom)[1].lower() in _IMG_EXT


def _placeholder_image(nom):
    ext = os.path.splitext(nom)[1].lower()
    if ext == ".gif":
        return _PH_GIF
    if ext == ".svg":
        return _PH_SVG
    return _PH_PNG


def _strip_jpeg(data):
    """Retire les segments porteurs de métadonnées : APP1 (EXIF/XMP),
    APP13 (IPTC/Photoshop) et COM (commentaire). Conserve l'image."""
    if data[:2] != b"\xff\xd8":
        return data
    out = bytearray(b"\xff\xd8")
    i, n = 2, len(data)
    drop = {0xE1, 0xED, 0xFE}   # APP1, APP13, COM
    while i + 1 < n:
        if data[i] != 0xFF:
            out += data[i:]
            break
        marker = data[i + 1]
        if marker in (0xD9, 0xDA):        # EOI / début des données compressées
            out += data[i:]
            break
        if 0xD0 <= marker <= 0xD7 or marker in (0x00, 0x01):
            out += data[i:i + 2]
            i += 2
            continue
        if i + 3 >= n:
            out += data[i:]
            break
        seg_end = i + 2 + ((data[i + 2] << 8) | data[i + 3])
        if seg_end > n:
            out += data[i:]
            break
        if marker not in drop:
            out += data[i:seg_end]
        i = seg_end
    return bytes(out)


def _strip_png(data):
    """Retire les chunks de métadonnées : tEXt/zTXt/iTXt (texte, XMP),
    eXIf (EXIF) et tIME (horodatage). Conserve l'image."""
    sig = b"\x89PNG\r\n\x1a\n"
    if data[:8] != sig:
        return data
    out = bytearray(sig)
    i, n = 8, len(data)
    drop = {b"tEXt", b"zTXt", b"iTXt", b"eXIf", b"tIME"}
    while i + 8 <= n:
        length = int.from_bytes(data[i:i + 4], "big")
        ctype = data[i + 4:i + 8]
        chunk_end = i + 12 + length
        if chunk_end > n:
            out += data[i:]
            break
        if ctype not in drop:
            out += data[i:chunk_end]
        i = chunk_end
        if ctype == b"IEND":
            break
    return bytes(out)


def _stripper_image(nom, data):
    """Strip best-effort des métadonnées d'une image, sans la corrompre :
    en cas d'imprévu on renvoie les octets d'origine."""
    ext = os.path.splitext(nom)[1].lower()
    try:
        if ext in (".jpg", ".jpeg"):
            return _strip_jpeg(data)
        if ext == ".png":
            return _strip_png(data)
    except Exception:
        return data
    return data   # gif/bmp/tif/emf/wmf/svg/… : laissés tels quels


def _zinfo_neutre(nom):
    """ZipInfo à horodatage fixe : empêche les dates réelles de chaque membre
    de survivre dans la structure de l'archive (fuite indépendante du XML)."""
    zi = zipfile.ZipInfo(nom, date_time=(1980, 1, 1, 0, 0, 0))
    zi.compress_type = zipfile.ZIP_DEFLATED
    return zi


def traiter_office(src, dst, termes, scrubbers, args):
    compteur = Counter()
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        zin = zipfile.ZipFile(src, "r")
    except zipfile.BadZipFile:
        print(f"  [erreur] {os.path.basename(src)} n'est pas un zip Office valide "
              f"(fichier .doc/.ppt hérité ?). Ignoré.")
        NON_LUS.append(src)
        return None
    except OSError as e:
        print(f"  [ignoré] {os.path.basename(src)} illisible (ouvert dans Office "
              f"ou OneDrive hors-ligne ?) : {e.strerror or e}")
        NON_LUS.append(src)
        return None

    with zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        # Références de membres supprimés à nettoyer aussi des rels / content-types
        supprimes = [i.filename for i in zin.infolist()
                     if _membre_a_supprimer(i.filename)]
        base_supprimes = {s.rsplit("/", 1)[-1].lower() for s in supprimes}

        for info in zin.infolist():
            nom = info.filename
            if nom in supprimes:
                continue
            data = zin.read(nom)

            # 0) Images : supprimées par défaut (placeholder 1×1 vierge pour
            #    garder le document valide), ou conservées mais strippées de
            #    leurs métadonnées (EXIF/XMP/IPTC…) si --garder-images.
            if _est_media_image(nom):
                if getattr(args, "garder_images", False):
                    data = _stripper_image(nom, data)
                else:
                    data = _placeholder_image(nom)
                zout.writestr(_zinfo_neutre(nom), data)
                continue

            est_xml = nom.lower().endswith((".xml", ".rels"))
            if est_xml:
                try:
                    xml = data.decode("utf-8")
                except UnicodeDecodeError:
                    zout.writestr(_zinfo_neutre(nom), data)
                    continue

                # 1) Suivi de modifications (document Word)
                if nom.lower().endswith("word/document.xml") or \
                        re.search(r"word/(header|footer)\d*\.xml$", nom.lower()):
                    if not args.garder_suivi:
                        xml = _accepter_suivi_modifs(xml)
                    xml = _supprimer_marqueurs_commentaires(xml)
                # 1b) Identifiants de session de sauvegarde (corrélation
                #     inter-documents) : on retire le registre central des rsid.
                if nom.lower().endswith("word/settings.xml"):
                    xml = re.sub(r"<w:rsids>.*?</w:rsids>", "", xml,
                                 flags=re.DOTALL)
                # 3) Nettoyage des références aux membres supprimés
                if nom.lower().endswith(".rels"):
                    for b in base_supprimes:
                        xml = re.sub(rf'<Relationship\b[^>]*Target="[^"]*{re.escape(b)}"[^>]*/>',
                                     "", xml, flags=re.IGNORECASE)
                if nom.lower().endswith("[content_types].xml"):
                    for b in base_supprimes:
                        xml = re.sub(rf'<Override\b[^>]*PartName="[^"]*{re.escape(b)}"[^>]*/>',
                                     "", xml, flags=re.IGNORECASE)
                # 4) Remplacement des termes dans le texte
                xml = _remplacer_dans_noeuds_texte(xml, termes, scrubbers, compteur)
                # 4b) Cibles des hyperliens (attributs Target dans les .rels)
                if nom.lower().endswith(".rels"):
                    xml = _nettoyer_liens_rels(xml, termes, scrubbers, compteur)
                zout.writestr(_zinfo_neutre(nom), xml.encode("utf-8"))
            else:
                zout.writestr(_zinfo_neutre(nom), data)
    return compteur


# --------------------------------------------------------------------------- #
#  Extraction de texte (pour vérification)
# --------------------------------------------------------------------------- #
def extraire_texte(chemin):
    ext = os.path.splitext(chemin)[1].lower()
    try:
        if ext in TEXT_EXT:
            with open(chemin, encoding="utf-8", errors="replace") as f:
                return f.read()
        if ext in PDF_EXT:
            # PDF : converti en .docx (Word) puis lu comme un .docx. Sert au
            # mode SCAN comme à la simulation. Retourne "" si Word indisponible.
            import tempfile
            tmp = tempfile.mktemp(suffix=".docx")
            if not pdf_vers_docx(chemin, tmp):
                NON_LUS.append(chemin)
                return ""
            try:
                return extraire_texte(tmp)
            finally:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        if ext in OFFICE_EXT:
            morceaux = []
            with zipfile.ZipFile(chemin) as z:
                for nom in z.namelist():
                    low = nom.lower()
                    if low.endswith(".xml") and (
                        "word/" in low or "ppt/slides" in low or
                        "ppt/notesslides" in low or "docprops/" in low or
                        "xl/sharedstrings" in low or "xl/worksheets" in low
                    ):
                        try:
                            xml = z.read(nom).decode("utf-8")
                        except UnicodeDecodeError:
                            continue
                        morceaux.append(re.sub(r"<[^>]+>", " ", xml))
            return " ".join(morceaux)
    except zipfile.BadZipFile:
        print(f"  [ignoré] {os.path.basename(chemin)} : fichier Office "
              f"illisible/corrompu (ou .doc/.ppt hérité renommé).")
        NON_LUS.append(chemin)
        return ""
    except OSError as e:
        print(f"  [ignoré] {os.path.basename(chemin)} illisible (ouvert dans Office "
              f"ou OneDrive hors-ligne ?) : {e.strerror or e}")
        NON_LUS.append(chemin)
        return ""
    return ""


# --------------------------------------------------------------------------- #
#  PDF — conversion en .docx via Microsoft Word (COM / pywin32)
#
#  DÉPENDANCE OPTIONNELLE : ce bloc n'est sollicité QUE si le lot contient des
#  PDF. Il exige Microsoft Word installé + le paquet « pywin32 »
#  (pip install pywin32). Le PDF est converti en .docx, qui passe ensuite par
#  tout le durcissement Office habituel (métadonnées, OLE, images, termes…).
#
#  LIMITE : un PDF SCANNÉ (image de texte, sans couche texte) se convertit en
#  pages-images ; l'anonymiseur retirera ces images (sortie « blanche »), car
#  le texte incrusté dans les pixels n'est pas anonymisable sans OCR.
# --------------------------------------------------------------------------- #
_WORD_APP = None       # instance Word réutilisée sur tout le lot (coûteux à lancer)
_WORD_VERSION = None   # ex. "16.0" — sert à cibler la bonne clé de registre
_PDF_REG_BACKUP = None  # (existait: bool, ancienne_valeur | None) pour restauration


def _popup_pdf_off(version):
    """Coche « Ne plus afficher » du dialogue « Word va convertir ce PDF » via la
    clé de registre HKCU\\...\\Word\\Options\\DisableConvertPdfWarning, en
    mémorisant l'état initial pour le restaurer ensuite (Word non altéré)."""
    global _PDF_REG_BACKUP
    try:
        import winreg
        path = rf"Software\Microsoft\Office\{version}\Word\Options"
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
        try:
            _PDF_REG_BACKUP = (True, winreg.QueryValueEx(key, "DisableConvertPdfWarning")[0])
        except FileNotFoundError:
            _PDF_REG_BACKUP = (False, None)
        winreg.SetValueEx(key, "DisableConvertPdfWarning", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
    except Exception:
        _PDF_REG_BACKUP = None   # registre inaccessible : on n'a rien changé


def _popup_pdf_restaurer(version):
    """Restaure la valeur d'origine de DisableConvertPdfWarning."""
    global _PDF_REG_BACKUP
    if _PDF_REG_BACKUP is None:
        return
    try:
        import winreg
        path = rf"Software\Microsoft\Office\{version}\Word\Options"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE)
        existait, val = _PDF_REG_BACKUP
        if existait:
            winreg.SetValueEx(key, "DisableConvertPdfWarning", 0, winreg.REG_DWORD, int(val))
        else:
            try:
                winreg.DeleteValue(key, "DisableConvertPdfWarning")
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception:
        pass
    finally:
        _PDF_REG_BACKUP = None


def _word_app():
    """Retourne une instance Word automation, créée à la demande. Lève
    ImportError si pywin32 manque, ou une autre exception si Word est absent."""
    global _WORD_APP, _WORD_VERSION
    if _WORD_APP is None:
        import win32com.client as win32   # pywin32 : dépendance optionnelle
        app = win32.DispatchEx("Word.Application")   # instance isolée et neuve
        app.Visible = False
        try:
            app.DisplayAlerts = 0          # wdAlertsNone : pas de boîte de dialogue
        except Exception:
            pass
        try:
            _WORD_VERSION = str(app.Version)            # ex. "16.0"
            _popup_pdf_off(_WORD_VERSION)               # supprime le popup de conversion
        except Exception:
            pass
        _WORD_APP = app
    return _WORD_APP


def fermer_word():
    """Ferme l'instance Word si elle a été ouverte (à appeler en fin de lot)."""
    global _WORD_APP, _WORD_VERSION
    if _WORD_APP is not None:
        try:
            _WORD_APP.Quit()
        except Exception:
            pass
        _WORD_APP = None
    if _WORD_VERSION is not None:
        _popup_pdf_restaurer(_WORD_VERSION)             # remet Word dans son état initial
        _WORD_VERSION = None


def pdf_vers_docx(src_pdf, dst_docx):
    """Convertit src_pdf -> dst_docx via Word. Retourne True si succès."""
    try:
        app = _word_app()
    except ImportError:
        print("  [pdf] pywin32 absent — `pip install pywin32` pour traiter les "
              "PDF (Microsoft Word requis). PDF ignoré.")
        return False
    except Exception as e:
        print(f"  [pdf] Microsoft Word indisponible ({e}). PDF ignoré.")
        return False
    doc = None
    try:
        # ConfirmConversions=False : supprime l'invite « convertir le PDF ».
        # wdOpenFormatAuto=0 ; FileFormat 16 = wdFormatDocumentDefault (.docx).
        doc = app.Documents.Open(os.path.abspath(src_pdf), ConfirmConversions=False,
                                 ReadOnly=True, AddToRecentFiles=False, Visible=False)
        doc.SaveAs2(os.path.abspath(dst_docx), FileFormat=16)
        return True
    except Exception as e:
        print(f"  [pdf] échec de conversion de {os.path.basename(src_pdf)} : {e}")
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)   # wdDoNotSaveChanges
            except Exception:
                pass


def traiter_pdf(src, dst, termes, scrubbers, args):
    """Convertit le PDF en .docx temporaire puis applique le traitement Office.
    'dst' doit déjà porter l'extension .docx."""
    import tempfile
    tmp = tempfile.mktemp(suffix=".docx")
    if not pdf_vers_docx(src, tmp):
        NON_LUS.append(src)
        return None
    try:
        return traiter_office(tmp, dst, termes, scrubbers, args)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


# --------------------------------------------------------------------------- #
#  Mode SCAN — pré-remplissage du fichier de correspondances
# --------------------------------------------------------------------------- #
RE_ACRONYME = re.compile(r"\b[A-ZÉÈÀÂÎÔÛ]{2,}(?:\d+)?\b")
RE_CAPS = re.compile(r"\b[A-ZÉÈÀÂÎÔÛ][\wÀ-ÿ]+(?:\s+[A-ZÉÈÀÂÎÔÛ][\wÀ-ÿ]+){0,3}")
STOPWORDS = {
    "Le", "La", "Les", "Un", "Une", "Des", "De", "Du", "Et", "Ou", "Mais",
    "Dans", "Pour", "Par", "Sur", "Avec", "Sans", "Sous", "Ce", "Cette",
    "Il", "Elle", "On", "Nous", "Vous", "Si", "Quand", "Mon", "Ma", "Mes",
}

def _chemin_unique(chemin):
    """Renvoie un chemin qui n'existe pas encore.

    Si 'chemin' existe déjà, insère un suffixe numérique avant l'extension
    (_2, _3, ...) afin de ne JAMAIS écraser un fichier de correspondances
    généré précédemment.
    """
    if not os.path.exists(chemin):
        return chemin
    racine, ext = os.path.splitext(chemin)
    n = 2
    while os.path.exists(f"{racine}_{n}{ext}"):
        n += 1
    return f"{racine}_{n}{ext}"


def _charger_termes_existants(chemin):
    """Renvoie l'ensemble (en minuscules) des termes littéraux déjà présents
    dans un fichier de correspondances. Les motifs regex (re:) sont ignorés.
    Sert à EXCLURE du scan ce qui est déjà connu."""
    termes = set()
    try:
        for orig, _repl, est_regex in charger_correspondances(chemin):
            if not est_regex:
                termes.add(orig.lower())
    except OSError:
        print(f"  [avert] fichier d'exclusion introuvable, ignoré : {chemin}")
    return termes


# Caractères interdits dans un nom de fichier Windows (au cas où un remplacement
# en introduirait). Les crochets [ ] sont VALIDES, ils ne sont pas listés ici.
_CARACTERES_INTERDITS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def anonymiser_nom_relatif(rel, termes_noms):
    """Applique les termes LITTÉRAUX à chaque composant d'un chemin relatif
    (sous-dossiers + nom de fichier), puis neutralise tout caractère interdit.
    Les motifs regex et les scrubbers techniques ne sont PAS appliqués aux noms."""
    propres = []
    jeton = Counter()
    for comp in re.split(r"[\\/]+", rel):
        nv = remplacer_texte(comp, termes_noms, [], jeton)
        nv = _CARACTERES_INTERDITS.sub("_", nv).rstrip(" .") or "_"
        propres.append(nv)
    return os.path.join(*propres)


def _dst_unique(dst, utilises):
    """Évite d'écraser un fichier déjà écrit pendant ce run (collision possible
    après anonymisation des noms) en insérant _2, _3... avant l'extension.
    Comparaison insensible à la casse (système de fichiers Windows)."""
    if dst.lower() not in utilises:
        utilises.add(dst.lower())
        return dst
    racine, ext = os.path.splitext(dst)
    n = 2
    while f"{racine}_{n}{ext}".lower() in utilises:
        n += 1
    dst = f"{racine}_{n}{ext}"
    utilises.add(dst.lower())
    return dst


def mode_scan(dossier, sortie_map, exclure=None):
    exclure = exclure or set()
    cand = Counter()
    try:
        for racine, _, fichiers in os.walk(dossier):
            for nom in fichiers:
                ext = os.path.splitext(nom)[1].lower()
                if ext in TEXT_EXT or ext in OFFICE_EXT or ext in PDF_EXT:
                    texte = extraire_texte(os.path.join(racine, nom))
                    for m in RE_ACRONYME.findall(texte):
                        cand[m] += 1
                    for m in RE_CAPS.findall(texte):
                        tete = m.split()[0]
                        if tete not in STOPWORDS and len(m) > 3:
                            cand[m] += 1
    finally:
        fermer_word()   # ferme Word si des PDF ont été lus
    # Filtre : >= 2 occurrences ET pas déjà connu du fichier de référence.
    retenus = sorted((c for c in cand.items()
                      if c[1] >= 2 and c[0].lower() not in exclure),
                     key=lambda x: (-x[1], x[0]))
    with open(sortie_map, "w", encoding="utf-8") as f:
        f.write("# Fichier de correspondances pré-rempli (mode --scan)\n")
        f.write("# Format : terme_original<TAB>remplacement\n")
        f.write("# Complète/corrige la colonne de droite, supprime les lignes inutiles.\n")
        f.write("# Les lignes commençant par # sont ignorées.\n\n")
        for terme, n in retenus:
            f.write(f"{terme}\t\t# vu {n}x — à compléter\n")
    suffixe = f" (termes déjà connus exclus)" if exclure else ""
    print(f"[scan] {len(retenus)} candidats écrits dans {sortie_map}{suffixe}")
    print("       -> Complète la colonne de droite, puis recopie les lignes utiles "
          "dans ton fichier de référence.")


# --------------------------------------------------------------------------- #
#  Vérification finale
# --------------------------------------------------------------------------- #
def verifier(sortie, paires, sensible_casse, verifier_noms=True):
    flags = 0 if sensible_casse else re.IGNORECASE
    motifs = []
    for orig, _repl, est_regex in paires:
        try:
            if est_regex:
                motifs.append((orig, re.compile(orig, flags)))
            else:
                motifs.append((orig, re.compile(r"\b" + _motif_litteral(orig) + r"\b", flags)))
        except re.error:
            continue
    survivants = []
    for racine, _, fichiers in os.walk(sortie):
        for nom in fichiers:
            chemin = os.path.join(racine, nom)
            if os.path.splitext(nom)[1].lower() not in (TEXT_EXT | OFFICE_EXT):
                continue
            rel_nom = os.path.relpath(chemin, sortie)
            texte = extraire_texte(chemin)
            for orig, motif in motifs:
                n = len(motif.findall(texte))
                if n:
                    survivants.append((rel_nom, orig, n))
                # Termes subsistant dans le NOM du fichier / dossier
                if verifier_noms:
                    nn = len(motif.findall(rel_nom))
                    if nn:
                        survivants.append((rel_nom + "  [NOM]", orig, nn))
    return survivants


# --------------------------------------------------------------------------- #
#  Cœur réutilisable (partagé par la ligne de commande ET le menu interactif)
# --------------------------------------------------------------------------- #
def _dossier_app():
    """Dossier de l'application : à côté de l'.exe si figé (PyInstaller),
    sinon à côté du script .py. Sert à localiser le .tsv et le dossier Scan."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _trouver_tsv(dossier):
    """Renvoie (premier_tsv | None, [tous les .tsv triés par nom]) d'un dossier."""
    fichiers = sorted(glob.glob(os.path.join(dossier, "*.tsv")),
                      key=lambda p: os.path.basename(p).lower())
    return (fichiers[0] if fichiers else None), fichiers


def lancer_scan(source, exclure_path=None):
    """Scanne 'source' et écrit un nouveau correspondances_<dossier>.tsv dans le
    sous-dossier 'Scan' de l'application. 'exclure_path' : .tsv dont les termes
    connus sont exclus des candidats."""
    scan_dir = os.path.join(_dossier_app(), "Scan")
    os.makedirs(scan_dir, exist_ok=True)
    nom_dossier = os.path.basename(os.path.normpath(source)) or "scan"
    sortie_map = _chemin_unique(
        os.path.join(scan_dir, f"correspondances_{nom_dossier}.tsv"))
    exclure = _charger_termes_existants(exclure_path) if exclure_path else set()
    mode_scan(source, sortie_map, exclure)
    _recap_non_lus()
    return sortie_map


def anonymiser_dossier(source, out, map_path, opts, report=None):
    """Cœur de l'anonymisation : charge 'map_path', traite tous les fichiers de
    'source' vers 'out', écrit le rapport CSV et lance la vérification.

    'opts' porte les attributs : sensible_casse, neutraliser_dates, garder_suivi,
    garder_noms, dry_run, no_email, no_url, no_ip, no_mac, no_port, no_path.
    """
    paires = charger_correspondances(map_path)
    if not paires:
        print("[avert] aucune correspondance valide chargée — seuls les motifs "
              "techniques seront appliqués.")
    termes = compiler_motifs_termes(paires, opts.sensible_casse)
    # Pour les NOMS de fichiers : uniquement les termes littéraux (pas les regex).
    termes_noms = compiler_motifs_termes([p for p in paires if not p[2]],
                                         opts.sensible_casse)
    scrubbers = construire_scrubbers(opts)

    total = Counter()
    nb_fichiers = 0
    rapport_lignes = []
    renommages = []
    dst_utilises = set()

    try:
      for racine, _, fichiers in os.walk(source):
        for nom in fichiers:
            src = os.path.join(racine, nom)
            ext = os.path.splitext(nom)[1].lower()
            rel = os.path.relpath(src, source)

            if ext in LEGACY_EXT:
                print(f"  [ignoré] {rel} : format hérité .doc/.ppt — convertir en "
                      f".docx/.pptx d'abord.")
                continue
            if ext not in (TEXT_EXT | OFFICE_EXT | PDF_EXT):
                continue

            # Nom de sortie : anonymisé par défaut (sauf garder_noms).
            rel_out = rel if opts.garder_noms else anonymiser_nom_relatif(rel, termes_noms)
            # Un PDF est converti : sa sortie est un .docx anonymisé.
            if ext in PDF_EXT:
                rel_out = os.path.splitext(rel_out)[0] + ".docx"
            dst = os.path.join(out, rel_out)
            if not opts.dry_run:
                dst = _dst_unique(dst, dst_utilises)
                rel_out = os.path.relpath(dst, out)

            if opts.dry_run:
                texte = extraire_texte(src)   # gère .txt/.md/.docx/.pptx/.xlsx/.pdf
                c = Counter()
                remplacer_texte(texte, termes, scrubbers, c)
                compteur = c
            elif ext in TEXT_EXT:
                compteur = traiter_texte(src, dst, termes, scrubbers)
            elif ext in PDF_EXT:
                compteur = traiter_pdf(src, dst, termes, scrubbers, opts)
            else:
                compteur = traiter_office(src, dst, termes, scrubbers, opts)
            if compteur is None:   # fichier illisible / verrouillé : on saute
                continue

            nb_fichiers += 1
            total.update(compteur)
            if rel_out != rel:
                renommages.append((rel, rel_out))
            for terme, n in sorted(compteur.items()):
                rapport_lignes.append((rel_out, terme, n))
            tag = "[simulé]" if opts.dry_run else "[ok]"
            renom = f"  →  {rel_out}" if rel_out != rel else ""
            print(f"  {tag} {rel}{renom} — {sum(compteur.values())} remplacement(s)")
    finally:
        fermer_word()   # ferme Word si des PDF ont été traités

    # --- Rapport ------------------------------------------------------------ #
    if not opts.dry_run:
        os.makedirs(out, exist_ok=True)
    rapport = report or os.path.join(
        out if not opts.dry_run else os.getcwd(), "_rapport_anonymisation.csv")
    with open(rapport, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["fichier", "terme_ou_motif", "occurrences_remplacees"])
        for ligne in rapport_lignes:
            w.writerow(ligne)
        w.writerow([])
        w.writerow(["TOTAL", "tous termes confondus", sum(total.values())])
        if renommages:
            w.writerow([])
            w.writerow(["RENOMMAGES (nom d'origine -> nom anonymisé)", "", ""])
            for avant, apres in renommages:
                w.writerow([avant, apres, ""])

    print(f"\n{nb_fichiers} fichier(s) traité(s) — "
          f"{sum(total.values())} remplacement(s) au total.")
    if renommages:
        print(f"{len(renommages)} fichier(s) renommé(s) (noms anonymisés).")
    print(f"Rapport : {rapport}")

    # --- Vérification ------------------------------------------------------- #
    if not opts.dry_run and paires:
        survivants = verifier(out, paires, opts.sensible_casse,
                              verifier_noms=not opts.garder_noms)
        if survivants:
            print("\n⚠️  VÉRIFICATION : des termes interdits SUBSISTENT dans la "
                  "sortie (probable coupure entre runs Word) :")
            for fichier, terme, n in survivants:
                print(f"    - {fichier} : « {terme} » ×{n}")
            print("    -> Ouvre/réenregistre le .docx dans Word puis relance, ou "
                  "corrige à la main.")
            with open(rapport, "a", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow([])
                w.writerow(["SURVIVANTS (à corriger)", "", ""])
                for ligne in survivants:
                    w.writerow(ligne)
        else:
            print("\n✅ VÉRIFICATION : aucun terme du fichier de correspondances ne "
                  "subsiste dans la sortie.")
            print("    (Rappel : ceci ne couvre PAS le contexte métier non listé.)")

    # --- Récapitulatif des fichiers non lus --------------------------------- #
    _recap_non_lus()
    if NON_LUS and not opts.dry_run:
        with open(rapport, "a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([])
            w.writerow(["FICHIERS NON LUS (absents du résultat)", "", ""])
            for c in sorted(set(NON_LUS)):
                w.writerow([c, "", ""])


def _opts_defaut():
    """Objet d'options avec les valeurs par défaut (tout actif, rien d'exclu)."""
    return types.SimpleNamespace(
        sensible_casse=False, neutraliser_dates=False, garder_suivi=False,
        garder_noms=False, garder_images=False, dry_run=False, no_email=False,
        no_url=False, no_ip=False, no_mac=False, no_port=False, no_path=False)


# --------------------------------------------------------------------------- #
#  Menu interactif (équivalent du lanceur .bat, intégré pour l'.exe)
# --------------------------------------------------------------------------- #
def _saisir(prompt):
    try:
        return input(prompt).strip().strip('"')
    except (EOFError, KeyboardInterrupt):
        return ""


def _oui(question):
    return _saisir(question + " [o/N] : ").lower() == "o"


def _pause():
    _saisir("\nAppuyez sur Entrée pour continuer...")


def _menu_scan(app):
    print("\n" + "-" * 60)
    print(" MODE SCAN")
    print("-" * 60)
    print(" Génère un correspondances_<dossier>.tsv dans le sous-dossier « Scan »")
    print(" (à côté de l'application). Les termes déjà présents dans le .tsv de")
    print(" référence sont exclus. Aucun document n'est modifié.")
    src = _saisir("\nDossier à scanner : ")
    if not src:
        return
    if not os.path.isdir(src):
        print(f"[ERREUR] Dossier introuvable : {src}")
        _pause()
        return
    ref, _ = _trouver_tsv(app)
    lancer_scan(src, ref)
    print("\nTerminé. Le .tsv généré est dans le sous-dossier « Scan ».")
    print("Complète la 2e colonne, puis recopie les lignes utiles dans ton .tsv.")
    _pause()


def _menu_anon(app):
    print("\n" + "-" * 60)
    print(" MODE ANONYMISATION")
    print("-" * 60)
    src = _saisir("\nDossier SOURCE : ")
    if not src:
        return
    if not os.path.isdir(src):
        print(f"[ERREUR] Dossier source introuvable : {src}")
        _pause()
        return
    out = _saisir("Dossier CIBLE (sortie, l'arborescence y sera recréée) : ")
    if not out:
        return
    if os.path.abspath(out) == os.path.abspath(src):
        print("[ERREUR] Le dossier de sortie doit être différent du source.")
        _pause()
        return

    mapfile, liste = _trouver_tsv(app)
    if not mapfile:
        print(f"[ERREUR] Aucun fichier .tsv trouvé dans : {app}")
        _pause()
        return
    if len(liste) > 1:
        noms = ", ".join(os.path.basename(p) for p in liste)
        print(f"\n[ATTENTION] {len(liste)} fichiers .tsv dans le dossier : {noms}")
        print("            Le premier est utilisé -- vérifie que c'est le bon.")
    import contextlib
    import io
    with contextlib.redirect_stdout(io.StringIO()):
        nb = len(charger_correspondances(mapfile))
    print(f"\nFichier de correspondances détecté automatiquement :")
    print(f"   {os.path.basename(mapfile)}")
    print(f"   -> {nb} correspondance(s) chargée(s).")

    opts = _opts_defaut()
    print("\n" + "-" * 60)
    print(" Par DÉFAUT, le traitement va :")
    print("   - remplacer les termes du .tsv (regex + termes exacts) ;")
    print("   - masquer emails, URLs, IP, MAC, ports et chemins ;")
    print("   - accepter le suivi de modifications Word, supprimer les commentaires ;")
    print("   - SUPPRIMER entièrement les propriétés du document (auteur, société,")
    print("     titres, propriétés custom) et neutraliser les horodatages internes ;")
    print("   - SUPPRIMER les images et les objets OLE embarqués (Excel/Word collés) ;")
    print("   - anonymiser AUSSI les noms de fichiers et de dossiers ;")
    print("   - insensible à la casse et aux accents.")
    print("-" * 60)
    print("\n  [1] Lancer avec les options PAR DÉFAUT")
    print("  [2] Configurer les options en détail")
    if _saisir("Votre choix [1-2] : ") == "2":
        opts.dry_run = _oui("Simulation, sans rien écrire (dry-run) ?")
        opts.sensible_casse = _oui("Respecter la casse ?")
        opts.garder_suivi = _oui("Garder le suivi de modifications Word ?")
        opts.garder_noms = _oui("Garder les noms d'origine (ne PAS anonymiser les noms) ?")
        opts.garder_images = _oui("Conserver les images ? (par défaut elles sont "
                                  "SUPPRIMÉES ; si gardées, leurs métadonnées sont "
                                  "strippées)")
        scrub = _saisir("Désactiver nettoyages (E=email U=url I=ip P=port C=chemin "
                        "M=mac) ou Entrée : ").upper()
        opts.no_email = "E" in scrub
        opts.no_url = "U" in scrub
        opts.no_ip = "I" in scrub
        opts.no_port = "P" in scrub
        opts.no_path = "C" in scrub
        opts.no_mac = "M" in scrub

    print("\n" + "-" * 60)
    print(" RÉCAPITULATIF")
    print(f"   Source : {src}")
    print(f"   Cible  : {out}")
    print(f"   Map    : {os.path.basename(mapfile)} ({nb} correspondances)")
    print("-" * 60)
    if _saisir("Lancer maintenant ? [O/n] : ").lower() == "n":
        return
    print()
    anonymiser_dossier(src, out, mapfile, opts)
    _pause()


def menu_interactif():
    app = _dossier_app()
    while True:
        print("\n" + "=" * 60)
        print("                ANONYMISEUR DE DOCUMENTS")
        print("=" * 60)
        print("\n  Formats traités : .txt .md .docx .pptx .xlsx .pdf\n")
        print("  1. Scanner un dossier  (générer un fichier de correspondances)")
        print("  2. Anonymiser un dossier")
        print("  3. Quitter")
        try:
            choix = input("\nVotre choix [1-3] : ").strip()
        except (EOFError, KeyboardInterrupt):
            return   # fin de flux / Ctrl+C : on quitte proprement
        if choix == "1":
            _menu_scan(app)
        elif choix == "2":
            _menu_anon(app)
        elif choix == "3":
            return


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def main():
    # Console Windows souvent en cp1252 : passer en UTF-8 (sortie + page de code)
    # évite un plantage / des accents cassés sur ✅ / ⚠️ / é è à...
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
            ctypes.windll.kernel32.SetConsoleCP(65001)
        except Exception:
            pass
    for flux in (sys.stdout, sys.stderr):
        try:
            flux.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

    # Aucun argument (double-clic sur l'.exe, ou « python anonymiseur.py ») :
    # on lance le menu interactif.
    if len(sys.argv) == 1:
        menu_interactif()
        return

    p = argparse.ArgumentParser(
        description="Anonymise par lot des documents .txt/.md/.docx/.pptx/.xlsx/.pdf "
                    "(.pdf : conversion via Word, nécessite pywin32).")
    p.add_argument("source", nargs="?",
                   help="Dossier source contenant les documents.")
    p.add_argument("--out", help="Dossier de sortie (obligatoire hors --scan, "
                                 "doit être différent du source).")
    p.add_argument("--map", help="Fichier de correspondances TSV "
                                 "(original<TAB>remplacement).")
    p.add_argument("--scan", action="store_true",
                   help="Mode pré-remplissage : génère un fichier de "
                        "correspondances et ne modifie rien.")
    p.add_argument("--exclure", help="(mode --scan) Fichier de correspondances "
                   "dont les termes connus seront EXCLUS des candidats générés.")
    p.add_argument("--report", help="Chemin du rapport CSV "
                                     "(défaut : <out>/_rapport_anonymisation.csv).")
    p.add_argument("--case-sensitive", dest="sensible_casse", action="store_true",
                   help="Respecte la casse (recommandé si un terme coïncide avec "
                        "un mot courant, ex. l'adjectif « simplicité »).")
    p.add_argument("--neutraliser-dates", action="store_true",
                   help="(sans effet : conservé pour compatibilité) Les "
                        "propriétés du document — dates incluses — sont "
                        "désormais TOUJOURS supprimées, et les horodatages "
                        "internes de l'archive neutralisés.")
    p.add_argument("--garder-suivi", action="store_true",
                   help="Ne touche pas au suivi de modifications Word.")
    p.add_argument("--garder-noms", action="store_true",
                   help="Ne PAS anonymiser les noms de fichiers/dossiers "
                        "(par défaut, les termes du .tsv y sont aussi appliqués).")
    p.add_argument("--garder-images", action="store_true",
                   help="Conserver les images des documents Office (leurs "
                        "métadonnées EXIF/XMP/IPTC sont alors strippées). "
                        "PAR DÉFAUT, les images sont SUPPRIMÉES.")
    p.add_argument("--dry-run", action="store_true",
                   help="Simule : compte les remplacements sans écrire les docs.")
    p.add_argument("--compter", action="store_true",
                   help="Affiche juste le nombre de correspondances valides de "
                        "--map (sur stdout) puis quitte. Sert au lanceur .bat.")
    # Désactivation des scrubbers techniques
    p.add_argument("--no-email", action="store_true")
    p.add_argument("--no-url", action="store_true")
    p.add_argument("--no-ip", action="store_true")
    p.add_argument("--no-mac", action="store_true")
    p.add_argument("--no-port", action="store_true")
    p.add_argument("--no-path", action="store_true")
    args = p.parse_args()

    # --- Mode comptage (pour le lanceur) : n'imprime QUE le nombre ---------- #
    if args.compter:
        if not args.map:
            sys.exit("[erreur] --compter requiert --map.")
        import io
        import contextlib
        with contextlib.redirect_stdout(io.StringIO()):   # tait les [avert]
            paires = charger_correspondances(args.map)
        print(len(paires))
        return

    if not args.source or not os.path.isdir(args.source):
        sys.exit(f"[erreur] dossier source introuvable : {args.source}")

    # --- Mode scan ---------------------------------------------------------- #
    if args.scan:
        lancer_scan(args.source, args.exclure)
        return

    # --- Mode anonymisation ------------------------------------------------- #
    if not args.out:
        sys.exit("[erreur] --out est obligatoire en mode anonymisation.")
    if os.path.abspath(args.out) == os.path.abspath(args.source):
        sys.exit("[erreur] le dossier de sortie doit être différent du source.")
    if not args.map:
        sys.exit("[erreur] --map est obligatoire (lance d'abord --scan pour le générer).")
    anonymiser_dossier(args.source, args.out, args.map, args, args.report)


if __name__ == "__main__":
    main()
