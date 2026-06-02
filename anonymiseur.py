#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
anonymiseur.py — Anonymisation par lot de documents .txt / .md / .docx / .pptx / .xlsx

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
- Ce script neutralise les noms ; il ne neutralise PAS le contexte métier
  (secteur, taille, budget, stack legacy) sauf si tu ajoutes ces expressions
  dans le fichier de correspondances. Le contexte seul peut rester identifiant.
"""

import argparse
import csv
import os
import re
import shutil
import sys
import zipfile
from collections import Counter

TEXT_EXT = {".txt", ".md", ".markdown"}
OFFICE_EXT = {".docx", ".pptx", ".xlsx"}
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
      1) règles REGEX du fichier (re:) — les plus spécifiques (ex. URLs) ;
      2) SCRUBBERS techniques (email, url, ip, mac, port, chemin) ;
      3) termes EXACTS du fichier.
    Les scrubbers passent AVANT les termes exacts pour qu'un terme (ex. SEMAE)
    ne fragmente pas un email/URL et n'empêche pas sa détection
    (ex. « x@semae.fr » doit devenir « [EMAIL] », pas « x@[ORG-S].fr »)."""
    # 1) Règles regex (re:)
    for orig, motif, repl, est_regex in termes:
        if not est_regex:
            continue
        texte, n = motif.subn(repl.replace("\\", "\\\\"), texte)
        if n:
            compteur[orig] += n
    # 2) Scrubbers techniques
    for nom, motif, repl in scrubbers:
        texte, n = motif.subn(repl, texte)
        if n:
            compteur[f"<{nom}>"] += n
    # 3) Termes exacts (littéraux)
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


def _vider_balises(xml, tags):
    """Vide le contenu textuel de balises données : <tag ...>X</tag> -> <tag .../>."""
    for tag in tags:
        t = re.escape(tag)
        xml = re.sub(rf"(<{t}(?:\s[^>]*)?>).*?(</{t}>)", r"\1\2", xml, flags=re.DOTALL)
    return xml


def _nettoyer_metadonnees(membre, xml, neutraliser_dates):
    """core.xml / app.xml / custom.xml : retire les champs identifiants."""
    base = membre.rsplit("/", 1)[-1]
    if base == "core.xml":
        xml = _vider_balises(xml, [
            "dc:creator", "cp:lastModifiedBy", "dc:title", "dc:subject",
            "cp:keywords", "dc:description", "cp:category", "cp:contentStatus",
        ])
        if neutraliser_dates:
            xml = _vider_balises(xml, ["dcterms:created", "dcterms:modified",
                                       "cp:lastPrinted"])
    elif base == "app.xml":
        xml = _vider_balises(xml, ["Company", "Manager"])
    elif base == "custom.xml":
        # Valeurs de propriétés personnalisées
        xml = _vider_balises(xml, ["vt:lpwstr", "vt:i4", "vt:filetime", "vt:bool"])
    return xml


# Membres à supprimer entièrement (commentaires, miniatures, personnes)
def _membre_a_supprimer(nom):
    n = nom.lower()
    motifs = (
        "word/comments", "word/people.xml", "word/commentsids.xml",
        "word/commentsextended.xml", "word/commentsextensible.xml",
        "ppt/comments/", "ppt/authors.xml", "ppt/cmauthors.xml",
        "xl/comments", "xl/threadedcomments/", "xl/persons/",
        "docprops/thumbnail",
    )
    return any(n.startswith(p) or n == p for p in motifs) or \
        ("/comments" in n and n.endswith(".xml") and "ppt/" in n)


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

            est_xml = nom.lower().endswith((".xml", ".rels"))
            if est_xml:
                try:
                    xml = data.decode("utf-8")
                except UnicodeDecodeError:
                    zout.writestr(info, data)
                    continue

                # 1) Métadonnées
                if "docprops/" in nom.lower():
                    xml = _nettoyer_metadonnees(nom, xml, args.neutraliser_dates)
                # 2) Suivi de modifications (document Word)
                if nom.lower().endswith("word/document.xml") or \
                        re.search(r"word/(header|footer)\d*\.xml$", nom.lower()):
                    if not args.garder_suivi:
                        xml = _accepter_suivi_modifs(xml)
                    xml = _supprimer_marqueurs_commentaires(xml)
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
                zout.writestr(info, xml.encode("utf-8"))
            else:
                zout.writestr(info, data)
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
    for racine, _, fichiers in os.walk(dossier):
        for nom in fichiers:
            ext = os.path.splitext(nom)[1].lower()
            if ext in TEXT_EXT or ext in OFFICE_EXT:
                texte = extraire_texte(os.path.join(racine, nom))
                for m in RE_ACRONYME.findall(texte):
                    cand[m] += 1
                for m in RE_CAPS.findall(texte):
                    tete = m.split()[0]
                    if tete not in STOPWORDS and len(m) > 3:
                        cand[m] += 1
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
#  Main
# --------------------------------------------------------------------------- #
def main():
    # Console Windows souvent en cp1252 : forcer l'UTF-8 en sortie évite un
    # plantage (UnicodeEncodeError) sur les caractères ✅ / ⚠️ / accents.
    for flux in (sys.stdout, sys.stderr):
        try:
            flux.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

    p = argparse.ArgumentParser(
        description="Anonymise par lot des documents .txt/.md/.docx/.pptx/.xlsx.")
    p.add_argument("source", help="Dossier source contenant les documents.")
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
                   help="Efface aussi les dates de création/modification Office.")
    p.add_argument("--garder-suivi", action="store_true",
                   help="Ne touche pas au suivi de modifications Word.")
    p.add_argument("--garder-noms", action="store_true",
                   help="Ne PAS anonymiser les noms de fichiers/dossiers "
                        "(par défaut, les termes du .tsv y sont aussi appliqués).")
    p.add_argument("--dry-run", action="store_true",
                   help="Simule : compte les remplacements sans écrire les docs.")
    # Désactivation des scrubbers techniques
    p.add_argument("--no-email", action="store_true")
    p.add_argument("--no-url", action="store_true")
    p.add_argument("--no-ip", action="store_true")
    p.add_argument("--no-mac", action="store_true")
    p.add_argument("--no-port", action="store_true")
    p.add_argument("--no-path", action="store_true")
    args = p.parse_args()

    if not os.path.isdir(args.source):
        sys.exit(f"[erreur] dossier source introuvable : {args.source}")

    # --- Mode scan ---------------------------------------------------------- #
    if args.scan:
        # Sortie dans un sous-dossier "Scan" du dossier de ce script, peu importe
        # d'où la commande est lancée. Un NOUVEAU fichier à chaque scan.
        dossier_script = os.path.dirname(os.path.abspath(__file__))
        scan_dir = os.path.join(dossier_script, "Scan")
        os.makedirs(scan_dir, exist_ok=True)
        nom_dossier = os.path.basename(os.path.normpath(args.source)) or "scan"
        base = f"correspondances_{nom_dossier}.tsv"
        # _chemin_unique garantit qu'aucun fichier existant n'est écrasé.
        sortie_map = _chemin_unique(os.path.join(scan_dir, base))
        # Exclut du scan les termes déjà présents dans le fichier de référence.
        exclure = _charger_termes_existants(args.exclure) if args.exclure else set()
        mode_scan(args.source, sortie_map, exclure)
        _recap_non_lus()
        return

    # --- Mode anonymisation ------------------------------------------------- #
    if not args.out:
        sys.exit("[erreur] --out est obligatoire en mode anonymisation.")
    if os.path.abspath(args.out) == os.path.abspath(args.source):
        sys.exit("[erreur] le dossier de sortie doit être différent du source.")
    if not args.map:
        sys.exit("[erreur] --map est obligatoire (lance d'abord --scan pour le générer).")

    paires = charger_correspondances(args.map)
    if not paires:
        print("[avert] aucune correspondance valide chargée — seuls les motifs "
              "techniques seront appliqués.")
    termes = compiler_motifs_termes(paires, args.sensible_casse)
    # Pour les NOMS de fichiers : uniquement les termes littéraux (pas les regex).
    termes_noms = compiler_motifs_termes([p for p in paires if not p[2]],
                                         args.sensible_casse)
    scrubbers = construire_scrubbers(args)

    total = Counter()
    nb_fichiers = 0
    rapport_lignes = []
    renommages = []
    dst_utilises = set()

    for racine, _, fichiers in os.walk(args.source):
        for nom in fichiers:
            src = os.path.join(racine, nom)
            ext = os.path.splitext(nom)[1].lower()
            rel = os.path.relpath(src, args.source)

            if ext in LEGACY_EXT:
                print(f"  [ignoré] {rel} : format hérité .doc/.ppt — convertir en "
                      f".docx/.pptx d'abord.")
                continue
            if ext not in (TEXT_EXT | OFFICE_EXT):
                continue

            # Nom de sortie : anonymisé par défaut (sauf --garder-noms).
            rel_out = rel if args.garder_noms else anonymiser_nom_relatif(rel, termes_noms)
            dst = os.path.join(args.out, rel_out)
            if not args.dry_run:
                dst = _dst_unique(dst, dst_utilises)
                rel_out = os.path.relpath(dst, args.out)

            if args.dry_run:
                # Compte uniquement
                texte = extraire_texte(src)
                c = Counter()
                remplacer_texte(texte, termes, scrubbers, c)
                compteur = c
            elif ext in TEXT_EXT:
                compteur = traiter_texte(src, dst, termes, scrubbers)
            else:
                compteur = traiter_office(src, dst, termes, scrubbers, args)
            if compteur is None:   # fichier illisible / verrouillé : on saute
                continue

            nb_fichiers += 1
            total.update(compteur)
            if rel_out != rel:
                renommages.append((rel, rel_out))
            for terme, n in sorted(compteur.items()):
                rapport_lignes.append((rel_out, terme, n))
            tag = "[simulé]" if args.dry_run else "[ok]"
            renom = f"  →  {rel_out}" if rel_out != rel else ""
            print(f"  {tag} {rel}{renom} — {sum(compteur.values())} remplacement(s)")

    # --- Rapport ------------------------------------------------------------ #
    if not args.dry_run:
        os.makedirs(args.out, exist_ok=True)
    rapport = args.report or os.path.join(
        args.out if not args.dry_run else os.getcwd(), "_rapport_anonymisation.csv")
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
    if not args.dry_run and paires:
        survivants = verifier(args.out, paires, args.sensible_casse,
                              verifier_noms=not args.garder_noms)
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
    if NON_LUS and not args.dry_run:
        with open(rapport, "a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([])
            w.writerow(["FICHIERS NON LUS (absents du résultat)", "", ""])
            for c in sorted(set(NON_LUS)):
                w.writerow([c, "", ""])


if __name__ == "__main__":
    main()
