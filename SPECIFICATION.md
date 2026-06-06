# Anonymiseur de documents — Spécification

## 1. Objectif

Anonymiser **par lot** des documents bureautiques en remplaçant les éléments
identifiants par des pseudonymes/jetons, et en supprimant les identifiants
techniques et les métadonnées. Pensé pour produire un livrable « anonymisé »
à partir d'un dossier de travail.

- **Formats traités** : `.txt`, `.md`, `.markdown`, `.docx`, `.pptx`, `.xlsx`
- **Formats hérités refusés** (à convertir d'abord) : `.doc`, `.ppt`, `.xls`
- **Zéro dépendance** : bibliothèque standard Python 3.8+ uniquement.

---

## 2. Composants du projet

| Fichier | Rôle |
|---|---|
| `anonymiseur.py` | Programme complet : cœur, ligne de commande **et** menu interactif. |
| `anonymiser.bat` | Lanceur interactif Windows (appelle `python`). Nécessite Python. |
| `dist/anonymiseur.exe` | Exécutable autonome (PyInstaller). **Aucun Python requis.** |
| `correspondances_reference.tsv` | Fichier de correspondances (règles de remplacement). |
| `anonymiseur.spec` | Recette de build PyInstaller (pour recompiler l'`.exe`). |
| `Scan/` | Dossier de sortie du mode scan (créé automatiquement). |

Les **trois points d'entrée** (`.exe`, `.bat`, `.py`) font la même chose et
partagent le même cœur de code.

---

## 3. Modes de fonctionnement

### 3.1 Mode SCAN (pré-remplissage)
Repère des candidats identifiants (acronymes en MAJUSCULES, séquences
capitalisées vues ≥ 2 fois) et écrit un fichier de correspondances pré-rempli
**à compléter**. **Ne modifie aucun document.**

- Sortie : un **nouveau** fichier `correspondances_<dossier>.tsv` dans le
  sous-dossier `Scan/` (jamais d'écrasement : suffixe `_2`, `_3`… si besoin).
- Les termes **déjà présents** dans le fichier de référence sont **exclus**
  des candidats.

### 3.2 Mode ANONYMISATION (par défaut)
Applique les correspondances + les motifs techniques, écrit les documents
nettoyés dans le **dossier de sortie**, produit un **rapport CSV**, puis lance
une **passe de vérification**.

---

## 4. Lancement

### 4.1 Menu interactif
- **`.exe`** : double-clic (ouvre une console) — aucun Python requis.
- **`.bat`** : double-clic — nécessite Python.
- **`python anonymiseur.py`** (sans argument) : ouvre le menu.

Le menu propose : **1. Scanner** · **2. Anonymiser** · **3. Quitter**.

En mode Anonymiser, le lanceur :
1. demande le dossier **source** et le dossier **cible** ;
2. détecte **automatiquement le 1ᵉʳ `.tsv`** présent dans le dossier du lanceur
   (avertissement non bloquant s'il y en a plusieurs) et affiche le **nombre de
   correspondances** ;
3. propose **[1] options par défaut** ou **[2] configuration détaillée** ;
4. affiche un **récapitulatif** et demande **confirmation** avant de lancer.

> L'`.exe` et son fichier `.tsv` doivent être **dans le même dossier**. Le
> sous-dossier `Scan/` se crée à côté de l'`.exe`.

### 4.2 Ligne de commande
```
python anonymiseur.py SOURCE --out SORTIE --map FICHIER.tsv [options]
python anonymiseur.py SOURCE --scan [--exclure reference.tsv]
```

| Option | Effet |
|---|---|
| `--out` | Dossier de sortie (obligatoire hors `--scan`, ≠ source). |
| `--map` | Fichier de correspondances `.tsv`. |
| `--scan` | Mode pré-remplissage (ne modifie rien). |
| `--exclure FIC` | (scan) Exclut les termes déjà connus d'un `.tsv`. |
| `--report` | Chemin du rapport CSV. |
| `--case-sensitive` | Respecte la casse (par défaut : insensible). |
| `--neutraliser-dates` | Efface aussi les dates Office. |
| `--garder-suivi` | Ne touche pas au suivi de modifications Word. |
| `--garder-noms` | N'anonymise **pas** les noms de fichiers/dossiers. |
| `--dry-run` | Simule (compte) sans rien écrire. |
| `--compter` | Affiche le nombre de correspondances de `--map` puis quitte. |
| `--no-email/url/ip/mac/port/path` | Désactive le scrubber correspondant. |

---

## 5. Fichier de correspondances (`.tsv`)

### 5.1 Format
- Une règle par ligne : `terme_original` **<TAB>** `remplacement`.
- Lignes vides et lignes commençant par `#` : ignorées.
- Encodage **UTF-8** (un éventuel BOM est toléré).
- **Tolérance tabulation** : si une ligne n'a pas de tabulation mais se termine
  par un remplacement entre crochets `[...]`, elle est quand même acceptée
  (ex. `plants [PRODUIT2]`).

### 5.2 Deux types de règles
- **Terme littéral** (défaut) : comparé avec **frontières de mot** (`\b`),
  **insensible à la casse** ET **insensible aux accents** (`Eric` attrape
  `Éric`, `Qualité` attrape `qualite`).
- **Règle regex** : préfixe **`re:`**. Le terme de gauche est une expression
  régulière (ni échappée, ni encadrée par `\b`). Pratique pour les URLs, les
  identifiants collés, etc.

### 5.3 Ordre des remplacements
Les règles sont triées par **longueur décroissante** : la plus longue
s'applique en premier (ex. « Marketo Engage » avant « Marketo »). L'ordre des
lignes dans le fichier n'a donc **pas d'importance**.

### 5.4 Exemples
```
# Règles regex (URLs, codes)
re:https?://semaefrance\.sharepoint\.com[^\s"'<>)]*	[LIEN OUTIL COLLABORATIF]
re:(?i)\bsemae\w*	[ORG-S]
re:(?i)(?<!\S)SMA\S*V(?!\S)	[NOM-SERVEUR]

# Termes littéraux
SEMAE	[ORG-S]
Hardis Group	[PRESTA-H]
Eric MICHAUD	[PERSONNE-EMD]
```

---

## 6. Chaîne de traitement d'un texte

Pour chaque fragment de texte, les remplacements sont appliqués dans **cet
ordre précis** (`remplacer_texte`) :

1. **Scrubber EMAIL** — un email est un bloc atomique, neutralisé en premier
   pour qu'aucune règle ne casse son domaine
   (`x@semae.fr` → `[EMAIL]`, pas `x@[ORG-S].fr`).
2. **Règles regex** (`re:`) du fichier — ex. URLs internes.
3. **Autres scrubbers techniques** (url, ip, mac, port, chemin).
4. **Termes exacts** (littéraux) du fichier.

### 6.1 Scrubbers techniques (actifs par défaut)
| Nom | Cible | Remplacement |
|---|---|---|
| `email` | adresses email | `[EMAIL]` |
| `url` | `http(s)://…` | `[URL]` |
| `ip` | IPv4 / IPv6 | `[IP]` / `[IPV6]` |
| `mac` | adresses MAC (avant IPv6) | `[MAC]` |
| `port` | `IP:port`, `port N`, `port https N` | `…:[PORT]`, `port [PORT]`, `port https [PORT]` |
| `path` | UNC, `C:\…`, `/mnt|home|srv|opt|var/…` | `[CHEMIN]` |

---

## 7. Traitement des fichiers Office (`.docx` / `.pptx` / `.xlsx`)

Un fichier Office est un **ZIP de fichiers XML**. Le traitement :

- **Texte** : remplacé uniquement dans les **nœuds texte** (entre `>` et `<`),
  jamais dans les balises. Pour `.xlsx` : `xl/sharedStrings.xml` et feuilles.
- **Métadonnées** (`docProps/core.xml`, `app.xml`, `custom.xml`) : vide
  `creator`, `lastModifiedBy`, `title`, `subject`, `keywords`, `description`,
  `category`, `Company`, `Manager`, propriétés personnalisées… ; dates
  effacées si `--neutraliser-dates`.
- **Suivi de modifications Word** (best-effort, sauf `--garder-suivi`) :
  garde les insertions, supprime les suppressions et les marqueurs de
  changement de format.
- **Commentaires** : parts et références supprimées (Word/PowerPoint/Excel),
  ainsi qu'auteurs, personnes, miniatures.
- **Hyperliens** (`.rels`) : les cibles externes (`http(s)`, `mailto`, `ftp`)
  sont anonymisées. Si le résultat n'est plus une URL absolue, la cible est
  forcée à `https://anonymise.invalid/` pour **éviter que Word la résolve en
  lien relatif** (ce qui révélerait le chemin OneDrive du document).

---

## 8. Anonymisation des noms de fichiers/dossiers

Activée **par défaut** (désactivable avec `--garder-noms`) :
- les **termes littéraux** du `.tsv` sont appliqués à chaque composant du
  chemin (sous-dossiers + nom de fichier) ;
- les caractères interdits par Windows (`< > : " / \ | ? *`) résiduels sont
  neutralisés (les crochets `[ ]` sont valides) ;
- en cas de **collision** (deux fichiers donnant le même nom), un suffixe
  `_2`, `_3`… est ajouté (jamais d'écrasement).

Les règles **regex** et les scrubbers ne sont **pas** appliqués aux noms.

---

## 9. Sortie et rapports

- L'**arborescence** du dossier source est recréée dans le dossier de sortie
  (seuls les fichiers traités y figurent).
- **Rapport CSV** `_rapport_anonymisation.csv` (dans le dossier de sortie) :
  - remplacements par fichier et par terme/motif ;
  - total ;
  - section **RENOMMAGES** (ancien nom → nom anonymisé) ;
  - section **SURVIVANTS** (termes subsistants) le cas échéant ;
  - section **FICHIERS NON LUS** le cas échéant.

### 9.1 Fichiers non lus
Un fichier **verrouillé** (ouvert dans Office), **OneDrive hors-ligne** ou
**corrompu** est **ignoré** (jamais bloquant) : avertissement immédiat +
**récapitulatif final** + ligne dans le CSV. ⚠️ Un fichier ignoré est **absent
du résultat** : le fermer/télécharger puis relancer.

### 9.2 Vérification finale
Après écriture, la sortie est re-balayée pour repérer tout terme du `.tsv`
**ayant survécu** (dans le contenu **et** les noms, sauf `--garder-noms`).
Cause typique d'un survivant : un terme **coupé entre deux « runs » Word**
(ouvrir/réenregistrer le `.docx` dans Word puis relancer corrige souvent).

---

## 10. Limites connues

- Un terme **coupé entre deux runs** Word peut échapper au remplacement
  (la vérification le signale).
- L'acceptation du **suivi de modifications** est best-effort (regex).
- Les formats **hérités** `.doc/.ppt/.xls` ne sont pas traités.
- Matching **insensible aux accents et à la casse**, mais un terme **collé à un
  `_`** (ex. `semae_x`) nécessite une règle **regex** (le `_` est un caractère
  de mot, donc pas de frontière `\b`).
- **Images / logos** non traités (pas d'analyse d'image, principe zéro
  dépendance).
- Le script neutralise les **noms**, pas le **contexte métier** (secteur,
  budget, stack) sauf si ajouté au `.tsv`.

---

## 11. Construction de l'exécutable

```
python -m pip install pyinstaller
python -m PyInstaller --onefile --console --name anonymiseur anonymiseur.py
```
Résultat : `dist/anonymiseur.exe` (autonome). Le fichier `anonymiseur.spec`
permet de reconstruire avec les mêmes options : `python -m PyInstaller anonymiseur.spec`.

> Au 1ᵉʳ lancement, **Windows SmartScreen / l'antivirus** peut avertir
> (`.exe` non signé) → « Informations complémentaires » → « Exécuter quand même ».

---

## 12. Workflow recommandé

1. **Scanner** le dossier cible → un `.tsv` de candidats dans `Scan/`.
2. **Compléter** la colonne de droite et **recopier** les lignes utiles dans
   `correspondances_reference.tsv` (le fichier de référence).
3. **Anonymiser** : source = dossier d'origine, cible = dossier vide
   (idéalement **hors OneDrive**), `.tsv` = la référence.
4. **Vérifier** le message final + le rapport CSV ; traiter les éventuels
   survivants et fichiers non lus.
