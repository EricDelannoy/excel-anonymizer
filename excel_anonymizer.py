#!/usr/bin/env python3
"""Anonymise un classeur Excel (.xlsx / .xlsm), détecte les données cachées et
supprime tout ce qui n'est pas nécessaire.

Le fichier est traité au niveau du paquet OOXML (zip + XML) plutôt que via
openpyxl : les graphiques, images, tableaux croisés dynamiques, styles, etc.
ne sont donc pas perdus lors du nettoyage.

Usage :
    python excel_anonymizer.py classeur.xlsx                 # -> classeur.clean.xlsx
    python excel_anonymizer.py classeur.xlsx --audit-only    # rapport seul
    python excel_anonymizer.py --help
"""

from __future__ import annotations

import argparse
import datetime as dt
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from lxml import etree

# --------------------------------------------------------------------------- #
# Constantes
# --------------------------------------------------------------------------- #

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

CT_XML = "[Content_Types].xml"
ROOT_RELS = "_rels/.rels"

# Types de relations toujours supprimés (clé = dernier segment du type).
ALWAYS_REMOVED = {
    "thumbnail": "miniature d'aperçu",
    "custom-properties": "propriétés personnalisées",
    "printerSettings": "paramètres d'imprimante",
    "calcChain": "chaîne de calcul (recalculée par Excel)",
    "customXml": "données XML personnalisées",
    "revisionHeaders": "historique des modifications suivies",
    "revisionLog": "journal des modifications suivies",
    "usernames": "noms d'utilisateurs des modifications suivies",
}

# Types de relations liés aux commentaires.
COMMENT_KINDS = {"comments", "vmlDrawing", "threadedComment"}

SECRET_SAFE_PARSER = etree.XMLParser(
    resolve_entities=False, no_network=True, huge_tree=False, remove_blank_text=False
)


class CleanError(Exception):
    """Erreur bloquante : le nettoyage ne peut pas être réalisé en sécurité."""


@dataclass
class Finding:
    level: str  # "info" | "warning"
    text: str


@dataclass
class Options:
    author: str = "Anonyme"
    keep_hidden_sheets: bool = False
    remove_comments: bool = False
    purge_hidden_cells: bool = False
    keep_dates: bool = False
    force: bool = False


@dataclass
class SheetInfo:
    index: int
    name: str
    state: str
    rid: str
    part: str | None


# --------------------------------------------------------------------------- #
# Utilitaires XML
# --------------------------------------------------------------------------- #


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def local(el: etree._Element) -> str:
    tag = el.tag
    return etree.QName(tag).localname if isinstance(tag, str) else ""


def rel_kind(rel_type: str) -> str:
    return rel_type.rstrip("/").rsplit("/", 1)[-1]


def col_to_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref: str) -> tuple[int, int] | None:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    return (col_to_index(m.group(1)), int(m.group(2))) if m else None


def is_true(value: str | None) -> bool:
    return value in ("1", "true")


def sheet_ref_regex(name: str) -> re.Pattern[str]:
    quoted = "'" + name.replace("'", "''") + "'!"
    return re.compile(
        "(?:" + re.escape(quoted) + r"|(?<![\w.'])" + re.escape(name) + "!)",
        re.IGNORECASE,
    )


# --------------------------------------------------------------------------- #
# Paquet OOXML
# --------------------------------------------------------------------------- #


class Package:
    """Contenu d'un classeur .xlsx chargé en mémoire."""

    def __init__(self, files: dict[str, bytes]):
        self.files = files
        self._trees: dict[str, etree._ElementTree] = {}

    # -- chargement / sauvegarde ------------------------------------------- #
    @classmethod
    def load(cls, path: Path) -> "Package":
        files: dict[str, bytes] = {}
        try:
            with zipfile.ZipFile(path) as zf:
                if zf.testzip() is not None:
                    raise CleanError("Archive zip corrompue.")
                for info in zf.infolist():
                    name = info.filename
                    if info.is_dir():
                        continue
                    if name.startswith("/") or ".." in name.split("/"):
                        raise CleanError(f"Nom de fichier suspect dans l'archive : {name}")
                    files[name] = zf.read(info)
        except zipfile.BadZipFile as exc:
            raise CleanError(f"{path} n'est pas un classeur Excel valide (.xlsx/.xlsm).") from exc
        if CT_XML not in files or ROOT_RELS not in files:
            raise CleanError(f"{path} n'est pas un classeur Excel OOXML (.xlsx/.xlsm).")
        return cls(files)

    def save(self, path: Path) -> None:
        for name, tree in self._trees.items():
            if name in self.files:
                self.files[name] = etree.tostring(
                    tree, xml_declaration=True, encoding="UTF-8", standalone=True
                )
        order = [CT_XML] + [n for n in self.files if n != CT_XML]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in order:
                zf.writestr(name, self.files[name])

    # -- accès XML --------------------------------------------------------- #
    def has(self, name: str) -> bool:
        return name in self.files

    def xml(self, name: str) -> etree._ElementTree | None:
        if name not in self.files:
            return None
        if name not in self._trees:
            try:
                self._trees[name] = etree.parse(BytesIO(self.files[name]), SECRET_SAFE_PARSER)
            except etree.XMLSyntaxError as exc:
                raise CleanError(f"XML invalide dans {name} : {exc}") from exc
        return self._trees[name]

    def root(self, name: str) -> etree._Element | None:
        tree = self.xml(name)
        return tree.getroot() if tree is not None else None

    def delete(self, name: str) -> None:
        self.files.pop(name, None)
        self._trees.pop(name, None)

    # -- relations --------------------------------------------------------- #
    @staticmethod
    def rels_path(part: str) -> str:
        if part == "":
            return ROOT_RELS
        d, b = posixpath.split(part)
        return posixpath.join(d, "_rels", b + ".rels")

    def rels(self, part: str) -> list[tuple[etree._Element, str | None]]:
        """Relations d'une partie : (élément, partie cible ou None si externe)."""
        root = self.root(self.rels_path(part))
        if root is None:
            return []
        out = []
        base = posixpath.dirname(part)
        for rel in list(root):
            if local(rel) != "Relationship":
                continue
            if rel.get("TargetMode") == "External":
                out.append((rel, None))
                continue
            target = rel.get("Target", "")
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(posixpath.join(base, target))
            out.append((rel, resolved))
        return out

    def rels_of_kind(self, part: str, *kinds: str) -> list[tuple[etree._Element, str | None]]:
        return [(r, t) for r, t in self.rels(part) if rel_kind(r.get("Type", "")) in kinds]

    def remove_rel(self, part: str, rel: etree._Element) -> None:
        """Supprime une relation et les références à son Id dans la partie source."""
        rid = rel.get("Id")
        parent = rel.getparent()
        if parent is not None:
            parent.remove(rel)
        tree = self.xml(part) if part and part.endswith(".xml") else None
        if tree is None or rid is None:
            return
        for el in list(tree.getroot().iter()):
            if not isinstance(el.tag, str):
                continue
            hits = [a for a, v in el.attrib.items() if a.startswith(f"{{{NS_REL}}}") and v == rid]
            if not hits:
                continue
            if local(el) == "pageSetup":
                for a in hits:
                    del el.attrib[a]
            elif el.getparent() is not None:
                el.getparent().remove(el)

    def remove_rels_of_kind(self, part: str, *kinds: str) -> int:
        n = 0
        for rel, _ in self.rels_of_kind(part, *kinds):
            self.remove_rel(part, rel)
            n += 1
        return n

    # -- ramasse-miettes --------------------------------------------------- #
    def collect_garbage(self) -> list[str]:
        """Supprime les parties qui ne sont plus atteignables par les relations."""
        reachable: set[str] = {CT_XML}
        stack = [""]
        seen_sources: set[str] = set()
        while stack:
            source = stack.pop()
            if source in seen_sources:
                continue
            seen_sources.add(source)
            if self.has(self.rels_path(source)):
                reachable.add(self.rels_path(source))
            for _, target in self.rels(source):
                if target and target in self.files and target not in reachable:
                    reachable.add(target)
                    stack.append(target)
        removed = [n for n in list(self.files) if n not in reachable]
        for name in removed:
            self.delete(name)
        ct = self.root(CT_XML)
        if ct is not None:
            for el in list(ct):
                if local(el) == "Override" and el.get("PartName", "").lstrip("/") not in self.files:
                    ct.remove(el)
        return removed

    # -- structure du classeur -------------------------------------------- #
    @property
    def workbook_part(self) -> str:
        for rel, target in self.rels(""):
            if rel_kind(rel.get("Type", "")) == "officeDocument" and target:
                return target
        raise CleanError("Partie workbook introuvable.")

    def sheets(self) -> list[SheetInfo]:
        wb = self.workbook_part
        root = self.root(wb)
        rel_targets = {r.get("Id"): t for r, t in self.rels(wb)}
        out = []
        sheets_el = root.find(q(NS_MAIN, "sheets"))
        if sheets_el is None:
            return out
        for i, sh in enumerate(sheets_el.findall(q(NS_MAIN, "sheet"))):
            rid = sh.get(q(NS_REL, "id"), "")
            out.append(SheetInfo(i, sh.get("name", ""), sh.get("state", "visible"), rid, rel_targets.get(rid)))
        return out

    def shared_strings_part(self) -> str | None:
        for _, target in self.rels_of_kind(self.workbook_part, "sharedStrings"):
            return target
        return None

    def shared_strings(self) -> list[str]:
        part = self.shared_strings_part()
        root = self.root(part) if part else None
        if root is None:
            return []
        return ["".join(si.itertext()) for si in root.findall(q(NS_MAIN, "si"))]


# --------------------------------------------------------------------------- #
# Détection (lecture seule)
# --------------------------------------------------------------------------- #


def hidden_rows_cols(sheet: etree._Element) -> tuple[list[int], list[tuple[int, int]], bool]:
    """Retourne (lignes masquées, plages de colonnes masquées, présence d'un filtre)."""
    rows = [
        int(r.get("r"))
        for r in sheet.iter(q(NS_MAIN, "row"))
        if is_true(r.get("hidden")) and r.get("r")
    ]
    cols = [
        (int(c.get("min")), int(c.get("max")))
        for c in sheet.iter(q(NS_MAIN, "col"))
        if is_true(c.get("hidden")) and c.get("min") and c.get("max")
    ]
    has_filter = sheet.find(q(NS_MAIN, "autoFilter")) is not None
    return rows, cols, has_filter


def invisible_numfmt_xfs(pkg: Package) -> set[int]:
    """Index des styles de cellule dont le format est « ;;; » (valeur invisible)."""
    styles = None
    for _, target in pkg.rels_of_kind(pkg.workbook_part, "styles"):
        styles = pkg.root(target) if target else None
    if styles is None:
        return set()
    hidden_ids = set()
    numfmts = styles.find(q(NS_MAIN, "numFmts"))
    if numfmts is not None:
        for nf in numfmts:
            code = nf.get("formatCode", "")
            if code and set(code) <= {";"}:
                hidden_ids.add(nf.get("numFmtId"))
    xfs = styles.find(q(NS_MAIN, "cellXfs"))
    if xfs is None or not hidden_ids:
        return set()
    return {i for i, xf in enumerate(xfs) if xf.get("numFmtId") in hidden_ids}


def invisible_cells(sheet: etree._Element, xf_ids: set[int]) -> int:
    if not xf_ids:
        return 0
    return sum(
        1
        for c in sheet.iter(q(NS_MAIN, "c"))
        if c.get("s", "").isdigit() and int(c.get("s")) in xf_ids and len(c)
    )


def hidden_drawing_anchors(drawing: etree._Element) -> list[etree._Element]:
    anchors = []
    for el in drawing.iter():
        if isinstance(el.tag, str) and local(el) == "cNvPr" and is_true(el.get("hidden")):
            node = el
            while node is not None and not (isinstance(node.tag, str) and local(node).endswith("Anchor")):
                node = node.getparent()
            if node is not None and node not in anchors:
                anchors.append(node)
    return anchors


def sheet_parts(pkg: Package, include_hidden: bool = True) -> list[SheetInfo]:
    return [s for s in pkg.sheets() if s.part and (include_hidden or s.state == "visible")]


def identity_values(pkg: Package) -> dict[str, str]:
    """Valeurs identifiantes trouvées dans le fichier : {valeur: origine}."""
    found: dict[str, str] = {}

    def add(value: str | None, origin: str) -> None:
        if value and value.strip():
            found.setdefault(value.strip(), origin)

    core = next((t for r, t in pkg.rels("") if rel_kind(r.get("Type", "")) == "core-properties"), None)
    root = pkg.root(core) if core else None
    if root is not None:
        add(root.findtext(q(NS_DC, "creator")), "auteur (propriétés du document)")
        add(root.findtext(q(NS_CP, "lastModifiedBy")), "dernier modificateur")
    app = next((t for r, t in pkg.rels("") if rel_kind(r.get("Type", "")) == "extended-properties"), None)
    root = pkg.root(app) if app else None
    if root is not None:
        add(root.findtext(q(NS_EXT_PROPS, "Company")), "société")
        add(root.findtext(q(NS_EXT_PROPS, "Manager")), "responsable")
    for s in sheet_parts(pkg):
        for _, target in pkg.rels_of_kind(s.part, "comments"):
            croot = pkg.root(target) if target else None
            if croot is not None:
                for a in croot.iter(q(NS_MAIN, "author")):
                    add(a.text, f"auteur de commentaire (« {s.name} »)")
    for rel, target in pkg.rels_of_kind(pkg.workbook_part, "person"):
        proot = pkg.root(target) if target else None
        if proot is not None:
            for p in proot:
                add(p.get("displayName"), "personne (commentaires modernes)")
    return found


def audit(pkg: Package, opts: Options | None = None) -> list[Finding]:
    """Inspecte le classeur et retourne la liste de ce qui a été trouvé."""
    findings: list[Finding] = []
    add = lambda level, text: findings.append(Finding(level, text))  # noqa: E731
    wb_part = pkg.workbook_part
    wb = pkg.root(wb_part)

    # --- Identité ---------------------------------------------------------- #
    author = opts.author if opts else None
    ids = {v: o for v, o in identity_values(pkg).items() if v != author}
    for value, origin in ids.items():
        add("warning", f"Identité présente : « {value} » — {origin}")
    if wb is not None:
        if wb.xpath("//*[local-name()='absPath']"):
            add("warning", "Chemin absolu du fichier d'origine enregistré (peut contenir un nom d'utilisateur).")
        fs = wb.find(q(NS_MAIN, "fileSharing"))
        if fs is not None and fs.get("userName"):
            add("warning", f"Nom d'utilisateur de partage : « {fs.get('userName')} »")

    # Noms propres retrouvés dans le contenu des cellules / commentaires
    texts = pkg.shared_strings()
    for s in sheet_parts(pkg):
        for _, target in pkg.rels_of_kind(s.part, "comments"):
            croot = pkg.root(target) if target else None
            if croot is not None:
                texts.extend("".join(c.itertext()) for c in croot.iter(q(NS_MAIN, "comment")))
    blob = "\n".join(texts).lower()
    for value in ids:
        if len(value) >= 3 and value.lower() in blob:
            add("warning", f"« {value} » apparaît aussi dans le contenu des cellules ou des commentaires "
                           "(non modifié automatiquement : à vérifier manuellement).")

    # --- Onglets masqués ---------------------------------------------------- #
    for s in pkg.sheets():
        if s.state in ("hidden", "veryHidden"):
            label = "très masqué (invisible depuis l'interface)" if s.state == "veryHidden" else "masqué"
            add("warning", f"Onglet {label} : « {s.name} »")

    # --- Par feuille -------------------------------------------------------- #
    xf_ids = invisible_numfmt_xfs(pkg)
    for s in sheet_parts(pkg):
        sroot = pkg.root(s.part)
        if sroot is None:
            continue
        rows, cols, has_filter = hidden_rows_cols(sroot)
        if rows:
            extra = " (probablement dû à un filtre)" if has_filter else ""
            add("warning", f"« {s.name} » : {len(rows)} ligne(s) masquée(s){extra}")
        if cols:
            n = sum(hi - lo + 1 for lo, hi in cols)
            add("warning", f"« {s.name} » : {n} colonne(s) masquée(s)")
        n_inv = invisible_cells(sroot, xf_ids)
        if n_inv:
            add("warning", f"« {s.name} » : {n_inv} cellule(s) au format « ;;; » (valeur invisible)")
        for _, target in pkg.rels_of_kind(s.part, "comments"):
            croot = pkg.root(target) if target else None
            if croot is not None:
                n = len(croot.findall(f".//{q(NS_MAIN, 'comment')}"))
                add("info", f"« {s.name} » : {n} commentaire(s)")
        for _, target in pkg.rels_of_kind(s.part, "drawing"):
            droot = pkg.root(target) if target else None
            if droot is not None:
                n = len(hidden_drawing_anchors(droot))
                if n:
                    add("warning", f"« {s.name} » : {n} objet(s) graphique(s) masqué(s)")

    # --- Noms définis ------------------------------------------------------- #
    for dn in iter_removable_names(pkg):
        add("warning", f"Nom défini inutilisé et masqué/cassé : « {dn.get('name')} »")

    # --- Références de formules vers des onglets masqués -------------------- #
    hidden = [s for s in pkg.sheets() if s.state != "visible"]
    for name, cells in formulas_referencing(pkg, hidden).items():
        add("warning", f"{len(cells)} formule(s) des onglets visibles dépendent de l'onglet masqué « {name} » "
                       f"(ex. {', '.join(cells[:3])})")

    # --- Éléments divers ---------------------------------------------------- #
    files = pkg.files
    checks = [
        ("vbaProject.bin", "Macros VBA présentes (non modifiées : peuvent contenir des noms/chemins)."),
        ("xl/externalLinks/", "Liens externes vers d'autres classeurs."),
        ("xl/pivotCache/", "Cache de tableau croisé dynamique (contient une copie des données sources)."),
        ("xl/connections.xml", "Connexions de données (chaînes de connexion possibles)."),
        ("xl/embeddings/", "Objets incorporés (OLE)."),
        ("xl/revisions/", "Historique de modifications suivies."),
        ("customXml/", "Données XML personnalisées."),
        ("docProps/thumbnail", "Miniature d'aperçu du document."),
        ("docProps/custom.xml", "Propriétés personnalisées."),
    ]
    for needle, message in checks:
        if any(needle in n for n in files):
            level = "info" if needle in ("docProps/thumbnail", "docProps/custom.xml", "customXml/", "xl/revisions/") else "warning"
            add(level, message)
    for part in list(files):
        if part.startswith("xl/externalLinks/_rels/"):
            root = pkg.root(part)
            for rel in root if root is not None else []:
                if rel.get("TargetMode") == "External":
                    add("warning", f"Lien externe : {rel.get('Target')}")
    return findings


def iter_removable_names(pkg: Package) -> list[etree._Element]:
    """Noms définis masqués (hors noms internes Excel) ou cassés, et non utilisés."""
    wb = pkg.root(pkg.workbook_part)
    dns = wb.find(q(NS_MAIN, "definedNames"))
    if dns is None:
        return []
    formulas = all_formula_texts(pkg)
    out = []
    for dn in dns:
        name = dn.get("name", "")
        value = dn.text or ""
        if name.startswith("_xlnm."):
            continue
        if not (is_true(dn.get("hidden")) or "#REF!" in value):
            continue
        pattern = re.compile(r"(?<![\w.])" + re.escape(name) + r"(?![\w.(])")
        used = any(pattern.search(f) for f in formulas) or any(
            pattern.search(o.text or "") for o in dns if o is not dn
        )
        if not used:
            out.append(dn)
    return out


def all_formula_texts(pkg: Package) -> list[str]:
    out = []
    for s in sheet_parts(pkg):
        root = pkg.root(s.part)
        if root is None:
            continue
        out.extend(f.text for f in root.iter(q(NS_MAIN, "f")) if f.text)
    return out


def formulas_referencing(pkg: Package, sheets: list[SheetInfo]) -> dict[str, list[str]]:
    """Cellules (feuilles visibles) dont la formule référence l'une des `sheets`."""
    result: dict[str, list[str]] = {}
    removed_parts = {s.part for s in sheets}
    for target in sheets:
        regex = sheet_ref_regex(target.name)
        for s in sheet_parts(pkg):
            if s.part in removed_parts:
                continue
            root = pkg.root(s.part)
            for c in root.iter(q(NS_MAIN, "c")):
                f = c.find(q(NS_MAIN, "f"))
                if f is not None and f.text and regex.search(f.text):
                    result.setdefault(target.name, []).append(f"{s.name}!{c.get('r')}")
    return result


# --------------------------------------------------------------------------- #
# Nettoyage (modifie le paquet)
# --------------------------------------------------------------------------- #


def anonymize_properties(pkg: Package, opts: Options, actions: list[str]) -> None:
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    for rel, target in pkg.rels(""):
        kind = rel_kind(rel.get("Type", ""))
        if kind == "core-properties" and target:
            root = pkg.root(target)
            if root is None:
                continue
            for tag_ns, tag in ((NS_DC, "creator"), (NS_CP, "lastModifiedBy")):
                el = root.find(q(tag_ns, tag))
                if el is not None:
                    el.text = opts.author
            for tag_ns, tag in (
                (NS_DC, "subject"), (NS_CP, "keywords"), (NS_DC, "description"),
                (NS_CP, "category"), (NS_CP, "contentStatus"), (NS_CP, "lastPrinted"),
            ):
                el = root.find(q(tag_ns, tag))
                if el is not None:
                    root.remove(el)
            if not opts.keep_dates:
                for tag in ("created", "modified"):
                    el = root.find(q(NS_DCTERMS, tag))
                    if el is not None:
                        el.text = now
            actions.append(f"Auteur et dernier modificateur remplacés par « {opts.author} »")
        elif kind == "extended-properties" and target:
            # Réécrit app.xml au strict minimum : supprime société, responsable
            # et la liste des noms d'onglets (qui révélerait les onglets masqués).
            pkg.files[target] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Properties xmlns="{NS_EXT_PROPS}"><Application>Microsoft Excel</Application></Properties>'
            ).encode()
            pkg._trees.pop(target, None)
            actions.append("Propriétés étendues (société, responsable, noms d'onglets) réduites au minimum")


def anonymize_comment_authors(pkg: Package, opts: Options, actions: list[str]) -> None:
    n = 0
    for s in sheet_parts(pkg):
        for _, target in pkg.rels_of_kind(s.part, "comments"):
            root = pkg.root(target) if target else None
            if root is None:
                continue
            for a in root.iter(q(NS_MAIN, "author")):
                if a.text != opts.author:
                    a.text = opts.author
                    n += 1
    for _, target in pkg.rels_of_kind(pkg.workbook_part, "person"):
        root = pkg.root(target) if target else None
        if root is None:
            continue
        for p in root:
            if not isinstance(p.tag, str):
                continue
            p.set("displayName", opts.author)
            p.set("userId", "anonymous")
            p.set("providerId", "None")
            n += 1
    if n:
        actions.append(f"{n} auteur(s) de commentaires anonymisé(s)")


def remove_always_unneeded(pkg: Package, actions: list[str]) -> None:
    sources = [""] + [pkg.workbook_part] + [s.part for s in sheet_parts(pkg)]
    counts: dict[str, int] = {}
    for src in sources:
        for rel, _ in list(pkg.rels(src)):
            kind = rel_kind(rel.get("Type", ""))
            if kind in ALWAYS_REMOVED:
                pkg.remove_rel(src, rel)
                counts[kind] = counts.get(kind, 0) + 1
    for kind, n in counts.items():
        actions.append(f"Supprimé : {ALWAYS_REMOVED[kind]} ({n})")


def remove_comments(pkg: Package, actions: list[str]) -> None:
    n = 0
    for s in sheet_parts(pkg):
        n += pkg.remove_rels_of_kind(s.part, *COMMENT_KINDS)
    n += pkg.remove_rels_of_kind(pkg.workbook_part, "person")
    if n:
        actions.append("Commentaires supprimés (notes et fils de discussion)")


def remove_hidden_drawing_objects(pkg: Package, actions: list[str]) -> None:
    total = 0
    for s in sheet_parts(pkg):
        for _, dpart in pkg.rels_of_kind(s.part, "drawing"):
            droot = pkg.root(dpart) if dpart else None
            if droot is None:
                continue
            for anchor in hidden_drawing_anchors(droot):
                rids = {v for el in anchor.iter() for a, v in el.attrib.items() if a.startswith(f"{{{NS_REL}}}")}
                top = anchor
                while top.getparent() is not None and local(top.getparent()) in ("Choice", "Fallback", "AlternateContent"):
                    top = top.getparent()
                top.getparent().remove(top)
                still_used = {v for el in droot.iter() for a, v in el.attrib.items() if a.startswith(f"{{{NS_REL}}}")}
                for rel, _ in pkg.rels(dpart):
                    if rel.get("Id") in rids and rel.get("Id") not in still_used:
                        pkg.remove_rel(dpart, rel)
                total += 1
    if total:
        actions.append(f"{total} objet(s) graphique(s) masqué(s) supprimé(s)")


def purge_hidden_cells(pkg: Package, actions: list[str]) -> None:
    for s in sheet_parts(pkg):
        root = pkg.root(s.part)
        rows, cols, has_filter = hidden_rows_cols(root)
        removed = 0
        sheet_data = root.find(q(NS_MAIN, "sheetData"))
        if sheet_data is None:
            continue
        hidden_rows = set() if has_filter else set(rows)
        for row in sheet_data:
            r_hidden = (int(row.get("r", 0)) in hidden_rows)
            for c in list(row):
                ref = split_ref(c.get("r", ""))
                in_hidden_col = ref and any(lo <= ref[0] <= hi for lo, hi in cols)
                if (r_hidden or in_hidden_col) and len(c):
                    for child in list(c):
                        c.remove(child)
                    c.attrib.pop("t", None)
                    removed += 1
        if removed:
            actions.append(f"« {s.name} » : contenu de {removed} cellule(s) masquée(s) vidé")
        if has_filter and rows:
            actions.append(f"« {s.name} » : lignes masquées conservées (filtre automatique actif)")


def remove_hidden_sheets(pkg: Package, opts: Options, actions: list[str]) -> None:
    sheets = pkg.sheets()
    hidden = [s for s in sheets if s.state in ("hidden", "veryHidden")]
    if not hidden:
        return
    if len(hidden) == len(sheets):
        raise CleanError("Tous les onglets sont masqués : impossible de les supprimer.")

    dependents = formulas_referencing(pkg, hidden)
    if dependents and not opts.force:
        lines = [f"  - « {n} » : {', '.join(c[:5])}{'…' if len(c) > 5 else ''}" for n, c in dependents.items()]
        raise CleanError(
            "Des formules d'onglets visibles dépendent d'onglets masqués :\n"
            + "\n".join(lines)
            + "\nUtilisez --force pour remplacer ces formules par leurs valeurs calculées, "
            "ou --keep-hidden-sheets pour conserver les onglets masqués."
        )

    hidden_parts = {s.part for s in hidden}
    if dependents:
        converted = convert_formulas_to_values(pkg, hidden, hidden_parts)
        actions.append(f"{converted} formule(s) remplacée(s) par leur valeur (dépendance à un onglet masqué)")

    wb_part = pkg.workbook_part
    wb = pkg.root(wb_part)
    sheets_el = wb.find(q(NS_MAIN, "sheets"))
    removed_idx = {s.index for s in hidden}
    # Décalage des index pour les noms définis à portée feuille.
    remap, new_i = {}, 0
    for s in sheets:
        if s.index not in removed_idx:
            remap[s.index] = new_i
            new_i += 1
    dns = wb.find(q(NS_MAIN, "definedNames"))
    if dns is not None:
        for dn in list(dns):
            local_id = dn.get("localSheetId")
            if local_id is None:
                continue
            if int(local_id) in removed_idx:
                dns.remove(dn)
            else:
                dn.set("localSheetId", str(remap[int(local_id)]))
        # Noms qui pointent vers un onglet supprimé.
        for dn in list(dns):
            if any(sheet_ref_regex(s.name).search(dn.text or "") for s in hidden):
                dns.remove(dn)
        if len(dns) == 0:
            wb.remove(dns)

    for el in list(sheets_el):
        if el.get(q(NS_REL, "id")) in {s.rid for s in hidden}:
            sheets_el.remove(el)
    for rel, _ in pkg.rels(wb_part):
        if rel.get("Id") in {s.rid for s in hidden}:
            pkg.remove_rel(wb_part, rel)

    for s in hidden:
        actions.append(f"Onglet masqué supprimé : « {s.name} »")


def select_first_sheet(pkg: Package, actions: list[str]) -> None:
    """Fait en sorte que le classeur s'ouvre sur la première feuille visible.

    Met à jour l'onglet actif du classeur et retire la sélection de toutes les
    autres feuilles (plusieurs feuilles sélectionnées = « mode groupe » dans
    Excel, où une saisie se répercute sur chacune d'elles).
    """
    sheets = [s for s in pkg.sheets() if s.part]
    first = next((s for s in sheets if s.state == "visible"), None)
    if first is None:
        return
    for view in pkg.root(pkg.workbook_part).iter(q(NS_MAIN, "workbookView")):
        view.set("activeTab", str(first.index))
        view.set("firstSheet", "0")
    for s in sheets:
        for sv in pkg.root(s.part).iter(q(NS_MAIN, "sheetView")):
            if s is first:
                sv.set("tabSelected", "1")
            else:
                sv.attrib.pop("tabSelected", None)
    actions.append(f"Le classeur s'ouvrira sur la première feuille : « {first.name} »")


def convert_formulas_to_values(pkg: Package, hidden: list[SheetInfo], hidden_parts: set[str | None]) -> int:
    regexes = [sheet_ref_regex(s.name) for s in hidden]
    count = 0
    for s in sheet_parts(pkg):
        if s.part in hidden_parts:
            continue
        root = pkg.root(s.part)
        shared_ids = set()
        for f in root.iter(q(NS_MAIN, "f")):
            if f.text and any(r.search(f.text) for r in regexes) and f.get("t") == "shared":
                shared_ids.add(f.get("si"))
        for c in root.iter(q(NS_MAIN, "c")):
            f = c.find(q(NS_MAIN, "f"))
            if f is None:
                continue
            matches = bool(f.text and any(r.search(f.text) for r in regexes))
            in_group = f.get("t") == "shared" and f.get("si") in shared_ids
            if matches or in_group:
                c.remove(f)
                count += 1
    return count


def clean_workbook_xml(pkg: Package, actions: list[str]) -> None:
    wb = pkg.root(pkg.workbook_part)
    removed = 0
    for el in list(wb.iter()):
        if not isinstance(el.tag, str):
            continue
        name = local(el)
        parent = el.getparent()
        if parent is None:
            continue
        # absPath est enveloppé dans un mc:AlternateContent : on retire l'enveloppe.
        if name == "AlternateContent" and el.xpath(".//*[local-name()='absPath']"):
            parent.remove(el)
            removed += 1
        elif name in ("revisionPtr", "fileSharing"):
            parent.remove(el)
            removed += 1
    if removed:
        actions.append("Chemin d'origine et identifiants de révision supprimés de workbook.xml")

    unused = iter_removable_names(pkg)
    dns = wb.find(q(NS_MAIN, "definedNames"))
    for dn in unused:
        actions.append(f"Nom défini inutile supprimé : « {dn.get('name')} »")
        dns.remove(dn)
    if dns is not None and len(dns) == 0:
        wb.remove(dns)


def compact_shared_strings(pkg: Package, actions: list[str]) -> None:
    """Retire de sharedStrings.xml les chaînes qui ne sont plus utilisées.

    Sans cela, le texte de cellules ou d'onglets supprimés resterait lisible
    dans le fichier même s'il n'apparaît plus dans Excel.
    """
    part = pkg.shared_strings_part()
    root = pkg.root(part) if part else None
    if root is None:
        return
    items = root.findall(q(NS_MAIN, "si"))
    used: list[int] = []
    cells = []
    for s in sheet_parts(pkg):
        for c in pkg.root(s.part).iter(q(NS_MAIN, "c")):
            if c.get("t") == "s":
                v = c.find(q(NS_MAIN, "v"))
                if v is not None and v.text and v.text.isdigit():
                    cells.append(v)
                    used.append(int(v.text))
    keep = sorted(set(used))
    if len(keep) == len(items):
        return
    mapping = {old: new for new, old in enumerate(keep)}
    for v in cells:
        v.text = str(mapping[int(v.text)])
    for i, si in enumerate(items):
        if i not in mapping:
            root.remove(si)
    root.set("uniqueCount", str(len(keep)))
    root.set("count", str(len(used)))
    actions.append(f"{len(items) - len(keep)} chaîne(s) de texte orpheline(s) purgée(s)")


def clean(pkg: Package, opts: Options) -> list[str]:
    actions: list[str] = []
    anonymize_properties(pkg, opts, actions)
    anonymize_comment_authors(pkg, opts, actions)
    remove_always_unneeded(pkg, actions)
    if opts.remove_comments:
        remove_comments(pkg, actions)
    if not opts.keep_hidden_sheets:
        remove_hidden_sheets(pkg, opts, actions)
    remove_hidden_drawing_objects(pkg, actions)
    if opts.purge_hidden_cells:
        purge_hidden_cells(pkg, actions)
    clean_workbook_xml(pkg, actions)
    select_first_sheet(pkg, actions)
    compact_shared_strings(pkg, actions)
    garbage = pkg.collect_garbage()
    if garbage:
        actions.append(f"{len(garbage)} fichier(s) interne(s) orphelin(s) retiré(s) de l'archive")
    return actions


# --------------------------------------------------------------------------- #
# Ligne de commande
# --------------------------------------------------------------------------- #


def print_findings(title: str, findings: list[Finding]) -> None:
    print(f"\n== {title} ==")
    if not findings:
        print("  Rien à signaler.")
    for f in findings:
        mark = "!" if f.level == "warning" else "-"
        print(f"  {mark} {f.text}")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Anonymise un fichier Excel, vérifie les données cachées / onglets masqués "
        "et supprime ce qui n'est pas nécessaire.",
    )
    p.add_argument("input", type=Path, help="classeur .xlsx ou .xlsm à traiter")
    p.add_argument("-o", "--output", type=Path, help="fichier de sortie (défaut : <nom>.clean.<ext>)")
    p.add_argument("--in-place", action="store_true", help="écrase le fichier d'origine (déconseillé)")
    p.add_argument("-a", "--author", default="Anonyme", metavar="NOM",
                   help="nom inscrit comme auteur et dernier modificateur du fichier "
                        "(et des commentaires) à la place de l'auteur d'origine (défaut : Anonyme)")
    p.add_argument("--audit-only", action="store_true", help="affiche seulement le rapport, sans rien écrire")
    p.add_argument("--keep-hidden-sheets", action="store_true", help="ne supprime pas les onglets masqués")
    p.add_argument("--remove-comments", action="store_true", help="supprime aussi les commentaires et notes")
    p.add_argument("--purge-hidden-cells", action="store_true",
                   help="vide le contenu des lignes/colonnes masquées (lignes masquées par filtre conservées)")
    p.add_argument("--keep-dates", action="store_true", help="conserve les dates de création/modification")
    p.add_argument("--force", action="store_true",
                   help="remplace par leur valeur les formules qui dépendent d'onglets masqués supprimés")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    args.author = args.author.strip()
    if not args.author:
        print("Erreur : le nom de l'auteur ne peut pas être vide.", file=sys.stderr)
        return 2
    src: Path = args.input
    if not src.is_file():
        print(f"Erreur : fichier introuvable : {src}", file=sys.stderr)
        return 2
    if src.suffix.lower() not in (".xlsx", ".xlsm"):
        print("Erreur : seuls les formats .xlsx et .xlsm sont pris en charge "
              "(enregistrez les .xls anciens au format .xlsx).", file=sys.stderr)
        return 2

    dest = src if args.in_place else (args.output or src.with_suffix(f".clean{src.suffix}"))
    if not args.audit_only and dest.resolve() == src.resolve() and not args.in_place:
        print("Erreur : le fichier de sortie est identique à l'entrée (utilisez --in-place).", file=sys.stderr)
        return 2

    opts = Options(
        author=args.author, keep_hidden_sheets=args.keep_hidden_sheets,
        remove_comments=args.remove_comments, purge_hidden_cells=args.purge_hidden_cells,
        keep_dates=args.keep_dates, force=args.force,
    )
    try:
        pkg = Package.load(src)
        print(f"Fichier analysé : {src}")
        print_findings("Audit avant nettoyage", audit(pkg, opts))
        if args.audit_only:
            return 0
        actions = clean(pkg, opts)
        print_findings("Actions réalisées", [Finding("info", a) for a in actions])
        pkg.save(dest)
    except CleanError as exc:
        print(f"\nErreur : {exc}", file=sys.stderr)
        return 1

    after = audit(Package.load(dest), opts)
    print_findings("Audit après nettoyage", after)
    print(f"\nFichier nettoyé : {dest}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
