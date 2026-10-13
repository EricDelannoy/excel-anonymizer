import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

import excel_anonymizer as ea

SECRET = "SECRET-HIDDEN-DATA"
PERSON = "Jean Dupont"


def all_text(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        return "\n".join(zf.read(n).decode("utf-8", "ignore") for n in zf.namelist())


def build(path: Path, *, dependent_formula: bool = False) -> Path:
    wb = Workbook()
    wb.properties.creator = PERSON
    wb.properties.lastModifiedBy = PERSON
    wb.properties.description = "note interne"
    ws = wb.active
    ws.title = "Public"
    ws.append(["Produit", "Prix"])
    for i in range(1, 6):
        ws.append([f"p{i}", i * 10])
    ws["D1"] = "visible"
    ws["A8"].comment = Comment("à vérifier", PERSON)

    hidden = wb.create_sheet("Cache")
    hidden["A1"] = SECRET
    hidden["B1"] = 42
    hidden.sheet_state = "hidden"
    very = wb.create_sheet("TresCache")
    very["A1"] = SECRET + "-2"
    very.sheet_state = "veryHidden"
    last = wb.create_sheet("Fin")
    last["A1"] = "dernier"

    wb.defined_names["zone"] = DefinedName("zone", attr_text="Public!$A$1:$B$6")
    last.defined_names["local"] = DefinedName("local", attr_text="Fin!$A$1")
    if dependent_formula:
        ws["F1"] = "=Cache!B1*2"
    wb.save(path)
    return path


@pytest.fixture
def src(tmp_path):
    return build(tmp_path / "in.xlsx")


def run(src, *args):
    out = src.with_name("out.xlsx")
    code = ea.main([str(src), "-o", str(out), *args])
    return code, out


def test_author_is_anonymized(src):
    assert PERSON in all_text(src)
    code, out = run(src)
    assert code == 0
    text = all_text(out)
    assert PERSON not in text
    wb = load_workbook(out)
    assert wb.properties.creator == "Anonyme"
    assert wb.properties.lastModifiedBy == "Anonyme"
    assert wb.properties.description is None


def test_custom_author(src):
    code, out = run(src, "--author", "Équipe")
    assert code == 0
    assert load_workbook(out).properties.creator == "Équipe"


def test_hidden_sheets_removed_and_data_purged(src):
    assert SECRET in all_text(src)
    code, out = run(src)
    assert code == 0
    assert SECRET not in all_text(out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["Public", "Fin"]
    assert wb["Public"]["B2"].value == 10
    assert wb["Fin"]["A1"].value == "dernier"
    assert wb["Fin"].defined_names["local"].attr_text == "Fin!$A$1"
    assert wb.defined_names["zone"].attr_text == "Public!$A$1:$B$6"


def test_keep_hidden_sheets(src):
    code, out = run(src, "--keep-hidden-sheets")
    assert code == 0
    wb = load_workbook(out)
    assert wb.sheetnames == ["Public", "Cache", "TresCache", "Fin"]
    assert wb["Cache"].sheet_state == "hidden"


def test_dependent_formula_blocks_then_force(tmp_path):
    src = build(tmp_path / "dep.xlsx", dependent_formula=True)
    code, out = run(src)
    assert code == 1 and not out.exists()
    # Le fichier doit avoir une valeur en cache : on passe par LibreOffice si dispo,
    # sinon on vérifie seulement que la formule a disparu.
    code, out = run(src, "--force")
    assert code == 0
    ws = load_workbook(out)["Public"]
    assert not str(ws["F1"].value or "").startswith("=")
    assert "Cache!" not in all_text(out)


def test_audit_only_writes_nothing(src, capsys):
    code, out = run(src, "--audit-only")
    assert code == 0 and not out.exists()
    printed = capsys.readouterr().out
    assert "Cache" in printed and "TresCache" in printed and PERSON in printed


def test_comments_authors_anonymized_and_removable(src):
    code, out = run(src)
    assert code == 0
    assert load_workbook(out)["Public"]["A8"].comment.author == "Anonyme"
    code, out = run(src, "--remove-comments")
    assert code == 0
    assert load_workbook(out)["Public"]["A8"].comment is None
    assert "comments" not in " ".join(zipfile.ZipFile(out).namelist())


def test_purge_hidden_rows_and_cols(tmp_path):
    p = tmp_path / "hid.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["B2"], ws["C3"] = "ligne1", "LIGNE_MASQUEE", "COL_MASQUEE", "ok"
    ws.row_dimensions[2].hidden = True
    ws.column_dimensions["B"].hidden = True
    wb.save(p)
    code, out = run(p)
    assert code == 0
    assert "LIGNE_MASQUEE" in all_text(out)  # signalé mais conservé par défaut
    code, out = run(p, "--purge-hidden-cells")
    assert code == 0
    text = all_text(out)
    assert "LIGNE_MASQUEE" not in text and "COL_MASQUEE" not in text
    ws = load_workbook(out).active
    assert ws["A1"].value == "ligne1" and ws["C3"].value == "ok"


def test_hidden_chart_object_removed(tmp_path):
    p = tmp_path / "chart.xlsx"
    wb = Workbook()
    ws = wb.active
    for r in ([1, 2], [3, 4]):
        ws.append(r)
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=1, min_row=1, max_row=2))
    ws.add_chart(chart, "E2")
    wb.save(p)
    patched = tmp_path / "chart_hidden.xlsx"
    with zipfile.ZipFile(p) as zin, zipfile.ZipFile(patched, "w") as zout:
        for n in zin.namelist():
            data = zin.read(n)
            if n.startswith("xl/drawings/drawing"):
                data = data.replace(b"<cNvPr ", b'<cNvPr hidden="1" ', 1)
                assert b'hidden="1"' in data
            zout.writestr(n, data)
    code, out = run(patched)
    assert code == 0
    assert not any(n.startswith("xl/charts/") for n in zipfile.ZipFile(out).namelist())
    load_workbook(out)


def test_orphan_shared_strings_purged(src, tmp_path):
    import shutil

    soffice = shutil.which("soffice")
    if not soffice:
        pytest.skip("LibreOffice absent")
    # LibreOffice réécrit le classeur avec un vrai sharedStrings.xml, comme Excel.
    outdir = tmp_path / "lo"
    subprocess.run(
        [soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
        capture_output=True, timeout=120, check=True,
    )
    shared = outdir / "in.xlsx"
    with zipfile.ZipFile(shared) as zf:
        assert SECRET in zf.read("xl/sharedStrings.xml").decode()
    code, out = run(shared)
    assert code == 0
    with zipfile.ZipFile(out) as zf:
        sst = zf.read("xl/sharedStrings.xml").decode()
    assert SECRET not in sst and "Produit" in sst
    ws = load_workbook(out)["Public"]
    assert ws["A1"].value == "Produit" and ws["A6"].value == "p5"


def test_not_an_excel_file(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"pas un zip")
    assert ea.main([str(bad)]) == 1
    assert ea.main([str(tmp_path / "absent.xlsx")]) == 2


def test_refuses_to_overwrite_input(src):
    assert ea.main([str(src), "-o", str(src)]) == 2


def test_libreoffice_can_open_result(src, tmp_path):
    import shutil

    soffice = shutil.which("soffice")
    if not soffice:
        pytest.skip("LibreOffice absent")
    code, out = run(src)
    conv = tmp_path / "conv"
    res = subprocess.run(
        [soffice, "--headless", "--convert-to", "csv", "--outdir", str(conv), str(out)],
        capture_output=True, timeout=120,
    )
    assert res.returncode == 0
    assert "Produit" in next(conv.glob("*.csv")).read_text()
