"""LibreOffice headless recalculation via UNO (used by migrate_v19.py and the tests).

    with LO() as lo:
        doc = lo.open(path)          # loads hidden, then forces a full recalc (calculateAll)
        vals = lo.read(doc, "Old P110", "A1:R1004")   # tuple of tuples: float / str ("" for blank)
        errs = lo.error_cells(doc)   # ["Sheet!A1", ...] cells whose formula result is an error
        arrays = read_all(doc)       # every sheet's computed values
        inject_cached_values(src_xlsx, out_xlsx, arrays)   # write them as the cached values
        doc.close(True)
"""
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from xml.sax.saxutils import escape as xml_escape, unescape as xml_unescape

import uno
from com.sun.star.beans import PropertyValue


def _pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


class LO:
    def __init__(self):
        self.profile = tempfile.mkdtemp(prefix="lo_profile_")
        self.pipe = "p110_%d" % os.getpid()
        self.proc = None
        self.desktop = None

    def __enter__(self):
        self.proc = subprocess.Popen(
            ["soffice", "--headless", "--invisible", "--norestore", "--nologo", "--nodefault",
             "-env:UserInstallation=file://" + self.profile,
             "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % self.pipe],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(120):
            try:
                ctx = resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % self.pipe)
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("LibreOffice did not start")
        self.desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        return self

    def __exit__(self, *a):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=30)
        except Exception:
            self.proc.kill()
        shutil.rmtree(self.profile, ignore_errors=True)

    def open(self, path, recalc=True):
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        doc = self.desktop.loadComponentFromURL(url, "_blank", 0, (_pv("Hidden", True),))
        if doc is None:
            raise RuntimeError("could not open " + path)
        if recalc:
            t = time.time()
            doc.calculateAll()
            self.last_recalc_s = time.time() - t
        return doc

    @staticmethod
    def read(doc, sheet, rng):
        return doc.Sheets.getByName(sheet).getCellRangeByName(rng).getDataArray()

    @staticmethod
    def formulas(doc, sheet, rng):
        return doc.Sheets.getByName(sheet).getCellRangeByName(rng).getFormulaArray()

    @staticmethod
    def error_cells(doc):
        out = []
        for i in range(doc.Sheets.Count):
            sh = doc.Sheets.getByIndex(i)
            ranges = sh.queryFormulaCells(4)          # com.sun.star.sheet.FormulaResult.ERROR
            for addr in ranges.getRangeAddresses():
                out.append("%s!R%dC%d:R%dC%d" % (sh.Name, addr.StartRow + 1, addr.StartColumn + 1,
                                                   addr.EndRow + 1, addr.EndColumn + 1))
        return out

    @staticmethod
    def save_xlsx(doc, path):
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(path)),
                       (_pv("FilterName", "Calc Office Open XML"),))


def read_all(doc):
    """{sheet name: data array from A1 to the end of the used area} for every sheet."""
    out = {}
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        a = cur.RangeAddress
        out[sh.Name] = sh.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).getDataArray()
    return out


# LibreOffice 24.2's xlsx export drops cell styles on later sheets once a workbook holds ~400k+ formula
# cells (observed: 850 rows/sheet OK, 950 rows lost Bucket Status/Settings/Dashboard styling). So the
# deliverable is the openpyxl-generated file (exact styles, validation, protection) and the cached
# results LibreOffice computed in a real recalc are written into its <v> elements.
_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>', re.S)


def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def inject_cached_values(src, dst, arrays):
    """Copy src -> dst, filling every formula cell's cached value from arrays (see read_all)."""
    zin = zipfile.ZipFile(src)
    wbxml = zin.read("xl/workbook.xml").decode("utf8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid_target = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid_target.update({b: a for a, b in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    part_sheet = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml):
        target = rid_target[rid].lstrip("/")
        part_sheet[target if target.startswith("xl/") else "xl/" + target] = xml_unescape(name)
    missing = []
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        sheet = part_sheet.get(item.filename)
        if sheet is not None:
            arr = arrays[sheet]

            def sub(m):
                col, row, attrs, f = m.group(1), int(m.group(2)), m.group(3), m.group(4)
                try:
                    v = arr[row - 1][_col_index(col) - 1]
                except IndexError:
                    missing.append("%s!%s%d" % (sheet, col, row))
                    return m.group(0)
                attrs = re.sub(r'\s*t="[^"]*"', "", attrs)
                if isinstance(v, str):
                    return '<c r="%s%d"%s t="str"><f>%s</f><v>%s</v></c>' % (col, row, attrs, f, xml_escape(v))
                return '<c r="%s%d"%s><f>%s</f><v>%s</v></c>' % (col, row, attrs, f, repr(float(v)))
            data = _CELL.sub(sub, data.decode("utf8")).encode("utf8")
        zout.writestr(item, data)
    zout.close()
    if missing:
        raise RuntimeError("no recalculated value for %d formula cells, e.g. %s" % (len(missing), missing[:5]))
