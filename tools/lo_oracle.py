#!/usr/bin/env python3
"""Dev-only oracle: compute JSON workbooks with LibreOffice headless.

Pipeline: write each workbook as a minimal .xlsx, let
`soffice --headless --convert-to xlsx` recalculate it, read cached values back.

Usage: python3 tools/lo_oracle.py cases.json > expected.json
  cases.json: {"name": workbook, ...}  ->  {"name": computed workbook, ...}
Not needed by ./test.sh (expected values are stored as fixtures).
"""
import json, os, re, shutil, subprocess, sys, tempfile, zipfile
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def cell_key(a):
    m = re.match(r"([A-Z]+)(\d+)$", a)
    c = 0
    for ch in m.group(1):
        c = c * 26 + ord(ch) - 64
    return int(m.group(2)), c


def write_xlsx(path, wb):
    names = list(wb["sheets"].keys())
    z = zipfile.ZipFile(path, "w")
    z.writestr("[Content_Types].xml",
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(len(names)))
        + "</Types>")
    z.writestr("_rels/.rels",
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    z.writestr("xl/workbook.xml",
        '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
        + "".join(f'<sheet name="{escape(n, {chr(34): "&quot;"})}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i, n in enumerate(names))
        + "</sheets></workbook>")
    z.writestr("xl/_rels/workbook.xml.rels",
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>' for i in range(len(names)))
        + "</Relationships>")
    for i, n in enumerate(names):
        rows = {}
        for a, v in wb["sheets"][n].items():
            r, c = cell_key(a)
            if isinstance(v, bool):
                x = f'<c r="{a}" t="b"><v>{int(v)}</v></c>'
            elif isinstance(v, (int, float)):
                x = f'<c r="{a}"><v>{repr(v)}</v></c>'
            elif isinstance(v, str) and v.startswith("="):
                x = f'<c r="{a}"><f>{escape(v[1:])}</f></c>'
            elif v is None:
                continue
            else:
                x = f'<c r="{a}" t="inlineStr"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
            rows.setdefault(r, []).append((c, x))
        body = "".join(f'<row r="{r}">' + "".join(x for _, x in sorted(cs)) + "</row>" for r, cs in sorted(rows.items()))
        z.writestr(f"xl/worksheets/sheet{i+1}.xml",
            '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<sheetData>{body}</sheetData></worksheet>")
    z.close()


def num(s):
    f = float(s)
    return int(f) if f.is_integer() and abs(f) < 2**53 else f


def read_xlsx(path, wb):
    z = zipfile.ZipFile(path)
    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            sst.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    out = {}
    for i, n in enumerate(wb["sheets"].keys()):
        root = ET.fromstring(z.read(f"xl/worksheets/sheet{i+1}.xml"))
        cells = {}
        for c in root.iter("{%s}c" % NS["m"]):
            t = c.get("t", "n")
            v = c.find("m:v", NS)
            isel = c.find("m:is", NS)
            if v is None and isel is None:
                continue
            txt = v.text if v is not None else "".join(x.text or "" for x in isel.iter("{%s}t" % NS["m"]))
            if t == "e":
                val = {"error": txt}
            elif t == "b":
                val = txt == "1"
            elif t == "s":
                val = sst[int(txt)]
            elif t in ("str", "inlineStr"):
                val = txt or ""
            else:
                val = num(txt)
            cells[c.get("r")] = val
        res = {}
        for a, v in wb["sheets"][n].items():
            if isinstance(v, str) and v.startswith("="):
                res[a] = cells.get(a, "")
            else:
                res[a] = v
        out[n] = res
    return {"sheets": out}


def main():
    cases = json.load(open(sys.argv[1]))
    tmp = tempfile.mkdtemp(prefix="lo_oracle_")
    try:
        names = list(cases.keys())
        files = []
        for k, name in enumerate(names):
            p = os.path.join(tmp, f"c{k}.xlsx")
            write_xlsx(p, cases[name])
            files.append(p)
        outdir = os.path.join(tmp, "out")
        for j in range(0, len(files), 200):
            subprocess.run(["soffice", "--headless", "--norestore",
                            "-env:UserInstallation=file:///tmp/lo_oracle_profile",
                            "--convert-to", "xlsx", "--outdir", outdir] + files[j:j + 200],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        out = {}
        for k, name in enumerate(names):
            out[name] = read_xlsx(os.path.join(outdir, f"c{k}.xlsx"), cases[name])
        json.dump(out, sys.stdout, indent=1, ensure_ascii=False)
        print()
    finally:
        shutil.rmtree(tmp)


main()
