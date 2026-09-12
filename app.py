from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse
from typing import List
import io, re, zipfile, json, os, math, uuid, threading, time, urllib.request, urllib.error
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed
from datetime import datetime
from copy import copy
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

app = FastAPI(title="NIRF Data Automation Portal", version="40.0")

_JOBS = {}
_JOBS_LOCK = threading.Lock()
_PARALLEL_WORKERS = max(1, min(4, int(os.getenv("NIRF_WORKERS", "2"))))
_USE_PROCESS_POOL = os.getenv("NIRF_PROCESS_POOL", "1").strip().lower() not in ("0", "false", "no")

YEAR3 = ["2024-25", "2023-24", "2022-23"]

# Reuse one immutable border object for all output cells (faster than recreating it per cell).
_THIN = Side(style="thin")
_ALL_BORDERS = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

# ---- Generic helpers ----
def n(s):
    if s is None: return None
    s = str(s).replace(",", "").replace("₹","").strip()
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m: return None
    x=float(m.group())
    return int(x) if x.is_integer() else x

def avg3(v):
    v=[n(x) for x in v]
    return sum(v)/3 if len(v)==3 and all(x is not None for x in v) else None

def nums(line):
    return [n(x) for x in re.findall(r"\d[\d,]*(?:\.\d+)?", line or "")]

_RAPID_OCR_ENGINE = None

def _ocr_image_text(png_bytes):
    """OCR one rendered PDF page.

    RapidOCR is preferred because it does not require a separate Windows
    Tesseract installation.  pytesseract is kept as a fallback when available.
    """
    global _RAPID_OCR_ENGINE

    # Pure Python / ONNX fallback used by the Windows package.
    try:
        import numpy as np
        from PIL import Image
        from rapidocr_onnxruntime import RapidOCR
        if _RAPID_OCR_ENGINE is None:
            _RAPID_OCR_ENGINE = RapidOCR()
        image=np.array(Image.open(io.BytesIO(png_bytes)).convert("RGB"))
        result,_=_RAPID_OCR_ENGINE(image)
        if result:
            items=[]
            for item in result:
                if not item or len(item)<2:
                    continue
                box,txt=item[0],str(item[1] or "").strip()
                if not txt:
                    continue
                xs=[pt[0] for pt in box]; ys=[pt[1] for pt in box]
                items.append({"x":min(xs),"y":sum(ys)/len(ys),"h":max(ys)-min(ys),"text":txt})
            if items:
                items.sort(key=lambda z:(z["y"],z["x"]))
                heights=sorted(max(1,z["h"]) for z in items)
                med=heights[len(heights)//2] if heights else 12
                lines=[]
                for it in items:
                    if not lines or abs(it["y"]-lines[-1]["y"]) > max(8,med*0.65):
                        lines.append({"y":it["y"],"items":[it]})
                    else:
                        lines[-1]["items"].append(it)
                out=[]
                for line in lines:
                    line["items"].sort(key=lambda z:z["x"])
                    out.append(" ".join(z["text"] for z in line["items"]))
                return "\n".join(out)
    except Exception:
        pass

    # Secondary fallback if Tesseract already exists on the machine.
    try:
        from PIL import Image
        import pytesseract
        image=Image.open(io.BytesIO(png_bytes)).convert("RGB")
        return pytesseract.image_to_string(image,config="--psm 3") or ""
    except Exception:
        return ""

def _ocr_pdf_pages(pdf_bytes, target_pages):
    """OCR only pages that have no usable embedded text."""
    if not target_pages:
        return {}
    try:
        import fitz
    except Exception:
        return {}
    out={}
    doc=fitz.open(stream=pdf_bytes,filetype="pdf")
    try:
        for page_no in target_pages:
            page=doc[page_no-1]
            # 3x rendering greatly improves NIRF table OCR.
            pix=page.get_pixmap(matrix=fitz.Matrix(3,3),alpha=False)
            txt=_ocr_image_text(pix.tobytes("png"))
            if txt and len(re.sub(r"\s+","",txt))>=20:
                out[page_no]=txt
    finally:
        doc.close()
    return out

def extract_pdf(pdf_bytes, category="Engineering"):
    # v35 adaptive fast mode: every category starts with the fastest plain-text
    # extraction. Slow table extraction is used only as a targeted fallback when
    # Engineering/Pharmacy text does not contain a complete program row or IPR
    # metric. This preserves reliability while avoiding unnecessary table work.
    pages=[]
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i,p in enumerate(pdf.pages,1):
            pages.append({"page":i,"text":p.extract_text() or "","tables":[],"ocr":False})

        cat=str(category or "").strip().lower()
        table_pages=set()
        if cat in ("engineering","pharmacy"):
            full_text="\n".join(x.get("text") or "" for x in pages)

            # Check whether the Sanctioned Intake and Actual Strength rows are
            # already complete in text. Most NIRF PDFs are, so no table parsing
            # is needed. Wrapped/broken rows trigger a first-page fallback.
            san_block=_section(full_text,"Sanctioned (Approved) Intake",("Total Actual Student Strength",))
            act_block=_section(full_text,"Total Actual Student Strength",("Placement & Higher Studies",))
            san_keys=set(); act_keys=set()
            for line in san_block.splitlines():
                key,label,tail=_program_line_parts(line)
                if key and nums(tail): san_keys.add(key)
            for line in act_block.splitlines():
                key,label,tail=_program_line_parts(line)
                if key and len(nums(tail))>=3: act_keys.add(key)
            if san_keys and (not act_keys or not san_keys.issubset(act_keys)):
                # Program tables are normally on page 1; locate marker page if not.
                for idx,info in enumerate(pages):
                    low=(info.get("text") or "").lower()
                    if "sanctioned intake" in low or "total actual student strength" in low:
                        table_pages.add(idx); break

            # IPR is also text-first. Only invoke table parsing when a patent
            # section exists but both published/granted numeric rows are not
            # readable from the extracted text.
            lowall=full_text.lower()
            if "patent" in lowall or "intellectual property right" in lowall:
                pub_ok=bool(re.search(r"(?i)patents?\s+published[^\n]*\d",full_text))
                gra_ok=bool(re.search(r"(?i)patents?\s+granted[^\n]*\d",full_text))
                if not (pub_ok and gra_ok):
                    for idx,info in enumerate(pages):
                        low=(info.get("text") or "").lower()
                        if "patent" in low or "intellectual property right" in low:
                            table_pages.add(idx)

        for idx in sorted(table_pages):
            try:
                pages[idx]["tables"]=pdf.pages[idx].extract_tables() or []
            except Exception:
                pages[idx]["tables"]=[]
    return pages

def text_of(pages): return "\n".join(p["text"] for p in pages)

def source_find(pages, label, year_window=False):
    """Find label and nearby numeric values, returning source page + values."""
    for p in pages:
        lines=p["text"].splitlines()
        for i,line in enumerate(lines):
            if label.lower() in line.lower():
                block=" ".join(lines[i:i+4])
                vals=nums(block)
                if year_window:
                    # Keep plausible values after the label, stripping obvious years.
                    vals=[x for x in vals if not (isinstance(x,int) and 1900 <= x <= 2100)]
                return {"label":label,"page":p["page"],"raw":block,"values":vals}
    return None

def institution_name(text):
    m=re.search(r"Institute Name:\s*(.*?)\s*\[", text, re.I)
    return m.group(1).strip() if m else None

# ---- NIRF field extraction ----
def _program_key(label):
    """Normalize every NIRF UG/PG program label to a template key.

    PG Integrated is treated as a normal program (like PG1, PG2, etc.) and
    accepts common PDF variants such as "PG Integrated", "PG-Integrated",
    "PG [Integrated Program(s)]" and "Integrated PG".
    """
    compact=re.sub(r"[^a-z0-9]","",str(label or "").lower())
    if ("pg" in compact and "integrat" in compact) or "integratedpg" in compact:
        return "pgintegrated"
    m=re.search(r"(ug|pg)([1-9])",compact)
    return f"{m.group(1)}{m.group(2)}" if m else None

def _program_line_parts(line):
    """Return (program_key, label, trailing-data-text) for a UG/PG data row.

    Supports square brackets, round brackets, plain "PG 2 Years" text and
    PG Integrated.  The program-duration number is kept out of the data tail.
    """
    line=str(line or "").strip()
    patterns=[
        # PDF text can wrap the label as: "UG [4 Years 3941 ..." then "Program(s)]"
        # on the next line.  In that case the values already start after "Years".
        r"^\s*((?:UG|PG)\s*\[\s*[1-9]\s*Years?)\s+(.*)$",
        # UG/PG [4 Years Program(s)]  ...data...
        r"^\s*((?:UG|PG)\s*\[\s*[1-9]\s*Years?[^\]]*\])\s*[:=-]?\s+(.*)$",
        # UG/PG (4 Years)  ...data...
        r"^\s*((?:UG|PG)\s*\(\s*[1-9]\s*Years?[^)]*\))\s*[:=-]?\s+(.*)$",
        # PG [Integrated Program(s)] / PG Integrated / PG-Integrated ...data...
        r"^\s*((?:PG)\s*(?:\[\s*Integrated[^\]]*\]|[- ]+Integrated(?:\s*Program\(s\))?))\s*[:=-]?\s+(.*)$",
        # Integrated PG ...data...
        r"^\s*((?:Integrated\s+PG)(?:\s*Program\(s\))?)\s*[:=-]?\s+(.*)$",
        # UG/PG 4 Years ...data...
        r"^\s*((?:UG|PG)\s+[1-9]\s*Years?(?:\s*Program\(s\))?)\s*[:=-]?\s+(.*)$",
    ]
    for pat in patterns:
        m=re.match(pat,line,re.I)
        if m and _program_key(m.group(1)):
            return _program_key(m.group(1)),m.group(1).strip(),m.group(2).strip()
    return None,None,None

def _placement_heading_matches(text):
    """Find placement-section program headings with broad NIRF label support."""
    pat=re.compile(r"(?mi)^\s*((?:UG|PG)[^\n:]{0,100}|Integrated\s+PG[^\n:]{0,80})\s*:\s*Placement\s*&\s*higher\s+studies(?:\s+for\s+previous\s+\d+\s+years?)?\s*$")
    out=[]
    for m in pat.finditer(text):
        label=m.group(1)
        compact=re.sub(r"[^a-z0-9]","",label.lower())
        if _program_key(label) and ("year" in label.lower() or "integrat" in label.lower() or re.fullmatch(r"(?:ug|pg)[1-9]",compact)):
            out.append(m)
    return out

def _sum3(values):
    vals=[n(v) for v in values]
    return sum(vals) if vals and all(v is not None for v in vals) else None

def _median3(values):
    # User rule for the Excel "Median" heading: total the 3 yearly values and divide by 3.
    vals=[n(v) for v in values]
    return sum(vals)/3 if len(vals)==3 and all(v is not None for v in vals) else None

def _extract_year_rows(block):
    """Extract [graduated, placed, median_salary, higher_studies] from each placement year row.

    NIRF placement rows contain several academic-year tokens.  The four required
    metrics always follow the LAST academic-year token in the row.
    """
    rows=[]
    for line in block.splitlines():
        years=list(re.finditer(r"20\d{2}-\d{2}", line))
        if not years:
            continue
        tail=line[years[-1].end():]
        v=nums(tail)
        if len(v)>=4:
            rows.append(v[:4])
    return rows

def _section(text, start_label, end_labels=()):
    low=text.lower()
    start=low.find(start_label.lower())
    if start < 0:
        return ""
    end=len(text)
    for label in end_labels:
        p=low.find(label.lower(), start+len(start_label))
        if p >= 0:
            end=min(end,p)
    return text[start:end]

def _three_after_label(section, label):
    m=re.search(re.escape(label)+r"\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)",section,re.I)
    if not m:
        return None
    return [n(m.group(1)),n(m.group(2)),n(m.group(3))]


def _cell_first_number(cell):
    """Return the first numeric value from a PDF table cell (ignoring year labels)."""
    if cell is None:
        return None
    txt=str(cell).strip()
    # A pure academic-year cell is a label, not a data value.
    if re.fullmatch(r"20\d{2}\s*[-–/]\s*(?:20)?\d{2}", txt):
        return None
    m=re.search(r"(?<!\d)(\d[\d,]*(?:\.\d+)?)", txt)
    return n(m.group(1)) if m else None

def _extract_program_tables(pages, result, put):
    """Recover Sanctioned Intake and Actual Strength from extracted PDF tables.

    This is intentionally table-based because many NIRF PDFs wrap program labels
    (especially PG Integrated) onto separate text lines.  Table cells preserve the
    row relationship much more reliably than plain extracted text.
    """
    # Sanctioned intake tables: first row is Academic Year, following rows are programs.
    for p in pages:
        for table in p.get("tables",[]) or []:
            if not table:
                continue
            first=" ".join(str(x or "") for x in table[0]).lower()
            has_program=any(_program_key((row[0] if row else "")) for row in table[1:])
            if "academic year" in first and has_program and "total students" not in first:
                for row in table[1:]:
                    if not row: continue
                    label=str(row[0] or "").replace("\n"," ").strip()
                    key=_program_key(label)
                    if not key: continue
                    vals=[]
                    for cell in row[1:]:
                        v=_cell_first_number(cell)
                        if v is not None: vals.append(v)
                    # Sanctioned intake may contain '-' cells; every actual numeric intake is summed.
                    if vals:
                        put("sanctioned_"+key,label,vals,p.get("page"),"Sum all academic-year sanctioned intake values shown in PDF")
                        result["values"]["sanctioned_"+key+"_total"]=sum(vals)

    # Actual strength tables: locate the column explicitly labelled Total Students.
    for p in pages:
        for table in p.get("tables",[]) or []:
            if not table:
                continue
            header=[str(x or "").replace("\n"," ").strip() for x in table[0]]
            total_idx=None
            for i,h in enumerate(header):
                if re.search(r"\btotal\s+students\b",h,re.I):
                    total_idx=i; break
            if total_idx is None:
                continue
            for row in table[1:]:
                if not row: continue
                label=str(row[0] or "").replace("\n"," ").strip()
                key=_program_key(label)
                if not key or total_idx>=len(row): continue
                v=_cell_first_number(row[total_idx])
                if v is not None:
                    put("actual_"+key,label,[v],p.get("page"),"Use Total Students from PDF program row")

def _capital_values_from_tables(pages):
    """Collect every capital-expenditure amount from the PDF tables.

    NIRF capital tables can continue onto another page.  Once the Capital
    Expenditure table starts, continuation tables are included until the
    Operational Expenditure table starts.  Each component normally contributes
    three yearly amounts, so formats can naturally yield 15, 12, 9, etc. values.
    """
    values=[]
    in_cap=False
    for p in pages:
        for table in p.get("tables",[]) or []:
            if not table:
                continue
            flat=" ".join(str(c or "") for row in table for c in (row or [])).lower()
            if "annual operational expenditure" in flat or ("operational expenditure" in flat and "annual capital expenditure" not in flat):
                if in_cap:
                    return values
            if "annual capital expenditure" in flat or ("capital expenditure" in flat and "operational expenditure" not in flat):
                in_cap=True
            if not in_cap:
                continue
            for row in table:
                if not row or len(row)<2:
                    continue
                label=str(row[0] or "").replace("\n"," ").strip().lower()
                if (not label or "academic year" in label or "annual capital expenditure" in label
                    or "utilised amount" in label):
                    continue
                rowvals=[]
                for cell in row[1:]:
                    v=_cell_first_number(cell)
                    if v is not None:
                        rowvals.append(v)
                # A genuine capital component row has yearly monetary values. Keep all
                # values present, not a fixed component count.
                if rowvals:
                    values.extend(rowvals)
    return values

def _placement_heading_lines(text):
    """Placement headings, including PG Integrated, with or without a colon."""
    matches=[]
    pat=re.compile(r"(?mi)^\s*((?:UG|PG|Integrated\s+PG)[^\n]{0,140}?)\s*:?[ \t]+Placement\s*&\s*higher\s+studies(?:\s+for\s+previous\s+\d+\s+years?)?\s*$")
    for m in pat.finditer(text):
        if _program_key(m.group(1)):
            matches.append(m)
    return matches

def _ipr_values_from_tables(pages):
    """Extract patent counts from IPR tables in any NIRF format.

    Supports both row-oriented tables (metric name followed by yearly values) and
    column-oriented tables (metric in header and one value per year).  If the PDF
    gives one direct total, that one number is kept; if it gives yearly values,
    every available patent count is kept and later summed.
    """
    out={"published":[], "granted":[]}
    for p in pages:
        for table in p.get("tables",[]) or []:
            if not table:
                continue
            flat=" ".join(str(c or "") for row in table for c in (row or [])).lower()
            if "patent" not in flat:
                continue

            # Row-oriented: "No. of Patents Published | 12 | 10 | 8"
            for row in table:
                if not row:
                    continue
                cells=[str(c or "").replace("\n"," ").strip() for c in row]
                rowtxt=" ".join(cells).lower()
                kind=None
                if "patent" in rowtxt and "publish" in rowtxt:
                    kind="published"
                elif "patent" in rowtxt and "grant" in rowtxt:
                    kind="granted"
                if kind:
                    vals=[]
                    for cell in cells:
                        # Do not treat calendar/academic years as patent counts.
                        cleaned=re.sub(r"\b20\d{2}\s*[-–/]\s*(?:20)?\d{2}\b"," ",cell)
                        if re.search(r"patent|publish|grant|calendar|academic|year", cleaned, re.I):
                            # A metric cell may still include a total after ':' or '='.
                            tail=re.split(r"[:=]",cleaned,maxsplit=1)
                            cleaned=tail[-1] if len(tail)>1 else ""
                        vals.extend([n(x) for x in re.findall(r"(?<!\d)\d[\d,]*(?:\.\d+)?",cleaned)])
                    vals=[v for v in vals if v is not None and not (isinstance(v,int) and 1900<=v<=2100)]
                    if vals:
                        out[kind]=vals

            # Column-oriented: header contains patents published/granted, rows contain years.
            header=[str(c or "").replace("\n"," ").strip() for c in table[0]]
            for kind,word in (("published","publish"),("granted","grant")):
                idx=None
                for i,h in enumerate(header):
                    low=h.lower()
                    if "patent" in low and word in low:
                        idx=i; break
                if idx is not None:
                    vals=[]
                    for row in table[1:]:
                        if not row or idx>=len(row):
                            continue
                        cell=str(row[idx] or "").strip()
                        # Only the value in the metric column; ignore year labels elsewhere.
                        v=_cell_first_number(cell)
                        if v is not None and not (isinstance(v,int) and 1900<=v<=2100):
                            vals.append(v)
                    if vals:
                        out[kind]=vals
    return out

def _nirf_year(text):
    """Detect the NIRF year from the PDF heading, especially:
    Data Submitted by Institution for India Rankings '2022'.
    """
    text=str(text or "")
    # pdfplumber can split spaces/newlines and can return straight or curly quotes.
    patterns=(
        r'Data\s+Submitted\s+by\s+Institution\s+for\s+India\s+Rankings?\s*[\"\'’‘]?\s*(20\d{2})',
        r"Data\s+Submitted\s+by\s+Institution[^\n]{0,120}?India\s+Rankings?[^\n]{0,40}?(20\d{2})",
        r'India\s+Rankings?\s*[\"\'’‘]?\s*(20\d{2})',
        r'Submitted\s+Institute\s+Data\s+for\s+NIRF\s*[\"\'’‘]?\s*(20\d{2})',
        r'NIRF\s*[\"\'’‘]?\s*(20\d{2})',
        r"Ranking\s+Framework[^\n]{0,80}?(20\d{2})",
    )
    for pat in patterns:
        m=re.search(pat,text,re.I)
        if m:
            return m.group(1)
    # Last-resort: only inspect the beginning of the report, where the NIRF heading is printed.
    head=text[:5000]
    m=re.search(r"India\s+Rankings?[^0-9]{0,80}(20\d{2})",head,re.I)
    return m.group(1) if m else None

def extract_data(pdf_bytes, category):
    pages=extract_pdf(pdf_bytes, category); text=text_of(pages)
    result={"institution":institution_name(text),"category":category,"nirf_year":_nirf_year(text),"source":{},"values":{},"review":[]}
    if any(p.get("ocr") for p in pages):
        result["review"].append("OCR fallback used for scanned/image-only PDF page(s)")

    def put(key,label,vals,page=None,rule="Direct extraction"):
        result["source"][key]={"label":label,"page":page,"values":vals,"rule":rule}
        result["values"][key]=vals

    # Sanctioned (Approved) Intake: SUM ALL academic-year values present for each program.
    san_block=_section(text,"Sanctioned (Approved) Intake",("Total Actual Student Strength",))
    for line in san_block.splitlines():
        key,label,tail=_program_line_parts(line)
        if not key:
            continue
        vals=nums(tail)
        # Dashes are not values; sum every numeric academic-year intake shown.
        if vals:
            put("sanctioned_"+key,label,vals,None,"Sum all academic-year sanctioned intake values shown in PDF")
            result["values"]["sanctioned_"+key+"_total"]=sum(vals)
    # Table recovery is applied below as well; it fixes wrapped program labels such as PG Integrated.

    # Total Actual Strength: use the PDF's Total Students value, not male+female recomputation.
    actual_block=_section(text,"Total Actual Student Strength",("Placement & Higher Studies",))
    for line in actual_block.splitlines():
        key,label,tail=_program_line_parts(line)
        if not key:
            continue
        vals=nums(tail)
        # The program duration is removed by _program_line_parts; remaining values are Male, Female, Total Students.
        if len(vals)>=3:
            total_students=vals[2]
            put("actual_"+key,label,[total_students],None,"Use Total Students from PDF program row")
    # Recover both sections from PDF tables. Table values overwrite equivalent text
    # values when found, because they are more reliable for wrapped labels.
    _extract_program_tables(pages,result,put)
    san_keys=[k for k in result["values"] if k.startswith("sanctioned_") and k.endswith("_total")]
    result["values"]["sanctioned_total"]=sum(result["values"][k] for k in san_keys if result["values"][k] is not None)
    actual_keys=[k for k in result["values"] if re.fullmatch(r"actual_(?:(?:ug|pg)[1-9]|pgintegrated)",k)]
    result["values"]["actual_total"]=sum(result["values"][k][0] for k in actual_keys if isinstance(result["values"][k],list) and result["values"][k])

    # Placement & Higher Studies: every program, including PG Integrated, is handled identically.
    # Sum Graduated/Placed/Higher Studies; average the three yearly median-salary values.
    placement_matches=_placement_heading_lines(text)
    for idx,m in enumerate(placement_matches):
        label=m.group(1).strip()
        block_start=m.end()
        if idx+1 < len(placement_matches):
            block_end=placement_matches[idx+1].start()
        else:
            stops=[]
            for stop_pat in (r"(?mi)^\s*Ph\.?D\s+Student\s+Details", r"(?mi)^\s*Financial\s+Resources"):
                sm=re.search(stop_pat,text[block_start:])
                if sm: stops.append(block_start+sm.start())
            block_end=min(stops) if stops else len(text)
        block=text[block_start:block_end]
        rows=_extract_year_rows(block)
        keybase=_program_key(label)
        if keybase and len(rows)>=3:
            rows=rows[:3]
            specs=[
                ("graduated",0,"No. of students graduating in minimum stipulated time","Sum of 3 years"),
                ("placed",1,"No. of students placed","Sum of 3 years"),
                ("median",2,"Median salary of placed graduates","(Y1 + Y2 + Y3) / 3"),
                ("higher_studies",3,"No. of students selected for Higher Studies","Sum of 3 years")
            ]
            for suffix,colidx,lab,rule in specs:
                vals=[r[colidx] for r in rows]
                put(f"{keybase}_{suffix}",lab,vals,None,rule)
                if suffix=="median":
                    result["values"][f"{keybase}_median_result"]=_median3(vals)
                else:
                    result["values"][f"{keybase}_{suffix}_total"]=_sum3(vals)
        elif keybase:
            result["review"].append(f"{label}: could not confidently extract all 3 placement years")

    # Ph.D Total Students: direct Full Time / Part Time values.
    phd_block=_section(text,"Ph.D Student Details",("Financial Resources",))
    m=re.search(r"Total Students\s+Full Time\s+([0-9,]+)\s+Part Time\s+([0-9,]+)",phd_block,re.I|re.S)
    if m:
        put("phd_full_time","Ph.D Total Students - Full Time",[n(m.group(1))])
        put("phd_part_time","Ph.D Total Students - Part Time",[n(m.group(2))])
    else:
        result["review"].append("Ph.D Total Students: Full Time / Part Time not confidently extracted")

    # Capital / Operational expenditure.
    # IMPORTANT: for Capital Expenditure, take EVERY numeric expenditure amount shown
    # in the Capital Expenditure section (irrespective of component names), add them,
    # and divide the grand total by 3.  Year labels such as 2024-25 are excluded.
    cap_section=_section(text,"Financial Resources: Utilised Amount for the Capital expenditure",("Financial Resources: Utilised Amount for the Operational expenditure",))
    if not cap_section:
        cap_section=_section(text,"Capital expenditure",("Operational expenditure",))
    op_section=_section(text,"Financial Resources: Utilised Amount for the Operational expenditure",("IPR", "Sponsored Research Details"))
    if not op_section:
        op_section=_section(text,"Operational expenditure",("IPR", "Sponsored Research Details"))

    def expenditure_amounts(section):
        amounts=[]
        for raw_line in (section or "").splitlines():
            line=raw_line
            low0=line.lower()
            # Section/table headers can contain "previous 3 years"; that 3 is not expenditure data.
            if "financial resources:" in low0 or "academic year" in low0 or "utilised amount" in low0:
                continue
            # Remove academic/calendar year tokens before reading amounts.
            line=re.sub(r"\b20\d{2}\s*[-–/]\s*(?:20)?\d{2}\b"," ",line)
            # Remove common table/header wording that may contain a standalone year count.
            low=line.lower()
            if not re.search(r"\d",line):
                continue
            vals=[n(x) for x in re.findall(r"(?<![A-Za-z])\d[\d,]*(?:\.\d+)?",line)]
            vals=[v for v in vals if v is not None]
            if not vals:
                continue
            # Exclude standalone 4-digit calendar years if any survived cleanup.
            vals=[v for v in vals if not (isinstance(v,int) and 1900 <= v <= 2100)]
            # Exclude isolated program/table serial-like numbers from headings. Monetary
            # rows normally contain one or more actual amounts; zeros are legitimate.
            if vals:
                amounts.extend(vals)
        return amounts

    # Prefer table extraction because Capital Expenditure frequently continues on the
    # next PDF page. This captures every component x every available year (e.g. 15,
    # 12 or 9 numeric amounts) and then divides their grand total by 3.
    cap_vals=_capital_values_from_tables(pages)
    if not cap_vals:
        cap_vals=expenditure_amounts(cap_section)
    if cap_vals:
        put("cap_1","All Capital Expenditure amounts",cap_vals,None,"Add every Capital Expenditure amount shown in PDF and divide by 3")
        result["values"]["CAP_EXP"]=sum(cap_vals)/3
        result["source"]["CAP_EXP"]={"label":"Capital Expenditure","page":None,"values":[cap_vals],"rule":"(all numbers in Capital Expenditure section) / 3"}
    else:
        result["review"].append("Capital Expenditure: no numeric expenditure amounts confidently extracted")

    # Operational expenditure keeps the same rule. Prefer the known component rows;
    # if wording differs, fall back to all numeric amounts in the section / 3.
    op_labels=["Salaries (Teaching and Non Teaching staff)","Maintenance of Academic Infrastructure or consumables and","Seminars/Conferences/Workshops"]
    def component_values(section,label,next_labels):
        i=section.lower().find(label.lower())
        if i<0: return None
        tail=section[i+len(label):]
        end=len(tail)
        for nl in next_labels:
            j=tail.lower().find(nl.lower())
            if j>=0: end=min(end,j)
        vals=[n(x) for x in re.findall(r"\b\d[\d,]{3,}\b",tail[:end])]
        return vals[:3] if len(vals)>=3 else None

    op_groups=[]
    for idx,label in enumerate(op_labels):
        vals=component_values(op_section,label,op_labels[idx+1:])
        if vals and len(vals)==3:
            op_groups.append(vals); put(f"op_{idx+1}",label,vals,None,"Source component values")
    if len(op_groups)==len(op_labels):
        result["values"]["OP_EXP"]=sum(sum(v) for v in op_groups)/3
        result["source"]["OP_EXP"]={"label":"OP_EXP","page":None,"values":op_groups,"rule":"(all component values for 3 years) / 3"}
    else:
        # Remove any partial op_* values so the proof formula does not mix extraction modes.
        for k in [k for k in list(result["values"]) if re.fullmatch(r"op_\d+",k)]:
            result["values"].pop(k,None); result["source"].pop(k,None)
        op_vals=expenditure_amounts(op_section)
        if op_vals:
            put("op_1","All Operational Expenditure amounts",op_vals,None,"Add every Operational Expenditure amount shown in PDF and divide by 3")
            result["values"]["OP_EXP"]=sum(op_vals)/3
            result["source"]["OP_EXP"]={"label":"Operational Expenditure","page":None,"values":[op_vals],"rule":"(all numbers in Operational Expenditure section) / 3"}
        else:
            result["review"].append("Operational Expenditure: no numeric expenditure amounts confidently extracted")

    # IPR (Engineering + Pharmacy, and any other template containing IPR columns):
    # use whatever patent-count numbers the PDF provides. A single direct total is
    # entered directly; multiple yearly values are added to produce the Excel total.
    ipr_section=_section(text,"IPR",("Sponsored Research Details",))
    def first_three_after(section, anchors, stop_anchors=()):
        for anchor in anchors:
            i=section.lower().find(anchor.lower())
            if i < 0:
                continue
            tail=section[i+len(anchor):]
            end=len(tail)
            for stop in stop_anchors:
                j=tail.lower().find(stop.lower())
                if j >= 0:
                    end=min(end,j)
            # The anchor is already the metric row, so the first three numbers after it are the three yearly values.
            # Do not discard values such as 2000, which may be legitimate earnings/count data.
            vals=[n(x) for x in re.findall(r"\b\d[\d,]*(?:\.\d+)?\b",tail[:end])]
            if len(vals)>=3:
                return vals[:3]
        return None

    ipr_table=_ipr_values_from_tables(pages) if category.lower() in ("pharmacy","engineering") else {"published": [], "granted": []}
    published=(ipr_table.get("published") or first_three_after(ipr_section,["No. of Patents Published","No of Patents Published"],["No. of Patents Granted","No of Patents Granted","Sponsored Research Details"])) if category.lower() in ("pharmacy","engineering") else None
    granted=(ipr_table.get("granted") or first_three_after(ipr_section,["No. of Patents Granted","No of Patents Granted"],["Sponsored Research Details"])) if category.lower() in ("pharmacy","engineering") else None
    if published:
        rule="Direct patent total from PDF" if len(published)==1 else "Add all patent-published counts shown in PDF"
        put("patents_published","No. of Patents Published",published,None,rule)
        result["values"]["patents_published_total"]=sum(published)
    elif category.lower() == "pharmacy":
        result["review"].append("IPR - No. of Patents Published: not confidently extracted")
    if granted:
        rule="Direct patent total from PDF" if len(granted)==1 else "Add all patent-granted counts shown in PDF"
        put("patents_granted","No. of Patents Granted",granted,None,rule)
        result["values"]["patents_granted_total"]=sum(granted)
    elif category.lower() == "pharmacy":
        result["review"].append("IPR - No. of Patents Granted: not confidently extracted")

    # Executive Development Program (EDP/MDP) -- Management / Overall.
    # NIRF PDFs use several headings/wordings for this section.  Prefer table
    # extraction because it is much less sensitive to line wrapping, then fall
    # back to flexible text matching.  Each metric is the SUM of the three
    # financial-year values shown in the PDF.
    def _clean_metric_nums(values):
        out=[]
        for x in values:
            v=n(x)
            if v is None:
                continue
            # Do NOT discard values merely because they fall between 1900 and 2100.
            # In EDP/MDP, legitimate participant counts can be values such as
            # 1907 or 1944. Financial-year ranges are already removed before
            # token collection, so these values must be preserved.
            out.append(v)
        return out

    def _metric_nums_from_text(s):
        if not s:
            return []
        # Remove academic/financial/calendar year ranges before collecting numbers.
        # Examples: 2024-25, 2023–24, 2022 / 23.
        s=re.sub(r"\b(?:19|20)\d{2}\s*[-–/]\s*\d{2,4}\b"," ",str(s))
        # Also remove isolated year headings such as 2024 2023 2022.
        toks=re.findall(r"\b\d[\d,]*(?:\.\d+)?\b",s)
        return _clean_metric_nums(toks)

    def _edp_metric_key(norm):
        norm=re.sub(r"[^a-z0-9]+"," ",(norm or "").lower()).strip()
        if not norm:
            return None
        # Accept wording variations used in NIRF Management/Overall PDFs.
        if "participant" in norm:
            return "participants"
        if ("annual earning" in norm or "annual revenue" in norm or "earning amount" in norm) and "word" not in norm:
            return "earnings"
        if (("executive development" in norm or "management development" in norm or "edp" in norm or "mdp" in norm)
            and ("program" in norm or "programme" in norm or "edp" in norm or "mdp" in norm)
            and ("participant" not in norm) and ("earning" not in norm) and ("revenue" not in norm)):
            # Avoid treating a section title as the metric unless the row also
            # looks like a number/count row.
            if any(w in norm.split() for w in ("total","no","number","programs","programmes")):
                return "count"
        return None

    def _edp_values_from_tables_local():
        got={"count":[],"participants":[],"earnings":[]}
        for pg in pages:
            for table in (pg.get("tables") or []):
                rows=list(table or [])
                for ri,row in enumerate(rows):
                    cells=[str(c or "").strip() for c in (row or [])]
                    joined=" ".join(cells)
                    key=_edp_metric_key(joined)
                    if not key:
                        continue

                    # Find the label cell when possible and read values to its right.
                    label_idx=0
                    for ci,c in enumerate(cells):
                        if _edp_metric_key(c)==key:
                            label_idx=ci
                            break
                    vals=[]
                    for c in cells[label_idx+1:]:
                        vals.extend(_metric_nums_from_text(c))

                    # Some NIRF PDFs put the metric label in one row and the 3
                    # yearly values in the next row(s), or split a wrapped label
                    # across two rows. Continue until the next EDP metric begins.
                    if len(vals)<3:
                        for rj in range(ri+1,min(len(rows),ri+4)):
                            nxt=[str(c or "").strip() for c in (rows[rj] or [])]
                            nxt_join=" ".join(nxt)
                            nxt_key=_edp_metric_key(nxt_join)
                            if nxt_key and nxt_key!=key:
                                break
                            vals.extend(_metric_nums_from_text(nxt_join))
                            if len(vals)>=3:
                                break
                    if len(vals)>=3 and not got[key]:
                        got[key]=vals[:3]
        return got

    def _three_after_flexible(section, patterns, stop_patterns=()):
        if not section:
            return None
        for pat in patterns:
            m=re.search(pat,section,re.I|re.S)
            if not m:
                continue
            tail=section[m.end():]
            end=len(tail)
            for sp in stop_patterns:
                sm=re.search(sp,tail,re.I|re.S)
                if sm:
                    end=min(end,sm.start())
            vals=_metric_nums_from_text(tail[:end])
            if len(vals)>=3:
                return vals[:3]
        return None

    edp_table=_edp_values_from_tables_local() if category.lower() in ("management","overall") else {"count":[],"participants":[],"earnings":[]}
    # Slice from any common EDP/MDP heading.  This includes the exact wording
    # "Executive Development Program (EDP/MDP)" used by some Management PDFs.
    em=re.search(r"Executive\s+Development\s+Program(?:me)?s?\s*(?:\([^)]*(?:EDP|MDP)[^)]*\)|[/&-]?\s*(?:EDP|MDP)(?:\s*/\s*(?:EDP|MDP))?)?",text,re.I)
    if not em:
        em=re.search(r"(?:EDP\s*/\s*MDP|MDP\s*/\s*EDP)",text,re.I)
    if em:
        tail=text[em.start():]
        stop=len(tail)
        for stop_pat in (r"PCS\s+Facilities",r"Faculty\s+Details",r"Sponsored\s+Research",r"Consultancy\s+Project"):
            sm=re.search(stop_pat,tail,re.I)
            if sm and sm.start()>0:
                stop=min(stop,sm.start())
        edp_section=tail[:stop]
    else:
        edp_section=""

    count_patterns=[
        r"Total\s*(?:no\.?|number)?\s*of\s*(?:Executive\s+Development\s+Program(?:me)?s?(?:\s*/\s*Management\s+Development\s+Program(?:me)?s?)?|EDP\s*/\s*MDP)",
        r"(?:No\.?|Number)\s*of\s*(?:EDP\s*/\s*MDP|Executive\s+Development\s+Program(?:me)?s?)",
    ]
    participant_patterns=[
        r"Total\s*(?:no\.?|number)?\s*(?:of\s*)?Participants?",
        r"(?:No\.?|Number)\s*(?:of\s*)?Participants?",
        r"Participants?\s*(?:Total)?",
    ]
    earning_patterns=[r"Total\s+Annual\s+Earnings?(?:\s*\([^)]*Rupees?[^)]*\))?",r"Annual\s+Earnings?(?:\s*\([^)]*Rupees?[^)]*\))?"]

    edp=(edp_table.get("count") or _three_after_flexible(edp_section,count_patterns,participant_patterns))
    participants=(edp_table.get("participants") or _three_after_flexible(edp_section,participant_patterns,earning_patterns))
    earnings=(edp_table.get("earnings") or _three_after_flexible(edp_section,earning_patterns,[r"Total\s+Annual\s+Earnings?\s+in\s+Words",r"PCS\s+Facilities",r"Faculty\s+Details"]))

    if edp:
        put("edp_count","Total no. of Executive Development Programs/ Management Development Programs",edp,None,"Sum of 3 years")
        result["values"]["edp_count_total"]=sum(edp)
    elif category.lower() in ("management","overall"):
        result["review"].append("EDP/MDP - Total no. of EDP/MDP: not confidently extracted")
    if participants:
        put("edp_participants","Total no. of Participants",participants,None,"Sum of 3 years")
        result["values"]["edp_participants_total"]=sum(participants)
    elif category.lower() in ("management","overall"):
        result["review"].append("EDP/MDP - Total Participants: not confidently extracted")
    if earnings:
        put("edp_earnings","Total Annual Earnings (Amount in Rupees)",earnings,None,"Sum of 3 years")
        result["values"]["edp_earnings_total"]=sum(earnings)
    elif category.lower() in ("management","overall"):
        result["review"].append("EDP/MDP - Total Annual Earnings: not confidently extracted")

    # Sponsored Research: section-specific totals, summed across 3 years.
    sponsored_section=_section(text,"Sponsored Research Details",("Consultancy Project Details",))
    for key,label,total_key in [
        ("sponsored","Total no. of Sponsored Projects","sponsored_total"),
        ("sponsored_agencies","Total no. of Funding Agencies","sponsored_agencies_total"),
        ("sponsored_amount","Total Amount Received (Amount in Rupees)","sponsored_amount_total")]:
        vals=_three_after_label(sponsored_section,label)
        if vals:
            put(key,label,vals,None,"Sum of 3 years")
            result["values"][total_key]=sum(vals)
        else:
            result["review"].append(f"Sponsored Research - {label}: not confidently extracted")

    # Consultancy: independently slice consultancy section so its 'Total Amount Received' cannot be confused with sponsored research.
    consultancy_section=_section(text,"Consultancy Project Details",("PCS Facilities", "Faculty Details"))
    for key,label,total_key in [
        ("consultancy","Total no. of Consultancy Projects","consultancy_total"),
        ("consultancy_clients","Total no. of Client Organizations","consultancy_clients_total"),
        ("consultancy_amount","Total Amount Received (Amount in Rupees)","consultancy_amount_total")]:
        vals=_three_after_label(consultancy_section,label)
        if vals:
            put(key,label,vals,None,"Sum of 3 years")
            result["values"][total_key]=sum(vals)
        else:
            result["review"].append(f"Consultancy - {label}: not confidently extracted")

    # Faculty Details: use every reliable method, in priority order.
    # 1) explicit faculty total/count in the Faculty Details section;
    # 2) otherwise highest valid Srno / Sr No / Serial No from the faculty table;
    # 3) otherwise leave blank (never insert 0).
    fpos=text.lower().find("faculty details")
    if fpos>=0:
        faculty_text=text[fpos:]
        faculty_count=None
        faculty_rule=None

        # First preference: an explicit total printed by NIRF.  Different PDFs use
        # slightly different wording, so accept the common variants but only inside
        # the Faculty Details section.
        explicit_patterns=[
            r"no\.?\s*of\s*faculty\s*members\s*entered\s*[:\-]?\s*(\d{1,4})",
            r"number\s*of\s*faculty\s*members\s*entered\s*[:\-]?\s*(\d{1,4})",
            r"total\s*(?:no\.?|number)?\s*of\s*faculty\s*members\s*[:\-]?\s*(\d{1,4})",
            r"total\s*faculty\s*(?:members)?\s*[:\-]?\s*(\d{1,4})",
            r"faculty\s*(?:members)?\s*count\s*[:\-]?\s*(\d{1,4})",
            r"number\s*of\s*faculty\s*(?:members)?\s*[:\-]?\s*(\d{1,4})",
            r"no\.?\s*of\s*faculty\s*(?:members)?\s*[:\-]?\s*(\d{1,4})",
        ]
        for pat in explicit_patterns:
            m=re.search(pat, faculty_text, re.I)
            if m:
                val=int(m.group(1))
                if val>0:
                    faculty_count=val
                    faculty_rule="Explicit faculty count in Faculty Details"
                    break

        # Otherwise read the Srno column from extracted tables.  Once a Faculty
        # Details table starts, continuation pages are included even when the heading
        # is not repeated.  We only inspect the first table column to avoid counting
        # Age, Experience, dates, etc.
        if faculty_count is None:
            srnos=[]
            faculty_table_started=False
            for pinfo in pages:
                page_text=(pinfo.get("text") or "")
                if "faculty details" in page_text.lower():
                    faculty_table_started=True
                if not faculty_table_started:
                    continue
                for table in (pinfo.get("tables") or []):
                    if not table:
                        continue
                    header=" ".join(str(c or "") for c in table[0]).lower()
                    has_srno_header=bool(re.search(r"\bsr\s*\.?\s*no\.?\b|\bsrno\b|serial\s*(?:no|number)", header, re.I))
                    wide_table=max((len(r) for r in table if r), default=0) >= 8
                    if not has_srno_header and not wide_table:
                        continue
                    for row in table:
                        if not row:
                            continue
                        first=str(row[0] or "").strip()
                        if re.fullmatch(r"\d{1,4}", first):
                            val=int(first)
                            if val>0:
                                srnos.append(val)
            if srnos:
                faculty_count=max(srnos)
                faculty_rule="Highest Srno in Faculty Details"

        # OCR/text fallback: follow a plausible serial sequence beginning at 1 from
        # the Faculty Details section.  This prevents ages and dates from being used
        # as faculty counts.
        if faculty_count is None:
            serial_candidates=[int(x) for x in re.findall(r"(?m)^\s*(\d{1,4})(?=\s+)",faculty_text)]
            if serial_candidates:
                expected=1
                highest=0
                for val in serial_candidates:
                    if val==expected:
                        highest=val
                        expected+=1
                    elif val==highest:
                        # tolerate OCR duplication of the same row number
                        continue
                if highest>0:
                    faculty_count=highest
                    faculty_rule="Srno sequence in Faculty Details (OCR/text fallback)"

        if faculty_count is not None:
            put("faculty_count","Faculty Details",[faculty_count],None,faculty_rule or "Faculty count")

    # Intentionally do not insert 0 when faculty data is absent.
    # A missing/unreliable faculty number must remain blank in the Excel output.
    return result

# ---- Dynamic template handling ----
def _merged_header_value(ws,row,col):
    v=ws.cell(row,col).value
    if v is not None:
        return v
    for rg in ws.merged_cells.ranges:
        if rg.min_row <= row <= rg.max_row and rg.min_col <= col <= rg.max_col:
            return ws.cell(rg.min_row,rg.min_col).value
    return None

def _header_bounds(ws):
    # Find the row containing the main column headings (usually Sl.No./Institution).
    start=1
    for r in range(1,min(ws.max_row,12)+1):
        vals=[str(ws.cell(r,c).value or "").lower() for c in range(1,ws.max_column+1)]
        if any(re.search(r"\b(?:sl|sr|s)\s*\.?\s*no\.?\b|\bslno\b|\bsrno\b|serial\s*(?:no|number)",v,re.I) for v in vals) or any("name of the institution" in v for v in vals):
            start=r; break
    # NIRF templates normally have parent/subgroup/leaf headings over 3 rows.
    return start,min(ws.max_row,start+2)

def get_headers(ws):
    start,end=_header_bounds(ws)
    headers={}

    # Some supplied NIRF templates visually group columns under a main heading but the
    # actual Excel merge stops early (for example Overall: Sanctioned Intake / Placement).
    # Forward-fill ONLY the main group row until the next explicit main heading.
    top_values={}
    last_group=None
    for c in range(1,ws.max_column+1):
        raw=_merged_header_value(ws,start,c)
        if raw is not None and str(raw).strip():
            sv=str(raw).strip()
            if sv.lower()=="total" and last_group and ("sanctioned intake" in last_group.lower() or "total actual strength" in last_group.lower()):
                top_values[c]=last_group+" | Total"
            else:
                top_values[c]=sv
                last_group=sv
        elif last_group:
            top_values[c]=last_group

    for c in range(1,ws.max_column+1):
        parts=[]
        for r in range(start,end+1):
            v=top_values.get(c) if r==start else _merged_header_value(ws,r,c)
            if v is not None and str(v).strip():
                # A contextualized main heading such as 'Sanctioned Intake | Total'
                # already contains two logical header levels.
                subparts=[x.strip() for x in str(v).split("|") if x.strip()]
                for sv in subparts:
                    if not parts or parts[-1].lower()!=sv.lower():
                        parts.append(sv)
        if parts:
            headers[c]=" | ".join(parts)
    return headers

def _program_from_heading(low):
    # Use the same normalization as PDF extraction so PG Integrated behaves exactly
    # like PG1/PG2/etc. in Sanctioned, Actual Strength and Placement sections.
    return _program_key(low)

def map_heading_to_value(h,data,sl_no=1):
    low=h.lower(); vals=data["values"]
    program=_program_from_heading(low)
    key=None; rule=None

    if "name of the institution" in low:
        return data.get("institution"),{},"Direct extraction"
    if re.search(r"\b(?:sl|sr|s)\s*\.?\s*no\.?\b|\bslno\b|\bsrno\b|serial\s*(?:no|number)",low,re.I):
        return sl_no,{},"Sequential number"

    if "sanctioned intake" in low:
        if program: key=f"sanctioned_{program}_total"
        elif "total" in low: key="sanctioned_total"
        rule="Sum all academic-year sanctioned intake values shown in PDF"
    elif "total actual strength" in low:
        if program: key=f"actual_{program}"
        else: key="actual_total"
        rule="Use PDF Total Students value; sum program totals for overall Total"
    elif "placements & higher studies" in low or "placement & higher studies" in low:
        if "median salary" in low: key=f"{program}_median_result" if program else None; rule="(Y1 + Y2 + Y3) / 3"
        elif "graduated" in low: key=f"{program}_graduated_total" if program else None; rule="Sum of 3 years"
        elif re.search(r"(^|\|)\s*placed\s*$",low) or " no. of students placed" in low: key=f"{program}_placed_total" if program else None; rule="Sum of 3 years"
        elif "higher studies" in low: key=f"{program}_higher_studies_total" if program else None; rule="Sum of 3 years"
    elif "ph.d" in low or "phd" in low:
        if "full time" in low: key="phd_full_time"
        elif "part time" in low: key="phd_part_time"
        rule="Direct Total Students value from PDF"
    elif "financial resouces" in low or "financial resources" in low:
        if "capital expenditure" in low: key="CAP_EXP"; rule="(all Capital component values for 3 years) / 3"
        elif "operational expenditure" in low: key="OP_EXP"; rule="(all Operational component values for 3 years) / 3"
    elif "sponsered research" in low or "sponsored research" in low:
        if "total amount received" in low: key="sponsored_amount_total"
        elif "funding agencies" in low: key="sponsored_agencies_total"
        elif "sponsered projects" in low or "sponsored projects" in low: key="sponsored_total"
        rule="Sum of 3 years"
    elif "consultancy project" in low:
        if "total amount received" in low: key="consultancy_amount_total"
        elif "client organizations" in low: key="consultancy_clients_total"
        elif "consultancy projects" in low: key="consultancy_total"
        rule="Sum of 3 years"
    elif "ipr" in low:
        if "patents published" in low: key="patents_published_total"
        elif "patents granted" in low: key="patents_granted_total"
        rule="Use direct total or add all patent counts shown in PDF"
    elif ("executive development program" in low or "management development program" in low or "edp/mdp" in low
          or "total no. of participants" in low or "total no of participants" in low or "total participants" in low
          or "annual earnings" in low):
        if "participant" in low: key="edp_participants_total"
        elif "annual earning" in low: key="edp_earnings_total"
        elif ("no. of edp/mdp" in low or "no of edp/mdp" in low or "executive development program" in low
              or "management development program" in low): key="edp_count_total"
        rule="Sum of 3 years"
    elif "faculty" in low or "staff" in low:
        key="faculty_count"; rule="Count faculty records in PDF"

    if key in vals:
        value=vals[key]
        if isinstance(value,(list,tuple)):
            value=value[0] if len(value)==1 else None
        return value,data["source"].get(key,{}),rule or data["source"].get(key,{}).get("rule","Direct extraction")
    return None,None,None

def _excel_number(x):
    x=n(x)
    if x is None: return None
    return str(int(x)) if isinstance(x,(int,float)) and float(x).is_integer() else str(x)

def _formula_from_source(values, divide_by=None):
    flat=[]
    def add(v):
        if isinstance(v,(list,tuple)):
            for z in v: add(z)
        else:
            z=_excel_number(v)
            if z is not None: flat.append(z)
    add(values)
    if not flat: return None
    expr="+".join(flat)
    return f"=({expr})/{divide_by}" if divide_by else f"={expr}"

def _proof_formula_for_heading(h,data,value):
    """Return an Excel formula that visibly proves the calculation in the cell/formula bar."""
    low=h.lower(); program=_program_from_heading(low); vals=data["values"]
    if "sl.no" in low or "sl. no" in low or "name of the institution" in low:
        return None
    if "sanctioned intake" in low:
        if program: return _formula_from_source(vals.get(f"sanctioned_{program}",[]))
        allv=[]
        for k,v in vals.items():
            if re.fullmatch(r"sanctioned_(?:(?:ug|pg)[1-9]|pgintegrated)",k): allv.extend(v if isinstance(v,list) else [v])
        return _formula_from_source(allv)
    if "total actual strength" in low:
        if program: return _formula_from_source(vals.get(f"actual_{program}",[]))
        allv=[]
        for k,v in vals.items():
            if re.fullmatch(r"actual_(?:(?:ug|pg)[1-9]|pgintegrated)",k): allv.extend(v if isinstance(v,list) else [v])
        return _formula_from_source(allv)
    if ("placements & higher studies" in low or "placement & higher studies" in low) and program:
        if "median salary" in low: return _formula_from_source(vals.get(f"{program}_median",[]),3)
        if "graduated" in low: return _formula_from_source(vals.get(f"{program}_graduated",[]))
        if "placed" in low: return _formula_from_source(vals.get(f"{program}_placed",[]))
        if "higher studies" in low: return _formula_from_source(vals.get(f"{program}_higher_studies",[]))
    if "ph.d" in low or "phd" in low:
        if "full time" in low: return _formula_from_source(vals.get("phd_full_time",[]))
        if "part time" in low: return _formula_from_source(vals.get("phd_part_time",[]))
    if "capital expenditure" in low:
        groups=[vals[k] for k in sorted(vals) if re.fullmatch(r"cap_\d+",k)]
        return _formula_from_source(groups,3)
    if "operational expenditure" in low:
        groups=[vals[k] for k in sorted(vals) if re.fullmatch(r"op_\d+",k)]
        return _formula_from_source(groups,3)
    if "sponsored" in low or "sponsered" in low:
        if "amount" in low: return _formula_from_source(vals.get("sponsored_amount",[]))
        if "agencies" in low: return _formula_from_source(vals.get("sponsored_agencies",[]))
        return _formula_from_source(vals.get("sponsored",[]))
    if "consultancy" in low:
        if "amount" in low: return _formula_from_source(vals.get("consultancy_amount",[]))
        if "client" in low: return _formula_from_source(vals.get("consultancy_clients",[]))
        return _formula_from_source(vals.get("consultancy",[]))
    if "ipr" in low:
        if "published" in low: return _formula_from_source(vals.get("patents_published",[]))
        if "granted" in low: return _formula_from_source(vals.get("patents_granted",[]))
    if ("executive development program" in low or "management development program" in low or "edp/mdp" in low
        or "total no. of participants" in low or "total no of participants" in low or "total participants" in low
        or "annual earnings" in low):
        if "participant" in low: return _formula_from_source(vals.get("edp_participants",[]))
        if "annual earning" in low: return _formula_from_source(vals.get("edp_earnings",[]))
        return _formula_from_source(vals.get("edp_count",[]))
    if "faculty" in low or "staff" in low:
        return _formula_from_source(vals.get("faculty_count",[]))
    if isinstance(value,(int,float)): return _formula_from_source([value])
    return None

def _proof_for_heading(h, data, value, rule):
    """Human-readable proof attached to each populated Excel cell."""
    low=h.lower(); program=_program_from_heading(low); vals=data["values"]
    source_values=None
    if "sanctioned intake" in low and program:
        source_values=vals.get(f"sanctioned_{program}")
    elif ("placement" in low or "higher studies" in low) and program:
        if "median salary" in low: source_values=vals.get(f"{program}_median")
        elif "graduated" in low: source_values=vals.get(f"{program}_graduated")
        elif "placed" in low: source_values=vals.get(f"{program}_placed")
        elif "higher studies" in low: source_values=vals.get(f"{program}_higher_studies")
    elif "sponsored" in low or "sponsered" in low:
        if "amount" in low: source_values=vals.get("sponsored_amount")
        elif "agencies" in low: source_values=vals.get("sponsored_agencies")
        else: source_values=vals.get("sponsored")
    elif "consultancy" in low:
        if "amount" in low: source_values=vals.get("consultancy_amount")
        elif "client" in low: source_values=vals.get("consultancy_clients")
        else: source_values=vals.get("consultancy")
    elif "capital expenditure" in low:
        src=data["source"].get("CAP_EXP",{}); source_values=src.get("values")
    elif "operational expenditure" in low:
        src=data["source"].get("OP_EXP",{}); source_values=src.get("values")

    if source_values is not None:
        flat=[]
        for x in source_values if isinstance(source_values,list) else [source_values]:
            if isinstance(x,list): flat.extend(x)
            else: flat.append(x)
        if flat:
            if rule and ("/ 3" in rule or "divide" in rule.lower()):
                return f"Proof: ({' + '.join(str(x) for x in flat)}) / 3 = {value}"
            if rule and "sum" in rule.lower():
                return f"Proof: {' + '.join(str(x) for x in flat)} = {value}"
            return f"Source values: {', '.join(str(x) for x in flat)}; Result = {value}"
    return f"Rule: {rule or 'Direct extraction'}; Result = {value}"

def _copy_row_style(ws,source_row,target_row):
    if source_row==target_row: return
    ws.row_dimensions[target_row].height=ws.row_dimensions[source_row].height
    for c in range(1,ws.max_column+1):
        a=ws.cell(source_row,c); b=ws.cell(target_row,c)
        if a.has_style: b._style=copy(a._style)
        if a.number_format: b.number_format=a.number_format
        if a.alignment: b.alignment=copy(a.alignment)
        if a.protection: b.protection=copy(a.protection)

def fill_workbook(template_bytes,data,category,data_row=None,wb=None,sl_no=1):
    if wb is None:
        wb=load_workbook(io.BytesIO(template_bytes))
    ws=wb[category] if category in wb.sheetnames else wb.active
    headers=get_headers(ws)
    header_start,header_end=_header_bounds(ws)
    first_data_row=header_end+1
    if data_row is None: data_row=first_data_row
    _copy_row_style(ws,first_data_row,data_row)
    try:
        wb.calculation.fullCalcOnLoad=True
        wb.calculation.forceFullCalc=True
        wb.calculation.calcMode="auto"
    except Exception:
        pass
    audit=[]
    # Program columns belonging to the two sections. Their Total cells must add the
    # values already entered in Excel, rather than rebuilding a separate total.
    sanctioned_program_cols=[c for c,h in headers.items() if "sanctioned intake" in h.lower() and _program_from_heading(h.lower())]
    actual_program_cols=[c for c,h in headers.items() if "total actual strength" in h.lower() and _program_from_heading(h.lower())]

    for c,h in headers.items():
        value,src,rule=map_heading_to_value(h,data,sl_no)
        cell=ws.cell(data_row,c)
        # Every output-format cell gets All Borders, even when the PDF has no value
        # for that section. This keeps the completed Excel table visually complete.
        cell.border=copy(_ALL_BORDERS)
        cell.comment=None
        if value is not None:
            formula=_proof_formula_for_heading(h,data,value)
            low=h.lower(); program=_program_from_heading(low)
            if "sanctioned intake" in low and not program and "total" in low and sanctioned_program_cols:
                refs=[ws.cell(data_row,pc).coordinate for pc in sanctioned_program_cols]
                formula="="+"+".join(refs)
            elif "total actual strength" in low and not program and actual_program_cols:
                refs=[ws.cell(data_row,pc).coordinate for pc in actual_program_cols]
                formula="="+"+".join(refs)
            # Numeric fields contain the calculation itself, so clicking the cell shows proof in Excel's formula bar.
            cell.value=formula if formula and isinstance(value,(int,float)) else value
            # No comments or extra proof text; the Excel formula itself is the proof.
            audit.append([data.get("institution"),cell.coordinate,h,value,formula or "",src.get("page") if src else "","","OK"])
        else:
            # Blank is intentional when the PDF has no confident value; border remains.
            audit.append([data.get("institution"),cell.coordinate,h,"","","","No confident value extracted","REVIEW"])
    return wb,audit

def _write_audit_sheet(wb, audit):
    if "Calculation Proof" in wb.sheetnames: del wb["Calculation Proof"]
    a=wb.create_sheet("Calculation Proof")
    a.append(["Institution","Cell","Excel Heading","Calculated Value","Excel Proof Formula","PDF Page","Calculation Proof / Source","Status"])
    for row in audit: a.append(row)
    a.freeze_panes="A2"
    for cell in a[1]: cell.font=Font(bold=True)
    for col,w in zip("ABCDEFGH",[35,12,70,24,50,12,90,15]): a.column_dimensions[col].width=w

FRONTEND_HTML = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>NIRF Data Automation Portal</title>
<style>
:root{--navy:#0b2f5b;--blue:#1677ff;--ink:#12233f;--muted:#6b7b93;--soft:#f5f9fe;--success:#169b62;--danger:#c2352a}*{box-sizing:border-box}html,body{margin:0;min-height:100%;font-family:Inter,"Segoe UI",Arial,sans-serif;color:var(--ink);background:#f4f7fb}.page{min-height:100vh;display:flex;flex-direction:column}.header{height:72px;background:#fff;border-bottom:1px solid #e7edf5;display:flex;align-items:center;justify-content:space-between;padding:0 clamp(22px,5vw,72px)}.brand{display:flex;align-items:center;gap:12px;font-weight:800;font-size:clamp(20px,2vw,27px);color:var(--navy)}.brand-mark{width:38px;height:38px;border-radius:11px;background:linear-gradient(145deg,var(--blue),#0758bd);display:grid;place-items:center}.brand-mark svg{width:22px;height:22px;stroke:#fff;fill:none;stroke-width:2}.header-note{font-size:13px;color:var(--muted);font-weight:600}.hero{background:linear-gradient(120deg,#092c59 0%,#0c4f9d 62%,#1689e8 100%);padding:20px clamp(22px,5vw,72px) 70px;color:#fff}.hero-inner{max-width:1120px;margin:auto;display:flex;justify-content:space-between;align-items:end;gap:28px}.eyebrow{letter-spacing:2.5px;font-size:10px;font-weight:800;color:#a9d8ff;margin-bottom:7px}h1{font-size:clamp(31px,3.7vw,50px);line-height:1.02;margin:0 0 8px;letter-spacing:-1.1px}.hero p{margin:0;max-width:720px;font-size:clamp(14px,1.2vw,18px);color:#e6f1ff}.hero-badges{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end}.badge{border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.09);padding:8px 11px;border-radius:9px;font-size:11px;font-weight:700;white-space:nowrap}.main{flex:1;width:min(1100px,calc(100% - 36px));margin:-46px auto 14px;position:relative;z-index:3}.card{background:#fff;border:1px solid #e2eaf4;border-radius:18px;box-shadow:0 18px 50px rgba(20,55,96,.13);padding:18px 20px}.mode-tabs{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}.mode-btn{border:1px solid #d4e1ef;background:#fff;color:#385574;border-radius:10px;height:42px;font-weight:800;cursor:pointer}.mode-btn.active{background:#eaf4ff;border-color:#79b8ff;color:#0b5fc2}.grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}.field label{display:block;font-size:12px;font-weight:800;color:#17365f;margin-bottom:6px}select,input[type="number"]{width:100%;height:42px;border:1px solid #cfdcec;border-radius:9px;padding:0 12px;background:#fff;font-size:14px;color:var(--ink)}.auto-box,.upload-section{background:var(--soft);border:1px solid #dfebf7;border-radius:13px;padding:13px}.auto-box{margin-top:12px}.auto-note{font-size:12px;color:var(--muted);margin-top:9px;line-height:1.45}.upload-section{margin-top:12px}.upload-head{display:flex;justify-content:space-between;gap:12px;margin-bottom:8px}.upload-title{font-weight:800;font-size:15px}.upload-limit{font-size:11px;color:var(--muted)}.drop{height:112px;border:1.5px dashed #79b8ff;border-radius:11px;background:#fff;display:flex;align-items:center;justify-content:center;text-align:center;cursor:pointer}.drop.drag{background:#f1f8ff}.upload-icon{width:34px;height:34px;margin:0 auto 4px;border-radius:10px;background:#eaf4ff;display:grid;place-items:center;color:var(--blue);font-size:20px;font-weight:800}.drop strong{font-size:13px}.or{font-size:10px;color:var(--muted);margin:2px 0 6px}.choose{display:inline-flex;background:var(--blue);color:#fff;padding:8px 17px;border-radius:8px;font-size:12px;font-weight:800}#pdfs{display:none}.file-status{text-align:center;font-size:12px;color:var(--muted);margin-top:6px;min-height:16px}.file-status.has{color:var(--success);font-weight:800}.actions{display:flex;align-items:center;justify-content:center;gap:11px;margin-top:13px}#go{height:43px;min-width:290px;border:0;border-radius:10px;background:linear-gradient(135deg,#126bd5,#0a4eaa);color:#fff;font-weight:800;font-size:14px;cursor:pointer}#go:disabled{opacity:.45;cursor:not-allowed}#clearPdfs{height:39px;border:1px solid #d4e0ee;background:#fff;color:#526987;border-radius:9px;padding:0 13px;font-weight:700;cursor:pointer;display:none}.status{text-align:center;font-size:12px;min-height:17px;margin-top:7px;color:var(--muted)}.status.ok{color:var(--success);font-weight:700}.status.err{color:var(--danger);font-weight:700}.progress-wrap{display:none;margin-top:12px}.progress-wrap.show{display:block}.progress-head{display:flex;justify-content:space-between;gap:16px;font-size:12px;font-weight:700;color:#17365d;margin-bottom:6px}.progress-track{height:11px;background:#e5eaf0;border-radius:999px;overflow:hidden}.progress-bar{height:100%;width:0;background:linear-gradient(90deg,#1769aa,#25a7df);transition:width .25s ease}.footer{height:58px;background:#fff;border-top:1px solid #e0e7f0;display:flex;align-items:center;justify-content:space-between;padding:0 clamp(22px,5vw,72px);font-size:11px;color:#708096;gap:20px}.footer strong{color:#294665}.hidden{display:none!important}@media (min-width:900px) and (min-height:650px){html,body{overflow:hidden}.page{height:100vh;min-height:0}.footer{margin-top:auto}}@media(max-width:760px){.header{height:auto;min-height:66px;padding:12px 16px}.header-note{display:none}.hero{padding:25px 18px 66px}.hero-inner{display:block}.hero-badges{justify-content:flex-start;margin-top:15px}.main{width:calc(100% - 22px);margin-top:-42px}.grid{grid-template-columns:1fr}.mode-tabs{grid-template-columns:1fr}.footer{height:auto;min-height:70px;padding:14px 16px;display:block;text-align:center;line-height:1.6}}
</style></head><body><div class="page">
<header class="header"><div class="brand"><div class="brand-mark"><svg viewBox="0 0 24 24"><path d="M4 4h10a3 3 0 0 1 3 3v13H7a3 3 0 0 1-3-3V4Z"/><path d="M8 8h5M8 12h5M8 16h3M17 9h3v11h-3"/></svg></div><span>NIRF Data Automation</span></div><div class="header-note">Official NIRF PDFs → Structured Excel</div></header>
<section class="hero"><div class="hero-inner"><div><div class="eyebrow">NIRF REPORT PROCESSING</div><h1>NIRF Data Automation Portal</h1><p>Generate structured Excel workbooks directly from official NIRF ranking PDFs or from PDFs you upload.</p></div><div class="hero-badges"><div class="badge">Official NIRF Source</div><div class="badge">Up to 100 PDFs</div><div class="badge">Parallel Processing</div></div></div></section>
<main class="main"><section class="card"><div class="mode-tabs"><button id="modeAuto" class="mode-btn active" type="button">Automatic from NIRF Website</button><button id="modeUpload" class="mode-btn" type="button">Upload NIRF PDFs</button></div>
<div id="autoPanel"><div class="auto-box"><div class="grid"><div class="field"><label for="year">NIRF Ranking Year</label><select id="year"><option>2026</option><option>2025</option><option>2024</option><option>2023</option><option>2022</option><option>2021</option><option>2020</option></select></div><div class="field"><label for="autoCat">Category</label><select id="autoCat"><option>Engineering</option><option>Pharmacy</option><option>Management</option><option>Overall</option></select></div><div class="field"><label for="topn">Institutions</label><input id="topn" type="number" min="1" max="100" step="1" value="100" inputmode="numeric" placeholder="e.g. 5, 18, 35, 89"></div></div><div class="auto-note">Enter any number from 1 to 100. The app retrieves that many ranked institutions and official PDFs from the NIRF website, then creates one Excel workbook.</div></div></div>
<div id="uploadPanel" class="hidden"><div class="field"><label for="cat">Category</label><select id="cat"><option>Engineering</option><option>Pharmacy</option><option>Management</option><option>Overall</option></select></div><div class="upload-section"><div class="upload-head"><div class="upload-title">Upload NIRF PDFs</div><div class="upload-limit">Maximum 100 PDF files per batch</div></div><label class="drop" id="pdfdrop"><div><div class="upload-icon">↑</div><strong>Drag & drop PDF files here</strong><div class="or">or</div><span class="choose">Choose PDFs</span><input id="pdfs" type="file" accept=".pdf,application/pdf" multiple></div></label><div id="pn" class="file-status">No PDF files selected</div></div></div>
<div class="actions"><button id="go">Generate Excel</button><button id="clearPdfs" type="button">Clear</button></div><div id="progressWrap" class="progress-wrap"><div class="progress-head"><span id="progressText">Preparing...</span><span id="progressPct">0%</span></div><div class="progress-track"><div id="progressBar" class="progress-bar"></div></div></div><div id="status" class="status"></div></section></main>
<footer class="footer"><div><strong>NIRF Data Automation Portal</strong> · Public PDF-to-Excel utility</div><div>Official NIRF source · Manual upload fallback</div></footer></div>
<script>
const autoBtn=document.getElementById('modeAuto'),uploadBtn=document.getElementById('modeUpload'),autoPanel=document.getElementById('autoPanel'),uploadPanel=document.getElementById('uploadPanel');const inp=document.getElementById('pdfs'),b=document.getElementById('go'),drop=document.getElementById('pdfdrop'),pn=document.getElementById('pn'),clearBtn=document.getElementById('clearPdfs'),st=document.getElementById('status');const pw=document.getElementById('progressWrap'),pt=document.getElementById('progressText'),pp=document.getElementById('progressPct'),pbar=document.getElementById('progressBar');let mode='auto',selectedPDFs=[];
function setMode(m){mode=m;autoBtn.classList.toggle('active',m==='auto');uploadBtn.classList.toggle('active',m==='upload');autoPanel.classList.toggle('hidden',m!=='auto');uploadPanel.classList.toggle('hidden',m!=='upload');clearBtn.style.display=(m==='upload'&&selectedPDFs.length)?'inline-block':'none';b.disabled=(m==='upload'&&!selectedPDFs.length);st.textContent='';pw.classList.remove('show')}autoBtn.onclick=()=>setMode('auto');uploadBtn.onclick=()=>setMode('upload');function keyFor(f){return [f.name,f.size,f.lastModified].join('|')}
function addFiles(list){for(const f of Array.from(list||[])){if((f.type==='application/pdf'||f.name.toLowerCase().endsWith('.pdf'))&&!selectedPDFs.some(e=>keyFor(e)===keyFor(f))){if(selectedPDFs.length>=100){alert('Maximum 100 PDFs per batch.');break}selectedPDFs.push(f)}}renderFiles()}function renderFiles(){const n=selectedPDFs.length;pn.className='file-status'+(n?' has':'');pn.textContent=n?(n+' PDF '+(n===1?'file':'files')+' selected'):'No PDF files selected';if(mode==='upload')b.disabled=!n;clearBtn.style.display=(mode==='upload'&&n)?'inline-block':'none'}
inp.addEventListener('change',e=>{addFiles(e.target.files);inp.value=''});['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')}));['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')}));drop.addEventListener('drop',e=>addFiles(e.dataTransfer.files));clearBtn.onclick=()=>{selectedPDFs=[];renderFiles()};function setProgress(done,total,message){const pct=total?Math.round(done*100/total):0;pw.classList.add('show');pbar.style.width=pct+'%';pp.textContent=pct+'%';pt.textContent=message||('Processing '+done+' of '+total+' PDFs')}async function sleep(ms){return new Promise(r=>setTimeout(r,ms))}
async function runJob(startUrl,fd,totalHint){b.disabled=true;b.textContent='Processing...';st.className='status';st.textContent='';setProgress(0,totalHint||1,'Starting...');try{const start=await fetch(startUrl,{method:'POST',body:fd});if(!start.ok)throw new Error(await start.text());const info=await start.json();const job=info.job_id;let state;while(true){const r=await fetch('/batch-status/'+job,{cache:'no-store'});if(!r.ok)throw new Error(await r.text());state=await r.json();const pct=(state.percent!=null)?state.percent:Math.min(95,Math.round((state.completed||0)*95/(state.total||totalHint||1)));pw.classList.add('show');pbar.style.width=pct+'%';pp.textContent=pct+'%';pt.textContent=state.message||'Processing...';if(state.status==='done')break;if(state.status==='error')throw new Error(state.error||'Processing failed');await sleep(350)}pbar.style.width='99%';pp.textContent='99%';pt.textContent='Preparing download...';const r=await fetch('/batch-download/'+job,{cache:'no-store'});if(!r.ok)throw new Error(await r.text());const blob=await r.blob();const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;const cd=r.headers.get('content-disposition')||'';const fm=cd.match(/filename="?([^";]+)"?/i);a.download=fm?fm[1]:'Nirf data.xlsx';document.body.appendChild(a);pbar.style.width='100%';pp.textContent='100%';pt.textContent='Excel ready — downloading now';a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);st.className='status ok';st.textContent='Excel generated successfully'+(state.failed?(' · '+state.failed+' PDF(s) could not be processed'):'')+'.'}catch(e){st.className='status err';st.textContent='Unable to generate Excel. '+String(e.message||e)}finally{b.textContent='Generate Excel';if(mode==='auto')b.disabled=false;else renderFiles()}}
b.onclick=async()=>{if(mode==='auto'){const n=Number(document.getElementById('topn').value);if(!Number.isInteger(n)||n<1||n>100){st.className='status err';st.textContent='Enter number of institutions from 1 to 100.';return}const fd=new FormData();fd.append('year',document.getElementById('year').value);fd.append('category',document.getElementById('autoCat').value);fd.append('count',String(n));await runJob('/start-auto',fd,n)}else{if(!selectedPDFs.length)return;const fd=new FormData();for(const f of selectedPDFs)fd.append('pdfs',f,f.name);fd.append('category',document.getElementById('cat').value);await runJob('/start-batch',fd,selectedPDFs.length)}};setMode('auto');renderFiles();
</script></body></html>
'''

TEMPLATE_FILES={"engineering":"Nirf Engineering Format.xlsx","pharmacy":"Nirf Pharmacy Format.xlsx","management":"Nirf Management Format.xlsx","overall":"Nirf Overall Format.xlsx"}
def _builtin_template_bytes(category):
    name=TEMPLATE_FILES.get(str(category or "").strip().lower())
    if not name: raise HTTPException(status_code=400,detail="Invalid category.")
    base=os.path.abspath(os.path.dirname(__file__))
    candidates=[os.path.join(base,name),os.path.join(base,"templates",name)]
    path=next((p for p in candidates if os.path.isfile(p)),None)
    if not path: raise HTTPException(status_code=500,detail=f"Built-in template missing: {name}")
    return open(path,"rb").read()

def _ensure_engineering_ipr(wb,category):
    if str(category).lower()!="engineering": return wb
    ws=wb[category] if category in wb.sheetnames else wb.active
    if any(str(ws.cell(r,c).value or "").strip().lower()=="ipr" for r in range(1,8) for c in range(1,ws.max_column+1)): return wb
    fc=next((c for c in range(1,ws.max_column+1) if any("faculty details" in str(ws.cell(r,c).value or "").lower() for r in range(1,8))),None)
    if not fc: return wb

    # openpyxl does not reliably shift merged ranges when inserting columns.
    # Remove the original Faculty Details merge first, then recreate it after
    # inserting the two IPR columns. This prevents the overlapping merge that
    # made the Engineering IPR header look different/misaligned.
    old_faculty_merges=[]
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col <= fc <= rng.max_col and rng.min_row <= 4 <= rng.max_row:
            old_faculty_merges.append(str(rng))
    for rng in old_faculty_merges:
        ws.unmerge_cells(rng)

    ws.insert_cols(fc,2)
    faculty_col=fc+2

    # Recreate Faculty Details exactly as the original Engineering section.
    ws.merge_cells(start_row=4,start_column=faculty_col,end_row=6,end_column=faculty_col)
    ws.cell(4,faculty_col).value="Faculty Details"

    # IPR should follow the same visual hierarchy as Pharmacy and the other
    # Engineering sections: one merged group heading over two sub-headings.
    # Use Consultancy Project as the style reference because it is the adjacent
    # standard section in the Engineering template.
    group_src=max(1,fc-3)   # AJ before insertion
    sub_src=max(1,fc-3)
    data_style_src=max(1,fc-3)

    # New IPR data columns inherit the neighbouring table style.
    for r in range(7,ws.max_row+1):
        src=ws.cell(r,data_style_src)
        for c in (fc,fc+1):
            dst=ws.cell(r,c)
            if src.has_style: dst._style=copy(src._style)
            dst.alignment=copy(src.alignment)

    # Group header occupies rows 4-5, matching Financial/Sponsored/Consultancy.
    for r in (4,5):
        for c in (fc,fc+1):
            src=ws.cell(r,group_src); dst=ws.cell(r,c)
            if src.has_style: dst._style=copy(src._style)
            dst.alignment=copy(src.alignment)
    ws.merge_cells(start_row=4,start_column=fc,end_row=5,end_column=fc+1)
    ws.cell(4,fc).value="IPR"

    # Two patent sub-headings on row 6, styled exactly like normal sub-headings.
    for c,offset in ((fc,0),(fc+1,1)):
        src=ws.cell(6,sub_src+offset); dst=ws.cell(6,c)
        if src.has_style: dst._style=copy(src._style)
        dst.alignment=copy(src.alignment)
    ws.cell(6,fc).value="No. of Patents Published"
    ws.cell(6,fc+1).value="No. of Patents Granted"

    # Keep both IPR columns balanced and readable.
    base_width=ws.column_dimensions[ws.cell(1,sub_src).column_letter].width
    if base_width:
        ws.column_dimensions[ws.cell(1,fc).column_letter].width=base_width
        ws.column_dimensions[ws.cell(1,fc+1).column_letter].width=base_width
    return wb

@app.get("/", include_in_schema=False)
def home():
    return HTMLResponse(content=FRONTEND_HTML)

@app.get("/health")
def health(): return {"status":"ok","version":"40.0","parallel_workers":_PARALLEL_WORKERS,"parallel_mode":"process" if _USE_PROCESS_POOL else "thread","built_in_templates":"enabled","ocr":"disabled","engineering_pharmacy_ipr":"enabled","management_overall_edp_mdp":"enabled","faculty_multi_rule":"explicit_count_then_srno"}

@app.post("/inspect")
async def inspect(pdf:UploadFile=File(...),category:str=Form("Engineering")):
    xb=_builtin_template_bytes(category); pb=await pdf.read(); wb=_ensure_engineering_ipr(load_workbook(io.BytesIO(xb),data_only=False),category); ws=wb[category] if category in wb.sheetnames else wb.active; d=extract_data(pb,category); return {"category":category,"template_sheet":ws.title,"headings":get_headers(ws),"extraction":d}

def _clean_engineering_title(wb,category):
    # The generated Engineering workbook should not display the old fixed
    # template title "NIRF Data 2026". Keep the template file itself intact.
    if str(category or "").strip().lower()!="engineering": return wb
    ws=wb[category] if category in wb.sheetnames else wb.active
    for row in ws.iter_rows(min_row=1,max_row=min(3,ws.max_row)):
        for cell in row:
            if re.search(r"\bNIRF\s+Data\s+20\d{2}\b",str(cell.value or ""),re.I):
                cell.value=None
    return wb

def _year_filename(category,count,years):
    years=sorted({str(y) for y in years if y and re.fullmatch(r"20\d{2}",str(y))})
    year_part=(years[0] if len(years)==1 else ("-".join(years) if years else ""))
    cat=re.sub(r"[^A-Za-z0-9 _-]","",str(category)).strip() or "NIRF"
    word="Entry" if count==1 else "Entries"
    # Required format: Nirf 2022 Pharmacy 5 Entries.xlsx
    return f"Nirf{(' '+year_part) if year_part else ''} {cat} {count} {word}.xlsx"

_NIRF_CATEGORY_CODE={"engineering":"E","pharmacy":"P","management":"M","overall":"O"}
_NIRF_CATEGORY_FOLDER={"engineering":"Engineering","pharmacy":"Pharmacy","management":"Management","overall":"Overall"}

def _http_get(url, timeout=40):
    req=urllib.request.Request(url,headers={"User-Agent":"Mozilla/5.0 (compatible; NIRF-Data-Automation/41.0)","Accept":"text/html,application/pdf,*/*","Referer":"https://www.nirfindia.org/"})
    try:
        with urllib.request.urlopen(req,timeout=timeout) as resp:
            return resp.read(),(resp.headers.get("Content-Type") or "").lower()
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"Official NIRF website returned HTTP {exc.code}.")
    except Exception as exc:
        raise RuntimeError(f"Could not connect to the official NIRF website: {exc}")

def _nirf_ranked_ids(year,category,limit):
    cat=str(category or "").strip().lower(); code=_NIRF_CATEGORY_CODE.get(cat); folder=_NIRF_CATEGORY_FOLDER.get(cat)
    if not code or not folder: raise ValueError("Invalid NIRF category.")
    url=f"https://www.nirfindia.org/Rankings/{int(year)}/{folder}Ranking.html"
    raw,_=_http_get(url,35); html=raw.decode("utf-8","ignore")
    pat=re.compile(rf"\bIR-{code}-[A-Z]-\d+\b",re.I); ids=[]; seen=set()
    for m in pat.finditer(html):
        iid=m.group(0).upper()
        if iid not in seen:
            seen.add(iid); ids.append(iid)
        if len(ids)>=int(limit): break
    if len(ids)<int(limit): raise RuntimeError(f"Only {len(ids)} ranked institution IDs were found for NIRF {year} {folder}; {limit} were requested.")
    return ids

def _nirf_pdf_url(year,category,institute_id):
    folder=_NIRF_CATEGORY_FOLDER.get(str(category or "").strip().lower())
    if not folder: raise ValueError("Invalid NIRF category.")
    return f"https://www.nirfindia.org/nirfpdfcdn/{int(year)}/pdf/{folder}/{institute_id}.pdf"

def _download_official_nirf_pdf(year,category,institute_id):
    data,_=_http_get(_nirf_pdf_url(year,category,institute_id),60)
    if not data.startswith(b"%PDF"): raise RuntimeError(f"Official PDF not available for {institute_id}.")
    return data

def _process_auto_job(job_id,year,category,count):
    try:
        year=int(year); count=int(count)
        _set_job(job_id,status="processing",completed=0,total=count,percent=0,message=f"Fetching official NIRF {year} ranking list...")
        ids=_nirf_ranked_ids(year,category,count)
        pdfs=[None]*len(ids)
        # Network downloads are I/O-bound, so download several official PDFs concurrently.
        download_workers=min(8,len(ids))
        with ThreadPoolExecutor(max_workers=download_workers,thread_name_prefix="nirf-download") as pool:
            futures={pool.submit(_download_official_nirf_pdf,year,category,iid):idx for idx,iid in enumerate(ids)}
            downloaded=0
            for fut in as_completed(futures):
                idx=futures[fut]
                try: pdfs[idx]=fut.result()
                except Exception: pdfs[idx]=None
                downloaded+=1
                pct=min(20,round(downloaded*20/max(1,len(ids))))
                _set_job(job_id,completed=0,percent=pct,message=f"Downloading {downloaded} of {len(ids)} official PDFs")
        valid=[(idx,data,category) for idx,data in enumerate(pdfs) if data]
        if not valid: raise RuntimeError("None of the official NIRF PDFs could be downloaded.")
        results=[None]*len(ids); workers=min(_PARALLEL_WORKERS,len(valid))
        Executor=ProcessPoolExecutor if (_USE_PROCESS_POOL and workers>1) else ThreadPoolExecutor
        with Executor(max_workers=workers) as pool:
            futures={pool.submit(_extract_pdf_worker,item):item[0] for item in valid}; done_count=0
            for fut in as_completed(futures):
                idx=futures[fut]
                try: _,results[idx]=fut.result()
                except Exception: results[idx]=None
                done_count+=1
                pct=min(95,20+round(done_count*75/max(1,len(valid))))
                _set_job(job_id,completed=done_count,total=len(ids),percent=pct,message=f"Processing {done_count} of {len(valid)} downloaded PDFs")
        xb=_builtin_template_bytes(category); wb=_ensure_engineering_ipr(load_workbook(io.BytesIO(xb)),category); wb=_clean_engineering_title(wb,category)
        ws=wb[category] if category in wb.sheetnames else wb.active; _,header_end=_header_bounds(ws); row=header_end+1; success=0
        _set_job(job_id,percent=97,message="Creating Excel file...")
        for d in results:
            if d is None: continue
            wb,_=fill_workbook(xb,d,category,data_row=row,wb=wb,sl_no=row-header_end); row+=1; success+=1
        if not success: raise RuntimeError("None of the official NIRF PDFs could be processed.")
        _set_job(job_id,percent=98,message="Finalizing Excel file...")
        out=io.BytesIO(); wb.save(out); name=_year_filename(category,success,[str(year)])
        _set_job(job_id,status="done",completed=len(ids),total=len(ids),percent=99,message="Preparing download...",filename=name,output=out.getvalue(),success=success,failed=len(ids)-success,source="official_nirf",requested=count)
    except Exception as exc:
        _set_job(job_id,status="error",error=str(exc),message="Processing failed")

def _set_job(job_id, **changes):
    with _JOBS_LOCK:
        if job_id in _JOBS: _JOBS[job_id].update(changes)

def _extract_pdf_worker(item_category):
    """Process-pool-safe PDF extraction worker."""
    idx, data, category = item_category
    return idx, extract_data(data, category)

def _process_batch_job(job_id, pdf_items, category):
    try:
        total=len(pdf_items); results=[None]*total
        _set_job(job_id,status="processing",completed=0,total=total,percent=0,message=f"Processing 0 of {total} PDFs")
        workers=min(_PARALLEL_WORKERS,total)
        Executor=ProcessPoolExecutor if (_USE_PROCESS_POOL and workers>1) else ThreadPoolExecutor
        with Executor(max_workers=workers) as pool:
            futures={pool.submit(_extract_pdf_worker,(idx,data,category)):idx for idx,_name,data in pdf_items}; done_count=0
            for fut in as_completed(futures):
                idx=futures[fut]
                try: _,results[idx]=fut.result()
                except Exception: results[idx]=None
                done_count+=1; pct=min(95,round(done_count*95/max(1,total)))
                _set_job(job_id,completed=done_count,percent=pct,message=f"Processing {done_count} of {total} PDFs")
        xb=_builtin_template_bytes(category); wb=_ensure_engineering_ipr(load_workbook(io.BytesIO(xb)),category); wb=_clean_engineering_title(wb,category)
        ws=wb[category] if category in wb.sheetnames else wb.active; _,header_end=_header_bounds(ws); row=header_end+1; years=[]; success=0
        _set_job(job_id,percent=97,message="Creating Excel file...")
        for d in results:
            if d is None: continue
            wb,_=fill_workbook(xb,d,category,data_row=row,wb=wb,sl_no=row-header_end); years.append(d.get("nirf_year")); row+=1; success+=1
        if not success: raise ValueError("No PDF could be processed.")
        _set_job(job_id,percent=98,message="Finalizing Excel file...")
        out=io.BytesIO(); wb.save(out); name=_year_filename(category,success,years)
        _set_job(job_id,status="done",completed=total,total=total,percent=99,message="Preparing download...",filename=name,output=out.getvalue(),success=success,failed=total-success)
    except Exception as exc: _set_job(job_id,status="error",error=str(exc),message="Processing failed")

@app.post("/start-auto")
async def start_auto(year:int=Form(...),category:str=Form("Engineering"),count:int=Form(100)):
    if year<2020 or year>2026: raise HTTPException(status_code=400,detail="Select a NIRF ranking year from 2020 to 2026.")
    if str(category or "").strip().lower() not in _NIRF_CATEGORY_CODE: raise HTTPException(status_code=400,detail="Invalid category.")
    if count<1 or count>100: raise HTTPException(status_code=400,detail="Enter number of institutions from 1 to 100.")
    job_id=uuid.uuid4().hex
    with _JOBS_LOCK:
        _JOBS[job_id]={"status":"queued","completed":0,"total":count,"message":"Starting...","created":time.time(),"source":"official_nirf"}
        cutoff=time.time()-3600
        for old in list(_JOBS):
            if old!=job_id and _JOBS[old].get("created",0)<cutoff: _JOBS.pop(old,None)
    threading.Thread(target=_process_auto_job,args=(job_id,year,category,count),daemon=True).start()
    return {"job_id":job_id,"total":count,"workers":min(_PARALLEL_WORKERS,count),"source":"official_nirf"}

@app.post("/start-batch")
async def start_batch(pdfs:List[UploadFile]=File(...),category:str=Form("Engineering")):
    if not pdfs: raise HTTPException(status_code=400,detail="Please upload at least one PDF.")
    if len(pdfs)>100: raise HTTPException(status_code=400,detail=f"Maximum 100 PDFs per batch. You selected {len(pdfs)}.")
    items=[]
    for idx,pdf in enumerate(pdfs): items.append((idx,pdf.filename or f"PDF {idx+1}",await pdf.read()))
    job_id=uuid.uuid4().hex
    with _JOBS_LOCK:
        _JOBS[job_id]={"status":"queued","completed":0,"total":len(items),"percent":0,"message":"Starting...","created":time.time()}
        cutoff=time.time()-3600
        for old in list(_JOBS):
            if old!=job_id and _JOBS[old].get("created",0)<cutoff: _JOBS.pop(old,None)
    threading.Thread(target=_process_batch_job,args=(job_id,items,category),daemon=True).start()
    return {"job_id":job_id,"total":len(items),"workers":min(_PARALLEL_WORKERS,len(items)),"parallel_mode":"process" if _USE_PROCESS_POOL else "thread"}

@app.get("/batch-status/{job_id}")
def batch_status(job_id:str):
    with _JOBS_LOCK:
        j=_JOBS.get(job_id)
        if not j: raise HTTPException(status_code=404,detail="Batch job not found or expired.")
        return {k:v for k,v in j.items() if k!="output"}

@app.get("/batch-download/{job_id}")
def batch_download(job_id:str):
    with _JOBS_LOCK:
        j=_JOBS.get(job_id)
        if not j: raise HTTPException(status_code=404,detail="Batch job not found or expired.")
        if j.get("status")!="done": raise HTTPException(status_code=409,detail="Excel is not ready yet.")
        data=j.get("output"); name=j.get("filename","Nirf data.xlsx")
    return StreamingResponse(io.BytesIO(data),media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":f'attachment; filename="{name}"'})

@app.post("/process-batch")
async def process_batch(pdfs:List[UploadFile]=File(...),category:str=Form("Engineering")):
    if not pdfs: raise HTTPException(status_code=400,detail="Please upload at least one PDF.")
    if len(pdfs)>100: raise HTTPException(status_code=400,detail=f"Maximum 100 PDFs per batch. You selected {len(pdfs)}.")
    xb=_builtin_template_bytes(category); wb=_ensure_engineering_ipr(load_workbook(io.BytesIO(xb)),category); wb=_clean_engineering_title(wb,category); ws=wb[category] if category in wb.sheetnames else wb.active; _,header_end=_header_bounds(ws); row=header_end+1; years=[]
    for pdf in pdfs:
        try:
            d=extract_data(await pdf.read(),category); wb,_=fill_workbook(xb,d,category,data_row=row,wb=wb,sl_no=row-header_end); years.append(d.get("nirf_year")); row+=1
        except Exception: continue
    if row==header_end+1: raise HTTPException(status_code=400,detail="No PDF could be processed.")
    out=io.BytesIO(); wb.save(out); out.seek(0); count=row-header_end-1; name=_year_filename(category,count,years)
    return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":f'attachment; filename="{name}"'})

@app.post("/process")
async def process(pdf:UploadFile=File(...),category:str=Form("Engineering")):
    xb=_builtin_template_bytes(category); wb=_ensure_engineering_ipr(load_workbook(io.BytesIO(xb)),category); wb=_clean_engineering_title(wb,category); d=extract_data(await pdf.read(),category); wb,_=fill_workbook(xb,d,category,wb=wb); out=io.BytesIO(); wb.save(out); out.seek(0); name=_year_filename(category,1,[d.get("nirf_year")]); return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":f'attachment; filename="{name}"'})
