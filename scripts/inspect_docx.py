# -*- coding: utf-8 -*-
"""
วัดผลจริงจากไฟล์ที่เรนเดอร์แล้ว — อย่าเชื่อว่าโค้ดถูก ให้ดูหน้าตาไฟล์จริง

ใช้:
    python inspect_docx.py ไฟล์.docx            ตรวจครบชุด
    python inspect_docx.py ไฟล์.docx --png out  เรนเดอร์เป็นภาพไว้ดูด้วยตา
    python inspect_docx.py ไฟล์.docx --breaks   แสดงรอยต่อบรรทัดทุกจุด ไว้ไล่อ่านเอง

ตรวจ 3 อย่าง
  1. จุดตัดบรรทัดกลางคำ  (สิ่งที่ Word ทำเองถ้าไม่ใส่ ZWSP — วัดจากเอกสารจริงเจอ 12.7%)
  2. ช่องว่างเฉลี่ยระหว่างอักขระ  (เกิน 0.8 pt = บรรทัดถ่าง ต้องบีบ)
  3. ขอบขวาของแต่ละบรรทัด  (ดูว่าชิดจริงไหม)

ข้อ 1 ใช้พจนานุกรมเดียวกับตัวตัดคำ จึงมองไม่เห็นคำที่พจนานุกรมไม่มี (เฝ้า|ระวัง · กรมควบคุม|โรค)
ต้องใช้ --breaks อ่านทุกรอยต่อด้วยตาเสมอ — ตอนทำตัวอย่างประเด็นถามตอบ จุดผิดจริงเจอจากการอ่านทั้งหมด

หมายเหตุเรื่อง Word COM: ห้ามให้ Word บันทึกไฟล์ลง OneDrive โดยตรง
สคริปต์นี้แค่ 'อ่าน' แล้วส่งออก PDF ลงโฟลเดอร์ชั่วคราว จึงปลอดภัย
และเปิดด้วย DispatchEx เพื่อไม่ไปยุ่งกับ Word ที่กิ๊ฟเปิดค้างอยู่
"""
from __future__ import annotations

import os
import sys
import tempfile

CM = 28.3465          # pt ต่อ 1 ซม.
WD_EXPORT_PDF = 17


# ─────────────────────────────────────────────────────── docx -> pdf

def to_pdf(docx_path: str, out_dir: str | None = None) -> str:
    """ส่งออก PDF ด้วย Word instance แยก (ไม่รบกวนไฟล์ที่กิ๊ฟเปิดอยู่)"""
    import pythoncom
    import win32com.client as win32

    docx_path = os.path.abspath(docx_path)
    # Word ต้องการ path เต็ม — ส่งโฟลเดอร์แบบ relative มาจะขึ้น "The directory name isn't valid"
    out_dir = os.path.abspath(out_dir) if out_dir else tempfile.mkdtemp(prefix="giftdoc_")
    os.makedirs(out_dir, exist_ok=True)
    dst = os.path.join(out_dir,
                       os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")

    pythoncom.CoInitialize()
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Open(docx_path, ReadOnly=True,
                                 AddToRecentFiles=False, Visible=False)
        try:
            doc.ExportAsFixedFormat(OutputFileName=dst,
                                    ExportFormat=WD_EXPORT_PDF,
                                    OpenAfterExport=False)
        finally:
            doc.Close(SaveChanges=0)
    finally:
        app.Quit()
        pythoncom.CoUninitialize()
    return dst


# ─────────────────────────────────────────────────────────── การวัด

def _line_text(ln) -> str:
    return "".join(s["text"] for s in ln["spans"])


def midword_breaks(pdf_path: str, *, right_slack_cm: float = 3.2):
    """
    หาบรรทัดที่ตัดกลางคำ — นับเฉพาะบรรทัดที่ยาวเกือบเต็ม (คือถูกห่อจริง)
    ข้ามคู่ที่จุดตัดตรงกับช่องว่างในต้นฉบับ ไม่งั้นได้ผลบวกลวง
    """
    import fitz
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from thai_break import find_midword_breaks

    d = fitz.open(pdf_path)
    pairs = []
    for page in d:
        limit = page.rect.width - right_slack_cm * CM
        for b in page.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            lines = b["lines"]
            for i in range(len(lines) - 1):
                if lines[i]["bbox"][2] < limit:
                    continue                    # บรรทัดสั้น = ท้ายย่อหน้า ไม่ใช่การห่อ
                pairs.append((_line_text(lines[i]), _line_text(lines[i + 1])))
    total = len(pairs)
    d.close()
    return find_midword_breaks(pairs), total


def char_gaps(pdf_path: str, *, size: float = 16.0, min_chars: int = 12):
    """
    ช่องว่างเฉลี่ยระหว่างอักขระของแต่ละบรรทัด (pt)
    บรรทัดที่เกิน MAX_GAP_PT คือบรรทัดถ่าง ต้องไปบีบอักษร
    """
    import fitz
    d = fitz.open(pdf_path)
    rows = []
    for pno, page in enumerate(d, 1):
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for ln in b["lines"]:
                chars = [c for s in ln["spans"] for c in s.get("chars", [])]
                if len(chars) < min_chars:
                    continue
                gaps = []
                for a, z in zip(chars, chars[1:]):
                    g = z["bbox"][0] - a["bbox"][2]
                    if 0 <= g < size:          # กันค่าประหลาดจากสระซ้อน
                        gaps.append(g)
                if gaps:
                    rows.append((pno, sum(gaps) / len(gaps),
                                 "".join(c["c"] for c in chars)[:40]))
    d.close()
    return rows


def line_fill(pdf_path: str, *, right_margin_cm: float = 2.0):
    """ขอบขวาของแต่ละบรรทัดคิดเป็น % ของกรอบข้อความ"""
    import fitz
    d = fitz.open(pdf_path)
    out = []
    for pno, page in enumerate(d, 1):
        edge = page.rect.width - right_margin_cm * CM
        for b in page.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            lines = b["lines"]
            for i, ln in enumerate(lines):
                if i < len(lines) - 1:
                    out.append((pno, ln["bbox"][2] / edge * 100))
    d.close()
    return out


def render_png(pdf_path: str, out_dir: str, dpi: int = 150, pages=None):
    """เรนเดอร์เป็นภาพเพื่อ 'ดูด้วยตาจริง' ก่อนส่ง"""
    import fitz
    os.makedirs(out_dir, exist_ok=True)
    d = fitz.open(pdf_path)
    made = []
    for i, page in enumerate(d):
        if pages and (i + 1) not in pages:
            continue
        dst = os.path.join(out_dir, "p%02d.png" % (i + 1))
        page.get_pixmap(dpi=dpi).save(dst)
        made.append(dst)
    d.close()
    return made


def line_breaks(pdf_path: str, *, right_margin_cm: float = 2.0, slack_pt: float = 6.0):
    """
    รอยต่อบรรทัดทุกจุดที่เกิดจากการห่อบรรทัด (บรรทัดบนยาวจนชนขอบขวา)
    คืน [(หน้า, ท้ายบรรทัดบน, ต้นบรรทัดล่าง)] ไว้ไล่อ่านด้วยตา

    รวมบรรทัดตามแนว y ก่อน เพราะ PyMuPDF มักแยกเลขข้อ "1." ออกเป็นบรรทัดของมันเอง
    ถ้าจับคู่ตาม block ตรง ๆ บรรทัดแรกของรายการเลขข้อจะหลุดการตรวจ (เคยพลาด "กรมควบคุม / โรค")
    """
    import fitz
    zw = chr(0x200B)
    out = []
    d = fitz.open(pdf_path)
    for pno, page in enumerate(d, 1):
        right = page.rect.width - right_margin_cm * CM
        rows = {}
        for b in page.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                y = round(ln["bbox"][3])
                key = next((k for k in rows if abs(k - y) <= 2), y)
                rows.setdefault(key, []).append(ln)
        vis = []
        for y in sorted(rows):
            lns = sorted(rows[y], key=lambda l: l["bbox"][0])
            txt = "".join(_line_text(l) for l in lns).replace(zw, "")
            vis.append((max(l["bbox"][2] for l in lns), txt))
        for (x1, a), (_, c) in zip(vis, vis[1:]):
            if x1 >= right - slack_pt and a.strip() and c.strip():
                out.append((pno, a.rstrip(), c.lstrip()))
    d.close()
    return out


# ──────────────────────────────────────────────────────────── รายงาน

def report(docx_path: str, png_dir: str | None = None):
    from thai_fit import MAX_GAP_PT

    pdf = to_pdf(docx_path)
    bad, total = midword_breaks(pdf)
    gaps = char_gaps(pdf)
    fills = line_fill(pdf)

    print("ไฟล์: %s" % os.path.basename(docx_path))
    pct = len(bad) / total * 100 if total else 0
    print("  ตัดกลางคำ      : %d / %d บรรทัด (%.1f%%)" % (len(bad), total, pct))
    for a, b in bad[:8]:
        print("      ...%s  |  %s..." % (a, b))

    wide = [g for g in gaps if g[1] > MAX_GAP_PT]
    print("  บรรทัดถ่างเกิน %.1f pt: %d / %d" % (MAX_GAP_PT, len(wide), len(gaps)))
    for pno, g, txt in wide[:8]:
        print("      หน้า %d  %.2f pt  %s" % (pno, g, txt))

    if fills:
        avg = sum(f for _, f in fills) / len(fills)
        print("  ขอบขวาเฉลี่ย   : %.1f%% ของกรอบข้อความ" % avg)

    if png_dir:
        made = render_png(pdf, png_dir)
        print("  ภาพ            : %d หน้า -> %s" % (len(made), png_dir))
    return {"midword": len(bad), "lines": total,
            "stretched": len(wide), "pdf": pdf}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    png = None
    if "--png" in sys.argv:
        png = sys.argv[sys.argv.index("--png") + 1]
    info = report(sys.argv[1], png)
    if "--breaks" in sys.argv:
        rows = line_breaks(info["pdf"])
        print("  รอยต่อบรรทัดทั้งหมด %d จุด — อ่านทีละจุดว่าสะดุดไหม:" % len(rows))
        for i, (pno, a, c) in enumerate(rows, 1):
            print("   %3d  หน้า %d  %16s | %s" % (i, pno, a[-16:], c[:16]))
