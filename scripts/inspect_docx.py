# -*- coding: utf-8 -*-
"""
วัดผลจริงจากไฟล์ที่เรนเดอร์แล้ว — อย่าเชื่อว่าโค้ดถูก ให้ดูหน้าตาไฟล์จริง

ใช้ (macOS พิมพ์ python3 แทน python):
    python inspect_docx.py ไฟล์.docx            ตรวจครบชุด
    python inspect_docx.py ไฟล์.docx --png out  เรนเดอร์เป็นภาพไว้ดูด้วยตา
    python inspect_docx.py ไฟล์.docx --breaks   แสดงรอยต่อบรรทัดทุกจุด ไว้ไล่อ่านเอง
    python inspect_docx.py ไฟล์.docx --engine libreoffice   เครื่องที่ไม่มี Word (ได้ผลคร่าว ๆ)

ตรวจ 4 อย่าง
  1. จุดตัดบรรทัดกลางคำ  (สิ่งที่ Word ทำเองถ้าไม่ใส่ ZWSP — วัดจากเอกสารจริงเจอ 12.7%)
  2. ช่องว่างเฉลี่ยระหว่างอักขระ  (เกิน 0.8 pt = บรรทัดถ่าง ต้องบีบ)
  3. ขอบขวาของแต่ละบรรทัด  (ดูว่าชิดจริงไหม)
  4. ฟอนต์ใน PDF ตรงกับที่เอกสารสั่งไหม  (เครื่องไม่มีฟอนต์ Word จะใช้ตัวอื่นแทนเงียบ ๆ ผลวัดใช้ไม่ได้)

ข้อ 1 ใช้พจนานุกรมเดียวกับตัวตัดคำ จึงมองไม่เห็นคำที่พจนานุกรมไม่มี (เฝ้า|ระวัง, กรมควบคุม|โรค)
ต้องใช้ --breaks อ่านทุกรอยต่อด้วยตาเสมอ — ตอนทำตัวอย่างประเด็นถามตอบ จุดผิดจริงเจอจากการอ่านทั้งหมด

ส่งออก PDF ด้วยอะไร (to_pdf)
  Windows  Word ผ่าน COM เปิดด้วย DispatchEx เป็น instance แยก ไม่ไปยุ่งกับ Word ที่กิ๊ฟเปิดค้างอยู่
  macOS    Word for Mac ผ่าน AppleScript (word_pdf_mac.applescript) — Mac มี Word ได้ตัวเดียว
           จึงเปิดสำเนาชื่อสุ่มใน ~/riabroi-pdf แล้วปิดเฉพาะไฟล์นั้น
  ไม่มี Word  LibreOffice — ตัดบรรทัดไทยไม่เหมือน Word ใช้ดูภาพรวมได้ ห้ามใช้ตัดสินว่าผ่าน
ทุกทางห้ามให้โปรแกรมเอกสารเขียนลง OneDrive ตรง ๆ — ส่งออกในโฟลเดอร์ทำงานแล้ว Python คัดลอกออกไป
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
import warnings

CM = 28.3465          # pt ต่อ 1 ซม.
WD_EXPORT_PDF = 17
HERE = os.path.dirname(os.path.abspath(__file__))
MAC_SCRIPT = os.path.join(HERE, "word_pdf_mac.applescript")
MAC_TIMEOUT = 300     # วินาที — เผื่อให้คนกดกล่องขออนุญาตครั้งแรกทัน
MAC_WORD_APPS = ("/Applications/Microsoft Word.app", "~/Applications/Microsoft Word.app")
ENGINES = ("auto", "word", "libreoffice")
ENGINE_LABEL = {"word": "Microsoft Word", "word-mac": "Microsoft Word for Mac",
                "libreoffice": "LibreOffice (ผลคร่าว ๆ)", "unknown": "ไม่ทราบ"}
LO_WARNING = ("PDF จาก LibreOffice ตัดบรรทัดไทยไม่เหมือน Word — ใช้ดูภาพรวมได้ "
              "แต่ห้ามใช้ตัดสินว่าผ่าน ต้องตรวจซ้ำบนเครื่องที่มี Word")


# ─────────────────────────────────────────────── หาตัวส่งออก PDF

def find_word() -> str | None:
    """Word ที่สั่งงานได้ หรือ None — Windows ดูทะเบียน COM · macOS ดูโฟลเดอร์แอปแล้วค่อยถาม Spotlight"""
    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
                return "Word.Application"
        except OSError:
            return None
    if sys.platform == "darwin":
        for app in MAC_WORD_APPS:
            app = os.path.expanduser(app)
            if os.path.isdir(app):
                return app
        try:
            out = subprocess.run(["mdfind", "kMDItemCFBundleIdentifier == 'com.microsoft.Word'"],
                                 capture_output=True, text=True, timeout=15).stdout
        except (OSError, subprocess.SubprocessError):
            return None
        return next((p for p in out.splitlines()
                     if p.endswith(".app") and os.path.isdir(p)), None)
    return None


def find_libreoffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    for path in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
                 r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.isfile(path):
            return path
    return None


def mac_workdir() -> str:
    """
    โฟลเดอร์ที่ Word for Mac ใช้แปลงไฟล์ — ต้องเป็นที่เดิมทุกครั้ง
    Word อยู่ใน sandbox ครั้งแรกจะถามสิทธิ์ (Grant File Access) แล้วจำโฟลเดอร์ไว้
    ถ้าสุ่มโฟลเดอร์ใหม่ทุกรอบจะโดนถามทุกรอบ · เปลี่ยนที่ได้ด้วยตัวแปร RIABROI_WORKDIR
    """
    return os.path.abspath(os.path.expanduser(
        os.environ.get("RIABROI_WORKDIR") or "~/riabroi-pdf"))


# ─────────────────────────────────────────────────────── docx -> pdf

def to_pdf(docx_path: str, out_dir: str | None = None, *, engine: str = "auto") -> str:
    """
    ส่งออก PDF แล้วคืน path ของไฟล์ PDF

    engine  "auto"         Word ถ้ามี — ไม่มีค่อยใช้ LibreOffice พร้อมคำเตือน
            "word"         Word เท่านั้น (Windows: COM · macOS: AppleScript)
            "libreoffice"  LibreOffice — ตัดบรรทัดไม่เหมือน Word ใช้ดูภาพรวมเท่านั้น
    """
    if engine not in ENGINES:
        raise ValueError("engine ต้องเป็น " + " / ".join(ENGINES))
    docx_path = os.path.abspath(docx_path)
    # Word ต้องการ path เต็ม — ส่งโฟลเดอร์แบบ relative มาจะขึ้น "The directory name isn't valid"
    out_dir = os.path.abspath(out_dir) if out_dir else tempfile.mkdtemp(prefix="giftdoc_")
    os.makedirs(out_dir, exist_ok=True)
    dst = os.path.join(out_dir,
                       os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")

    if engine != "libreoffice":
        if find_word():
            if sys.platform == "win32":
                _word_pdf_windows(docx_path, dst)
            else:
                _word_pdf_mac(docx_path, dst)
            return dst
        if engine == "word":
            raise RuntimeError("ไม่พบ Microsoft Word บนเครื่องนี้ — รัน scripts/doctor.py ดูว่าขาดอะไร")
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("ไม่พบ Microsoft Word หรือ LibreOffice สำหรับส่งออก PDF — "
                           "ติดตั้งอย่างใดอย่างหนึ่งแล้วรัน scripts/doctor.py")
    if engine == "auto":
        warnings.warn("ไม่พบ Microsoft Word จึงใช้ LibreOffice แทน — " + LO_WARNING,
                      RuntimeWarning, stacklevel=2)
    _libreoffice_pdf(soffice, docx_path, dst)
    return dst


def _word_pdf_windows(src: str, dst: str) -> None:
    """Word instance แยก (DispatchEx) ไม่รบกวนไฟล์ที่กิ๊ฟเปิดอยู่ · ส่งออกในเทมป์แล้วค่อยคัดลอก"""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        raise RuntimeError("ต้องติดตั้ง pywin32 ก่อน: pip install pywin32") from None

    work = tempfile.mkdtemp(prefix="riabroi_word_")
    tmp_pdf = os.path.join(work, "out.pdf")
    try:
        pythoncom.CoInitialize()
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            doc = app.Documents.Open(src, ReadOnly=True,
                                     AddToRecentFiles=False, Visible=False)
            try:
                doc.ExportAsFixedFormat(OutputFileName=tmp_pdf,
                                        ExportFormat=WD_EXPORT_PDF,
                                        OpenAfterExport=False)
            finally:
                doc.Close(SaveChanges=0)
        finally:
            app.Quit()
            pythoncom.CoUninitialize()
        shutil.copyfile(tmp_pdf, dst)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _word_pdf_mac(src: str, dst: str) -> None:
    """
    Word for Mac ผ่าน AppleScript — Mac มี Word ได้ตัวเดียว ต่างจาก DispatchEx บน Windows จึงต้อง
    - เปิดสำเนาชื่อสุ่ม ไม่เปิดไฟล์จริง: ถ้าไฟล์นั้นเปิดค้างอยู่ Word จะคืนหน้าต่างเดิมมา
      แล้วการปิดแบบไม่บันทึกจะทำให้งานที่ยังไม่ได้บันทึกหาย
    - แปลงในโฟลเดอร์เดิมทุกครั้ง (mac_workdir) Word จะได้ถามสิทธิ์แค่ครั้งแรก
    - คัดลอก PDF ออกไปด้วย Python — Word ไม่ได้เขียนลง OneDrive ตรง ๆ
    """
    work = mac_workdir()
    os.makedirs(work, exist_ok=True)
    _mac_readme(work)
    tag = uuid.uuid4().hex
    stem = "riabroi-" + tag
    copy_docx = os.path.join(work, stem + ".docx")
    copy_pdf = os.path.join(work, stem + ".pdf")
    shutil.copyfile(src, copy_docx)      # คัดลอกแต่เนื้อไฟล์ ไม่ติดธงดาวน์โหลดที่ทำให้ Word เปิดแบบป้องกัน
    try:
        try:
            proc = subprocess.run(
                ["osascript", MAC_SCRIPT, copy_docx, copy_pdf, stem, str(MAC_TIMEOUT)],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=MAC_TIMEOUT * 2 + 60)
        except subprocess.TimeoutExpired:
            raise RuntimeError(_mac_error("Python หยุดรอ osascript (-1712)")) from None
        if proc.returncode != 0:
            raise RuntimeError(_mac_error(proc.stderr or proc.stdout))
        if not os.path.isfile(copy_pdf):
            raise RuntimeError(
                "Word for Mac ทำงานจบแต่ไม่มีไฟล์ PDF — เปิด Word แล้ว Save As เป็น PDF "
                "ด้วยมือหนึ่งครั้ง เลือก Best for printing แล้วรันใหม่")
        shutil.copyfile(copy_pdf, dst)
    finally:
        for name in os.listdir(work):
            if tag in name:                   # สำเนา · PDF · ไฟล์ล็อก ~$ ของ Word
                try:
                    os.remove(os.path.join(work, name))
                except OSError:
                    pass


def _mac_readme(work: str) -> None:
    note = os.path.join(work, "อ่านก่อน.txt")
    if os.path.exists(note):
        return
    with open(note, "w", encoding="utf-8") as f:
        f.write("โฟลเดอร์ทำงานของ riabroi-ratchakan-typeset\n"
                "Word for Mac แปลง .docx เป็น PDF ในนี้ แล้วสคริปต์คัดลอก PDF ออกไปและลบไฟล์ชั่วคราวเอง\n"
                "อย่าลบโฟลเดอร์นี้ — Word จำสิทธิ์เข้าถึงไว้ที่นี่ ลบแล้วครั้งหน้าอาจถามสิทธิ์ใหม่\n")


_MAC_HINTS = {
    -1743: ("macOS ยังไม่อนุญาตให้สั่งงาน Word — เปิด System Settings > Privacy & Security > "
            "Automation แล้วเปิดสวิตช์ Microsoft Word ใต้แอปที่รันคำสั่งนี้ "
            "(Terminal, iTerm หรือ Claude) แล้วรันใหม่"),
    -1712: ("Word ไม่ตอบภายในเวลาที่รอ — มักมีกล่องค้างอยู่บนหน้าจอ Word: Grant File Access "
            "(กด Select... แล้วกด Grant Access ที่โฟลเดอร์ %s) · ลงชื่อเข้าใช้ Office · "
            "What's New — ปิดกล่องแล้วรันใหม่"),
    9001: "Word เปิดสำเนาเอกสารไม่สำเร็จ — ลองเปิดไฟล์นี้ใน Word ด้วยมือดูว่ามีข้อความเตือนอะไร",
    -600: "Word ปิดตัวหรือค้างระหว่างทำงาน — เปิด Word ด้วยมือให้พร้อมใช้ แล้วรันใหม่",
    -609: "Word ปิดตัวหรือค้างระหว่างทำงาน — เปิด Word ด้วยมือให้พร้อมใช้ แล้วรันใหม่",
}


def _mac_error(stderr: str) -> str:
    """แปลงรหัสข้อผิดพลาดของ AppleScript เป็นวิธีแก้ภาษาไทย"""
    raw = (stderr or "").strip()
    m = re.search(r"\((-?\d+)\)$", raw)
    hint = _MAC_HINTS.get(int(m.group(1)) if m else None, "Word for Mac แปลง PDF ไม่สำเร็จ")
    if "%s" in hint:
        hint = hint % mac_workdir()
    return hint + ("\n  ข้อความจาก osascript: " + raw if raw else "")


def _libreoffice_pdf(soffice: str, src: str, dst: str) -> None:
    """แปลงในเทมป์ด้วยโปรไฟล์แยก — ถ้า LibreOffice เปิดค้างอยู่ คำสั่ง headless จะเงียบไม่ทำอะไร"""
    import pathlib
    work = tempfile.mkdtemp(prefix="riabroi_lo_")
    try:
        profile = pathlib.Path(work, "profile").as_uri()
        proc = subprocess.run(
            [soffice, "-env:UserInstallation=" + profile, "--headless", "--norestore",
             "--convert-to", "pdf", "--outdir", work, src],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
        made = os.path.join(work, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        if proc.returncode != 0 or not os.path.isfile(made):
            raise RuntimeError("LibreOffice แปลง PDF ไม่สำเร็จ: "
                               + (proc.stderr or proc.stdout or "").strip())
        shutil.copyfile(made, dst)
    finally:
        shutil.rmtree(work, ignore_errors=True)


# ─────────────────────────────────────── อักษรเพี้ยนตอนดึงข้อความจาก PDF

# TH SarabunPSK เก็บรูปวรรณยุกต์/สระที่เลื่อนตำแหน่งไว้ในช่อง PUA (U+F700-F71A)
# ตัวสร้าง PDF บางตัวคืนรหัสพวกนี้แทนตัวจริง ตัวตรวจตัดกลางคำจะอ่านคำผิด — ต้อง map กลับก่อน
_PUA_THAI = {
    0xF700: 0x0E10, 0xF701: 0x0E34, 0xF702: 0x0E35, 0xF703: 0x0E36, 0xF704: 0x0E37,
    0xF705: 0x0E48, 0xF706: 0x0E49, 0xF707: 0x0E4A, 0xF708: 0x0E4B, 0xF709: 0x0E4C,
    0xF70A: 0x0E48, 0xF70B: 0x0E49, 0xF70C: 0x0E4A, 0xF70D: 0x0E4B, 0xF70E: 0x0E4C,
    0xF70F: 0x0E0D, 0xF710: 0x0E31, 0xF711: 0x0E4D, 0xF712: 0x0E47, 0xF713: 0x0E48,
    0xF714: 0x0E49, 0xF715: 0x0E4A, 0xF716: 0x0E4B, 0xF717: 0x0E4C, 0xF718: 0x0E38,
    0xF719: 0x0E39, 0xF71A: 0x0E3A,
}
_TONE = "[" + chr(0x0E48) + "-" + chr(0x0E4B) + "]"
_CONSONANT = "[" + chr(0x0E01) + "-" + chr(0x0E2E) + "]"
_NIKHAHIT, _SARA_AA, _SARA_AM = chr(0x0E4D), chr(0x0E32), chr(0x0E33)


def fix_thai_text(text: str) -> str:
    """แก้ข้อความที่ดึงจาก PDF: รหัส PUA → ตัวจริง · นิคหิต + า → สระอำ · 'ด าเนิน' → 'ดำเนิน'"""
    text = text.translate(_PUA_THAI)
    text = re.sub(_NIKHAHIT + "(" + _TONE + "?)" + _SARA_AA,
                  lambda m: m.group(1) + _SARA_AM, text)
    return re.sub("(" + _CONSONANT + ") (" + _TONE + "?)" + _SARA_AA,
                  lambda m: m.group(1) + m.group(2) + _SARA_AM, text)


# ─────────────────────────────────────────────────────────── การวัด

def _line_text(ln) -> str:
    return fix_thai_text("".join(s["text"] for s in ln["spans"]))


def midword_breaks(pdf_path: str, *, right_slack_cm: float = 3.2):
    """
    หาบรรทัดที่ตัดกลางคำ — นับเฉพาะบรรทัดที่ยาวเกือบเต็ม (คือถูกห่อจริง)
    ข้ามคู่ที่จุดตัดตรงกับช่องว่างในต้นฉบับ ไม่งั้นได้ผลบวกลวง
    """
    import fitz
    sys.path.insert(0, HERE)
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
                                 fix_thai_text("".join(c["c"] for c in chars))[:40]))
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


# ────────────────────────────────────────────── PDF มาจากไหน · ฟอนต์ถูกแทนไหม

def pdf_engine(pdf_path: str) -> str:
    """ดูจาก metadata ว่า PDF มาจากอะไร: word · word-mac · libreoffice · unknown"""
    import fitz
    with fitz.open(pdf_path) as d:
        meta = d.metadata or {}
    who = " ".join(filter(None, (meta.get("creator"), meta.get("producer"))))
    if "LibreOffice" in who or "OpenOffice" in who:
        return "libreoffice"
    if "Word" in who:
        mac = any(k in who for k in ("Quartz", "macOS", "Mac OS X"))
        return "word-mac" if mac else "word"
    return "unknown"


def pdf_fonts(pdf_path: str) -> list[str]:
    """ฟอนต์ที่ฝังใน PDF (ตัดคำนำหน้า subset แบบ BCDEEE+ ออก)"""
    import fitz
    names = set()
    with fitz.open(pdf_path) as d:
        for page in d:
            for f in page.get_fonts():
                names.add(re.sub(r"^[A-Z]{6}\+", "", f[3]))
    return sorted(names)


_FONT_TAILS = ("bolditalic", "boldoblique", "italic", "oblique", "bold",
               "regular", "psmt", "mt", "ps")


def _norm_font(name: str) -> str:
    """
    ชื่อฟอนต์ในเอกสารกับใน PDF เขียนต่างกัน — ทำให้เทียบกันได้
    TH SarabunPSK = BCDEEE+THSarabunPSK-Bold · TH SarabunIT๙ = THSarabunIT9 · Arial = ArialMT
    """
    name = re.sub(r"^[A-Z]{6}\+", "", (name or "").strip())
    name = name.translate({0x0E50 + i: 0x30 + i for i in range(10)})    # เลขไทย → อารบิก
    name = re.sub(r"[^0-9a-z]", "", name.lower())
    trimmed = True
    while trimmed:
        trimmed = False
        for tail in _FONT_TAILS:
            if name.endswith(tail) and len(name) > len(tail) + 2:
                name = name[:-len(tail)]
                trimmed = True
    return name


_DOCX_PARTS = re.compile(r"word/(document|styles|numbering|header\d*|footer\d*|"
                         r"footnotes|endnotes|theme/theme\d*)\.xml$")


def _rfonts(xml: str, slots: str = "ascii|hAnsi|cs|eastAsia") -> list[str]:
    """ชื่อฟอนต์ใน <w:rFonts> เท่านั้น — w:eastAsia ใน <w:lang> เป็นรหัสภาษา ไม่ใช่ฟอนต์"""
    return [name for tag in re.findall(r"<w:rFonts\b[^>]*>", xml)
            for name in re.findall(r'\bw:(?:%s)="([^"]+)"' % slots, tag)]


def font_check(docx_path: str, pdf_path: str) -> dict:
    """
    เทียบฟอนต์ที่เอกสารสั่งกับที่ฝังใน PDF

    main_missing  ฟอนต์ไทยหลักของเอกสารไม่อยู่ใน PDF = เครื่องนี้ไม่มีฟอนต์ Word ใช้ตัวอื่นแทน
                  ตัดบรรทัดและช่องไฟจะไม่ตรงกับเครื่องที่มีฟอนต์ ผลวัดทั้งหมดใช้ไม่ได้
                  (เจอง่ายบน Mac เพราะไม่มี TH Sarabun มาให้)
    others        ฟอนต์ที่เอกสารไม่ได้สั่ง — Word หยิบมาแทนบางตัวอักษร เช่น สัญลักษณ์ ขีด
    """
    import zipfile
    asked = set()
    with zipfile.ZipFile(docx_path) as z:
        for part in z.namelist():
            if not _DOCX_PARTS.match(part):
                continue
            xml = z.read(part).decode("utf-8", "replace")
            if "/theme/" in part:
                asked.update(n for n in re.findall(r'\btypeface="([^"]+)"', xml))
            else:
                asked.update(_rfonts(xml))
        cs = _rfonts(z.read("word/document.xml").decode("utf-8", "replace"), "cs")
        if not cs and "word/styles.xml" in z.namelist():
            cs = _rfonts(z.read("word/styles.xml").decode("utf-8", "replace"), "cs")
    main = max(set(cs), key=cs.count) if cs else None
    embedded = pdf_fonts(pdf_path)
    got = {_norm_font(f) for f in embedded}
    known = {_norm_font(f) for f in asked}
    return {"main": main,
            "main_missing": bool(main) and _norm_font(main) not in got,
            "embedded": embedded,
            "others": [f for f in embedded if _norm_font(f) not in known]}


# ──────────────────────────────────────────────────────────── รายงาน

def report(docx_path: str, png_dir: str | None = None, *, engine: str = "auto"):
    from thai_fit import MAX_GAP_PT

    pdf = to_pdf(docx_path, engine=engine)
    made_by = pdf_engine(pdf)
    fonts = font_check(docx_path, pdf)
    bad, total = midword_breaks(pdf)
    gaps = char_gaps(pdf)
    fills = line_fill(pdf)

    print("ไฟล์: %s" % os.path.basename(docx_path))
    print("  ส่งออกด้วย     : %s" % ENGINE_LABEL.get(made_by, made_by))
    if made_by == "libreoffice":
        print("  เตือน: " + LO_WARNING)
    print("  ฟอนต์ใน PDF    : %s" % (", ".join(fonts["embedded"]) or "-"))
    if fonts["main_missing"]:
        print("  เตือน: ไม่มี %s ใน PDF — เครื่องนี้ไม่มีฟอนต์นี้ Word จึงใช้ตัวอื่นแทน "
              "ผลวัดด้านล่างใช้ไม่ได้ ติดตั้งฟอนต์แล้วตรวจใหม่" % fonts["main"])
    elif fonts["others"]:
        print("  ฟอนต์ที่ไม่ได้สั่ง: %s (Word ใช้แทนบางตัวอักษร)" % ", ".join(fonts["others"]))

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
    return {"midword": len(bad), "lines": total, "stretched": len(wide), "pdf": pdf,
            "engine": made_by, "fonts": fonts}


if __name__ == "__main__":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    sys.path.insert(0, HERE)
    png = sys.argv[sys.argv.index("--png") + 1] if "--png" in sys.argv else None
    eng = sys.argv[sys.argv.index("--engine") + 1] if "--engine" in sys.argv else "auto"
    info = report(sys.argv[1], png, engine=eng)
    if "--breaks" in sys.argv:
        rows = line_breaks(info["pdf"])
        print("  รอยต่อบรรทัดทั้งหมด %d จุด — อ่านทีละจุดว่าสะดุดไหม:" % len(rows))
        for i, (pno, a, c) in enumerate(rows, 1):
            print("   %3d  หน้า %d  %16s | %s" % (i, pno, a[-16:], c[:16]))
