# -*- coding: utf-8 -*-
"""
ตรวจเครื่องก่อนใช้ riabroi — ใช้ได้ทั้ง Windows และ macOS

    python scripts/doctor.py          ตรวจ Python ไลบรารี ตัวตัดคำ ฟอนต์ และ Word
    python scripts/doctor.py --pdf    ลองส่งออก PDF จริง 1 หน้าแล้ววัดผล
    (macOS พิมพ์ python3 แทน python)

บน Mac ให้รัน --pdf ครั้งแรกตอนนั่งอยู่หน้าเครื่อง จะมีกล่องขออนุญาต 2 กล่อง
  1. macOS ถามว่าให้แอปที่รันคำสั่งนี้ควบคุม Microsoft Word ได้ไหม -> กด OK
  2. Word ถามสิทธิ์โฟลเดอร์ ~/riabroi-pdf (Grant File Access) -> Select... แล้ว Grant Access
กดครั้งเดียว ครั้งต่อไปไม่ถามอีก
"""
from __future__ import annotations

import os
import platform
import re
import shutil
import sys
import tempfile
import warnings

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import inspect_docx as ix  # noqa: E402

OK, WARN, MISSING = "ผ่าน", "เตือน", "ขาด"
SAMPLE = ("สถาบันวัคซีนแห่งชาติได้รับงบประมาณรายจ่ายประจำปีงบประมาณ พ.ศ. 2570 "
          "เพื่อดำเนินการตามแผนงานโครงการ ขอเรียนชี้แจงว่าการดำเนินการตามตัวชี้วัด"
          "และการก่อหนี้ผูกพันเป็นไปตามแนวทางที่สำนักงบประมาณกำหนด "
          "มีวงเงินงบประมาณรวมทั้งสิ้น 50,038,500 บาท ตามที่ขอทำความตกลงไปได้ "
          "ทั้งนี้ หน่วยรับงบประมาณต้องรายงานผลการดำเนินงานและผลการใช้จ่ายงบประมาณ"
          "ให้สำนักงบประมาณทราบทุกไตรมาส เพื่อประกอบการพิจารณาจัดสรรงบประมาณในงวดถัดไป")
_results: list[str] = []


def say(status: str, topic: str, detail: str = "") -> None:
    _results.append(status)
    print("  [%s] %s  %s" % (status, topic, detail))


def _has(module: str) -> bool:
    try:
        __import__(module)
        return True
    except ImportError:
        return False


# ──────────────────────────────────────────────────────────── ตรวจทีละเรื่อง

def check_system() -> None:
    if sys.platform == "darwin":
        name = "macOS " + platform.mac_ver()[0]
    elif sys.platform == "win32":
        name = "Windows " + platform.release()
    else:
        name = platform.system() + " " + platform.release()
    good = sys.version_info >= (3, 9)
    say(OK if good else MISSING, "ระบบ",
        "%s · Python %s%s" % (name, platform.python_version(), "" if good else " (ต้อง 3.9 ขึ้นไป)"))


def check_libraries() -> None:
    need = [("docx", "python-docx"), ("pythainlp", "pythainlp"), ("fitz", "PyMuPDF")]
    if sys.platform == "win32":
        need.append(("pythoncom", "pywin32"))
    for module, package in need:
        try:
            m = __import__(module)
        except ImportError:
            say(MISSING, package, "ติดตั้งด้วย: pip install -r requirements.txt")
            continue
        say(OK, package, str(getattr(m, "__version__", "") or getattr(m, "VersionBind", "")))


def check_segmenter() -> None:
    if not _has("pythainlp"):
        return                              # แจ้งไปแล้วในหัวข้อไลบรารี
    import thai_break
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        out = thai_break.insert_zwsp(SAMPLE)
    if thai_break.ZWSP in out:
        say(OK, "ตัดคำ", "ได้จุดตัด %d จุดจากประโยคทดสอบ" % out.count(thai_break.ZWSP))
    else:
        say(MISSING, "ตัดคำ", str(caught[0].message) if caught else "ไม่ได้จุดตัดเลย")


def font_dirs() -> list[str]:
    if sys.platform == "win32":
        return [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
                os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]
    if sys.platform == "darwin":
        return [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts",
                "/System/Library/Fonts", "/System/Library/Fonts/Supplemental",
                "/Applications/Microsoft Word.app/Contents/Resources/DFonts"]
    return [os.path.expanduser("~/.fonts"), os.path.expanduser("~/.local/share/fonts"),
            "/usr/share/fonts", "/usr/local/share/fonts"]


def font_names(data: bytes) -> list[tuple[int, int, int, str]]:
    """(platform, ภาษา, name ID, ข้อความ) ของชื่อ ID 1 และ 16 จากตาราง name ในไฟล์ฟอนต์"""
    import struct
    try:
        base = struct.unpack(">I", data[12:16])[0] if data[:4] == b"ttcf" else 0
        for i in range(struct.unpack(">H", data[base + 4:base + 6])[0]):
            tag, _, off, _ = struct.unpack(">4sIII", data[base + 12 + 16 * i:base + 28 + 16 * i])
            if tag != b"name":
                continue
            count, strings = struct.unpack(">HH", data[off + 2:off + 6])
            out = []
            for j in range(count):
                pid, _, lang, nid, size, at = struct.unpack(
                    ">HHHHHH", data[off + 6 + 12 * j:off + 18 + 12 * j])
                if nid in (1, 16) and pid in (0, 1, 3):
                    raw = data[off + strings + at:off + strings + at + size]
                    text = raw.decode("mac_roman" if pid == 1 else "utf-16-be", "replace")
                    out.append((pid, lang, nid, text.strip()))
            return out
    except (struct.error, IndexError):
        pass
    return []


def font_family(data: bytes, *, mac: bool = False) -> str | None:
    """
    ชื่อตระกูลของฟอนต์
      mac=False  ชื่อที่ Word บน Windows ใช้ = name ID 1 ฝั่ง Windows
      mac=True   ชื่อที่ macOS ใช้จัดกลุ่ม = ID 16 ถ้ามี ไม่มีค่อยใช้ ID 1

    ห้ามใช้ fitz.Font(...).name — ไฟล์ TH SarabunIT๙ เขียนชื่อเต็มไว้ว่า "TH SarabunPSK"
    และไฟล์เดียวกันใส่ ID 16 กับชื่อฝั่ง Mac ไว้ว่า "TH SarabunPSK" ทั้งที่ ID 1 ฝั่ง Windows คือ
    "TH SarabunIT๙" — บน Windows จึงเห็นเป็นคนละฟอนต์ แต่บน Mac อาจถูกรวมเป็นตระกูลเดียวกับ PSK
    """
    recs = font_names(data)
    for nid in ((16, 1) if mac else (1,)):
        for pid in (3, 0, 1):
            hits = [(lang != 0x409, t) for p, lang, n, t in recs if n == nid and p == pid and t]
            if hits:
                return min(hits)[1]
    return None


def installed_sarabun() -> dict[str, str]:
    """ฟอนต์ตระกูล Sarabun ที่ติดตั้งอยู่ {ชื่อที่ Word บน Windows เห็น: ชื่อที่ Mac ใช้จัดกลุ่ม}"""
    found: dict[str, str] = {}
    for top in font_dirs():
        for root, _, files in os.walk(top):
            for f in files:
                if "sarabun" not in f.lower() or not f.lower().endswith((".ttf", ".otf", ".ttc")):
                    continue
                try:
                    with open(os.path.join(root, f), "rb") as fh:
                        data = fh.read()
                except OSError:
                    continue
                plain = os.path.splitext(f)[0]
                win = font_family(data) or plain
                found.setdefault(win, font_family(data, mac=True) or win)
    return found


def check_fonts() -> None:
    try:
        from thai_fit import DEFAULT_FONT
    except ImportError:                     # ยังไม่มี python-docx — ใช้ค่าเดียวกับ thai_fit
        DEFAULT_FONT = "TH SarabunPSK"
    found = installed_sarabun()
    if any(ix._norm_font(f) == ix._norm_font(DEFAULT_FONT) for f in found):
        say(OK, "ฟอนต์", ", ".join(sorted(found)))
    elif found:
        say(WARN, "ฟอนต์", "มี %s แต่ไม่มี %s ที่เอกสารของกิ๊ฟใช้ — Word จะใช้ฟอนต์อื่นแทน"
            % (", ".join(sorted(found)), DEFAULT_FONT))
    else:
        where = ("ดับเบิลคลิกไฟล์ .ttf แล้วกด Install Font" if sys.platform == "darwin"
                 else "คลิกขวาที่ไฟล์ .ttf แล้วเลือก Install")
        say(MISSING, "ฟอนต์", "ไม่พบ TH Sarabun — ติดตั้ง %s ก่อน (%s)" % (DEFAULT_FONT, where))
    # ชื่อฝั่ง Mac ไม่ตรงกับชื่อจริง — เจอกับ TH SarabunIT๙ รุ่น Windows ที่บอก Mac ว่าตัวเองคือ PSK
    # IT๙ วาดเลขอารบิกเป็นรูปเลขไทยและกว้างกว่า ถ้า Mac หยิบมาแทน PSK ตัวเลขทั้งเอกสารจะเปลี่ยนหน้าตา
    clash = sorted("%s (Mac เห็นเป็น %s)" % (f, mac) for f, mac in found.items()
                   if ix._norm_font(mac) != ix._norm_font(f))
    if clash and sys.platform == "darwin":
        say(WARN, "ชื่อฟอนต์ซ้อน", "; ".join(clash) + " — Mac อาจรวมเป็นตระกูลเดียวกัน "
            "ถ้าไม่ได้ใช้ให้ถอดออกจาก Font Book แล้วตรวจ PDF ด้วย --pdf")


def _word_version(word: str) -> str:
    if sys.platform == "darwin":
        import plistlib
        try:
            with open(os.path.join(word, "Contents", "Info.plist"), "rb") as f:
                return "Microsoft Word %s (%s)" % (plistlib.load(f).get("CFBundleShortVersionString", "?"), word)
        except OSError:
            return word
    if sys.platform == "win32":
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as h:
                return winreg.QueryValue(h, None)
        except OSError:
            return word
    return word


def check_exporters() -> None:
    word = ix.find_word()
    office = ix.find_libreoffice()
    if word:
        say(OK, "Word", _word_version(word))
        if sys.platform == "darwin":
            say(OK, "โฟลเดอร์ทำงาน", ix.mac_workdir() + " (Word ถามสิทธิ์แค่ครั้งแรก)")
    elif office:
        say(WARN, "Word", "ไม่พบ Microsoft Word — จะใช้ LibreOffice แทน แต่ตัดบรรทัดไม่เหมือน Word")
    else:
        say(MISSING, "Word", "ไม่พบ Microsoft Word หรือ LibreOffice — ส่งออก PDF เพื่อตรวจผลไม่ได้")
    if office:
        say(OK, "LibreOffice", office + ("" if word else " (ใช้ดูภาพรวมเท่านั้น)"))


def pdf_test() -> None:
    """สร้างเอกสารทดสอบในโฟลเดอร์ชั่วคราว ส่งออก PDF จริง แล้ววัด"""
    if not all(_has(m) for m in ("docx", "fitz", "pythainlp")):
        say(MISSING, "ส่งออก PDF", "ติดตั้งไลบรารีให้ครบก่อน")
        return
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import thai_break
    import thai_fit

    tmp = tempfile.mkdtemp(prefix="riabroi_doctor_")
    try:
        src = os.path.join(tmp, "riabroi-doctor.docx")
        doc = Document()
        para = doc.add_paragraph(SAMPLE)
        para.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
        thai_fit.normalise_document(doc)
        thai_break.insert_zwsp_document(doc)
        doc.save(src)
        try:
            pdf = ix.to_pdf(src, tmp)
        except Exception as e:              # รายงานแทนการล้ม — ข้อความมีวิธีแก้อยู่แล้ว
            say(MISSING, "ส่งออก PDF", str(e))
            return
        made_by = ix.pdf_engine(pdf)
        say(OK if made_by in ("word", "word-mac") else WARN, "ส่งออก PDF",
            ix.ENGINE_LABEL.get(made_by, made_by))
        fonts = ix.font_check(src, pdf)
        if fonts["main_missing"]:
            say(MISSING, "ฟอนต์ใน PDF", "%s ถูกแทนด้วย %s — ติดตั้งฟอนต์แล้วรันใหม่"
                % (fonts["main"], ", ".join(fonts["embedded"]) or "-"))
        else:
            say(OK, "ฟอนต์ใน PDF", "%s ฝังอยู่ใน PDF" % fonts["main"])
        got = breaks_where_allowed(para.text, pdf)
        if got is None:
            say(WARN, "จุดตัดบรรทัด", "เทียบข้อความใน PDF กับต้นฉบับไม่ได้ — ดูภาพ PDF เอง")
        else:
            ok, total = got
            say(OK if ok == total else MISSING, "จุดตัดบรรทัด",
                "Word ตัดตรงจุดที่อนุญาต %d / %d บรรทัด%s" % (
                    ok, total, "" if ok == total else " — Word บนเครื่องนี้ไม่ตัดตาม ZWSP"))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _visual_lines(pdf_path: str) -> list[str]:
    """ข้อความทีละบรรทัดตามที่เห็นบนหน้า (รวมชิ้นที่อยู่แนว y เดียวกัน)"""
    import fitz
    out = []
    with fitz.open(pdf_path) as d:
        for page in d:
            rows: dict[int, list] = {}
            for b in page.get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    y = round(ln["bbox"][3])
                    key = next((k for k in rows if abs(k - y) <= 2), y)
                    rows.setdefault(key, []).append(ln)
            for y in sorted(rows):
                parts = sorted(rows[y], key=lambda ln: ln["bbox"][0])
                out.append("".join(ix._line_text(ln) for ln in parts))
    return out


def breaks_where_allowed(source: str, pdf_path: str):
    """
    Word บนเครื่องนี้ตัดบรรทัดตรงจุดที่อนุญาตไหม (ZWSP หรือช่องว่างในต้นฉบับ — NBSP ไม่นับ)
    คืน (ตัดตรงจุด, บรรทัดที่ห่อทั้งหมด) หรือ None ถ้าจับข้อความ PDF กับต้นฉบับไม่ติด

    ใช้ตัวนี้แทนตัวตรวจตัดกลางคำตอนตรวจเครื่อง เพราะตัวนั้นเดาจากพจนานุกรมและมีผลบวกลวง
    (รอยตัด "ดำเนินการ | ตาม" ที่ถูกต้อง ถูกอ่านเป็น ดำเนิน|การตาม)
    ส่วนตัวนี้วัดตรง ๆ ว่า Word เคารพ ZWSP — สิ่งที่ยังไม่รู้เมื่อย้ายไป Word for Mac
    """
    import thai_break
    zw, nb = thai_break.ZWSP, thai_break.NBSP
    plain, cuts = [], set()
    for ch in source:
        if ch == zw:
            cuts.add(len(plain))
        else:
            plain.append(ch)
    text = "".join(plain)
    spaces = {i for i, c in enumerate(text) if c == " "}
    flat = text.replace(nb, " ")
    lines = [ln.replace(zw, "").replace(nb, " ").strip() for ln in _visual_lines(pdf_path)]
    lines = [ln for ln in lines if ln]
    pos = ok = 0
    for i, line in enumerate(lines):
        while pos < len(flat) and flat[pos] == " ":
            pos += 1
        if not flat.startswith(line, pos):
            return None
        pos += len(line)
        if i < len(lines) - 1:
            ok += pos in cuts or pos in spaces or (pos - 1) in spaces
    return ok, len(lines) - 1


def main() -> int:
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    print("riabroi — ตรวจเครื่อง")
    check_system()
    check_libraries()
    check_segmenter()
    check_fonts()
    check_exporters()
    if "--pdf" in sys.argv:
        pdf_test()
    else:
        print("  (ยังไม่ได้ลองส่งออก PDF จริง — เพิ่ม --pdf)")
    missing, warned = _results.count(MISSING), _results.count(WARN)
    print()
    if not missing and not warned:
        print("พร้อมใช้")
    else:
        print("ขาด %d · เตือน %d — แก้ตามข้อความด้านบนแล้วรันใหม่" % (missing, warned))
    return 1 if missing else 0


if __name__ == "__main__":
    raise SystemExit(main())
