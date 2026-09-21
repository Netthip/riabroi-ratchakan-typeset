# -*- coding: utf-8 -*-
"""
ตัดคำไทยสำหรับเอกสารราชการ — ละเอียดกว่าตัวตัดคำมาตรฐาน

ต่างจาก thai-docx (ตัวสาธารณะ) 5 เรื่อง ซึ่งเป็นบทเรียนจากงานจริง

1. ตัดคำจาก "ข้อความทั้งย่อหน้า" แล้วค่อยแมปกลับเข้า run
   เอกสารจาก Word มักมี run ที่ผ่ากลางคำ (เจอจริง: 'ประชาชนไม่โด' + 'นทอดทิ้ง')
   ถ้าตัดทีละ run จะได้ขอบเขตคำผิดตรงรอยต่อ
2. NO_BREAK_BEFORE ไม่รวมสระหน้า เ แ โ ใ ไ
   สระหน้าเขียนนำหน้าพยัญชนะและเป็นตัวแรกของคำอยู่แล้ว ตัดบรรทัดหน้ามันได้
   ถ้ารวมไว้ คำอย่าง และ/ได้/เป็น/ให้ จะเกาะคำก่อนหน้าเป็นก้อนยาว 25-30 ตัว
3. min_token = 2 (ไม่ใช่ 4) สำหรับงานที่จัดชิดขอบขวา — ให้จุดตัดถี่พอจะไม่ถ่าง
4. split_long_chunks ซอยก้อนที่ยาวเกิน 12 ตัว
   'งบประมาณรายจ่าย' เป็นคำเดียวในพจนานุกรม พอไปอยู่ต้นบรรทัดจะดันบรรทัดก่อนถ่าง 3-4 ซม.
5. KEEP_TOGETHER / LEAD_WORDS — วลีที่ห้ามขาด และคำนำวลีที่ห้ามค้างท้ายบรรทัด

หมายเหตุสำหรับคนแก้ไฟล์นี้: อย่าพิมพ์อักขระ ZWSP ลงซอร์สตรง ๆ และอย่าเขียน escape ยูนิโค้ด
ผ่านเครื่องมือแก้ไฟล์ เพราะจะกลายเป็นตัวล่องหนที่มองไม่เห็นใน diff — ใช้ chr(0x200B) เสมอ
"""
from __future__ import annotations

import re
import unicodedata

ZWSP = chr(0x200B)
NBSP = chr(0x00A0)

# ─────────────────────────────────────────────────────────── ตัวอักษรไทย

THAI_CHAR = re.compile(r"[฀-๿]")

# อักขระที่ห้ามขึ้นต้นบรรทัด = สระบน/ล่าง/หลัง วรรณยุกต์ ไม้ยมก ฯลฯ
# เจตนา: ไม่มีสระหน้า (เ แ โ ใ ไ) อยู่ในชุดนี้ ดูเหตุผลข้อ 2 ด้านบน
NO_BREAK_BEFORE = set("ๆฯะัาำิีึืุูๅ็่้๊๋์ํฺ")

# วลีที่ห้าม ZWSP ไปตัดกลาง — ตัวตัดคำมองเป็นคนละคำ แต่ในหนังสือราชการต้องอยู่ด้วยกัน
# ใส่เฉพาะที่กิ๊ฟสั่งจริง ห้ามเติมวลียาวเอง: ยิ่งมัดยาว จุดตัดยิ่งน้อย Word จะยืดอักษรจนผิดปกติ
# (เคยใส่ 'แผนการปฏิบัติงาน' แล้วช่องไฟพุ่งเป็น 7 pt)
KEEP_TOGETHER = (
    "ประจำปีงบประมาณ", "ตามที่ขอทำความตกลงไปได้",
    "ความละเอียด", "สงป.",
    "แผนงาน", "โครงการ",
    "ตัวชี้วัด", "ก่อหนี้", "แนวทาง", "รวมทั้งสิ้น",
)

# คำนำวลีที่ห้ามค้างท้ายบรรทัด ต้องลงไปขึ้นบรรทัดใหม่พร้อมคำที่ตามมา
# หลักของกิ๊ฟไม่ใช่รายการคำบังคับ แต่คือ "คำที่ค้างบรรทัดบนเดี่ยว ๆ แล้วอ่านสะดุด"
# เจอคำใหม่ที่อ่านแล้วสะดุดแบบเดียวกันเพิ่มได้ ถ้าไม่ทำให้ความหมายเปลี่ยน
LEAD_WORDS = ("เรื่อง", "ให้", "เมื่อ", "ที่เป็น", "ตาม", "มี", "ต่อ",
              "วงเงิน", "แผนการ", "สถาบัน", "เพื่อ")
# คำนำที่เป็นคำนาม เป็นกรรมของคำนำตัวหน้าได้ จึงติดกันเป็นทอด (มี/วงเงิน · ให้/สถาบัน)
LEAD_NOUNS = ("วงเงิน", "แผนการ", "สถาบัน")
# ติดคำถัดไปเฉพาะเมื่อตามด้วย LEAD_NOUNS — '…๑ โครงการ รวม / วงเงิน ๕๓,๓๙๓,๔๐๐ บาท'
LEAD_BEFORE_NOUNS = ("รวม",)
# ห้ามขึ้นต้นบรรทัด ต้องติดคำข้างหน้า — 'สถาบันวัคซีน / แห่งชาติ'
TAIL_WORDS = ("แห่งชาติ",)

# คำที่กิ๊ฟเองปล่อยให้ค้างท้ายบรรทัดได้ ไม่ต้องมัด
ALLOW_DANGLING = ("และ", "หรือ", "ของ", "ใน", "ตั้งแต่")

MAX_CHUNK = 12          # ก้อนยาวกว่านี้ให้หาจุดตัดย่อยเพิ่ม
MIN_TOKEN = 2           # ชิ้นสั้นกว่านี้เกาะคำก่อนหน้า (2 สำหรับงานจัดชิดขอบขวา)

_TOKENIZER = None
_WORDS = None

# คำที่มักถูกตัวตัดคำผลักไปเกาะคำถัดไปทั้งที่เป็นหางของคำหน้า
# 'ดำเนิน|การตาม' ต้องเป็น 'ดำเนินการ|ตาม' — เช็กกับพจนานุกรมก่อนย้ายเสมอ
NOMINAL_PREFIX = ("การ", "ความ")


def _tokenize(text: str) -> list[str]:
    """ตัดคำด้วย pythainlp ถ้ามี ไม่มีก็คืนทั้งก้อน (ไม่ทำให้พัง แค่ไม่ได้จุดตัด)"""
    global _TOKENIZER
    if _TOKENIZER is None:
        try:
            from pythainlp.tokenize import word_tokenize
            _TOKENIZER = word_tokenize
        except ImportError:
            _TOKENIZER = False
    if not _TOKENIZER:
        return [text]
    return _TOKENIZER(text, engine="newmm")


def _words() -> set:
    """พจนานุกรมไว้ยืนยันว่าจุดที่จะซอย/ย้าย ได้คำจริงทั้งสองฝั่ง"""
    global _WORDS
    if _WORDS is None:
        try:
            from pythainlp.corpus import thai_words
            _WORDS = set(thai_words())
        except ImportError:
            _WORDS = set()
    return _WORDS


def is_thai(ch: str) -> bool:
    return bool(THAI_CHAR.match(ch))


# ─────────────────────────────────────────────────────── ซอยก้อนที่ยาวเกิน

def _best_split(tok: str) -> int:
    """
    หาจุดซอยที่ดีที่สุดในคำยาว: เอาจุดที่ 'ทั้งสองฝั่งเป็นคำจริง' และใกล้กลางที่สุด
    'งบประมาณรายจ่าย' -> งบประมาณ|รายจ่าย ไม่ใช่ งบประมาณราย|จ่าย
    ถ้าไม่มีจุดไหนได้คำจริงทั้งคู่ คืน 0 (แปลว่าอย่าซอย ปล่อยยาวดีกว่าซอยมั่ว)
    """
    words = _words()
    n = len(tok)
    mid = n // 2
    # ไล่จากกลางออกไปสองข้าง เพื่อให้ได้สองฝั่งที่สมดุล
    for off in range(0, n):
        for cut in {mid - off, mid + off}:
            if not (2 <= cut <= n - 2):
                continue
            if tok[cut] in NO_BREAK_BEFORE or unicodedata.combining(tok[cut]):
                continue
            if tok[:cut] in words and tok[cut:] in words:
                return cut
    return 0


def split_long_chunks(tokens: list[str], max_chars: int = MAX_CHUNK) -> list[str]:
    """
    ซอยคำประสมยาว ๆ ที่พจนานุกรมมองเป็นคำเดียว
    'งบประมาณรายจ่าย' (15 ตัว) -> 'งบประมาณ' + 'รายจ่าย'

    ซอยเฉพาะจุดที่ได้คำจริงทั้งสองฝั่ง — ซอยมั่วกลางคำแย่กว่าปล่อยให้ยาว
    และไม่แตะวลีใน KEEP_TOGETHER ที่ตั้งใจมัดไว้
    """
    out: list[str] = []
    for tok in tokens:
        if any(p in tok for p in KEEP_TOGETHER):
            out.append(tok)
            continue
        while len(tok) > max_chars:
            cut = _best_split(tok)
            if not cut:
                break
            out.append(tok[:cut])
            tok = tok[cut:]
        out.append(tok)
    return out


def _fix_prefix_split(tokens: list[str]) -> list[str]:
    """
    ย้าย 'การ'/'ความ' ที่ตัวตัดคำผลักไปเกาะคำถัดไป กลับมาเป็นหางของคำหน้า
    'ดำเนิน' + 'การตาม'  ->  'ดำเนินการ' + 'ตาม'   (เพราะ 'ดำเนินการ' เป็นคำจริง)
    ย้ายต่อเมื่อพจนานุกรมยืนยันว่าคำหน้า+คำนำหน้านั้นเป็นคำจริง ไม่งั้นปล่อยไว้
    """
    words = _words()
    out = list(tokens)
    for i in range(len(out) - 1):
        nxt = out[i + 1]
        for pre in NOMINAL_PREFIX:
            if nxt.startswith(pre) and len(nxt) > len(pre):
                if (out[i] + pre) in words:
                    out[i] = out[i] + pre
                    out[i + 1] = nxt[len(pre):]
                break
    return [t for t in out if t]


# ─────────────────────────────────────────── รวมชิ้นสั้นและวลีที่ห้ามขาด

def _merge_small(tokens: list[str], min_token: int) -> list[str]:
    """ชิ้นที่สั้นกว่า min_token ให้เกาะคำก่อนหน้า กันตัดกลางคำประสม"""
    out: list[str] = []
    for tok in tokens:
        if out and len(tok.strip()) < min_token and is_thai(tok[:1] or " "):
            out[-1] += tok
        else:
            out.append(tok)
    return out


# 'จำนวน <เลข> <หน่วย>' และ 'พ.ศ. <ปี>' — เลขกับหน่วยต้องไม่ขาดจากกัน
_AUTO_BIND = (
    re.compile(r"พ\.ศ\.\s*[0-9๐-๙]{2,4}"),
    re.compile(r"จำนวน\s*[0-9๐-๙,\.]+\s*\S{0,8}"),
)


def _protected_spans(text: str) -> list[tuple[int, int]]:
    """
    ช่วงตัวอักษรที่ห้ามมีจุดตัดอยู่ข้างใน

    ทำที่ระดับตัวอักษร ไม่ใช่ระดับ token เพราะขอบเขต token ไม่ตรงกับวลีเสมอ
    (ตัวตัดคำคืน 'ตามตัว|ชี้|วัด' วลี 'ตัวชี้วัด' จึงคร่อมกลาง token แรก)
    """
    spans = []
    for phrase in KEEP_TOGETHER:
        start = 0
        while True:
            k = text.find(phrase, start)
            if k < 0:
                break
            spans.append((k, k + len(phrase)))
            start = k + 1
    for rx in _AUTO_BIND:
        for m in rx.finditer(text):
            spans.append(m.span())
    return spans


def _glue_leads(tokens: list[str]) -> list[str]:
    """
    มัดคำนำวลีกับคำที่ตามมา เพื่อไม่ให้ค้างท้ายบรรทัดเดี่ยว ๆ
    ตัดบรรทัด 'ก่อน' คำพวกนี้ได้ตามปกติ ห้ามตัด 'หลัง'
    """
    def needs_glue(cur: str, nxt: str) -> bool:
        cur, nxt = cur.strip(), nxt.strip()
        if not nxt:
            return False
        if nxt in TAIL_WORDS:                       # 'แห่งชาติ' ห้ามขึ้นต้นบรรทัด
            return True
        if cur in LEAD_WORDS and cur not in ALLOW_DANGLING:
            return True
        if cur in LEAD_BEFORE_NOUNS and nxt in LEAD_NOUNS:
            return True
        return False

    out: list[str] = []
    i = 0
    while i < len(tokens):
        cur = tokens[i]
        # ดูดคำถัดไปเป็นทอด ๆ ('มี'+'วงเงิน'+... ) และที่สำคัญคือ หลังดูดแล้ว
        # ต้องตรวจคำถัดไปอีกที ไม่งั้น 'สถาบัน'+'วัคซีน' จะกระโดดข้ามการตรวจ 'แห่งชาติ'
        while i + 1 < len(tokens) and needs_glue(cur, tokens[i + 1]):
            cur += tokens[i + 1]
            i += 1
        out.append(cur)
        i += 1
    return out


# ─────────────────────────────────────────────────── หาตำแหน่งที่ตัดได้

def break_offsets(text: str, *, min_token: int = MIN_TOKEN,
                  max_chunk: int = MAX_CHUNK) -> set[int]:
    """
    คืนตำแหน่ง (ดัชนีตัวอักษร) ที่ใส่ ZWSP ได้ โดยผ่านกฎทั้งชุดแล้ว
    ใส่เฉพาะจุดที่ 'ไทยชนไทย' จะได้ไม่ไปแหกคำอังกฤษ ตัวเลข หรือเลขหมายโทรศัพท์
    """
    if not text or not THAI_CHAR.search(text):
        return set()

    toks = _tokenize(text)
    toks = _fix_prefix_split(toks)      # แก้ที่ตัวตัดคำผลัก การ/ความ ไปผิดฝั่ง
    toks = split_long_chunks(toks, max_chunk)
    toks = _merge_small(toks, min_token)
    toks = _glue_leads(toks)

    cuts, pos = set(), 0
    for tok in toks:
        pos += len(tok)
        cuts.add(pos)
    cuts.discard(0)
    cuts.discard(len(text))

    spans = _protected_spans(text)

    ok = set()
    for c in cuts:
        if not (0 < c < len(text)):
            continue
        prev, nxt = text[c - 1], text[c]
        if not (is_thai(prev) and is_thai(nxt)):
            continue            # ไทยชนไทยเท่านั้น จะได้ไม่แหกคำอังกฤษ/ตัวเลข
        if nxt in NO_BREAK_BEFORE or unicodedata.combining(nxt):
            continue            # ห้ามขึ้นบรรทัดด้วยสระบน/ล่าง/วรรณยุกต์
        if any(s < c < e for s, e in spans):
            continue            # อยู่กลางวลีที่ห้ามขาด
        ok.add(c)
    return ok


def insert_zwsp(text: str, *, min_token: int = MIN_TOKEN,
                max_chunk: int = MAX_CHUNK) -> str:
    """ใส่ ZWSP ลงข้อความเดี่ยว ๆ (ถ้ามี run ให้ใช้ insert_zwsp_paragraph แทน)"""
    cuts = break_offsets(text, min_token=min_token, max_chunk=max_chunk)
    if not cuts:
        return text
    out, prev = [], 0
    for c in sorted(cuts):
        out.append(text[prev:c])
        prev = c
    out.append(text[prev:])
    return ZWSP.join(out)


# ──────────────────────────────────────── ระดับย่อหน้า (ของจริงที่ต้องใช้)

def insert_zwsp_paragraph(paragraph, *, min_token: int = MIN_TOKEN,
                          max_chunk: int = MAX_CHUNK) -> int:
    """
    ใส่ ZWSP ให้ย่อหน้าของ python-docx โดยตัดคำจากข้อความรวมทั้งย่อหน้า
    แล้วแมปจุดตัดกลับเข้าแต่ละ run — นี่คือจุดต่างสำคัญจาก thai-docx

    คืนจำนวนจุดที่ใส่
    """
    runs = [r for r in paragraph.runs if r.text]
    if not runs:
        return 0

    texts = [r.text.replace(ZWSP, "") for r in runs]
    full = "".join(texts)
    cuts = break_offsets(full, min_token=min_token, max_chunk=max_chunk)
    if not cuts:
        for r, t in zip(runs, texts):
            r.text = t
        return 0

    added, base = 0, 0
    for run, txt in zip(runs, texts):
        end = base + len(txt)
        local = sorted(c - base for c in cuts if base < c < end)
        if local:
            parts, prev = [], 0
            for c in local:
                parts.append(txt[prev:c])
                prev = c
            parts.append(txt[prev:])
            txt = ZWSP.join(parts)
            added += len(local)
        # จุดตัดที่ตรงกับรอยต่อ run พอดี ต้องใส่ไว้ต้น run ถัดไป
        # (Word ไม่ถือว่ารอยต่อ run เป็นจุดตัดบรรทัดโดยอัตโนมัติ)
        if base in cuts and base > 0:
            txt = ZWSP + txt
            added += 1
        run.text = txt
        base = end
    return added


def insert_zwsp_document(doc, **kw) -> int:
    """ใส่ ZWSP ทั้งเอกสาร รวมย่อหน้าในตาราง"""
    n = 0
    for p in doc.paragraphs:
        n += insert_zwsp_paragraph(p, **kw)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    n += insert_zwsp_paragraph(p, **kw)
    return n


def strip_zwsp(text: str) -> str:
    """เอา ZWSP ออก (ใช้ตอนเทียบข้อความว่าไม่เปลี่ยน หรือตอนกิ๊ฟอยากได้ไฟล์สะอาด)"""
    return text.replace(ZWSP, "")


# ─────────────────────────────────────────────────────────── ตรวจผลจริง

def find_midword_breaks(line_pairs) -> list[tuple[str, str]]:
    """
    รับคู่ (ท้ายบรรทัดบน, ต้นบรรทัดล่าง) ที่อ่านจาก PDF จริง
    คืนคู่ที่ตัดกลางคำ — ใช้ยืนยันว่าจัดหน้าแล้วได้ผลจริง ไม่ใช่เชื่อว่าโค้ดถูก

    ข้ามคู่ที่จุดตัดเป็นช่องว่างจริงในต้นฉบับ ไม่งั้นจะได้ผลบวกลวง
    """
    bad = []
    for tail, head in line_pairs:
        tail, head = tail.rstrip(), head.lstrip()
        if not tail or not head:
            continue
        if not (is_thai(tail[-1]) and is_thai(head[0])):
            continue
        joined = tail + head
        pos, cuts = 0, set()
        for tok in _tokenize(joined):
            pos += len(tok)
            cuts.add(pos)
        if len(tail) not in cuts:
            bad.append((tail[-14:], head[:14]))
    return bad
