# -*- coding: utf-8 -*-
"""
จัดตัวหนังสือไทยในไฟล์ .docx — ส่วนที่ละเอียดกว่าตัวสร้างทั่วไป

รวมของ 4 อย่างที่ต้องทำทุกครั้งแต่ python-docx ไม่ได้ให้มา

1. ตั้งฟอนต์ให้ครบฝั่ง complex script (w:cs / w:szCs / w:bCs / w:iCs)
   python-docx ตั้งให้แค่ w:ascii กับ w:hAnsi ซึ่งอักษรไทยไม่ได้ใช้
   และต้องวาง element ตามลำดับ schema ไม่งั้น Word ทิ้ง property นั้นเงียบ ๆ
2. บีบระยะห่างอักขระ (w:spacing ค่าลบ) เพื่อดึงคำขึ้นบรรทัดบนโดยไม่เลิกจัดชิดขอบขวา
3. ตัดบรรทัดด้วย <w:br/> ในย่อหน้าเดิม ไม่ใช่ขึ้นย่อหน้าใหม่
4. ระยะบรรทัดขั้นต่ำ 1.0 — ต่ำกว่านี้วรรณยุกต์ไทยชนบรรทัดบน (วัดแล้ว 0.95 ชน 4 จุด)
"""
from __future__ import annotations

from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt

# ───────────────────────────────────────────────────────────── ค่าคงที่

# ขั้นการบีบที่กิ๊ฟใช้จริง — เกิน 0.7 pt ตาเริ่มจับได้ว่าถูกบีบ
CONDENSE_STEPS = (0.1, 0.2, 0.3, 0.5, 0.7)
# ช่องว่างเฉลี่ยระหว่างอักขระที่ยอมรับได้ วัดจาก PDF จริง
MAX_GAP_PT = 0.8
# ระยะบรรทัดต่ำสุดที่ปลอดภัยสำหรับ TH Sarabun (วัดแล้ว: 1.0 = pitch 18.14 pt ไม่ชน)
MIN_LINE_SPACING = 1.0

DEFAULT_FONT = "TH SarabunPSK"

# ลำดับลูกของ w:rPr ตาม schema — เรียงผิดแล้ว Word อาจไม่สนใจ property นั้น
_RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
    "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
    "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
    "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
    "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
]
_IDX = {qn(t): i for i, t in enumerate(_RPR_ORDER)}


def _set(rpr, tag: str, **attrs):
    """สร้าง/แก้ element ใน rPr แล้ววางในตำแหน่งที่ schema กำหนด"""
    el = rpr.find(qn(tag))
    if el is None:
        el = rpr.makeelement(qn(tag), {})
        want = _IDX.get(qn(tag), len(_RPR_ORDER))
        for child in rpr:
            if _IDX.get(child.tag, len(_RPR_ORDER)) > want:
                child.addprevious(el)
                break
        else:
            rpr.append(el)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


def _drop(rpr, tag: str):
    el = rpr.find(qn(tag))
    if el is not None:
        rpr.remove(el)


# ──────────────────────────────────────────────────────── ฟอนต์และภาษา

def effective_size(run, default: float = 16.0) -> float:
    """ขนาดที่ใช้จริงกับอักษรไทย — ยึด szCs ก่อน เพราะไทยคือ complex script"""
    rpr = run._element.rPr
    if rpr is not None:
        szcs = rpr.find(qn("w:szCs"))
        if szcs is not None:
            return int(szcs.get(qn("w:val"))) / 2
        if rpr.sz is not None:
            return int(rpr.sz.get(qn("w:val"))) / 2
    return default


def set_run_font(run, name: str = DEFAULT_FONT, size: float | None = None,
                 *, lang: str = "th-TH"):
    """
    ตั้งฟอนต์/ขนาด/ตัวหนา/ภาษา ให้ครบทั้งสองฝั่ง
    ถ้าไม่ตั้ง w:szCs อักษรไทยจะตกไปใช้ค่าปริยาย 11 pt บนเครื่องที่ไม่มี style เดียวกัน
    """
    rpr = run._element.get_or_add_rPr()
    pt = size if size is not None else effective_size(run)
    half = str(int(round(pt * 2)))

    _set(rpr, "w:rFonts", **{"w:ascii": name, "w:hAnsi": name, "w:cs": name})
    _set(rpr, "w:sz", **{"w:val": half})
    _set(rpr, "w:szCs", **{"w:val": half})
    if run.bold:
        _set(rpr, "w:bCs")
    else:
        _drop(rpr, "w:bCs")
    if run.italic:
        _set(rpr, "w:iCs")
    else:
        _drop(rpr, "w:iCs")
    if lang:
        _set(rpr, "w:lang", **{"w:bidi": lang})
    return run


def set_thai_language(doc, lang: str = "th-TH"):
    """
    ตั้งภาษา complex script เป็นไทยที่ docDefaults

    หมายเหตุสำคัญ: อันนี้ 'ไม่ได้' แก้การตัดบรรทัด — ทดลองแยกตัวแปรแล้วผลเท่าเดิมทุกประการ
    Word ไม่ได้ใช้พจนานุกรมไทยตัดบรรทัดให้ ต้องใส่ ZWSP เอง (ดู thai_break.py)
    ตั้งไว้เพราะเป็น metadata ที่ถูกต้อง และช่วยเรื่องตรวจคำผิด
    """
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is None:
        return
    rdd = dd.find(qn("w:rPrDefault"))
    if rdd is None:
        return
    rpr = rdd.find(qn("w:rPr"))
    if rpr is not None:
        _set(rpr, "w:lang", **{"w:bidi": lang})


def fix_doc_defaults(doc, *, line_240: bool = True, space_after_0: bool = True):
    """
    docDefaults ของ Word ตั้งระยะบรรทัด 1.15 (w:line="276") กับ space after 10 pt มาให้
    ซึ่งทำให้หน้ายืดโดยที่ย่อหน้าไม่ได้สั่งเอง — เคลียร์ทิ้งแล้วไปคุมรายย่อหน้าแทน
    """
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is None:
        return
    pdd = dd.find(qn("w:pPrDefault"))
    if pdd is None:
        return
    ppr = pdd.find(qn("w:pPr"))
    if ppr is None:
        return
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        return
    if space_after_0:
        sp.set(qn("w:after"), "0")
    if line_240:
        sp.set(qn("w:line"), "240")
        sp.set(qn("w:lineRule"), "auto")


# ──────────────────────────────────────────────────────── บีบระยะอักขระ

def set_char_spacing(run, condense_pt: float):
    """
    บีบ (ค่าบวก = บีบเข้า) หรือขยาย (ค่าลบ) ระยะห่างอักขระ
    ใน XML ใช้หน่วย 1/20 pt และ 'ค่าลบคือบีบ' — 0.7 pt จึงเป็น w:val="-14"
    """
    rpr = run._element.get_or_add_rPr()
    if not condense_pt:
        _drop(rpr, "w:spacing")
        return run
    _set(rpr, "w:spacing", **{"w:val": str(int(round(-condense_pt * 20)))})
    return run


def set_paragraph_condense(paragraph, condense_pt: float):
    for r in paragraph.runs:
        set_char_spacing(r, condense_pt)
    return paragraph


def pick_condense(tried: dict[float, float]) -> float:
    """
    เลือกขั้นการบีบที่ 'ถ่างน้อยที่สุด' จากผลที่วัดมาแล้ว
    tried = {ขั้นที่บีบ (pt): ช่องว่างเฉลี่ยที่วัดได้ (pt)}

    บทเรียน 16 ก.ย. 2569: บีบมากขึ้นไม่ได้ดีกว่าเสมอไป — บรรทัดบนดึงคำขึ้นไปเยอะ
    บรรทัดล่างที่รอก้อนคำยาวจะยิ่งสั้นและถ่างหนักกว่าเดิม จึงต้องวัดทุกขั้นแล้วเทียบ
    ไม่ใช่ไล่บีบจนสุด 0.7
    """
    if not tried:
        return 0.0
    best = min(tried.items(), key=lambda kv: (kv[1], kv[0]))
    return best[0]


# ───────────────────────────────────────────────────────────── ย่อหน้า

def line_break(paragraph):
    """
    ตัดบรรทัดในย่อหน้าเดิม (เท่ากับที่กิ๊ฟกด Alt+Enter)
    ถ้าเผลอขึ้นย่อหน้าใหม่แทน จะได้ space_before/after กับย่อหน้าบรรทัดแรกติดมาด้วย
    บรรทัดที่ตัดจะเยื้องผิดและกินที่เพิ่ม
    """
    run = paragraph.add_run()
    run._element.append(run._element.makeelement(qn("w:br"), {}))
    return run


def set_line_spacing(paragraph, value: float = 1.0):
    """ระยะบรรทัด — กันไม่ให้ตั้งต่ำกว่า 1.0 เพราะวรรณยุกต์จะชนบรรทัดบน"""
    if value < MIN_LINE_SPACING:
        raise ValueError(
            "ระยะบรรทัด %.2f ต่ำกว่า %.2f — วัดแล้ววรรณยุกต์ไทยชนบรรทัดบน "
            "ถ้าต้องการที่เพิ่ม ให้ลด space_after หรือลบย่อหน้าว่างแทน"
            % (value, MIN_LINE_SPACING))
    pf = paragraph.paragraph_format
    if value == 1.0:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing = value
    return paragraph


def normalise_document(doc, *, font: str = DEFAULT_FONT, lang: str = "th-TH"):
    """ตั้งฟอนต์ให้ครบทุก run ทั้งเอกสาร รวมในตาราง แล้วเคลียร์ docDefaults ที่ทำให้หน้ายืด"""
    fix_doc_defaults(doc)
    set_thai_language(doc, lang)
    n = 0
    for p in doc.paragraphs:
        for r in p.runs:
            set_run_font(r, font, lang=lang)
            n += 1
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        set_run_font(r, font, lang=lang)
                        n += 1
    return n
