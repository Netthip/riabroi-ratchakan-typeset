# -*- coding: utf-8 -*-
"""
เทสต์ตัวตัดคำ — ทุกข้อมาจากจุดที่เคยพังจริงในเอกสารที่ส่งไปแล้ว
รัน: python -m pytest tests -q   หรือ   python tests/test_thai_break.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "scripts"))
import thai_break as tb  # noqa: E402

Z = tb.ZWSP


def cut(text):
    """ใส่ ZWSP แล้วแทนด้วย | เพื่ออ่านง่ายในเทสต์"""
    return tb.insert_zwsp(text).replace(Z, "|")


# ─────────────────────────────────── ข้อความจริงจากเอกสารที่เคยส่ง

VACCINE = ("สถาบันวัคซีนแห่งชาติได้รับงบประมาณรายจ่ายประจำปีงบประมาณ พ.ศ. 2570 "
           "เพื่อดำเนินการตามแผนงานโครงการ")
KPI = ("ขอเรียนชี้แจงว่าการดำเนินการตามตัวชี้วัดและการก่อหนี้ผูกพัน"
       "เป็นไปตามแนวทางที่สำนักงบประมาณกำหนด")
MONEY = ("มีวงเงินงบประมาณรวมทั้งสิ้น 50,038,500 บาท ตามที่ขอทำความตกลงไปได้ "
         "ความละเอียดแจ้งแล้ว")


def test_text_unchanged():
    """
    ตัวตัดคำเปลี่ยนได้แค่ 2 อย่าง: ใส่ ZWSP และเปลี่ยนช่องว่างเป็น NBSP
    ถอดทั้งสองอย่างออกแล้วต้องได้ของเดิมเป๊ะ
    """
    for t in (VACCINE, KPI, MONEY):
        assert tb.plain_text(tb.insert_zwsp(t)) == t


def test_tail_word_never_starts_line():
    """'แห่งชาติ' ห้ามขึ้นต้นบรรทัด — เคยหลุดเพราะการ glue กระโดดข้ามการตรวจ"""
    assert "วัคซีน|แห่งชาติ" not in cut(VACCINE)


def test_keep_together_phrases():
    """วลีที่กิ๊ฟสั่งมัดไว้ ห้ามมีจุดตัดคั่นกลาง"""
    assert "ประจำปีงบประ|มาณ" not in cut(VACCINE)     # เคยถูกตัวซอยผ่า
    assert "ตัว|ชี้" not in cut(KPI) and "ชี้|วัด" not in cut(KPI)
    assert "ก่อ|หนี้" not in cut(KPI)
    assert "แนว|ทาง" not in cut(KPI)
    assert "รวม|ทั้งสิ้น" not in cut(MONEY)
    assert "ความ|ละเอียด" not in cut(MONEY)


def test_nominal_prefix_stays_with_verb():
    """ตัวตัดคำชอบผลัก 'การ' ไปเกาะคำถัดไป — 'ดำเนิน|การตาม' ต้องกลายเป็น 'ดำเนินการ|ตาม'"""
    assert "ดำเนิน|การ" not in cut(VACCINE)


def test_numbers_and_latin_untouched():
    """ใส่เฉพาะจุดไทยชนไทย ตัวเลขและอังกฤษต้องไม่ถูกแหก"""
    assert "50,038,500" in cut(MONEY)
    t = "ระบบ VIMS บนคลาวด์ GDCC รองรับ Mobile device ของหน่วยงาน"
    assert "VIMS" in cut(t) and "Mobile" in cut(t) and "GDCC" in cut(t)


def test_phone_number_intact():
    """เบอร์โทรห้ามขาด"""
    assert "012-345-6789" in cut("ผู้ชี้แจงหมายเลขโทรศัพท์ 012-345-6789 ติดต่อได้")


def test_split_long_chunks_uses_dictionary():
    """ซอยคำยาวต้องได้คำจริงทั้งสองฝั่ง ไม่ใช่ซอยตรงกลางมั่ว ๆ"""
    assert tb.split_long_chunks(["งบประมาณรายจ่าย"]) == ["งบประมาณ", "รายจ่าย"]
    # วลีที่มัดไว้ห้ามโดนซอย แม้จะยาวเกิน MAX_CHUNK
    assert tb.split_long_chunks(["ประจำปีงบประมาณ"]) == ["ประจำปีงบประมาณ"]


def test_no_break_before_vowel_marks():
    """ห้ามขึ้นบรรทัดใหม่ด้วยสระบน/ล่าง/วรรณยุกต์"""
    out = tb.insert_zwsp(KPI)
    for i, ch in enumerate(out):
        if ch == Z and i + 1 < len(out):
            assert out[i + 1] not in tb.NO_BREAK_BEFORE


def test_leading_vowels_are_breakable():
    """
    สระหน้า เ แ โ ใ ไ ต้องตัดข้างหน้าได้ — ถ้าห้าม คำอย่าง และ/ได้/เป็น
    จะเกาะคำก่อนหน้าเป็นก้อนยาว 25-30 ตัว แล้วบรรทัดจะถ่าง
    """
    for ch in "เแโใไ":
        assert ch not in tb.NO_BREAK_BEFORE
    assert "|และ" in cut("การดำเนินงานและการติดตามผลการปฏิบัติงานของหน่วยรับงบประมาณ")


def test_paragraph_level_beats_per_run():
    """
    หัวใจของรีโปนี้: เอกสารจาก Word มี run ที่ผ่ากลางคำ
    ตัดคำทีละ run จะได้ขอบเขตผิด ('โด' + 'น' -> ตัดกลาง 'โดน')
    """
    from docx import Document
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("ประชาชนไม่โด")      # run ผ่ากลางคำแบบที่เจอจริง
    p.add_run("นทอดทิ้ง")
    tb.insert_zwsp_paragraph(p)
    joined = "".join(r.text for r in p.runs).replace(Z, "|")
    assert "โด|น" not in joined
    assert tb.strip_zwsp("".join(r.text for r in p.runs)) == "ประชาชนไม่โดนทอดทิ้ง"


# ─────────────────────────── v1.1: กฎ keep_together ที่ถอดจากโม 09

N = tb.NBSP


def test_kwam_and_kan_glue_forward():
    """"ความ" ติดคำถัดไปเสมอ · "การ" ติดเฉพาะเมื่อยืนเดี่ยว ("ดำเนินการ" ไม่ติด)"""
    assert "ความ|" not in cut("ขอความเห็นชอบในหลักการตามความจำเป็นของหน่วยงาน")
    assert "การ|ก่อหนี้" not in cut("ขออนุมัติการก่อหนี้ผูกพันข้ามปีงบประมาณ")
    # "การ" ที่เป็นหางคำต้องไม่ดึงคำถัดไปมาติด
    assert "ดำเนินการ|" in cut("หน่วยงานได้ดำเนินการเพื่อให้เป็นไปตามแผน")


def test_number_and_unit_bound():
    """เลขไม่เกิน 3 หลักติดหน่วยนับ และคำว่าจำนวน/รวม ติดตัวเลข"""
    out = tb.insert_zwsp("มีทั้งหมด จำนวน 4 แผนงาน 2 ผลผลิต")
    assert "4" + N + "แผนงาน" in out and "2" + N + "ผลผลิต" in out
    assert "จำนวน" + N + "4" in out


def test_amount_and_baht_bound():
    """จำนวนเงินกับคำว่าบาทต้องอยู่บรรทัดเดียวกัน"""
    assert "50,038,500" + N + "บาท" in tb.insert_zwsp("วงเงินงบประมาณ 50,038,500 บาท")


def test_por_sor_bound():
    """พ.ศ. กับปีห้ามขาดจากกัน"""
    assert "พ.ศ." + N + "2570" in tb.insert_zwsp("ประจำปีงบประมาณ พ.ศ. 2570 ของหน่วยงาน")


def test_short_parentheses_whole():
    """วงเล็บสั้นไม่เกิน 12 ตัวอยู่บรรทัดเดียวทั้งก้อน และห้ามตัดหลังวงเล็บเปิด"""
    out = tb.insert_zwsp("ก่อสร้างอาคารสำนักงาน (เฟส 2) ตามแผน")
    assert "(เฟส" + N + "2)" in out
    assert "(" + Z not in out


def test_lead_word_does_not_swallow_next_phrase():
    """
    บั๊กเดิมของ v1.0 และของโม 09 รุ่นแรก: มัดคำนำต่อกันเป็นทอดจนตัดไม่ได้ทั้งก้อน
    "…ให้เต็มวงเงิน / เมื่อสำนักงบประมาณ" ต้องยังตัดหน้า "เมื่อ" ได้
    """
    out = tb.insert_zwsp("จัดสรรให้เต็มวงเงิน เมื่อสำนักงบประมาณให้ความเห็นชอบ")
    assert "วงเงิน เมื่อ" in out              # ช่องว่างหน้า "เมื่อ" ยังเป็นช่องว่างธรรมดา ตัดได้
    assert "วงเงิน" + N + "เมื่อ" not in out


def test_phrase_does_not_eat_following_break():
    """
    บั๊กที่เจอตอนถอดโค้ดโม 09: รูปแบบวลีต่อ ZWSP? หลังตัวสุดท้ายด้วย
    จุดตัดหลังวลีจึงหายไปเสมอ — "แผนงาน|โครงการ" ต้องยังตัดตรงกลางได้
    """
    assert "แผนงาน|โครงการ" in cut("ประกอบด้วยแผนงานโครงการต่าง ๆ")


# ─────────────────────────── v1.1: กฎจากการอ่านไล่รอยต่อบรรทัด

def test_abbreviations_never_split():
    """ชื่อย่อห้ามขาดกลาง — เจอจริง 'สว / รส.' และ 'ม. / มหิดล'"""
    assert "สวรส." in cut("ให้กว้างขวางมากยิ่งขึ้น สวรส. ได้บรรจุแผนงาน")
    assert "สวรส." in cut("ตามข้อเสนอของสวรส. ที่ผ่านมา")
    assert "ม.มหิดล" in cut("ตามบันทึกข้อตกลงกับ ม.มหิดล ในปีนี้")


def test_dictionary_gaps_kept_whole():
    """คำจริงที่พจนานุกรมไม่มี ตัวตัดคำเคยผ่ากลาง"""
    for w in ("เฝ้าระวัง", "กรมควบคุมโรค", "รวมถึง", "สัญชาติไทย"):
        assert w in cut("หน่วยงานมีระบบ" + w + "ที่ครอบคลุม")


def test_whole_compounds_not_split():
    """คำประสมที่ซอยแล้วความหมายเพี้ยนห้ามซอย แต่ 'งบประมาณรายจ่าย' ยังซอยได้เหมือนเดิม"""
    for w in ("โทรศัพท์มือถือ", "สิทธิประโยชน์", "จัดซื้อจัดจ้าง"):
        assert tb.split_long_chunks([w]) == [w]
        assert w in cut("รองรับการใช้งาน" + w + "ของประชาชน")
    assert tb.split_long_chunks(["งบประมาณรายจ่าย"]) == ["งบประมาณ", "รายจ่าย"]


def test_compound_leads_glue_forward():
    """"ให้มี" "อย่างมี" "เพื่อให้" ต้องไปพร้อมคำที่ตามมา — เจอ 'โรคให้ / มีประสิทธิภาพ'"""
    assert "ให้มีประสิทธิภาพ" in cut("สร้างเสริมภูมิคุ้มกันโรคให้มีประสิทธิภาพและยั่งยืน")
    assert "อย่างมีประสิทธิภาพ" in cut("รับมือกับวิกฤตในอนาคตอย่างมีประสิทธิภาพ")
    assert "เพื่อให้ประเทศ" in cut("ปรับปรุงกฎหมายเพื่อให้ประเทศไทยพึ่งพาตนเองได้")


def test_prefix_words_glue_forward():
    """"เชิง" "ทาง" ยืนเดี่ยวติดคำถัดไป แต่ "ทาง" ท้าย "แนวทาง" ห้ามดึงคำถัดไปมาติด"""
    assert "เชิงนโยบาย" in cut("หากมุ่งเน้นการสนับสนุนเชิงนโยบายโดยจัดสรรงบ")
    assert "ทางพันธุกรรม" in cut("คัดกรองมะเร็งและโรคติดต่อทางพันธุกรรมล่วงหน้า")
    assert "แนวทาง|" in cut("เป็นไปตามแนวทางที่สำนักงบประมาณกำหนด")


def test_document_only_justified_paragraphs():
    """
    ใส่ ZWSP เฉพาะย่อหน้าที่จัดชิดขอบขวา — หัวเรื่องไม่ต้องมี
    (บทเรียนโม 09 ก.ค. 2569: ใส่ ZWSP ทุกย่อหน้าแล้ว "แยกถูกผิดไม่ออก")
    """
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc = Document()
    title = doc.add_paragraph("ร่างพระราชบัญญัติงบประมาณรายจ่ายประจำปีงบประมาณ")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    body = doc.add_paragraph(VACCINE)
    body.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
    n = tb.insert_zwsp_document(doc)
    assert n > 0
    assert Z not in title.text
    assert Z in body.text


def test_document_warns_when_nothing_justified():
    """ไม่มีย่อหน้าจัดชิดขอบขวาเลย ต้องเตือน ไม่ใช่เงียบ"""
    import warnings
    from docx import Document
    doc = Document()
    doc.add_paragraph(VACCINE)
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        assert tb.insert_zwsp_document(doc) == 0
    assert any("ชิดขอบขวา" in str(x.message) for x in w)


def test_mapping_keeps_run_boundaries_with_nbsp():
    """เปลี่ยนช่องว่างเป็น NBSP ข้ามรอยต่อ run แล้ว ข้อความและขอบเขต run ต้องเหมือนเดิม"""
    from docx import Document
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("วงเงินงบประมาณรวมทั้งสิ้น ")
    p.add_run("50,038,500")
    p.add_run(" บาท ประจำปีงบประมาณ พ.ศ. 2570")
    before = [r.text for r in p.runs]
    tb.insert_zwsp_paragraph(p)
    after = [tb.plain_text(r.text) for r in p.runs]
    assert after == before
    assert N + "บาท" in p.runs[2].text


if __name__ == "__main__":
    import io
    # คอนโซล Windows เป็น cp874/cp1252 ภาษาไทยจะเพี้ยน ต้องบังคับ utf-8
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    fails = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print("  ผ่าน  %s" % name)
            except AssertionError as e:
                fails += 1
                print("  ตก    %s  %s" % (name, e))
    print("\n%s" % ("ผ่านทั้งหมด" if not fails else "ตก %d ข้อ" % fails))
    sys.exit(1 if fails else 0)
