# riabroi-ratchakan-typeset

**เรียบร้อย** — ตัดคำและจัดตัวหนังสือเอกสารราชการไทย ในแบบของกิ๊ฟ
*Riabroi — Thai government document typesetting: words break where they should, and the right edge stays flush.*

[ภาษาไทย](#ภาษาไทย) · [English](#english)

---

## ภาษาไทย

### คืออะไร

สกิลสำหรับผู้ช่วย AI (Claude, GPT และตัวอื่น ๆ) พร้อมโค้ด Python ที่ใช้งานได้จริง
สำหรับสร้างและจัดหน้าเอกสารราชการไทยในไฟล์ `.docx` ให้เรียบร้อยเหมือนคนจัดเอง

กลั่นจากเอกสารที่ส่งผู้บริหารลงนามไปแล้ว และจากจุดที่เจ้าของงานกดแก้มือเองทีละจุด

### ปัญหาที่แก้

- **Word ตัดบรรทัดไทยกลางคำ** — มันไล่เติมบรรทัดให้เต็มโดยไม่สนขอบเขตคำ
  วัดจากเอกสารจริงเจอ 12.7% ของบรรทัดที่ห่อ เช่น `ระบบสาก|ล` `กำหน|ด` `พื้น|ที่`
- **ขอบขวาไม่ชิด หรือชิดแล้วช่องไฟถ่าง** — แก้ด้วยการบีบระยะอักขระทีละขั้น ไม่ใช่เลิกจัดชิด
- **คำค้างท้ายบรรทัดแล้วอ่านสะดุด** เช่น "ให้" "เมื่อ" "วงเงิน" ค้างอยู่บรรทัดบนเดี่ยว ๆ
- **ฟอนต์ไทยไม่เปลี่ยนตามที่สั่ง** — python-docx ไม่ตั้งฝั่ง complex script ที่อักษรไทยใช้จริง
- **หน้ายืด** — ค่าปริยายของ Word ตั้งระยะบรรทัด 1.15 กับเว้นท้ายย่อหน้า 10 pt มาให้

### ผลที่วัดได้

จากเอกสารประเด็นถามตอบ 4 ไฟล์

| | ก่อน | หลัง |
|---|---|---|
| ตัดกลางคำ | 12.7% ของบรรทัดที่ห่อ | ~0% |
| จำนวนหน้า | 25 | 20 |
| ข้อความ | — | ไม่เปลี่ยนแม้แต่ตัวเดียว |

### ต่างจาก thai-docx

[thai-docx](https://github.com/Netthip/thai-docx-for-a-thai-civil-servant-eager-to-learn) เป็นตัวพื้นฐานสำหรับเอกสารไทยทั่วไป
รีโปนี้ลงรายละเอียดมากกว่า ทั้งการตัดคำและการจัดตัวหนังสือ สำหรับงานราชการโดยเฉพาะ

| เรื่อง | thai-docx | riabroi |
|---|---|---|
| ขอบเขตการตัดคำ | ทีละ run | ทั้งย่อหน้า แล้วแมปกลับเข้า run |
| สระหน้า เ แ โ ใ ไ | ห้ามตัดข้างหน้า | ตัดได้ — ไม่งั้นคำเกาะกันเป็นก้อนยาวจนบรรทัดถ่าง |
| คำประสมยาว | ปล่อยยาว | ซอยตรงจุดที่ได้คำจริงทั้งสองฝั่ง |
| วลีที่ห้ามขาด / คำค้างท้ายบรรทัด | ไม่มี | มี |
| ขอบขวา | ชิดซ้าย | ชิดขอบ + บีบอักษร 0.1–0.7 pt เลือกขั้นที่ถ่างน้อยสุด |

### ใช้งาน

**ติดตั้ง**

```bash
git clone https://github.com/Netthip/riabroi-ratchakan-typeset.git
pip install -r riabroi-ratchakan-typeset/requirements.txt
```

**ใช้เป็นสกิลใน Claude Code** — วางโฟลเดอร์ไว้ที่ `~/.claude/skills/riabroi-ratchakan-typeset`

**ใช้กับผู้ช่วย AI อื่น (เช่น GPT)** — อัปโหลด `SKILL.md` และโฟลเดอร์ `references/` เป็นความรู้ของผู้ช่วย
และอัปโหลด `scripts/` ถ้าผู้ช่วยรันโค้ด Python ได้
สคริปต์ตัดคำต้องมี `pythainlp` — ถ้าผู้ช่วยติดตั้งไม่ได้ (เช่น sandbox ไม่มีอินเทอร์เน็ต)
จะยังได้กฎใน `SKILL.md` อยู่ และสคริปต์จะขึ้นคำเตือนแทนที่จะเงียบ

**เรียกจากโค้ด**

```python
import sys; sys.path.insert(0, "riabroi-ratchakan-typeset/scripts")
from docx import Document
import thai_break, thai_fit

doc = Document("ต้นฉบับ.docx")
thai_fit.normalise_document(doc)       # ฟอนต์ฝั่งไทยครบ + เคลียร์ค่าที่ทำให้หน้ายืด
thai_break.insert_zwsp_document(doc)   # บอก Word ว่าตัดบรรทัดตรงไหนได้
doc.save("ผลลัพธ์.docx")
```

**ตรวจผล** (ต้องมี Microsoft Word บน Windows เพื่อส่งออก PDF)

```bash
python scripts/inspect_docx.py ผลลัพธ์.docx --png ภาพ
python tests/test_thai_break.py
```

### โครงรีโป

```
SKILL.md                กฎและลำดับงาน (ผู้ช่วย AI อ่านไฟล์นี้)
references/
  typesetting.md        จัดหน้า บีบอักษร คำค้าง วิธีวัดผล
  forms.md              ขอบกระดาษและผังของเอกสารแต่ละแบบ
  toolchain.md          Word COM · อ่าน PDF · กับดักที่เจอมาแล้ว
scripts/
  thai_break.py         ตัดคำ + ใส่ ZWSP
  thai_fit.py           ฟอนต์ · บีบอักษร · line break · ระยะบรรทัด
  inspect_docx.py       ส่งออก PDF แล้ววัดผลจริง
tests/                  เทสต์จากจุดที่เคยพังจริง
evals/                  ทดสอบว่าผู้ช่วย AI หยิบสกิลนี้ถูกจังหวะ (14/14)
```

---

## English

### What it is

A skill for AI assistants (Claude, GPT and others) plus working Python code
for producing and typesetting Thai government documents in `.docx` — so they look hand-set.

Distilled from documents already sent up for signature, and from the exact spots
the author fixed by hand, one line break at a time.

### Problems it solves

- **Word breaks Thai lines mid-word.** It fills each line to the edge regardless of word boundaries.
  Measured on real documents: 12.7% of wrapped lines broke inside a word.
- **Ragged right edge, or a flush edge with stretched spacing.** Fixed by condensing
  character spacing step by step — not by giving up on justification.
- **Dangling words at line ends** that make a sentence stumble.
- **Thai fonts that don't change.** python-docx never sets the complex-script font slot Thai actually uses.
- **Bloated pages.** Word's defaults add 1.15 line spacing and 10 pt after every paragraph.

### Measured results

On 4 committee Q&A documents:

| | Before | After |
|---|---|---|
| Mid-word breaks | 12.7% of wrapped lines | ~0% |
| Pages | 25 | 20 |
| Text | — | unchanged, character for character |

### How it differs from thai-docx

[thai-docx](https://github.com/Netthip/thai-docx-for-a-thai-civil-servant-eager-to-learn) is the general-purpose base for Thai documents.
This repo goes further on both word breaking and typesetting, specifically for government work.

| | thai-docx | riabroi |
|---|---|---|
| Segmentation scope | per run | whole paragraph, mapped back into runs |
| Leading vowels เ แ โ ใ ไ | never break before | breakable — otherwise words clump into 25–30 character blocks |
| Long compounds | left whole | split only where both halves are real dictionary words |
| Protected phrases / dangling lead words | — | yes |
| Right edge | left-aligned | justified, condensed 0.1–0.7 pt, least-stretched step wins |

### Usage

**Install**

```bash
git clone https://github.com/Netthip/riabroi-ratchakan-typeset.git
pip install -r riabroi-ratchakan-typeset/requirements.txt
```

**As a Claude Code skill** — place the folder at `~/.claude/skills/riabroi-ratchakan-typeset`

**With other AI assistants (e.g. GPT)** — upload `SKILL.md` and `references/` as knowledge,
plus `scripts/` if the assistant can run Python.
The segmentation script needs `pythainlp`. If the assistant can't install it (e.g. an offline sandbox),
you still get the rules in `SKILL.md`, and the script warns instead of silently doing nothing.

**From code**

```python
import sys; sys.path.insert(0, "riabroi-ratchakan-typeset/scripts")
from docx import Document
import thai_break, thai_fit

doc = Document("input.docx")
thai_fit.normalise_document(doc)       # complex-script fonts + clear page-bloating defaults
thai_break.insert_zwsp_document(doc)   # tell Word where Thai lines may break
doc.save("output.docx")
```

**Verify** (needs Microsoft Word on Windows to export PDF)

```bash
python scripts/inspect_docx.py output.docx --png images
python tests/test_thai_break.py
```

### Key findings

- Setting `w:lang w:bidi="th-TH"` does **not** fix line breaking — tested in isolation, identical results.
  Only zero-width spaces (U+200B) at word boundaries work.
- Line spacing below 1.0 makes Thai tone marks collide with the line above (0.95 → 4 collisions).
- Child elements of `w:rPr` must follow schema order, or Word silently drops the property.
- More condensing is not always better: pulling words up can leave the next line shorter and wider.
  Measure every step and keep the least-stretched one.

### Repository layout

```
SKILL.md                rules and workflow (the AI assistant reads this)
references/
  typesetting.md        condensing, dangling words, measurement
  forms.md              margins and layouts per document type
  toolchain.md          Word COM, reading PDFs, known pitfalls
scripts/
  thai_break.py         segmentation + ZWSP insertion
  thai_fit.py           fonts, condensing, line breaks, line spacing
  inspect_docx.py       export PDF and measure the real result
tests/                  regression tests from real failures
evals/                  checks that assistants pick this skill at the right time (14/14)
```
