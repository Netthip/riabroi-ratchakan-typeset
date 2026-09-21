# เครื่องมือและกับดักที่เจอมาแล้ว

สารบัญ
1. สคริปต์ในรีโปนี้
2. Word COM กับ OneDrive
3. Word for Mac (macOS)
4. อ่าน PDF แบบ สงป. / หนังสือเวียน
5. กับดักเครื่องมือแก้ไฟล์
6. โค้ดที่มีอยู่แล้วใน Gift Command Center

---

## 1. สคริปต์ในรีโปนี้

| ไฟล์ | ใช้ทำอะไร |
|---|---|
| `scripts/thai_break.py` | ตัดคำ + ใส่ ZWSP + `keep_together` (มัดคำด้วย NBSP) — ตัวหลักของรีโป |
| `scripts/thai_fit.py` | ฟอนต์ฝั่ง complex script · จัดเนื้อความชิดขอบขวา · บีบอักษร · line break · ระยะบรรทัด |
| `scripts/inspect_docx.py` | ส่งออก PDF แล้ววัดผลจริง (ตัดกลางคำ ช่องไฟ ขอบขวา ฟอนต์ถูกแทน) · `--breaks` รอยต่อบรรทัดทุกจุด · เรนเดอร์ภาพ |
| `scripts/word_pdf_mac.applescript` | ให้ `to_pdf` สั่ง Word for Mac ส่งออก PDF (หัวข้อ 3) |
| `scripts/doctor.py` | ตรวจเครื่องก่อนใช้ — ไลบรารี ตัวตัดคำ ฟอนต์ Word · `--pdf` ลองส่งออกจริงแล้วดูว่า Word ตัดตรง ZWSP ไหม |
| `tests/test_thai_break.py` | เทสต์จากจุดที่เคยพังจริง — รันก่อนแก้ตัวตัดคำทุกครั้ง |
| `tests/test_platform.py` | เทสต์ส่วนที่ขึ้นกับเครื่อง (จำลอง osascript/LibreOffice) + ล็อกกติกาความปลอดภัยของ AppleScript |

```python
import sys; sys.path.insert(0, "scripts")
from docx import Document
import thai_break, thai_fit

doc = Document("ต้นฉบับ.docx")
thai_fit.normalise_document(doc)          # ฟอนต์ครบฝั่ง cs + เคลียร์ docDefaults ที่ทำให้หน้ายืด
thai_fit.justify_body(doc)                # เนื้อความชิดขอบขวา
n = thai_break.insert_zwsp_document(doc)  # ZWSP + keep_together เฉพาะย่อหน้าที่ชิดขอบขวา
doc.save("ผลลัพธ์.docx")
```

ตรวจผล (macOS พิมพ์ `python3`):
```bash
python scripts/inspect_docx.py ผลลัพธ์.docx --png ภาพ --breaks
python scripts/inspect_docx.py ผลลัพธ์.docx --engine libreoffice   # เครื่องที่ไม่มี Word — ดูภาพรวมเท่านั้น
```

`to_pdf` เลือกตัวส่งออกเอง: Windows ใช้ Word ผ่าน COM · macOS ใช้ Word for Mac ผ่าน AppleScript ·
ไม่มี Word ใช้ LibreOffice แล้วเตือน (`engine="word"` บังคับว่าต้องเป็น Word)
รายงานอ่าน metadata ของ PDF บอกว่ามาจากโปรแกรมไหน และเทียบฟอนต์ที่เอกสารสั่งกับที่ฝังใน PDF
รับโฟลเดอร์ปลายทางแบบ relative ได้ (v1.0 พังเพราะ Word ต้องการ path เต็ม)

---

## 2. Word COM กับ OneDrive

**ห้ามให้ Word (win32com) บันทึกไฟล์ลงโฟลเดอร์ OneDrive โดยตรง**
สร้างใน `tempfile.mkdtemp()` แล้วค่อยคัดลอกไปปลายทาง

เจอมาแล้ว 2 อาการเมื่อบันทึกทับไฟล์ใน OneDrive ตรง ๆ (8 ก.ย. 2569)
1. marker แบบ `@@TABLE_ALLOC@@` ค้างเป็นตัวหนังสือ ทั้งที่ฝังวัตถุไปแล้ว (สุ่ม 2-3 ใน 4 ไฟล์ต่อรอบ)
2. ไฟล์ .docx ไม่ใช่ zip เปิดไม่ได้ ทั้งที่ตัวตรวจรายงานว่าผ่าน

**ต้องตรวจทุกครั้งก่อนส่ง:** `zipfile.is_zipfile(path)` และอ่าน `word/document.xml`
ยืนยันว่าไม่มี marker ค้าง — ถ้ามีให้ล้มทันที ห้ามส่ง

**เปิดด้วย `DispatchEx`** เสมอ เพื่อได้ instance แยก ไม่ไปยุ่งกับ Word ที่กิ๊ฟเปิดค้างอยู่
`to_pdf` ให้ Word ส่งออกลงเทมป์ก่อนแล้ว Python ค่อยคัดลอก PDF ไปปลายทาง (ตั้งแต่ v1.2)

**ถ้าเขียนไฟล์ปลายทางไม่ได้ (PermissionError) = กิ๊ฟเปิดค้างใน Word**
บอกกิ๊ฟให้ปิดก่อน — **ห้ามฆ่าโปรเซส Word ของกิ๊ฟ**

**สร้างซ้ำคนละวันได้หน้าตาไม่เหมือนเดิม** แม้โค้ดและข้อมูลเดิมทุกอย่าง
(17 ก.ย. 2569 ตาราง Excel ที่ฝังวาดสูงขึ้น 0.85 pt ทั้งที่ไฟล์ xlsx เหมือนเดิมทุกไบต์
แล้วข้อความใต้ตารางเลื่อนลงจนทับกล่องคำสั่ง)
→ **ก่อนวางทับไฟล์ที่ส่งกิ๊ฟไปแล้ว ให้สร้างลงโฟลเดอร์ทดลองแล้วเทียบตำแหน่งบรรทัด/รอยตัดก่อนเสมอ**

---

## 3. Word for Mac (macOS)

`to_pdf` บน Mac เรียก `osascript scripts/word_pdf_mac.applescript` — ส่วนตัดคำและจัดหน้าไม่ต่างจาก Windows

**Mac เปิด Word ได้ตัวเดียว** ไม่มี `DispatchEx` แบบ Windows จึงกันไม่ให้ไปยุ่งกับงานที่เปิดค้างด้วยกติกานี้
- Python คัดลอกไฟล์ไปเป็น `riabroi-<uuid>.docx` ใน `~/riabroi-pdf` แล้วสั่ง Word เปิดสำเนานั้น ไม่เปิดไฟล์จริง
  (ถ้าเปิดไฟล์ที่ผู้ใช้เปิดค้างอยู่ Word จะคืนหน้าต่างเดิมมา แล้วการปิดแบบไม่บันทึก = งานที่ยังไม่บันทึกหาย)
- หาเอกสารจากชื่อสำเนา ไม่ใช้ `active document` (ผู้ใช้คลิกหน้าต่างอื่นระหว่างทาง = ไปปิดผิดไฟล์)
- ปิดเฉพาะเอกสารที่ชื่อมี uuid นั้น แบบไม่บันทึก · ปิด Word เฉพาะเมื่อสคริปต์เป็นคนเปิด **และ** ไม่มีเอกสารอื่นค้าง
- **ห้าม** `killall`/`pkill` Word และห้ามเขียน AppleScript สั่ง Word เองนอกไฟล์นี้
- `tests/test_platform.py` ล็อกกติกาเหล่านี้ไว้ — แก้สคริปต์แล้วข้อไหนหลุดเทสต์จะตก

**ทำไมต้องโฟลเดอร์ `~/riabroi-pdf` ที่เดิมทุกครั้ง:** Word for Mac อยู่ใน sandbox เขียนไฟล์ได้เฉพาะโฟลเดอร์
ที่ผู้ใช้อนุญาต ครั้งแรกจะขึ้นกล่อง *Grant File Access* → Select… → Grant Access แล้ว Word จำโฟลเดอร์ไว้
ถ้าใช้โฟลเดอร์ชั่วคราวชื่อใหม่ทุกครั้งจะโดนถามทุกครั้ง · เปลี่ยนที่ได้ด้วยตัวแปร `RIABROI_WORKDIR`
PDF ถูกคัดลอกออกไปปลายทางด้วย Python — Word ไม่ได้เขียนลง OneDrive ตรง ๆ (กติกาเดียวกับหัวข้อ 2)

**สิทธิ์ครั้งแรก 2 กล่อง** ให้รัน `python3 scripts/doctor.py --pdf` ตอนผู้ใช้นั่งหน้าเครื่อง
1. macOS ถามให้แอปที่รันคำสั่ง (Terminal / iTerm / Claude) ควบคุม Microsoft Word — ถ้าเผลอกดไม่อนุญาต
   จะได้ error -1743 แก้ที่ System Settings > Privacy & Security > Automation
2. Word ขอสิทธิ์โฟลเดอร์ `~/riabroi-pdf` — ถ้าไม่มีคนกด AppleScript จะรอจนหมดเวลา (error -1712)

`to_pdf` แปลงรหัสพวกนี้เป็นวิธีแก้ภาษาไทยให้แล้ว (`_MAC_HINTS`) — เจอแล้วให้บอกผู้ใช้ อย่าวนรันซ้ำ

**ตั้ง Save as PDF เป็น "Best for printing"** — การส่งออกด้วยสคริปต์ใช้ตัวเลือกล่าสุดที่ผู้ใช้เลือกด้วยมือ
ถ้าค้างที่ "Best for electronic distribution and accessibility" Word จะส่งเอกสารไปแปลงบนบริการออนไลน์
ของ Microsoft (ต้องมีเน็ต และหน้าตาอาจไม่ตรงกับที่เห็นในเครื่อง) · `pdf_engine` ให้ค่า `word-mac`
เมื่อ metadata ของ PDF มีทั้งคำว่า Word และ Quartz/macOS (ยังไม่ได้เห็น PDF จาก Mac จริง — ถ้าได้ `word`
หรือ `unknown` บน Mac ให้ดู creator/producer แล้วแก้ `pdf_engine`)

**ฟอนต์** — Mac ไม่มี TH Sarabun มาให้ ถ้าไม่ได้ติดตั้ง Word จะใช้ฟอนต์อื่นแทนเงียบ ๆ ตัดบรรทัดและช่องไฟ
เปลี่ยนหมด → รายงานขึ้นเตือน "ไม่มี TH SarabunPSK ใน PDF" ให้หยุดแล้วติดตั้งฟอนต์ก่อน
ติดตั้ง: ดับเบิลคลิก `.ttf` → Install Font (ลง `~/Library/Fonts`)

**ไฟล์ TH SarabunIT๙ รุ่น Windows บอก Mac ว่าตัวเองชื่อ TH SarabunPSK** (วัดจากไฟล์จริง 21 ก.ย. 2569)
- ตาราง name: ID 1 ฝั่ง Windows = "TH SarabunIT๙" แต่ ID 16 และชื่อฝั่ง Mac ทุกช่อง = "TH SarabunPSK"
- IT๙ วาดเลขอารบิก `2570` เป็นรูปเลขไทย `๒๕๗๐` (ภาพตรงกันทุกพิกเซล) และกว้างกว่า PSK ~24%
- Windows ใช้ ID 1 จึงเห็นเป็นคนละฟอนต์ · Mac จัดกลุ่มตาม ID 16 จึง **อาจ** รวมเป็นตระกูลเดียวกับ PSK
  แล้วตัวเลขทั้งเอกสารเปลี่ยนหน้าตา — ยังไม่ได้ยืนยันบน Mac จริง `doctor.py` เตือนเมื่อเจอไฟล์แบบนี้
- อย่าใช้ `fitz.Font(...).name` อ่านชื่อฟอนต์ — ไฟล์ IT๙ เขียนชื่อเต็มไว้ว่า PSK จะรายงานผิดว่ามี PSK

**Word for Mac อาจตัดบรรทัดไม่ตรงกับ Windows** (จัดวางตัวอักษรคนละระบบ) — เอกสารที่จะพิมพ์หรือส่งต่อ
จากเครื่อง Windows ให้ตรวจรอบสุดท้ายบน Windows · `doctor.py --pdf` วัดตรง ๆ ว่า Word บนเครื่องนั้น
ตัดบรรทัดตรง ZWSP/ช่องว่างทุกบรรทัดไหม (ไม่ใช้พจนานุกรม จึงไม่มีผลบวกลวง)

**Python บน Mac** ใช้ `python3` · Python จาก Homebrew ห้าม `pip install` ลงระบบ (`externally-managed-environment`)
ให้ใช้ venv นอกโฟลเดอร์ OneDrive เช่น `~/.venvs/riabroi` (venv ในโฟลเดอร์ที่ซิงก์ = ไฟล์หลายพันไฟล์ขึ้นคลาวด์)

**OneDrive บน Mac** อยู่ที่ `~/Library/CloudStorage/OneDrive-…` (Finder แสดงเป็น "OneDrive - ชื่อองค์กร")
ดูชื่อจริงด้วย `ls ~/Library/CloudStorage`

**LibreOffice** (`--engine libreoffice`) ใช้ได้ทุกเครื่องที่ไม่มี Word แต่ตัดบรรทัดไทยด้วยพจนานุกรมของตัวเอง
ไม่ใช่แบบ Word — ตัวเลขตัดกลางคำจาก PDF ของ LibreOffice จึงบอกอะไรเกี่ยวกับ Word ไม่ได้ ใช้ดูภาพรวมเท่านั้น

---

## 4. อ่าน PDF แบบ สงป. / หนังสือเวียน

**เรนเดอร์หน้า:** เครื่องนี้ไม่มี pdftoppm การ Read PDF แบบระบุ `pages` จึงใช้ไม่ได้
ให้ใช้ PyMuPDF `page.get_pixmap(dpi=150)` บันทึก PNG ใน scratchpad แล้วอ่านภาพ
ครอปจุดเล็กด้วย `clip=` ที่ dpi 200+

**อักษรเพี้ยนตอนดึงข้อความ:**
- ฟอนต์ TH SarabunPSK คืนรหัส PUA (U+F700–U+F71A) ต้อง map กลับเป็นวรรณยุกต์/สระ
- สระอำหลุดเป็น "ช่องว่าง + า" (เช่น "ด าเนิน") แก้ด้วย regex `([ก-ฮ])\s([่้๊๋]?)า` → `\1\2ำ`
- หนังสือ ว 238 ดึงออกมาได้ "ำ" แทน "า" ทั้งฉบับ (ประมำณ)
  → **อ้างถ้อยคำหลักเกณฑ์ให้ดูจากภาพ ห้ามคัดจากข้อความที่ดึง**

**แยกคอลัมน์ตัวเลข:** จัดคำเป็นแถวตาม y แล้วจับคอลัมน์ "แผน (ขั้น พ.ร.บ.)"
จากกึ่งกลางหัว "(ขั้น พ.ร.บ.)" 5 ช่อง (รวม + ไตรมาส 1-4)

**จับข้อความที่ถูกตัดในแบบ 301:** ตัวชี้วัดใต้ "เป้าหมายการให้บริการกระทรวง"
เป็นชุดเดียวกันทุกหน่วยในกระทรวง เทียบกับแบบของหน่วยอื่นจะเห็นท้ายข้อความที่หาย
(เจอใน สวรส.: "พ.ศ. [2562]", "รับรองตาม[มาตรฐาน]", "(จำนวน 4,200 ทีม)")
เส้นที่ตัดผ่านตัวหนังสืออาจเป็นเส้นใต้หัวข้อ ไม่ใช่ข้อความขาด — ต้องดูภาพยืนยัน

---

## 5. กับดักเครื่องมือแก้ไฟล์

**อย่าเขียน escape ยูนิโค้ดของ ZWSP ผ่านเครื่องมือแก้ไฟล์** — มันจะกลายเป็นอักขระล่องหนจริง
ในซอร์ส (diff เห็นเป็นสตริงว่าง) และ editor อาจลบทิ้งโดยไม่รู้ตัว
ใช้ `chr(0x200B)` เสมอ แล้วสแกนไฟล์หาอักขระหมวด `Cf`/`Cc` หลังเขียน:

```python
import unicodedata
hits = [(i, hex(ord(c))) for i, c in enumerate(open(f, encoding="utf-8").read())
        if unicodedata.category(c) in ("Cf", "Cc") and c not in "\n\r\t"]
```

**แคช win32com ของ Excel บนเครื่อง Netthip.w เสีย** — เทสต์อาจขึ้นว่า "ไม่มี Excel" ทั้งที่มี
ให้รันผ่าน `gen_path` ที่ชี้ไปโฟลเดอร์ชั่วคราว · **การลบแคชต้องถามกิ๊ฟก่อน**

**คอนโซล Windows เป็น cp874/cp1252** ภาษาไทยที่ print จะเพี้ยน
ต้องครอบ `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")`
(Mac/Linux เป็น UTF-8 อยู่แล้ว ครอบไว้ก็ไม่เสียหาย)

---

## 6. โค้ดที่มีอยู่แล้วใน Gift Command Center

รีโปนี้เป็นตัวกลั่นของกฎ ส่วนเครื่องกำเนิดใบเสนองานตัวเต็มอยู่ที่
`Budget Bureau/Gift Command Center/modules/09_เครื่องกำเนิดใบเสนองาน/post_allocation/`

| โมดูล | หน้าที่ |
|---|---|
| `thai_text.py` | ตัดคำ · `set_char_spacing` · `line_break` · `keep_together` |
| `word_fit.py` | `tighten_justified` · `condense_to_fit` · `avoid_small_spill` · `center_signature` · `avoid_short_start` |
| `catchword.py` | คำโปรยท้ายหน้า |
| `build.py` / `generator.py` | สร้างในเทมป์แล้ว `deliver()` ไปปลายทาง · `_assert_no_marker` |
| `embed_xl.py` | ฝังตาราง Excel · `BLEED_SLACK_PT` |
| `ตรวจหน้าตา.py` | ตรวจขอบกระดาษ ตำแหน่งกล่องคำสั่ง ขนาดฟอนต์ |
| `core/thai_docx.py` | เอนจินกลาง — `insert_zwsp` · `_NO_BREAK_BEFORE` · presets |

กฎเหล็ก 3 ข้ออยู่ที่ `Gift Command Center/กฎเหล็ก.md` และบังคับจริงใน `core/hard_rules.py`
