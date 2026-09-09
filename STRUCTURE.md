# STRUCTURE — Bot_Transport_V2

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เพิ่ม/ลบ/ย้ายไฟล์, เปลี่ยน logic ฟังก์ชัน/คลาส, เปลี่ยน config key, เพิ่ม/แก้ประเภท report หรือ JMS endpoint, หรือเปลี่ยน flow — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
**เวอร์ชัน 2 ของ `Bot_Transport`** — รีไรต์ `Bot_Fei_Main.py` ให้รองรับ **หลาย report** และส่งได้ทั้งรูปและไฟล์เข้า Feishu Chat:
- **Main Line Transport** export (JMS) — `report.xlsx`
- **Branch Line Transport** export (JMS) — `Report2.xlsx` / `Branch_Task_Tracking.xlsx`
- แปลง Excel → PNG แล้วส่งเข้า Feishu Chat (รองรับ job sheets/range หลายชุด)

## วิธีรัน / Entry point
- รัน: `python Bot_Fei_Main.py` → คลาส `App(ctk.CTk)` (แท็บ Home / Setting)
- มาพร้อม template Excel ในโฟลเดอร์ `Excel/`

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `Bot_Fei_Main.py` | **ตัวหลัก (~2,200 บรรทัด)** — UI + scheduler + Feishu (token, `send_feishu_image_by_chat_id`, `send_feishu_file_by_chat_id`) + export Main/Branch transport + job config (sheets/range/output หลายชุด). ทั้งสอง export ใช้ header ชุดเดียวกันจาก `_jms_headers()`, ถาม `total` จาก endpoint ค้นหาของเมนูก่อน แล้วค่อยสั่ง export และกรอง task ใน `/export/selectTask` ด้วย `url` + `createTime` ของตัวเอง |
| `dws_collect.py` | อ่านไฟล์ realtime ของเครื่อง DWS แต่ละช่องแล้วสะสมลง SQLite (`dws_scans.db`) กันซ้ำด้วย (ช่อง, บาร์โค้ด, เวลา, ใบงาน) |
| `dws_report.py` | upsert เวลาลงพัสดุจาก Report ลงคลัง แล้ว join กับสแกนเพื่อสรุปเวลารวมต่อช่อง |
| `createpng.py` | Excel → PNG ด้วย win32com. มี `repoint_queries()` เขียน path ใน Power Query ให้ชี้ โฟลเดอร์ `jms_save_path`, `force_sync_refresh()` ปิด BackgroundQuery ก่อน refresh, `ExcelWatchdog` รายงานความคืบหน้าและฆ่า EXCEL.EXE เมื่อค้างเกินเวลา และ `_kill_excel()` เก็บกวาด process ผี |
| `sendfeishu.py` | ส่ง Feishu แบบ webhook + HMAC (เหมือน `Bot_Transport`) |
| `pyi_rth_tkinter_paths.py` | PyInstaller runtime hook แก้ path tkinter ตอน build เป็น exe |
| `Excel/` | template: `1Main_Line_Transport_Report.xlsx`, `2Branch_Line_Transport_Report.xlsx`, `report.xlsx`, `Report2.xlsx` |

## ความต่างจาก `Bot_Transport` (V1)
- รองรับ **2 ชนิด transport export** (Main + Branch) แยก checkbox: `run_main_transport`, `run_branch_transport`, `run_feishu_chat`
- เพิ่มการส่ง **ไฟล์** เข้า Feishu (`send_feishu_file_by_chat_id`) ไม่ใช่แค่รูป
- job config ยืดหยุ่นขึ้น (`_empty_sheet_rule`: `sheet`/`range`/`output` หลายชุดต่อ report)
- `createpng.py` / `sendfeishu.py` เนื้อในเหมือน V1

## Config (`config.ini`, `[SETTING]`)
- Scheduler: `run_hour_interval`, `run_minute_interval`, `start_hour`, `end_hour`
- `start_date` / `end_date` เป็นแค่ค่าที่เขียนกลับไว้ดู — ตอนเปิดโปรแกรมจะถูกคำนวณใหม่จากเวลาจริงเสมอ
- เปิด/ปิดงาน: `run_main_transport`, `run_branch_transport`, `run_feishu_chat`
- ไฟล์: `main_transport_filename` (report.xlsx), `branch_transport_filename` (Report2.xlsx)
- Feishu/PNG: `excel_file`, `excel_sheet_index`, `excel_range`, `chat_id`, `app_id`, `app_secret` ฯลฯ

## Dependencies / บริการภายนอก
- `customtkinter`, `tkcalendar`, `pywin32` (ต้องมี Microsoft Excel), `requests`
- JMS J&T (Main + Branch endpoints), Feishu OpenAPI

## เมนู JMS ที่แต่ละ export ดึงมา
> ⚠️ ต้องตรงกับเมนูบนเว็บ JMS จริง — ถ้าดึงผิดเมนู จำนวนคอลัมน์/แถวจะไม่ตรงกับที่ผู้ใช้ export มือ

| | เส้นหลัก (Main) | เส้นรอง (Branch) |
|---|---|---|
| เมนูบนเว็บ | รายงานสรุปการขนส่ง | รายงานขนส่งรวมสายย่อย |
| `routeName` | `monitoringReport` | `transportSynthesizeReport` |
| endpoint นับจำนวน | `/tmsShipmentEvent/report` | `/tmsBranchShipmentEvent/report` |
| endpoint export | `/tmsExportTransportReport/reportExport` | `/tmsExportTransportReport/branchReportExport` |
| taskName ใน selectTask | `运输综合查询导出` | `支线运输综合报表导出` |
| จำนวนคอลัมน์ | 63 | 53 |
| ดาวน์โหลด | `https://yl-file.jtexpress.co.th/{ossUrl}` ตรง ๆ | ผ่าน `/file/oss/getDownloadSignedUrl` |
| ตัวกรองเวลา | `timeType: 2` + `newTimeType: 1` | `timeType: 1` + `source: 2` |
| ตัวกรองสาขา | `arriveNetworkCodeList: ["999004"]` | `arriveNetworkCodeList: ["999004"]` |
| ต้องส่ง `count` | ไม่ต้อง | **ต้อง** (ไม่งั้นได้ไม่ครบ) |

**ห้ามใช้** `/tmsBranchTrackingDetail/page` + `/tmsnewBranchShipment/trackDetailExport` (เมนู "ติดตามสถานะการขนส่งเส้นรอง")
สำหรับเส้นรอง — คนละรายงาน คนละชุดใบงาน

## หมายเหตุเรื่องข้อมูลที่ export
- `_jms_headers()` — header ชุดเดียวที่ใช้ทั้ง Main และ Branch. `lang`/`langtype` = `TH` คือตัวที่ทำให้ JMS
  คืน **หัวคอลัมน์เป็นภาษาไทย** (ถ้าขาดไปจะได้หัวคอลัมน์จีน) และ `timezone: GMT+0700` ทำให้ตีความช่วงเวลาเป็นเวลาไทย
- ทั้งสอง export กรอง task จาก `/export/selectTask` ด้วย `exportUrls` + เทียบ `row["url"]` และ `createTime`
  กับเวลาที่เพิ่งสั่งงาน — ป้องกันไม่ให้ Main ไปหยิบไฟล์ของ Branch (หรือ task เก่า) มาเป็นผลลัพธ์
- ตัวกรองเวลา: Main ใช้ `timeType: 2` + `newTimeType: 1`, Branch ใช้ `timeType: 1` + `startDepartureTimeStr`/
  `endDepartureTimeStr` — ทั้งคู่กรองที่ **เวลาที่คาดว่าจะปล่อยรถ** (`plannedDepartureTime`)
- **ช่วงวันที่ของสองเส้นเป็นคนละรอบกัน** — มีตัวกรองเวลาแยก 2 ชุดบนหน้า Home
  ขอบช่วงเป็น `12:00 → 12:00` ทั้งคู่ ต่างกันที่ **วันพลิกตอนไหน**

  | | ช่วงข้อมูล | พลิกวันตอน | เหตุผล |
  |---|---|---|---|
  | **เส้นหลัก** | เมื่อวาน 12:00 → วันนี้ 12:00 | `MAIN_ROLLOVER` = **01:30** | เตรียมรถเมื่อวานเที่ยง–วันนี้เที่ยง ลงของหมดไม่เกินตี 1 พอถึง 01:30 ข้อมูลรอบนั้นครบแล้ว |
  | **เส้นรอง** | วันนี้ 12:00 → พรุ่งนี้ 12:00 | `BRANCH_ROLLOVER` = **12:00** | เตรียมรถวันนี้เที่ยง–พรุ่งนี้เที่ยง ลงของหมดไม่เกินเที่ยงของอีกวัน |

  ไทม์ไลน์ (สมมติวันที่ 10):
  ```
  09-10 01:29   หลัก 09-08 12:00→09-09 12:00   รอง 09-09 12:00→09-10 12:00
  09-10 01:30   หลัก 09-09 12:00→09-10 12:00   รอง 09-09 12:00→09-10 12:00  ← หลักพลิก (ตรงกันชั่วคราว)
  09-10 12:00   หลัก 09-09 12:00→09-10 12:00   รอง 09-10 12:00→09-11 12:00  ← รองพลิก
  ```
  ช่วง 01:30–12:00 ทั้งสองเส้นชี้ช่วงเดียวกัน — เป็นจังหวะที่เหมาะกับการสรุปรายวัน

  `_cycle_dates(rollover, lag_days)` เป็นตัวคำนวณกลาง เรียกผ่าน `_main_cycle_dates()` /
  `_branch_cycle_dates()`; `get_datetime_range("main"|"branch")` อ่านค่าจากช่องบน UI ตรง ๆ
  ปลายช่วงหักออก 1 วินาทีเสมอ (`11:59:59`) เพื่อไม่ให้คาบเกี่ยวกับรอบถัดไป
- **วันที่ขยับเองอัตโนมัติ** — `refresh_cycle_dates()` เช็คทั้งสองเส้นแยกกัน ขยับเฉพาะเส้นที่รอบเปลี่ยนจริง
  - `_watch_cycle_rollover()` ตั้ง `after(60s)` วนเช็คตัวเอง เปิดโปรแกรมค้างข้ามวันก็ขยับให้เอง
  - จุดที่เรียก: ตอนเปิดโปรแกรม (`_load_values_to_ui`), ทุก 60 วิ (watcher) และก่อนทุกรอบของ `scheduler_loop`
  - `_applied_cycle` เป็น dict แยกตามเส้น จำรอบล่าสุดที่เขียนลง UI ไปแล้ว
    ค่าที่ผู้ใช้ตั้งเองเพื่อดึงย้อนหลังจึงอยู่รอดจนกว่าจะพ้นเวลาพลิกของรอบถัดไป
- **config**: `main_start_date`/`main_end_date` + `start_hour`/`end_hour` = เส้นหลัก,
  `start_date`/`end_date` + `branch_start_hour`/`branch_end_hour` = เส้นรอง
  (ชื่อ `start_date`/`end_date` คงไว้ให้ config เก่าใช้ต่อได้)
- ตัวเลือกเวลาใน dropdown เป็นทีละครึ่งชั่วโมง (48 ตัว) เพราะต้องเลือก `01:30` ได้

## Excel → PNG: กับดักที่เจอมาแล้ว
- **template ฝัง path เต็มของเครื่องที่สร้างมัน** — Power Query ในไฟล์ `Excel/*.xlsx` เขียนไว้เป็น
  `Excel.Workbook(File.Contents("C:/Users/.../Transport_V2/code/report.xlsx"))`
  - ย้ายไปเครื่องอื่นแล้วหาไฟล์ไม่เจอ → Power Query เด้ง **modal dialog ที่ `DisplayAlerts=False` ปิดไม่ได้**
    → COM call บล็อก → Excel "ไม่มีการตอบสนอง" และบอทแขวนถาวร (ลูป `stop_checker` ไม่ถูกเรียกด้วยซ้ำ)
  - บนเครื่องเดิม path เก่ายังมีไฟล์ค้างอยู่ เลย refresh ผ่านแบบเงียบ ๆ **แต่อ่านข้อมูลเก่าผิดไฟล์**
  - แก้: `repoint_queries()` เขียน path ใหม่ทุกครั้งก่อน refresh (เปลี่ยนเฉพาะโฟลเดอร์ เก็บชื่อไฟล์เดิม)
- **`BackgroundQuery = True` ทำให้ refresh ผ่าน COM ค้าง** — ค่าเริ่มต้นของ Power Query
  `RefreshAll()` จะคืนค่าทันทีแล้วงานจริงไปทำเบื้องหลัง โค้ดต้องไปรอที่
  `CalculateUntilAsyncQueriesDone()` ซึ่งค้างยาวเพราะ Mashup engine ไม่ส่งสัญญาณจบ
  กลับมาในบริบท automation (กดรีเฟรชด้วยมือไม่เจอ เพราะ Excel มี message loop ของตัวเอง)
  - อาการ: log หยุดที่ `Refreshing data...` แล้วไม่ไปต่อ แต่เปิดไฟล์เองกดรีเฟรชกลับผ่านปกติ
  - แก้: `force_sync_refresh()` ปิด `BackgroundQuery` ทุก connection ก่อน `RefreshAll()`
    ทำให้บล็อกจนจบจริง ไม่ต้องพึ่ง `CalculateUntilAsyncQueriesDone` เลย
- **หน้าต่าง Excel ต้องอยู่ในจอตอน refresh** — โค้ดย้าย Excel ไป `-32000,-32000` เพื่อไม่ให้เกะกะ
  แต่ dialog ที่เด้งระหว่าง refresh จะไปโผล่นอกจอด้วย มองไม่เห็นและกดไม่ได้
  `_move_excel_onscreen()` ย้ายกลับเข้าจอช่วง refresh แล้วค่อยย้ายออกตอนถ่ายรูป
- **modal dialog ของ Excel หลุด `DisplayAlerts` ได้** — ต้องตั้ง `AskToUpdateLinks=False` และ
  `AutomationSecurity=3` เพิ่ม และมี `ExcelWatchdog` ฆ่า process จาก thread อื่นเป็นทางหนีสุดท้าย
- **retry ตอนเปิด Excel ต้องเก็บกวาดของเดิม** — `DispatchEx` อาจสำเร็จแล้วไปพังตอน set property
  ถ้าไม่ Quit/kill ก่อนวนใหม่ จะได้ EXCEL.EXE ผีสะสม และตัวแปร `excel` ที่ค้างจะหลุดเช็ค `if not excel`
- template เส้นรองยังเป็น query ของเมนูเก่า (61 คอลัมน์) — `Report2.xlsx` ตอนนี้ 53 คอลัมน์ ต้องรื้อ template

## DWS — รวมเวลาลงพัสดุแยกตามช่อง
เป้าหมาย: รู้ว่าแต่ละช่องลงพัสดุ (เครื่อง DWS 1-11) ใช้เวลารวมเท่าไหร่

**ตัวเชื่อม** — ไฟล์ DWS มี `taskNo` / `任务编号` ซึ่งคือ **หมายเลขใบงานตัวเดียวกับใน Report**
join ตรง ๆ ได้เลย ไม่ต้อง match ด้วยเวลาหรือยอดพัสดุ

**แหล่งข้อมูล**
| ช่อง | ที่มา | ระบุช่องจาก |
|------|-------|-------------|
| 1-8 | share ของแต่ละเครื่อง `\10.30.32.(31+N)\...ealtime\dwsN.xlsx` | ชื่อไฟล์ (1 ไฟล์ = 1 ช่อง) |
| 9-11 | ไฟล์รวมที่ bot อีกตัวสร้างไว้ในเครื่องนี้ | คอลัมน์ `dws序号` = IP `.220/.221/.222` |

path ทั้ง 9 ตั้งได้ในหน้า Setting การ์ด "DWS" (browse ทีละช่อง) เก็บใน `config.ini` หมวด `[DWS]`
ใช้งานจริงตอนนี้: ช่อง 2/3/4/8 + 9-11 — ที่เหลือเครื่องปิดหรือเลิกใช้

**ทำไมต้องสะสมลงคลัง** — ไฟล์ realtime เป็น **rolling window ~24 ชั่วโมง** ข้อมูลเก่ากว่านั้น
หลุดถาวร (พิสูจน์แล้วโดยเทียบไฟล์เดียวกัน 2 ช่วงเวลา แถวเก่าหายไปจริง)
พลาดรอบไหน = เสียข้อมูลวันนั้นถาวร จึงอ่านสะสมทุกรอบ scheduler

**ทำไมต้องเก็บ Report สะสมด้วย** — คอลัมน์เวลาลงพัสดุกรอกก็ต่อเมื่อรถลงของเสร็จแล้ว
ตอนดึงรอบแรกของ window ค่ายังว่าง ต้องดึงซ้ำรอบหลัง ๆ ถึงได้ค่า `upsert_tasks()` จึงเขียนทับ
เฉพาะใบที่มีเวลาแล้ว และสะสมข้ามวัน — cold start ครอบคลุม ~83% พอเข้าวันที่สองขึ้นไปได้ ~97%

**ช่วงเวลาของ Report ที่ใช้** — งาน DWS ดึง Report ที่ `MAIN_DAY_OFFSET` (-1) **ทั้งสองเส้น**
คนละช่วงกับที่ใช้ทำรายงาน Feishu ซึ่งเส้นรองมองไปข้างหน้า (offset 0) และยังไม่มีใครลงของ

**หน่วยเวลาไม่ตรงกัน** — เส้นหลัก `ระยะเวลาลงพัสดุ (ชั่วโมง)` เป็นชั่วโมง
เส้นรอง `ระยะเวลาขนถ่ายลงรถ(min）` เป็นนาที ต้องหาร 60 ก่อนรวม
ระวังอย่าไปจับ `รายะเวลาเข้าแถวขนถ่ายลงรถ(min）` ที่เป็นเวลารอคิว ชื่อใกล้กันมาก

**ใบงานที่ลงมากกว่า 1 ช่อง** — ~5% ของใบงาน (แต่ ~13% ของปริมาณสแกน) เป็นรถที่จอดลง 2 ช่อง
พร้อมกันจริง หารเวลาตามสัดส่วนจำนวนชิ้นที่สแกนได้แต่ละช่อง

**ค่า AY ปัดเศษ** — JMS ปัดวินาทีขึ้นเป็นนาทีเต็มก่อนแปลงเป็นชั่วโมง
(`AY = round(ceil(วินาที/60)/60, 2)` ตรง 57/57 ใบที่ตรวจ) บวกเกินจริงสูงสุด 59 วินาที/ใบ (~0.9%)
ถ้าต้องการละเอียดถึงวินาทีให้คำนวณจาก `จบงาน − เริ่มงาน` แทน

## ข้อควรระวัง
- ต้องรันบน Windows + Excel
- `Bot_Fei_Main.py` ของ V2 **ต่างจาก V1 มาก** — อย่า diff/merge ข้าม repo โดยไม่ตรวจ
