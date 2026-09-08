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
- **ช่วงวันที่ของสองเส้นไม่เท่ากัน** — `get_datetime_range(day_offset)`
  - **เส้นรอง (Branch)** = ช่วงที่ตั้งบน UI ตรง ๆ → *วันนี้ 12:00 → พรุ่งนี้ 12:00*
  - **เส้นหลัก (Main)** = เลื่อนย้อนหลัง `MAIN_DAY_OFFSET` (= -1) วัน → *เมื่อวาน 12:00 → วันนี้ 12:00*
  ปลายช่วงหักออก 1 วินาทีเสมอ (`11:59:59`) เพื่อไม่ให้คาบเกี่ยวกับรอบถัดไป
- **วันที่ขยับเองอัตโนมัติ** — `_cycle_dates()` คิดวันของรอบปัจจุบันจากเวลาจริง โดยถือว่า
  หนึ่งรอบเริ่มที่ `start_hour` (ปกติ 12:00) ของแต่ละวัน ก่อนถึงเวลานั้นยังนับเป็นรอบของเมื่อวาน
  - `refresh_cycle_dates()` เลื่อนวันบน DateEntry + เขียนกลับ config (main thread เท่านั้น)
  - `refresh_cycle_dates_threadsafe()` เวอร์ชันที่เรียกจาก scheduler thread ได้
  - `_watch_cycle_rollover()` ตั้ง `after(60s)` วนเช็คตัวเอง เปิดโปรแกรมค้างข้ามวันก็ขยับให้เอง
  - จุดที่เรียก: ตอนเปิดโปรแกรม (`_load_values_to_ui`), ทุก 60 วิ (watcher) และก่อนทุกรอบของ `scheduler_loop`
  - `_applied_cycle` จำรอบล่าสุดที่เขียนลง UI ไปแล้ว — ทำให้ขยับเฉพาะตอนรอบเปลี่ยนจริง
    ค่าที่ผู้ใช้ตั้งเองเพื่อดึงย้อนหลังจึงอยู่รอดจนกว่าจะพ้นเที่ยงรอบถัดไป
  ถ้าเทียบกับไฟล์ export มือแล้วจำนวนแถวไม่ตรง ให้ตรวจช่วงวันที่ใน Home ก่อนเสมอ
- Main ไม่ต้องส่ง `count` (ต่างจาก Branch) — `size: 100` ไม่ตัดจำนวนแถวที่ export (ทดสอบแล้วได้ครบ 116 แถว)

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

## ข้อควรระวัง
- ต้องรันบน Windows + Excel
- `Bot_Fei_Main.py` ของ V2 **ต่างจาก V1 มาก** — อย่า diff/merge ข้าม repo โดยไม่ตรวจ
