# STRUCTURE — Bot_Transport_V2

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เพิ่ม/ลบ/ย้ายไฟล์, เปลี่ยน logic ฟังก์ชัน/คลาส, เปลี่ยน config key, เพิ่ม/แก้ประเภท report หรือ JMS endpoint, หรือเปลี่ยน flow — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
**เวอร์ชัน 2 ของ `Bot_Transport`** — รีไรต์ `Bot_Fei_Main.py` ให้รองรับ **หลาย report** และส่งได้ทั้งรูปและไฟล์เข้า Feishu Chat:
- **Main Line Transport** export (JMS) — `report.xlsx`
- **Branch Line / Task Tracking** export (JMS อีก endpoint) — `Report2.xlsx` / `Branch_Task_Tracking.xlsx`
- แปลง Excel → PNG แล้วส่งเข้า Feishu Chat (รองรับ job sheets/range หลายชุด)

## วิธีรัน / Entry point
- รัน: `python Bot_Fei_Main.py` → คลาส `App(ctk.CTk)` (แท็บ Home / Setting)
- มาพร้อม template Excel ในโฟลเดอร์ `Excel/`

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `Bot_Fei_Main.py` | **ตัวหลัก (~2,200 บรรทัด)** — UI + scheduler + Feishu (token, `send_feishu_image_by_chat_id`, `send_feishu_file_by_chat_id`) + export Main/Branch transport + job config (sheets/range/output หลายชุด) |
| `createpng.py` | Excel → PNG ด้วย win32com (เหมือน `Bot_Transport/createpng.py`) |
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
- เปิด/ปิดงาน: `run_main_transport`, `run_branch_transport`, `run_feishu_chat`
- ไฟล์: `main_transport_filename` (report.xlsx), `branch_transport_filename` (Report2.xlsx)
- Feishu/PNG: `excel_file`, `excel_sheet_index`, `excel_range`, `chat_id`, `app_id`, `app_secret` ฯลฯ

## Dependencies / บริการภายนอก
- `customtkinter`, `tkcalendar`, `pywin32` (ต้องมี Microsoft Excel), `requests`
- JMS J&T (Main + Branch endpoints), Feishu OpenAPI

## ข้อควรระวัง
- ต้องรันบน Windows + Excel
- `Bot_Fei_Main.py` ของ V2 **ต่างจาก V1 มาก** — อย่า diff/merge ข้าม repo โดยไม่ตรวจ
