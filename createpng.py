import win32com.client as win32
import win32gui
import win32con
import pythoncom
import time, os

MIN_PNG_BYTES = 10_240   # 10 KB — ตาม Createphoto.py


def _pump(seconds=0.5):
    """Pump Windows message queue ให้ GDI flush (แบบ Createphoto.py)"""
    end = time.time() + max(0, seconds)
    while time.time() < end:
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        time.sleep(0.05)


def _move_excel_offscreen(excel):
    """
    ย้าย Excel window ออกนอกจอ (Left=-32000, Top=-32000)
    GDI ยังคง render ได้เต็มที่ แต่ผู้ใช้ไม่เห็น window บนหน้าจอ
    (เทคนิคจาก Createphoto.py — ดีกว่า Minimize/SW_HIDE ที่ทำให้รูปขาว)
    """
    try:
        excel.Visible     = True
        excel.WindowState = 2       # xlNormal
        excel.Left        = -32000
        excel.Top         = -32000
        excel.Width       = 800
        excel.Height      = 600
    except Exception:
        pass
    _pump(0.2)


def _hide_from_taskbar(excel):
    """
    ซ่อน Excel ออกจาก Taskbar โดยใช้ WS_EX_TOOLWINDOW
    (แสดงแค่ว่า Excel กำลังเปิดอยู่ใน Tab ไม่ขึ้น taskbar)
    """
    try:
        hwnd = excel.Hwnd
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(
            hwnd,
            win32con.GWL_EXSTYLE,
            ex_style | win32con.WS_EX_TOOLWINDOW
        )
    except Exception:
        pass


def _is_blank_image(path):
    """ตรวจว่ารูปขาวหรือเปล่า — ใช้ Pillow ถ้ามี ไม่งั้นใช้ file size"""
    if not os.path.exists(path):
        return True
    if os.path.getsize(path) < MIN_PNG_BYTES:
        return True
    try:
        from PIL import Image, ImageStat
        with Image.open(path) as img:
            img = img.convert("RGB")
            img.thumbnail((96, 96))
            stat = ImageStat.Stat(img)
            mean = sum(stat.mean) / 3
            variance = sum(stat.var) / 3
            return mean > 246 and variance < 45
    except Exception:
        return False


def run_create(
    excel_path,
    sheet_name,
    cell_range,
    output_path,
    report_date=None,
    log=None,
    stop_checker=None
):
    def write(msg):
        if log:
            log(msg)
        else:
            print(msg)

    def should_stop():
        return stop_checker() if stop_checker else False

    # ------------------------------------------------------------------
    # 0. Validate
    # ------------------------------------------------------------------
    if not excel_path or not os.path.exists(excel_path):
        raise FileNotFoundError(f"[createpng] ไม่พบไฟล์ Excel: {excel_path}")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    try:
        sheet_index = int(sheet_name)
    except (ValueError, TypeError):
        sheet_index = 2

    write("── Excel → PNG ─────────────────────")
    write(f"  Excel      : {os.path.basename(excel_path)}")
    write(f"  Sheet      : #{sheet_index}")
    write(f"  Range      : {cell_range}")
    write(f"  Output     : {os.path.basename(output_path)}")
    if report_date:
        write(f"  Date       : {report_date}")

    pythoncom.CoInitialize()
    excel = None
    wb    = None

    try:
        # ------------------------------------------------------------------
        # 1. เปิด Excel — Visible=True แต่ย้ายออกนอกจอทันที
        #    ไม่ใช้ Minimize/SW_HIDE เพราะทำให้ GDI ไม่ render → รูปขาว
        # ------------------------------------------------------------------
        for attempt in range(1, 4):
            try:
                excel = win32.DispatchEx("Excel.Application")
                excel.DisplayAlerts  = False
                excel.Visible        = True
                excel.ScreenUpdating = True
                excel.EnableEvents   = False
                _move_excel_offscreen(excel)   # ย้ายออกนอกจอก่อนเปิดไฟล์
                break
            except Exception as e:
                write(f"  Excel init failed (attempt {attempt}/3): {e}")
                time.sleep(2)

        if not excel:
            raise RuntimeError("[createpng] Excel ไม่สามารถเปิดได้")

        write("  Opening workbook...")
        wb = excel.Workbooks.Open(
            excel_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )

        _move_excel_offscreen(excel)
        _hide_from_taskbar(excel)   # ซ่อนออกจาก taskbar หลังเปิดไฟล์

        # ------------------------------------------------------------------
        # 2. เลือก Sheet
        # ------------------------------------------------------------------
        ws = wb.Worksheets(sheet_index)
        write(f"  Sheet      : {ws.Name}")

        # ------------------------------------------------------------------
        # 3. เปลี่ยนวันที่ B2
        # ------------------------------------------------------------------
        if report_date:
            ws.Range("B2").Value = report_date
            write(f"  Set date B2 → {report_date}")

        wb.Saved = True   # กันไม่ให้ Excel ถาม save dialog ตอนปิด

        if should_stop():
            return

        # ------------------------------------------------------------------
        # 4. Refresh + คำนวณ
        # ------------------------------------------------------------------
        write("  Refreshing data...")
        wb.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception as e:
            write(f"  CalculateUntilAsync warning: {e}")

        # รอ Excel คำนวณเสร็จ (แบบ wait_excel ใน Createphoto.py)
        write("  Waiting for Excel to calculate...")
        deadline = time.time() + 180
        last_log  = time.time()
        while time.time() < deadline:
            if should_stop():
                return
            if time.time() - last_log > 5:
                write("  Still calculating...")
                last_log = time.time()
            try:
                if excel.CalculateState == 0:
                    break
            except Exception:
                break
            time.sleep(0.5)

        _pump(1.0)

        if should_stop():
            return

        # ------------------------------------------------------------------
        # 5. ซ่อนแถว #N/A คอลัมน์ C
        # ------------------------------------------------------------------
        HEADER_ROWS = 4
        last_row = ws.Cells(ws.Rows.Count, 3).End(-4162).Row
        write(f"  Rows found : {last_row}")

        hidden_count = 0
        for row in range(HEADER_ROWS + 1, last_row + 1):
            if str(ws.Cells(row, 3).Text).strip() == "#N/A":
                ws.Rows(row).Hidden = True
                hidden_count += 1

        write(f"  Hidden rows: {hidden_count} (#N/A)")

        if should_stop():
            return

        # ------------------------------------------------------------------
        # 6. Export พร้อม retry (แบบ Createphoto.py)
        # ------------------------------------------------------------------
        tmp_path  = output_path + ".__tmp.png"
        base_path = output_path

        # ลบไฟล์เก่าก่อน
        for p in (base_path, tmp_path):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass

        copy_modes = [
            (1, 2, "screen-picture"),
            (2, 2, "printer-picture"),
            (1, 1, "screen-bitmap"),
            (2, 1, "printer-bitmap"),
        ]

        last_error = None
        export_ok  = False

        for attempt in range(1, 7):
            if should_stop():
                return

            appearance, fmt, mode_name = copy_modes[(attempt - 1) % len(copy_modes)]
            chart = None

            try:
                write(f"  Export [{attempt}/6] mode={mode_name}")

                _move_excel_offscreen(excel)
                ws.Activate()
                target = ws.Range(cell_range)
                target.Select()
                _pump(0.4)

                target.CopyPicture(Appearance=appearance, Format=fmt)
                _pump(0.8)

                # วาง chart ใกล้ range จริง (ไม่ใช้ 0,0 เพราะอาจ render ขาว)
                chart = ws.ChartObjects().Add(
                    target.Left,
                    target.Top,
                    max(float(target.Width) + 8, 120),
                    max(float(target.Height) + 8, 80),
                )
                chart.Activate()
                _pump(0.3)

                chart.Chart.Paste()
                _pump(0.8)

                # ตรวจว่า paste มีเนื้อหาจริง
                try:
                    if chart.Chart.Shapes.Count < 1:
                        raise Exception("Paste produced 0 shapes")
                except Exception as e:
                    raise Exception(f"Shape check failed: {e}")

                # ลบ border กันรูปเพี้ยน
                try:
                    chart.Chart.ChartArea.Border.LineStyle = 0
                except Exception:
                    pass

                ok = chart.Chart.Export(tmp_path, "PNG")
                _pump(0.4)

                if ok is False or not os.path.exists(tmp_path):
                    raise Exception("Chart.Export ไม่ได้ไฟล์")

                if _is_blank_image(tmp_path):
                    size = os.path.getsize(tmp_path) if os.path.exists(tmp_path) else 0
                    raise Exception(f"รูปขาว/blank ({size} bytes)")

                os.replace(tmp_path, base_path)
                png_size = os.path.getsize(base_path)
                write(f"  PNG ready  : {os.path.basename(base_path)} ({png_size:,} bytes)")
                export_ok = True
                break

            except Exception as e:
                last_error = e
                write(f"  Retry [{attempt}/6] failed — {e}")
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
                _pump(0.8)

            finally:
                try:
                    if chart is not None:
                        chart.Delete()
                except Exception:
                    pass

        if not export_ok:
            raise RuntimeError(f"[createpng] Export ล้มเหลวทุก attempt: {last_error}")

    finally:
        try:
            if wb:
                wb.Close(False)
        except Exception:
            pass  # workbook อาจถูกปิดไปแล้ว — ไม่ต้อง log
        try:
            if excel:
                time.sleep(0.5)
                excel.Quit()
                del excel
        except Exception as e:
            write(f"  excel.Quit warning: {e}")
        pythoncom.CoUninitialize()
