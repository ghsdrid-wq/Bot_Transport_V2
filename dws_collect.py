# -*- coding: utf-8 -*-
"""อ่านไฟล์ realtime ของเครื่อง DWS แล้วสะสมลงคลังในเครื่อง

ทำไมต้องสะสม
    ไฟล์ realtime ของแต่ละเครื่องเป็น rolling window ~24 ชั่วโมง ข้อมูลที่เก่ากว่านั้น
    หลุดออกไปถาวร ถ้าอ่านทีเดียวตอนสรุปแล้วบังเอิญพลาดรอบ วันนั้นจะหายเลย
    บอทจึงอ่านทุกรอบ scheduler แล้ว append เข้าคลัง กันซ้ำด้วย
    (ช่อง, บาร์โค้ด, เวลาสแกน, ใบงาน)

โครงสร้างไฟล์ต้นทาง
    ช่อง 1-8   share ของแต่ละเครื่อง 1 ไฟล์ = 1 ช่อง
               คอลัมน์ 单号 (บาร์โค้ด) / taskNo (ใบงาน) / 出秤时间 (เวลา)
    ช่อง 9-11  ไฟล์รวมที่ bot อีกตัวสร้างไว้ แยกช่องด้วย dws序号 ที่เป็น IP
               คอลัมน์ 条码 / 任务编号 / 扫描时间
"""

import os
import sqlite3
from datetime import datetime

import pandas as pd

# IP ของเครื่องเวอร์ชันใหม่ → หมายเลขช่อง
IP_TO_BAY = {
    "10.30.32.220": 9,
    "10.30.32.221": 10,
    "10.30.32.222": 11,
}

# (คอลัมน์บาร์โค้ด, คอลัมน์ใบงาน, คอลัมน์เวลา) ของแต่ละรูปแบบไฟล์
SCHEMAS = [
    ("单号", "taskNo", "出秤时间"),      # ช่อง 1-8 เครื่องเวอร์ชันเก่า
    ("条码", "任务编号", "扫描时间"),     # ช่อง 9-11 เครื่องเวอร์ชันใหม่
]

BAY_IP_COLUMN = "dws序号"
STORE_FILENAME = "dws_scans.db"


def store_path(folder):
    return os.path.join(folder, STORE_FILENAME)


def _connect(path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    con = sqlite3.connect(path, timeout=30)
    con.execute(
        """
        CREATE TABLE IF NOT EXISTS scans (
            bay     INTEGER NOT NULL,
            barcode TEXT    NOT NULL,
            task    TEXT    NOT NULL,
            scanned TEXT    NOT NULL,
            PRIMARY KEY (bay, barcode, scanned, task)
        )
        """
    )
    con.execute("CREATE INDEX IF NOT EXISTS idx_scans_time ON scans(scanned)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_scans_task ON scans(task)")
    con.commit()
    return con


def _read_source(path, fixed_bay=None):
    """อ่านไฟล์ 1 ไฟล์ คืน DataFrame: bay, barcode, task, scanned"""
    df = pd.read_excel(path, dtype=str).fillna("")

    schema = next(
        (s for s in SCHEMAS if s[0] in df.columns and s[2] in df.columns),
        None,
    )
    if schema is None:
        raise ValueError(f"ไม่รู้จักรูปแบบคอลัมน์ของไฟล์: {os.path.basename(path)}")
    barcode_col, task_col, time_col = schema

    if BAY_IP_COLUMN in df.columns:
        bays = df[BAY_IP_COLUMN].str.strip().map(IP_TO_BAY)
    elif fixed_bay is not None:
        bays = pd.Series(fixed_bay, index=df.index)
    else:
        raise ValueError(f"ระบุช่องของไฟล์ไม่ได้: {os.path.basename(path)}")

    scanned = pd.to_datetime(df[time_col], errors="coerce")
    out = pd.DataFrame({
        "bay": bays,
        "barcode": df[barcode_col].str.strip(),
        "task": df[task_col].str.strip() if task_col in df.columns else "",
        "scanned": scanned.dt.strftime("%Y-%m-%d %H:%M:%S"),
    })
    return out[out["bay"].notna() & scanned.notna()].astype({"bay": int})


def collect(paths, store_folder, log=print, stop_checker=None):
    """อ่านทุก path ที่ระบุแล้วสะสมลงคลัง

    ``paths`` เป็น dict {ช่อง: path} โดยช่อง 9-11 ใช้คีย์ ``"9-11"``
    คืน dict สรุปจำนวนแถวใหม่ต่อช่อง
    """
    db = store_path(store_folder)
    con = _connect(db)
    added = {}

    try:
        for key, path in sorted(paths.items(), key=lambda kv: str(kv[0])):
            if stop_checker and stop_checker():
                log("หยุดอ่านไฟล์ DWS ตามคำสั่งผู้ใช้")
                break
            path = (path or "").strip()
            if not path:
                continue

            label = f"ช่อง {key}"
            if not os.path.exists(path):
                # เครื่องปิด / share ล่ม — ข้ามไป ไม่ให้ทั้ง job พัง
                log(f"  {label}: เข้าไม่ถึง {path}")
                continue

            try:
                fixed = None if key == "9-11" else int(key)
                rows = _read_source(path, fixed_bay=fixed)
            except Exception as exc:
                log(f"  {label}: อ่านไม่ได้ — {exc}")
                continue

            before = con.execute("SELECT COUNT(*) FROM scans").fetchone()[0]
            con.executemany(
                "INSERT OR IGNORE INTO scans (bay, barcode, task, scanned) VALUES (?,?,?,?)",
                rows[["bay", "barcode", "task", "scanned"]].itertuples(index=False, name=None),
            )
            con.commit()
            after = con.execute("SELECT COUNT(*) FROM scans").fetchone()[0]

            new = after - before
            added[key] = new
            stamp = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%m-%d %H:%M")
            log(f"  {label}: อ่าน {len(rows):,} แถว (ไฟล์อัปเดต {stamp}) → เพิ่มใหม่ {new:,}")
    finally:
        con.close()

    return added


def load_scans(store_folder, start_time=None, end_time=None):
    """ดึงข้อมูลจากคลังตามช่วงเวลา คืน DataFrame: task, bay, scans"""
    db = store_path(store_folder)
    if not os.path.exists(db):
        return pd.DataFrame(columns=["task", "bay", "scans"])

    sql = "SELECT bay, task, COUNT(*) AS scans FROM scans WHERE task <> ''"
    args = []
    if start_time:
        sql += " AND scanned >= ?"
        args.append(start_time)
    if end_time:
        sql += " AND scanned <= ?"
        args.append(end_time)
    sql += " GROUP BY bay, task"

    con = _connect(db)
    try:
        df = pd.read_sql_query(sql, con, params=args)
    finally:
        con.close()
    return df[["task", "bay", "scans"]]


def store_summary(store_folder):
    """สรุปสถานะคลัง — ไว้แสดงใน log"""
    db = store_path(store_folder)
    if not os.path.exists(db):
        return "ยังไม่มีคลังข้อมูล"
    con = _connect(db)
    try:
        n, lo, hi = con.execute(
            "SELECT COUNT(*), MIN(scanned), MAX(scanned) FROM scans"
        ).fetchone()
    finally:
        con.close()
    if not n:
        return "คลังข้อมูลว่าง"
    size = os.path.getsize(db) / 1024 / 1024
    return f"คลังสะสม {n:,} สแกน | {lo} → {hi} | {size:.1f} MB"
