# -*- coding: utf-8 -*-
"""
Bot_Fei_Main.py - Combined JMS Export + Feishu Chat Sender
รวมจาก Bot_T.py + Bot_Fei_Main.py

สิ่งที่รวมแล้ว:
1) CustomTkinter UI: Home / Setting
2) Home มี DateEntry + เวลา + Run minute + checkbox Export JMS / Feishu Chat
3) Log แสดงผลผ่าน Scrollbar เดียวกัน
4) Config ทั้งหมดอยู่ใน Tab Setting และบันทึกลง config.ini
5) Feishu Chat เปลี่ยนจาก WebHook เป็น Chat ID ผ่าน OpenAPI
6) ถ้าเลือกรันทั้ง 2 งาน จะรัน Export JMS ก่อน แล้วค่อย Feishu Chat
"""

import configparser
import json
import logging
import os
import re
import sys
import threading
import time
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox
from urllib.parse import unquote


import customtkinter as ctk
import requests
from tkcalendar import DateEntry

try:
    from createpng import run_create
except Exception as exc:  # pragma: no cover
    run_create = None
    CREATEPNG_IMPORT_ERROR = exc
else:
    CREATEPNG_IMPORT_ERROR = None


APP_TITLE = "Feishu Auto Report - JMS + Chat"
LOG_FILE = "bot_fei_main.log"

MINUTES = [str(i) for i in range(60)]
RUN_HOURS = [str(i) for i in range(1, 25)]

HOURS = [f"{i:02d}:00" for i in range(24)]


# =========================
# A) PATH / CONFIG MODULE
# =========================
def resource_path(filename: str) -> str:
    """รองรับทั้งตอนรัน .py และตอน Pack เป็น .exe ด้วย PyInstaller"""
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(sys.executable), filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


CONFIG_FILE = resource_path("config.ini")


DEFAULT_CONFIG = {
    # Home / Scheduler
    "run_hour_interval": "1",
    "run_minute_interval": "5",
    "start_date": "",
    "end_date": "",
    "start_hour": "13:00",
    "end_hour": "23:00",
    "run_main_transport": "1",
    "run_branch_transport": "1",
    "run_feishu_chat": "1",

    # JMS Export
    "jms_auth_token": "",
    "jms_save_path": "",
    "main_transport_filename": "report.xlsx",
    "branch_transport_filename": "Report2.xlsx",

    # Feishu Chat / Excel to PNG
    # feishu_excel_jobs เก็บเป็น JSON เพื่อรองรับหลาย Excel และหลาย Sheet ต่อไฟล์
    "feishu_excel_jobs": "[]",
    "png_output_folder": "",
    "app_id": "",
    "app_secret": "",
    "chat_id": "",
    "jobs_file": "excel_jobs.json",

    # Legacy keys: เก็บไว้เพื่อ migrate config.ini เก่าอัตโนมัติ
    "excel_file": "",
    "excel_sheet_index": "2",
    "excel_range": "B2:V110",
    "png_filename": "report.png",
}



def _config_value(cfg: configparser.ConfigParser, section: str, key: str, fallback: str = "") -> str:
    """อ่าน config แบบปลอดภัย รองรับทั้ง section ใหม่และ config.ini รุ่นเก่า"""
    try:
        return cfg.get(section, key, fallback=fallback)
    except Exception:
        return fallback


def _resolve_sidecar_path(path_value: str, default_filename: str) -> str:
    """แปลง path ของไฟล์ config เสริมให้รองรับ relative path ข้าง .py/.exe"""
    path_value = (path_value or default_filename).strip() or default_filename
    if os.path.isabs(path_value):
        return path_value
    return resource_path(path_value)


def load_config() -> configparser.ConfigParser:
    """
    โหลด config แล้ว normalize กลับมาไว้ใน section SETTING เหมือนเดิม
    เพื่อไม่ต้องรื้อ UI/logic ทั้งไฟล์

    รองรับ:
    - config.ini แบบเก่า [SETTING]
    - config.ini แบบใหม่ [HOME] [JMS] [FEISHU]
    - excel_jobs.json แยกไฟล์
    """
    raw = configparser.ConfigParser()
    raw.read(CONFIG_FILE, encoding="utf-8")

    cfg = configparser.ConfigParser()
    cfg["SETTING"] = {}

    # เริ่มจากค่า default ทั้งหมด
    for key, value in DEFAULT_CONFIG.items():
        cfg["SETTING"][key] = value

    # อ่าน config.ini แบบเก่า [SETTING] ก่อน เพื่อ migration
    if raw.has_section("SETTING"):
        for key, value in raw.items("SETTING"):
            cfg["SETTING"][key] = value

    # อ่าน config.ini แบบใหม่ แล้ว map กลับเป็น key เดิม
    section_map = {
        "HOME": {
            "run_hour_interval": "run_hour_interval",
            "run_minute_interval": "run_minute_interval",
            "start_date": "start_date",
            "end_date": "end_date",
            "start_hour": "start_hour",
            "end_hour": "end_hour",
            "run_main_transport": "run_main_transport",
            "run_branch_transport": "run_branch_transport",
            "run_feishu_chat": "run_feishu_chat",
        },
        "JMS": {
            "auth_token": "jms_auth_token",
            "save_path": "jms_save_path",
            "main_transport_filename": "main_transport_filename",
            "branch_transport_filename": "branch_transport_filename",
        },
        "FEISHU": {
            "png_output_folder": "png_output_folder",
            "app_id": "app_id",
            "app_secret": "app_secret",
            "chat_id": "chat_id",
            "jobs_file": "jobs_file",
        },
    }

    for section, mapping in section_map.items():
        if raw.has_section(section):
            for new_key, old_key in mapping.items():
                if raw.has_option(section, new_key):
                    cfg["SETTING"][old_key] = raw.get(section, new_key)

    # ถ้ามี excel_jobs.json ให้ใช้เป็นแหล่งหลัก
    jobs_file = _resolve_sidecar_path(
        cfg["SETTING"].get("jobs_file", "excel_jobs.json"),
        "excel_jobs.json"
    )
    if os.path.exists(jobs_file):
        try:
            with open(jobs_file, "r", encoding="utf-8") as file:
                data = json.load(file)
            files = data.get("files", [])
            cfg["SETTING"]["feishu_excel_jobs"] = json.dumps(files, ensure_ascii=False)
        except Exception:
            # ถ้า JSON เสีย ให้ fallback ไปใช้ค่าจาก config.ini รุ่นเก่า
            pass

    return cfg


def _write_pretty_config(cfg: configparser.ConfigParser) -> None:
    """เขียน config.ini ให้อ่านง่าย แยกหมวด และไม่ยัด Excel Jobs ไว้ใน config.ini"""
    s = cfg["SETTING"]
    jobs_file = s.get("jobs_file", "excel_jobs.json").strip() or "excel_jobs.json"

    lines = [
        "# -*- coding: utf-8 -*-",
        "# Auto generated by Feishu Auto Report - JMS + Chat",
        "",
        "# ============================================================",
        "# FEISHU AUTO REPORT - JMS + CHAT",
        "# ============================================================",
        "",
        "[HOME]",
        "",
        "# Scheduler",
        f"run_hour_interval = {s.get('run_hour_interval', DEFAULT_CONFIG['run_hour_interval'])}",
        f"run_minute_interval = {s.get('run_minute_interval', DEFAULT_CONFIG['run_minute_interval'])}",
        "",
        "# Date Range",
        f"start_date = {s.get('start_date', '')}",
        f"end_date = {s.get('end_date', '')}",
        f"start_hour = {s.get('start_hour', DEFAULT_CONFIG['start_hour'])}",
        f"end_hour = {s.get('end_hour', DEFAULT_CONFIG['end_hour'])}",
        "",
        "# Jobs",
        f"run_main_transport = {s.get('run_main_transport', DEFAULT_CONFIG['run_main_transport'])}",
        f"run_branch_transport = {s.get('run_branch_transport', DEFAULT_CONFIG['run_branch_transport'])}",
        f"run_feishu_chat = {s.get('run_feishu_chat', DEFAULT_CONFIG['run_feishu_chat'])}",
        "",
        "",
        "[JMS]",
        "",
        f"auth_token = {s.get('jms_auth_token', '')}",
        f"save_path = {s.get('jms_save_path', '')}",
        "",
        f"main_transport_filename = {s.get('main_transport_filename', DEFAULT_CONFIG['main_transport_filename'])}",
        f"branch_transport_filename = {s.get('branch_transport_filename', DEFAULT_CONFIG['branch_transport_filename'])}",
        "",
        "",
        "[FEISHU]",
        "",
        f"png_output_folder = {s.get('png_output_folder', '')}",
        "",
        f"app_id = {s.get('app_id', '')}",
        f"app_secret = {s.get('app_secret', '')}",
        f"chat_id = {s.get('chat_id', '')}",
        "",
        "# Excel jobs are stored separately for readability",
        f"jobs_file = {jobs_file}",
        "",
    ]

    with open(CONFIG_FILE, "w", encoding="utf-8") as file:
        file.write("\n".join(lines))


def save_config(cfg: configparser.ConfigParser) -> None:
    folder = os.path.dirname(CONFIG_FILE)
    if folder:
        os.makedirs(folder, exist_ok=True)
    _write_pretty_config(cfg)


# =========================
# B) COMMON / FEISHU MODULE
# =========================
def clean_text(text: str) -> str:
    """ลบช่องว่าง/ขึ้นบรรทัดที่มักติดมากับ token, app secret, chat id"""
    return re.sub(r"\s+", "", str(text or "").replace("\n", "").replace("\r", "")).strip()


def request_with_retry(func, retries: int = 3, delay: int = 2, log=None, name: str = "request"):
    for attempt in range(1, retries + 1):
        try:
            return func()
        except Exception as exc:
            if log:
                log(f"{name} failed ({attempt}/{retries}): {exc}")
            if attempt >= retries:
                raise
            time.sleep(delay * attempt)


def get_tenant_access_token(app_id: str, app_secret: str, log=None) -> str:
    app_id = clean_text(app_id)
    app_secret = clean_text(app_secret)
    if not app_id or not app_secret:
        raise ValueError("กรุณาใส่ App ID และ App Secret ในหน้า Setting")

    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"

    def do_request():
        res = requests.post(url, json={"app_id": app_id, "app_secret": app_secret}, timeout=20)
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Get Feishu token")
    token = data.get("tenant_access_token")
    if not token:
        raise RuntimeError(f"Get Feishu token failed: {data}")
    if log:
        log("Feishu token OK")
    return token


def upload_feishu_image(token: str, image_path: str, log=None) -> str:
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"ไม่พบไฟล์รูปภาพ: {image_path}")

    url = "https://open.feishu.cn/open-apis/im/v1/images"
    headers = {"Authorization": f"Bearer {token}"}

    def do_request():
        with open(image_path, "rb") as file:
            res = requests.post(
                url,
                headers=headers,
                files={"image": file},
                data={"image_type": "message"},
                timeout=30,
            )
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Upload image")
    if data.get("code") != 0:
        raise RuntimeError(f"Upload image failed: {data}")
    image_key = data.get("data", {}).get("image_key")
    if not image_key:
        raise RuntimeError(f"Upload image missing image_key: {data}")
    return image_key


def send_feishu_image_by_chat_id(token: str, chat_id: str, image_key: str, log=None) -> None:
    chat_id = clean_text(chat_id)
    if not chat_id:
        raise ValueError("กรุณาใส่ Feishu Chat ID ในหน้า Setting")

    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }
    payload = {
        "receive_id": chat_id,
        "msg_type": "image",
        "content": json.dumps({"image_key": image_key}, ensure_ascii=False),
    }

    def do_request():
        res = requests.post(url, headers=headers, json=payload, timeout=30)
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Send image")
    if data.get("code") != 0:
        raise RuntimeError(f"Send image failed: {data}")
    if log:
        log("ส่งรูปเข้า Feishu Chat สำเร็จ ✅")


def upload_feishu_file(token: str, file_path: str, log=None) -> str:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"ไม่พบไฟล์ Excel: {file_path}")

    url = "https://open.feishu.cn/open-apis/im/v1/files"
    headers = {"Authorization": f"Bearer {token}"}
    filename = os.path.basename(file_path)

    def do_request():
        with open(file_path, "rb") as file:
            res = requests.post(
                url,
                headers=headers,
                files={"file": (filename, file)},
                data={"file_type": "stream", "file_name": filename},
                timeout=120,
            )
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Upload file")
    if data.get("code") != 0:
        raise RuntimeError(f"Upload file failed: {data}")
    file_key = data.get("data", {}).get("file_key")
    if not file_key:
        raise RuntimeError(f"Upload file missing file_key: {data}")
    return file_key


def send_feishu_file_by_chat_id(token: str, chat_id: str, file_key: str, log=None) -> None:
    chat_id = clean_text(chat_id)
    if not chat_id:
        raise ValueError("กรุณาใส่ Feishu Chat ID ในหน้า Setting")

    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }
    payload = {
        "receive_id": chat_id,
        "msg_type": "file",
        "content": json.dumps({"file_key": file_key}, ensure_ascii=False),
    }

    def do_request():
        res = requests.post(url, headers=headers, json=payload, timeout=30)
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Send file")
    if data.get("code") != 0:
        raise RuntimeError(f"Send file failed: {data}")
    if log:
        log("ส่งไฟล์ Excel เข้า Feishu Chat สำเร็จ ✅")


# =========================
# C) JMS EXPORT MODULE
# =========================
MAIN_TRACKING_COLUMNS = [
    "shipmentNo", "shipmentState", "shipmentName", "gxType", "businessAttribute",
    "shifts", "operationModel", "billingWay", "shipmentType", "plateNumber",
    "plateNumberProvince", "trailerNumber", "vehicleBelongName", "vehicleOrigin",
    "driverName", "sendNetworkCode", "sendNetworkName", "loadingScanStartTime",
    "loadingScanEndTime", "loadCount", "loadingScanTotalTimeShow", "scanTime",
    "standByTime", "plannedDepartureTime", "actualDepartureTime", "delayTimeShow",
    "departureLate", "trackOutTime", "stopTimeShow", "actualStopTimeShow",
    "delayStopTimeShow", "arriveNetworkCode", "arriveNetworkName", "plannedArrivalTime",
    "predictArriveTime", "actualArrivalTime", "tardyTimeShow", "arrivelLate",
    "trackInTime", "carrierName", "carrierShortName", "carrierType", "useTimeShow",
    "actualUseTimeShow", "useWayTimeShow", "runningLate", "unScanTime",
    "unLoadLineTimeShow", "unLoadingScanStartTime", "unLoadingScanEndTime",
    "unLoadCount", "unLoadingScanTotalTimeShow", "unLoadTime", "vehiclelineCode",
    "vehiclelineName", "isAssistLine", "vehicleTypegroup", "vehicletypeName",
    "loadWeight", "loadCapacity", "vehicleDoorCnt", "mileage", "overtimeType",
    "overtimeReasons", "quotationModel", "freightCode", "arriveProvince",
    "oriRegShiftCarrierName", "auditStatus", "auditRemark", "auditer",
]

BRANCH_TRACKING_COLUMNS = [
    "shipmentNo", "shipmentState", "shipmentName", "shipRegionName",
    "refLineRegionName", "sendRegionName", "arriveRegionName",
    "businessAttribute", "distributionType", "shifts", "mileage",
    "plateNumber", "plateNumberProvince", "trailerNumberProvince",
    "vehicletypeName", "carrierName", "driverName", "driverContact",
    "sendNetworkCode", "sendNetworkName", "sendAllianceBusiness",
    "sendRegionalAgent", "startProvince", "loadingScanStartTime",
    "loadingScanEndTime", "loadingScanTotalTime", "isLoadingExceedTwenty",
    "scanTime", "plannedDepartureTime", "standByTime", "goodsSignTime",
    "actualDepartureTime", "delayTime", "departureLate", "loadStartTime",
    "driverSignOutTime", "arriveNetworkCode", "arriveNetworkName",
    "arriveAllianceBusiness", "arriveRegionalAgent", "endProvince",
    "plannedArrivalTime", "actualArrivalTime", "tardyTime", "arrivelLate",
    "useWayTime", "runningLate", "driverPunchClockTime", "unScanTime",
    "unLoadLineTime", "unLoadingScanStartTime", "unLoadingScanEndTime",
    "unLoadingScanTotalTime",
]

def export_jms_excel(auth_token: str, save_folder: str, filename: str, start_time: str, end_time: str, stop_checker, log) -> str:
    auth_token = clean_text(auth_token)
    if not auth_token:
        raise ValueError("กรุณาใส่ JMS Auth Token ในหน้า Setting")
    if not save_folder:
        raise ValueError("กรุณาเลือก JMS Save Path ในหน้า Setting")

    os.makedirs(save_folder, exist_ok=True)
    filename = filename.strip() or "report.xlsx"
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    base = "https://jmsgw.jtexpress.co.th/transportation"
    headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": auth_token}
    payload = {
        "current": 1,
        "size": 100,
        "timeType": 2,
        "newTimeType": 1,
        "startTime": start_time,
        "endTime": end_time,
        "countryId": "1",
        "arriveNetworkCodeList": ["999004"],
        "sendNetworkCodeList": [],
        "columnList": MAIN_TRACKING_COLUMNS,
    }

    log(f"สร้างงาน Export Main Transport: {start_time} → {end_time}")

    def create_job():
        res = requests.post(f"{base}/tmsExportTransportReport/reportExport", json=payload, headers=headers, timeout=30)
        res.raise_for_status()
        return res

    request_with_retry(create_job, log=log, name="Create JMS export job")
    time.sleep(5)

    task = None
    for _ in range(30):
        if stop_checker():
            log("หยุด Export JMS ตามคำสั่งผู้ใช้")
            return ""
        time.sleep(1)

        def get_tasks():
            res = requests.post(f"{base}/export/selectTask", json={"current": 1, "size": 20}, headers=headers, timeout=30)
            res.raise_for_status()
            return res.json()

        data = request_with_retry(get_tasks, log=log, name="Check JMS export task")
        for row in data.get("data", {}).get("records", []):
            if row.get("state") == 2 and row.get("ossUrl"):
                task = row
                break
        if task:
            break

    if not task:
        raise RuntimeError("ไม่พบไฟล์ JMS ที่ Export เสร็จภายในเวลาที่กำหนด")

    download_url = f"https://yl-file.jtexpress.co.th/{task.get('ossUrl')}"
    output_path = os.path.join(save_folder, filename)

    log("ดาวน์โหลดไฟล์ Main Transport..")

    def download_file():
        res = requests.get(download_url, timeout=120)
        res.raise_for_status()
        return res.content

    content = request_with_retry(download_file, log=log, name="Download JMS file")
    with open(output_path, "wb") as file:
        file.write(content)

    log(f"Export Main Transport สำเร็จ: {output_path}")
    return output_path

def export_branch_tracking(
    auth_token,
    start_time,
    end_time,
    save_folder="output",
    filename="Branch_Task_Tracking.xlsx",
    stop_checker=None,
    log=print
):
    auth_token = clean_text(auth_token)
    if not auth_token:
        raise ValueError("กรุณาใส่ JMS Auth Token ในหน้า Setting")
    if not save_folder:
        raise ValueError("กรุณาเลือก JMS Save Path ในหน้า Setting")

    os.makedirs(save_folder, exist_ok=True)
    filename = filename.strip() or "Branch_Task_Tracking.xlsx"
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    base = "https://jmsgw.jtexpress.co.th/transportation"
    export_url = "/transportation/tmsExportTransportReport/branchReportExport"
    export_task_name = "支线运输综合报表导出"
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "authtoken": auth_token,
        "lang": "TH",
        "langtype": "TH",
        "routename": "transportSynthesizeReport",
        "origin": "https://jms.jtexpress.co.th",
        "referer": "https://jms.jtexpress.co.th/",
    }
    count_payload = {
        "current": 1,
        "size": 100,
        "timeType": 1,
        "startTime": start_time,
        "endTime": end_time,
        "countryId": "1",
        "arriveNetworkCodeList": ["999004"],
        "arriveRegionalAgentList": [],
        "carrierId": "",
        "carrierName": "",
        "sendNetworkCodeList": [],
        "sendRegionalAgentList": [],
        "source": 2,
    }
    payload = {
        **count_payload,
        "columnList": BRANCH_TRACKING_COLUMNS,
    }

    log(f"สร้างงาน Export Branch Transport: {start_time} → {end_time}")
    job_started_at = (datetime.now() - timedelta(seconds=10)).strftime("%Y-%m-%d %H:%M:%S")

    def create_job():
        res = requests.post(
            f"{base}/tmsExportTransportReport/branchReportExport",
            json=payload,
            headers=headers,
            timeout=60,
        )
        res.raise_for_status()
        return res

    request_with_retry(create_job, log=log, name="Create Branch export job")
    time.sleep(5)

    task = None
    for _ in range(300):
        if stop_checker and stop_checker():
            log("หยุด Export Branch Transport ตามคำสั่งผู้ใช้")
            return ""
        time.sleep(2)

        def get_tasks():
            task_payload = {
                "current": 1,
                "size": 20,
                "countryId": "1",
                "startTime": datetime.now().strftime("%Y-%m-%d 00:00:00"),
                "endTime": datetime.now().strftime("%Y-%m-%d 23:59:59"),
                "exportUrls": [export_url],
            }
            res = requests.post(f"{base}/export/selectTask", json=task_payload, headers=headers, timeout=60)
            res.raise_for_status()
            return res.json()

        data = request_with_retry(get_tasks, log=log, name="Check Branch export task")
        for row in data.get("data", {}).get("records", []):
            create_time = row.get("createTime") or ""
            is_current_job = not create_time or create_time >= job_started_at
            if row.get("state") == 2 and row.get("ossUrl") and row.get("url") == export_url and is_current_job:
                task = row
                break
        if task:
            break

    if not task:
        raise RuntimeError("Export Branch Transport timeout - task not completed")

    output_path = os.path.join(save_folder, filename)
    log("ดาวน์โหลดไฟล์ Branch Transport...")

    def get_download_url():
        signed_payload = {
            "countryId": "1",
            "data": task.get("ossUrl"),
            "taskName": task.get("taskName") or export_task_name,
        }
        res = requests.post(f"{base}/file/oss/getDownloadSignedUrl", json=signed_payload, headers=headers, timeout=60)
        res.raise_for_status()
        data = res.json()
        signed_url = data.get("data")
        if not signed_url:
            raise RuntimeError(f"Get Branch download URL failed: {data}")
        return signed_url

    download_url = request_with_retry(get_download_url, log=log, name="Get Branch download URL")

    def download_file():
        res = requests.get(download_url, timeout=300)
        res.raise_for_status()
        return res.content

    content = request_with_retry(download_file, log=log, name="Download Branch file")
    with open(output_path, "wb") as f:
        f.write(content)

    log(f"Export Branch Transport สำเร็จ: {output_path}")
    return output_path
# =========================
# D) UI / MAIN APP MODULE
# =========================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("1200x780")
        #self.resizable(False, False)

        self.cfg = load_config()
        if not os.path.exists(CONFIG_FILE):
            save_config(self.cfg)
        self.scheduler_running = False
        self.job_running = False
        self.stop_requested = False
        self.last_auto_key = ""
        self.active_mode = None

        logging.basicConfig(filename=resource_path(LOG_FILE), level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", encoding="utf-8")

        self.widgets = {}
        self.feishu_jobs = []
        self.feishu_job_widgets = []
        self.feishu_jobs_frame = None
        self._build_ui()
        self._load_values_to_ui()

    # ---------- UI helpers ----------
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(
            self,
            corner_radius=0,
            anchor="w",
            fg_color="#0a0f1e",
            segmented_button_fg_color="#0a0f1e",
            segmented_button_selected_color="#1a2744",
            segmented_button_selected_hover_color="#1e2f55",
            segmented_button_unselected_color="#0a0f1e",
            segmented_button_unselected_hover_color="#111827",
        )
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.tabview.add("  Home  ")
        self.tabview.add("  Setting  ")

        self._build_home(self.tabview.tab("  Home  "))
        self._build_setting(self.tabview.tab("  Setting  "))

    # ── colour tokens ────────────────────────────────────────────────
    C = {
        "bg"       : "#0a0f1e",
        "surface"  : "#0f1629",
        "card"     : "#131d35",
        "border"   : "#1e2d4a",
        "accent"   : "#3b82f6",
        "accent2"  : "#06b6d4",
        "success"  : "#10b981",
        "danger"   : "#ef4444",
        "warn"     : "#f59e0b",
        "text"     : "#e2e8f0",
        "muted"    : "#64748b",
        "log_bg"   : "#060c1a",
    }

    def _card(self, parent, **kw):
        return ctk.CTkFrame(
            parent,
            fg_color=self.C["card"],
            corner_radius=16,
            border_width=1,
            border_color=self.C["border"],
            **kw
        )

    def _label(self, parent, text, size=13, weight="normal", color=None, **kw):
        return ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(family="Segoe UI", size=size, weight=weight),
            text_color=color or self.C["text"],
            **kw
        )

    def _btn(self, parent, text, command, color=None, hover=None, width=120, height=38):
        c = color or self.C["accent"]
        h = hover or "#2563eb"
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=height,
            fg_color=c,
            hover_color=h,
            corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        )

    def _build_home(self, parent):
        parent.configure(fg_color=self.C["bg"])
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        # ── Header bar ─────────────────────────────────────────────
        header = ctk.CTkFrame(parent, fg_color=self.C["surface"], corner_radius=0, height=64)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        dot_frame = ctk.CTkFrame(header, fg_color="transparent")
        dot_frame.grid(row=0, column=0, padx=20, pady=0, sticky="w")
        for i, col in enumerate(["#ef4444", "#f59e0b", "#10b981"]):
            ctk.CTkFrame(dot_frame, width=12, height=12, corner_radius=6, fg_color=col).grid(row=0, column=i, padx=3, pady=26)

        self._label(
            header,
            "BOT  JMSKKN",
            size=15, weight="bold",
            color=self.C["text"],
        ).grid(row=0, column=1, padx=0, pady=20, sticky="w")

        self.status_badge = ctk.CTkLabel(
            header,
            text="● IDLE",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.C["success"],
            fg_color="#0d2318",
            corner_radius=20,
            width=100,
            height=28,
        )
        self.status_badge.grid(row=0, column=2, padx=20, pady=18, sticky="e")

        # ── Control card — single row ──────────────────────────────
        # Interval | divider | [Export JMS] [Feishu Chat] | spacer | [Start Auto] [Run Now]
        ctrl = self._card(parent)
        ctrl.grid(row=1, column=0, sticky="ew", padx=20, pady=(16, 8))
        ctrl.grid_columnconfigure(9, weight=1)   # spacer pushes buttons to right

        # Interval
        self._label(ctrl, "Interval", size=12, color=self.C["muted"]).grid(
            row=0, column=0, padx=(18, 6), pady=16, sticky="w")
        self.run_hour_interval = ctk.CTkComboBox(
            ctrl, values=RUN_HOURS, width=66, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.run_hour_interval.grid(row=0, column=1, padx=3, pady=16)
        self._label(ctrl, "hr", size=12, color=self.C["muted"]).grid(
            row=0, column=2, padx=(2, 4), pady=16)
        self.run_minute_interval = ctk.CTkComboBox(
            ctrl, values=MINUTES, width=66, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.run_minute_interval.grid(row=0, column=3, padx=3, pady=16)
        self._label(ctrl, "min", size=12, color=self.C["muted"]).grid(
            row=0, column=4, padx=(2, 16), pady=16)

        # Divider
        ctk.CTkFrame(ctrl, width=1, height=28, fg_color=self.C["border"]).grid(
            row=0, column=5, padx=4, pady=16)

        # Checkboxes (ก่อน buttons)
        self.var_main_transport = ctk.BooleanVar(value=True)
        self.var_branch_transport = ctk.BooleanVar(value=True)
        self.var_feishu = ctk.BooleanVar(value=True)
        chk_style = dict(
            font=ctk.CTkFont(size=13),
            checkbox_width=18, checkbox_height=18,
            corner_radius=5,
            fg_color=self.C["accent"],
            border_color=self.C["border"],
            hover_color="#2563eb",
            text_color=self.C["text"],
        )
        ctk.CTkCheckBox(
            ctrl,
            text="Main Line Transport",
            variable=self.var_main_transport,
            **chk_style
        ).grid(row=0, column=6, padx=(16,10), pady=16)

        ctk.CTkCheckBox(
            ctrl,
            text="Branch Line Transport",
            variable=self.var_branch_transport,
            **chk_style
        ).grid(row=0, column=7, padx=(0,10), pady=16)

        ctk.CTkCheckBox(
            ctrl,
            text="Feishu Chat",
            variable=self.var_feishu,
            **chk_style
        ).grid(row=0, column=8, padx=(0,8), pady=16)

        # column=8 = spacer (weight=1)

        # Buttons (ขวาสุด)
        self.btn_start = self._btn(ctrl, "▶  Start Auto", self.start_scheduler, width=140)
        self.btn_start.grid(row=0, column=10, padx=(0, 8), pady=16)
        self.btn_run = self._btn(ctrl, "⚡  Run Now", self.run_now,
                                 color=self.C["success"], hover="#059669", width=130)
        self.btn_run.grid(row=0, column=11, padx=(0, 18), pady=16)

        # ── Date range card ────────────────────────────────────────
        date_card = self._card(parent)
        date_card.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        date_card.grid_columnconfigure(7, weight=1)

        self._label(date_card, "DATE RANGE", size=10, weight="bold", color=self.C["muted"]).grid(
            row=0, column=0, columnspan=8, sticky="w", padx=18, pady=(14, 2))

        # Start
        self._label(date_card, "Start", size=12, color=self.C["muted"]).grid(row=1, column=0, padx=(18, 6), pady=(4, 14), sticky="w")
        self.start_date = DateEntry(
            date_card, width=13, date_pattern="yyyy-mm-dd", state="readonly",
            background="#1a2744", foreground="#e2e8f0",
            selectbackground="#3b82f6", selectforeground="#ffffff",
            font=("Segoe UI", 12),
        )
        self.start_date.grid(row=1, column=1, padx=4, pady=(4, 14))
        self.start_hour = ctk.CTkComboBox(
            date_card, values=HOURS, width=96, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.start_hour.grid(row=1, column=2, padx=4, pady=(4, 14))

        # Arrow
        self._label(date_card, "→", size=16, color=self.C["accent2"]).grid(row=1, column=3, padx=12, pady=(4, 14))

        # End
        self._label(date_card, "End", size=12, color=self.C["muted"]).grid(row=1, column=4, padx=(0, 6), pady=(4, 14), sticky="w")
        self.end_date = DateEntry(
            date_card, width=13, date_pattern="yyyy-mm-dd", state="readonly",
            background="#1a2744", foreground="#e2e8f0",
            selectbackground="#3b82f6", selectforeground="#ffffff",
            font=("Segoe UI", 12),
        )
        self.end_date.grid(row=1, column=5, padx=4, pady=(4, 14))
        self.end_hour = ctk.CTkComboBox(
            date_card, values=HOURS, width=96, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.end_hour.grid(row=1, column=6, padx=4, pady=(4, 14))

        # Status pill
        self.status = ctk.CTkLabel(
            date_card,
            text="Status: Idle",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=self.C["muted"],
            anchor="w",
        )
        self.status.grid(row=1, column=7, sticky="ew", padx=(16, 18), pady=(4, 14))

        # ── Live Log card ──────────────────────────────────────────
        log_card = self._card(parent)
        log_card.grid(row=3, column=0, sticky="nsew", padx=20, pady=(8, 20))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 4))
        log_header.grid_columnconfigure(0, weight=1)

        self._label(log_header, "LIVE LOG", size=10, weight="bold", color=self.C["muted"]).grid(
            row=0, column=0, sticky="w")

        ctk.CTkButton(
            log_header,
            text="Clear",
            width=60, height=26,
            fg_color=self.C["surface"],
            hover_color=self.C["border"],
            text_color=self.C["muted"],
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            command=lambda: self.log_box.delete("1.0", "end"),
        ).grid(row=0, column=1, sticky="e")

        self.log_box = ctk.CTkTextbox(
            log_card,
            wrap="none",
            font=("Consolas", 12),
            fg_color=self.C["log_bg"],
            text_color="#93c5fd",
            corner_radius=10,
            scrollbar_button_color=self.C["border"],
            scrollbar_button_hover_color=self.C["accent"],
        )
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))

        # tag colours
        try:
            self.log_box.tag_config("INFO",    foreground="#93c5fd")
            self.log_box.tag_config("SUCCESS", foreground="#86efac")
            self.log_box.tag_config("WARN",    foreground="#fbbf24")
            self.log_box.tag_config("ERROR",   foreground="#fca5a5")
            self.log_box.tag_config("START",   foreground="#67e8f9")
        except Exception:
            pass

    def _setting_entry(self, parent, row, label, key, browse=None, show=None):
        self._label(parent, label, size=12, color=self.C["muted"], width=160, anchor="w").grid(
            row=row, column=0, padx=(18, 8), pady=7, sticky="w")
        entry = ctk.CTkEntry(
            parent,
            show=show,
            fg_color=self.C["surface"],
            border_color=self.C["border"],
            text_color=self.C["text"],
            placeholder_text_color=self.C["muted"],
            corner_radius=8,
            height=36,
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        entry.grid(row=row, column=1, padx=8, pady=7, sticky="ew")
        self.widgets[key] = entry

        def paste_clean(event=None):
            try:
                text = clean_text(self.clipboard_get())
                entry.delete(0, "end")
                entry.insert(0, text)
            except Exception:
                pass
            return "break"

        entry.bind("<Control-v>", paste_clean)
        entry.bind("<Return>", lambda _e: self.save_from_ui())
        entry.bind("<FocusOut>", lambda _e: self.save_from_ui(silent=True))

        if browse:
            ctk.CTkButton(
                parent, text="Browse", width=88, height=36,
                fg_color=self.C["surface"],
                hover_color=self.C["border"],
                text_color=self.C["muted"],
                border_width=1,
                border_color=self.C["border"],
                corner_radius=8,
                font=ctk.CTkFont(size=12),
                command=lambda: self._browse_to_entry(key, browse),
            ).grid(row=row, column=2, padx=(8, 18), pady=7)
        return entry

    def _section_title(self, parent, row, icon, title):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=row, column=0, columnspan=3, sticky="ew", padx=14, pady=(18, 4))
        self._label(f, icon, size=16, color=self.C["accent"]).grid(row=0, column=0, padx=(4, 8))
        self._label(f, title, size=14, weight="bold", color=self.C["text"]).grid(row=0, column=1, sticky="w")
        ctk.CTkFrame(f, height=1, fg_color=self.C["border"]).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))

    # ---------- Feishu Excel Manager ----------
    def _empty_sheet_rule(self, index: int = 1):
        return {
            "enabled": True,
            "no": str(index),
            "sheet": "1",
            "range": "A1:Z50",
            "output": f"report_{index}.png",
            "hide": False,
            "send": True,
        }

    def _normalize_feishu_jobs(self, jobs):
        normalized = []
        if not isinstance(jobs, list):
            return normalized

        for job in jobs:
            if not isinstance(job, dict):
                continue
            sheets = job.get("sheets", [])
            if not isinstance(sheets, list):
                sheets = []
            norm_sheets = []
            for idx, sheet in enumerate(sheets, start=1):
                if not isinstance(sheet, dict):
                    continue
                norm_sheets.append({
                    "enabled": bool(sheet.get("enabled", True)),
                    "no": str(sheet.get("no", idx)),
                    "sheet": str(sheet.get("sheet", "1")),
                    "range": str(sheet.get("range", "A1:Z50")),
                    "output": str(sheet.get("output", f"report_{idx}.png")),
                    "hide": bool(sheet.get("hide", sheet.get("delete", False))),
                    "send": bool(sheet.get("send", True)),
                })
            normalized.append({
                "enabled": bool(job.get("enabled", True)),
                "send_excel": bool(job.get("send_excel", False)),
                "excel_file": str(job.get("excel_file", "")),
                "sheets": norm_sheets or [self._empty_sheet_rule(1)],
            })
        return normalized

    def _load_feishu_jobs_from_config(self):
        s = self.cfg["SETTING"]

        jobs_file = _resolve_sidecar_path(
            s.get("jobs_file", "excel_jobs.json"),
            "excel_jobs.json"
        )

        jobs = []

        # 1) ใช้ excel_jobs.json เป็นหลัก
        if os.path.exists(jobs_file):
            try:
                with open(jobs_file, "r", encoding="utf-8") as file:
                    data = json.load(file)
                jobs = data.get("files", [])
            except Exception as exc:
                self.write_log(f"อ่าน excel_jobs.json ไม่สำเร็จ: {exc}", level="WARN")

        # 2) fallback config.ini รุ่นเก่า ที่เคยเก็บ JSON ไว้ใน feishu_excel_jobs
        if not jobs:
            raw = s.get("feishu_excel_jobs", "[]")
            try:
                jobs = json.loads(raw) if raw else []
            except Exception:
                jobs = []

        jobs = self._normalize_feishu_jobs(jobs)

        # 3) migrate config เก่าอัตโนมัติ ถ้าเคยตั้ง Source Excel เดิมไว้
        legacy_excel = s.get("excel_file", "").strip()
        if not jobs and legacy_excel:
            jobs = [{
                "enabled": True,
                "send_excel": False,
                "excel_file": legacy_excel,
                "sheets": [{
                    "enabled": True,
                    "no": "1",
                    "sheet": s.get("excel_sheet_index", "2"),
                    "range": s.get("excel_range", "B2:V110"),
                    "output": s.get("png_filename", "report.png"),
                    "hide": False,
                    "send": True,
                }]
            }]

        self.feishu_jobs = jobs


    def _collect_feishu_jobs_from_ui(self):
        jobs = []
        for job_ui in self.feishu_job_widgets:
            excel_file = job_ui["excel_entry"].get().strip()
            sheets = []
            for sheet_ui in job_ui["sheets"]:
                output = sheet_ui["output"].get().strip() or "report.png"
                if not output.lower().endswith(".png"):
                    output += ".png"
                sheets.append({
                    "enabled": bool(sheet_ui["enabled"].get()),
                    "no": sheet_ui["no"].get().strip() or "1",
                    "sheet": sheet_ui["sheet"].get().strip() or "1",
                    "range": sheet_ui["range"].get().strip() or "A1:Z50",
                    "output": output,
                    "hide": bool(sheet_ui["hide"].get()),
                    "send": bool(sheet_ui["send"].get()),
                })
            jobs.append({
                "enabled": bool(job_ui["enabled"].get()),
                "send_excel": bool(job_ui["send_excel"].get()),
                "excel_file": excel_file,
                "sheets": sheets,
            })
        self.feishu_jobs = self._normalize_feishu_jobs(jobs)
        return self.feishu_jobs

    def _save_feishu_jobs_to_config(self):
        jobs = self._collect_feishu_jobs_from_ui() if self.feishu_job_widgets else self.feishu_jobs
        jobs = self._normalize_feishu_jobs(jobs)

        self.feishu_jobs = jobs

        jobs_filename = self.cfg["SETTING"].get("jobs_file", "excel_jobs.json").strip() or "excel_jobs.json"
        self.cfg["SETTING"]["jobs_file"] = jobs_filename

        jobs_path = _resolve_sidecar_path(
            jobs_filename,
            "excel_jobs.json"
        )

        folder = os.path.dirname(jobs_path)
        if folder:
            os.makedirs(folder, exist_ok=True)

        payload = {
            "files": jobs
        }

        with open(jobs_path, "w", encoding="utf-8") as file:
            json.dump(
                payload,
                file,
                ensure_ascii=False,
                indent=4
            )

        # ไม่เขียน JSON ก้อนใหญ่ลง config.ini แล้ว
        self.cfg["SETTING"]["feishu_excel_jobs"] = "[]"


    def _browse_excel_for_job(self, index: int):
        paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        )
        if not paths:
            return
        self._collect_feishu_jobs_from_ui()
        first = paths[0]
        if 0 <= index < len(self.feishu_jobs):
            self.feishu_jobs[index]["excel_file"] = first
        for extra in paths[1:]:
            self.feishu_jobs.append({
                "enabled": True,
                "send_excel": False,
                "excel_file": extra,
                "sheets": [self._empty_sheet_rule(1)],
            })
        self._render_feishu_jobs()
        self.save_from_ui(silent=True)

    def _add_feishu_excel(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        )
        if not paths:
            return
        self._collect_feishu_jobs_from_ui()
        for path in paths:
            self.feishu_jobs.append({
                "enabled": True,
                "send_excel": False,
                "excel_file": path,
                "sheets": [self._empty_sheet_rule(1)],
            })
        self._render_feishu_jobs()
        self.save_from_ui(silent=True)

    def _remove_feishu_excel(self, index: int):
        self._collect_feishu_jobs_from_ui()
        if 0 <= index < len(self.feishu_jobs):
            del self.feishu_jobs[index]
        self._render_feishu_jobs()
        self.save_from_ui(silent=True)

    def _add_feishu_sheet(self, index: int):
        self._collect_feishu_jobs_from_ui()
        if 0 <= index < len(self.feishu_jobs):
            next_no = len(self.feishu_jobs[index].get("sheets", [])) + 1
            self.feishu_jobs[index].setdefault("sheets", []).append(self._empty_sheet_rule(next_no))
        self._render_feishu_jobs()
        self.save_from_ui(silent=True)

    def _remove_feishu_sheet(self, job_index: int, sheet_index: int):
        self._collect_feishu_jobs_from_ui()
        if 0 <= job_index < len(self.feishu_jobs):
            sheets = self.feishu_jobs[job_index].get("sheets", [])
            if 0 <= sheet_index < len(sheets):
                del sheets[sheet_index]
            if not sheets:
                sheets.append(self._empty_sheet_rule(1))
        self._render_feishu_jobs()
        self.save_from_ui(silent=True)

    def _short_path(self, path: str, max_len: int = 95):
        path = str(path or "")
        if len(path) <= max_len:
            return path
        return "..." + path[-max_len:]

    def _entry_in_grid(self, parent, row, column, width=120):
        entry = ctk.CTkEntry(
            parent,
            width=width,
            height=30,
            fg_color=self.C["surface"],
            border_color=self.C["border"],
            text_color=self.C["text"],
            corner_radius=6,
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        entry.grid(row=row, column=column, padx=4, pady=4, sticky="ew")
        entry.bind("<FocusOut>", lambda _e: self.save_from_ui(silent=True))
        entry.bind("<Return>", lambda _e: self.save_from_ui(silent=True))
        return entry

    def _render_feishu_jobs(self):
        if not self.feishu_jobs_frame:
            return
        for child in self.feishu_jobs_frame.winfo_children():
            child.destroy()
        self.feishu_job_widgets = []

        if not self.feishu_jobs:
            empty = ctk.CTkFrame(self.feishu_jobs_frame, fg_color=self.C["surface"], corner_radius=12)
            empty.grid(row=0, column=0, sticky="ew", padx=12, pady=12)
            empty.grid_columnconfigure(0, weight=1)
            self._label(
                empty,
                "ยังไม่มีรายการ Excel — กด + Add Excel เพื่อเพิ่มไฟล์สำหรับ Export/Send",
                size=13,
                color=self.C["muted"],
            ).grid(row=0, column=0, padx=16, pady=18, sticky="w")
            return

        for job_index, job in enumerate(self.feishu_jobs):
            file_card = ctk.CTkFrame(self.feishu_jobs_frame, fg_color=self.C["surface"], corner_radius=14)
            file_card.grid(row=job_index, column=0, sticky="ew", padx=12, pady=(10, 14))
            file_card.grid_columnconfigure(2, weight=1)

            enabled_var = ctk.BooleanVar(value=bool(job.get("enabled", True)))
            ctk.CTkCheckBox(
                file_card,
                text="Use",
                variable=enabled_var,
                checkbox_width=20,
                checkbox_height=20,
                fg_color=self.C["accent"],
                border_color=self.C["border"],
                text_color=self.C["text"],
                command=lambda: self.save_from_ui(silent=True),
            ).grid(row=0, column=0, padx=(12, 8), pady=(12, 2), sticky="w")

            self._label(file_card, "▧", size=18, color=self.C["accent2"]).grid(row=0, column=1, padx=(8, 4), pady=(12, 2))

            excel_entry = ctk.CTkEntry(
                file_card,
                height=32,
                fg_color=self.C["card"],
                border_color=self.C["border"],
                text_color=self.C["text"],
                corner_radius=8,
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            )
            excel_entry.grid(row=0, column=2, padx=6, pady=(12, 2), sticky="ew")
            excel_entry.insert(0, job.get("excel_file", ""))
            excel_entry.bind("<FocusOut>", lambda _e: self.save_from_ui(silent=True))
            excel_entry.bind("<Return>", lambda _e: self.save_from_ui(silent=True))

            send_excel_var = ctk.BooleanVar(value=bool(job.get("send_excel", False)))
            ctk.CTkCheckBox(
                file_card,
                text="Send Excel",
                variable=send_excel_var,
                checkbox_width=20,
                checkbox_height=20,
                fg_color=self.C["accent"],
                border_color=self.C["border"],
                text_color=self.C["text"],
                command=lambda: self.save_from_ui(silent=True),
            ).grid(row=0, column=3, padx=8, pady=(12, 2))

            self._btn(file_card, "Browse", lambda i=job_index: self._browse_excel_for_job(i), width=78, height=30).grid(row=0, column=4, padx=4, pady=(12, 2))
            self._btn(file_card, "+ Add Sheet", lambda i=job_index: self._add_feishu_sheet(i), width=110, height=30).grid(row=0, column=5, padx=4, pady=(12, 2))
            ctk.CTkButton(
                file_card,
                text="Remove",
                command=lambda i=job_index: self._remove_feishu_excel(i),
                width=84,
                height=30,
                fg_color="#991b1b",
                hover_color="#7f1d1d",
                corner_radius=8,
                font=ctk.CTkFont(size=12),
            ).grid(row=0, column=6, padx=(4, 12), pady=(12, 2))

            header = ctk.CTkFrame(file_card, fg_color="transparent")
            header.grid(row=1, column=0, columnspan=7, sticky="ew", padx=12, pady=(10, 2))
            for col, weight in enumerate([0, 0, 0, 1, 1, 1, 0, 0, 0]):
                header.grid_columnconfigure(col, weight=weight)
            headers = ["", "Use", "No.", "Sheet", "Range", "Output File", "Hide", "Send", ""]
            for col, text in enumerate(headers):
                self._label(header, text, size=11, color="#93c5fd").grid(row=0, column=col, padx=4, sticky="w")

            sheet_widgets = []
            for sheet_index, sheet in enumerate(job.get("sheets", [])):
                row_frame = ctk.CTkFrame(file_card, fg_color="#0b1222", corner_radius=10)
                row_frame.grid(row=2 + sheet_index, column=0, columnspan=7, sticky="ew", padx=12, pady=4)
                for col, weight in enumerate([0, 0, 0, 1, 1, 1, 0, 0, 0]):
                    row_frame.grid_columnconfigure(col, weight=weight)

                enabled_sheet_var = ctk.BooleanVar(value=bool(sheet.get("enabled", True)))
                ctk.CTkLabel(row_frame, text="").grid(row=0, column=0, padx=2)
                ctk.CTkCheckBox(
                    row_frame,
                    text="Use",
                    variable=enabled_sheet_var,
                    checkbox_width=20,
                    checkbox_height=20,
                    fg_color=self.C["accent"],
                    border_color=self.C["border"],
                    text_color=self.C["text"],
                    command=lambda: self.save_from_ui(silent=True),
                ).grid(row=0, column=1, padx=4, pady=6, sticky="w")

                no_entry = self._entry_in_grid(row_frame, 0, 2, width=56)
                no_entry.insert(0, str(sheet.get("no", sheet_index + 1)))
                sheet_entry = self._entry_in_grid(row_frame, 0, 3, width=160)
                sheet_entry.insert(0, str(sheet.get("sheet", "1")))
                range_entry = self._entry_in_grid(row_frame, 0, 4, width=160)
                range_entry.insert(0, str(sheet.get("range", "A1:Z50")))
                output_entry = self._entry_in_grid(row_frame, 0, 5, width=180)
                output_entry.insert(0, str(sheet.get("output", f"report_{sheet_index+1}.png")))

                hide_var = ctk.BooleanVar(value=bool(sheet.get("hide", False)))
                ctk.CTkCheckBox(
                    row_frame,
                    text="Hide",
                    variable=hide_var,
                    checkbox_width=20,
                    checkbox_height=20,
                    fg_color=self.C["accent"],
                    border_color=self.C["border"],
                    text_color=self.C["text"],
                    command=lambda: self.save_from_ui(silent=True),
                ).grid(row=0, column=6, padx=4, pady=6)

                send_var = ctk.BooleanVar(value=bool(sheet.get("send", True)))
                ctk.CTkCheckBox(
                    row_frame,
                    text="Send",
                    variable=send_var,
                    checkbox_width=20,
                    checkbox_height=20,
                    fg_color=self.C["accent"],
                    border_color=self.C["border"],
                    text_color=self.C["text"],
                    command=lambda: self.save_from_ui(silent=True),
                ).grid(row=0, column=7, padx=4, pady=6)

                ctk.CTkButton(
                    row_frame,
                    text="×",
                    command=lambda i=job_index, j=sheet_index: self._remove_feishu_sheet(i, j),
                    width=34,
                    height=30,
                    fg_color="#991b1b",
                    hover_color="#7f1d1d",
                    corner_radius=8,
                    font=ctk.CTkFont(size=16),
                ).grid(row=0, column=8, padx=(4, 8), pady=6)

                sheet_widgets.append({
                    "enabled": enabled_sheet_var,
                    "no": no_entry,
                    "sheet": sheet_entry,
                    "range": range_entry,
                    "output": output_entry,
                    "hide": hide_var,
                    "send": send_var,
                })

            self.feishu_job_widgets.append({
                "enabled": enabled_var,
                "send_excel": send_excel_var,
                "excel_entry": excel_entry,
                "sheets": sheet_widgets,
            })

    def _build_setting(self, parent):
        parent.configure(fg_color=self.C["bg"])
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(
            parent,
            corner_radius=0,
            fg_color=self.C["bg"],
            scrollbar_button_color=self.C["border"],
            scrollbar_button_hover_color=self.C["accent"],
        )
        scroll.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll.grid_columnconfigure(0, weight=1)

        # ── JMS card ────────────────────────────────────────────────
        jms = self._card(scroll)
        jms.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 8))
        jms.grid_columnconfigure(1, weight=1)
        self._section_title(jms, 0, "⬇", "JMS Export")
        self._setting_entry(jms, 1, "Auth Token", "jms_auth_token", show="*")
        self._setting_entry(jms, 2, "Save Path", "jms_save_path", browse="folder")
        self._setting_entry(
            jms,
            3,
            "Main Line Transport",
            "main_transport_filename"
        )

        self._setting_entry(
            jms,
            4,
            "Branch Line Transport",
            "branch_transport_filename"
        )

        # ── Feishu card ─────────────────────────────────────────────
        feishu = self._card(scroll)
        feishu.grid(row=1, column=0, sticky="ew", padx=20, pady=8)
        feishu.grid_columnconfigure(1, weight=1)
        self._section_title(feishu, 0, "🚀", "Feishu Chat")
        self._setting_entry(feishu, 1, "PNG Output Folder", "png_output_folder", browse="folder")
        self._setting_entry(feishu, 2, "App ID",         "app_id")
        self._setting_entry(feishu, 3, "App Secret",     "app_secret",         show="*")
        self._setting_entry(feishu, 4, "Chat ID",        "chat_id")

        # ── Dynamic Excel Manager ───────────────────────────────────
        manager = self._card(scroll)
        manager.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        manager.grid_columnconfigure(0, weight=1)
        top = ctk.CTkFrame(manager, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        top.grid_columnconfigure(1, weight=1)
        self._btn(top, "+ Add Excel", self._add_feishu_excel, width=130, height=36).grid(row=0, column=0, padx=(0, 14), pady=2, sticky="w")
        self._label(
            top,
            "เลือกได้หลายไฟล์ในครั้งเดียว | กด + Add Sheet เพื่อเพิ่มรายการ Export",
            size=12,
            weight="bold",
            color=self.C["text"],
        ).grid(row=0, column=1, sticky="w")
        self._btn(top, "💾 Save", self.save_from_ui, width=100, height=32).grid(row=0, column=2, padx=(8, 0), sticky="e")

        self.feishu_jobs_frame = ctk.CTkFrame(manager, fg_color="#081120", corner_radius=14)
        self.feishu_jobs_frame.grid(row=1, column=0, sticky="ew", padx=16, pady=(8, 16))
        self.feishu_jobs_frame.grid_columnconfigure(0, weight=1)

        # ── Save button ─────────────────────────────────────────────
        btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_row.grid(row=3, column=0, sticky="ew", padx=20, pady=(8, 28))
        self._btn(btn_row, "💾  Save Settings", self.save_from_ui, width=160).pack(side="left", padx=4)

    def _browse_to_entry(self, key: str, mode: str):
        if mode == "folder":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if path:
            entry = self.widgets[key]
            entry.delete(0, "end")
            entry.insert(0, path)
            self.save_from_ui(silent=True)

    # ---------- Config ----------
    def _load_values_to_ui(self):
        s = self.cfg["SETTING"]
        today = datetime.now().date()
        self.run_hour_interval.set(
            s.get("run_hour_interval", "1")
        )

        self.run_minute_interval.set(
            s.get("run_minute_interval", "5")
        )
        self.start_hour.set(s.get("start_hour", "12:00"))
        self.end_hour.set(s.get("end_hour", "12:00"))
        self.var_main_transport.set(
            s.get("run_main_transport", "1") == "1"
        )

        self.var_branch_transport.set(
            s.get("run_branch_transport", "1") == "1"
        )

        self.var_feishu.set(
            s.get("run_feishu_chat", "1") == "1"
        )

        for widget, key, default_date in [
            (self.start_date, "start_date", today),
            (self.end_date, "end_date", today + timedelta(days=1)),
        ]:
            value = s.get(key, "")
            try:
                widget.set_date(datetime.strptime(value, "%Y-%m-%d").date() if value else default_date)
            except Exception:
                widget.set_date(default_date)

        for key, entry in self.widgets.items():
            entry.delete(0, "end")
            entry.insert(0, s.get(key, DEFAULT_CONFIG.get(key, "")))

        self._load_feishu_jobs_from_config()
        self._render_feishu_jobs()

    def save_from_ui(self, silent: bool = False):
        s = self.cfg["SETTING"]
        s["run_hour_interval"] = self.run_hour_interval.get()
        s["run_minute_interval"] = self.run_minute_interval.get()
        s["start_date"] = str(self.start_date.get_date())
        s["end_date"] = str(self.end_date.get_date())
        s["start_hour"] = self.start_hour.get()
        s["end_hour"] = self.end_hour.get()
        s["run_main_transport"] = (
            "1" if self.var_main_transport.get() else "0"
        )

        s["run_branch_transport"] = (
            "1" if self.var_branch_transport.get() else "0"
        )

        s["run_feishu_chat"] = (
            "1" if self.var_feishu.get() else "0"
        )
        for key, entry in self.widgets.items():
            s[key] = entry.get().strip()
        self._save_feishu_jobs_to_config()
        save_config(self.cfg)
        if not silent:
            self.write_log("บันทึก Config แล้ว")

    def get_setting(self, key: str) -> str:
        return self.cfg["SETTING"].get(key, DEFAULT_CONFIG.get(key, "")).strip()

    def get_datetime_range(self):
        start = f"{self.start_date.get_date()} {self.start_hour.get()}:00"
        end = f"{self.end_date.get_date()} {self.end_hour.get()}:00"
        return start, end

    # ---------- Log / State ----------
    # ── Log level → tag map ──────────────────────────────────────
    _LOG_TAGS = {
        "INFO"    : ("INFO",    "◉"),
        "OK"      : ("SUCCESS", "✔"),
        "WARN"    : ("WARN",    "⚠"),
        "ERROR"   : ("ERROR",   "✖"),
        "START"   : ("START",   "▶"),
        "SECTION" : ("INFO",    "─"),
    }

    def write_log(self, message: str, level: str = "INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        tag, icon = self._LOG_TAGS.get(level.upper(), ("INFO", "◉"))
        text = f"[{ts}]  {icon}  {message}"
        logging.info(text)
        self.after(0, lambda t=text, g=tag: self._append_log(t, g))

    def _append_log(self, text: str, tag: str = "INFO"):
        try:
            self.log_box.tag_config("INFO",    foreground="#93c5fd")
            self.log_box.tag_config("SUCCESS", foreground="#86efac")
            self.log_box.tag_config("WARN",    foreground="#fbbf24")
            self.log_box.tag_config("ERROR",   foreground="#fca5a5")
            self.log_box.tag_config("START",   foreground="#67e8f9")
        except Exception:
            pass
        self.log_box.insert("end", text + "\n", tag)
        self.log_box.see("end")

    # badge config map: keyword → (dot_color, bg, text_color, label)
    _BADGE = {
        "Running"  : ("#f59e0b", "#2d1f0a", "#fbbf24", "● RUNNING"),
        "Success"  : ("#10b981", "#0d2318", "#86efac", "● DONE"),
        "Error"    : ("#ef4444", "#2d0f0f", "#fca5a5", "● ERROR"),
        "Stopped"  : ("#64748b", "#111827", "#94a3b8", "● STOPPED"),
        "Next run" : ("#3b82f6", "#0e1f3d", "#93c5fd", "● AUTO"),
        "Idle"     : ("#10b981", "#0d2318", "#86efac", "● IDLE"),
    }

    def set_status(self, text: str):
        def _update():
            self.status.configure(text=text)
            for kw, (_, bg, fg, badge_text) in self._BADGE.items():
                if kw.lower() in text.lower():
                    self.status_badge.configure(
                        text=badge_text,
                        text_color=fg,
                        fg_color=bg,
                    )
                    return
            # default
            self.status_badge.configure(text="● IDLE", text_color=self.C["success"], fg_color="#0d2318")
        self.after(0, _update)

    def stop_checker(self) -> bool:
        return self.stop_requested

    # ---------- Jobs ----------
    def run_now(self):
        if self.job_running:
            self.write_log("มีงานกำลังทำงานอยู่")
            return
        self.save_from_ui(silent=True)
        self.active_mode = "run"
        self.lock_ui()

        self.btn_start.configure(
            state="disabled"
        )

        self.btn_run.configure(
            state="normal"
        )
        threading.Thread(target=self.run_selected_jobs, daemon=True).start()

    def run_selected_jobs(self):
        if self.job_running:
            return
        self.job_running = True
        self.after(0, self.toggle_run_button)
        self.stop_requested = False
        self.current_mode = None
        self.set_status("Status: Running...")
        self.write_log("เริ่มทำงาน")

        try:
            if (
                not self.var_main_transport.get()
                and not self.var_branch_transport.get()
                and not self.var_feishu.get()
            ):
                raise ValueError("กรุณาเลือกอย่างน้อย 1 งาน: Export JMS หรือ Feishu Chat")

            # Main Line
            if self.var_main_transport.get():

                start, end = self.get_datetime_range()

                export_jms_excel(
                    auth_token=self.get_setting("jms_auth_token"),
                    save_folder=self.get_setting("jms_save_path"),
                    filename=self.get_setting("main_transport_filename"),
                    start_time=start,
                    end_time=end,
                    stop_checker=self.stop_checker,
                    log=self.write_log,
                )

            # Branch Line
            if self.var_branch_transport.get():

                start, end = self.get_datetime_range()

                export_branch_tracking(
                    auth_token=self.get_setting("jms_auth_token"),
                    save_folder=self.get_setting("jms_save_path"),
                    filename=self.get_setting("branch_transport_filename"),
                    start_time=start,
                    end_time=end,
                    stop_checker=self.stop_checker,
                    log=self.write_log,
                )

            if self.stop_requested:
                self.write_log("หยุดก่อนส่ง Feishu Chat")
                self.set_status("Status: Stopped")
                return

            if self.var_feishu.get():
                self.run_feishu_chat()

            if self.stop_requested:
                self.set_status("Status: Stopped")
                return

            self.set_status("Status: Success")
            self.write_log("งานทั้งหมดเสร็จสมบูรณ์ ✅")
        except Exception as exc:
            self.set_status("Status: Error")
            self.write_log(f"ERROR: {exc}")
            messagebox.showerror("Error", str(exc))
        finally:
            self.job_running = False
            self.after(0, self.toggle_run_button)
            self.unlock_ui()

    def _resolve_excel_path(self, excel_file: str) -> str:
        excel_file = os.path.normpath(unquote(str(excel_file or "").strip()))
        if not excel_file:
            return ""
        candidates = [excel_file]
        if not os.path.isabs(excel_file):
            save_path = self.get_setting("jms_save_path")
            if save_path:
                candidates.append(os.path.join(save_path, excel_file))
            candidates.append(resource_path(excel_file))
        for candidate in candidates:
            if os.path.exists(candidate):
                return os.path.abspath(candidate)
        return os.path.abspath(candidates[0])

    def _dated_excel_copy_path(self, excel_path: str, output_folder: str) -> str:
        report_date = str(self.start_date.get_date())[:10]
        excel_path = os.path.normpath(unquote(str(excel_path or "")))
        folder = os.path.normpath(unquote(str(output_folder or os.path.dirname(excel_path))))
        os.makedirs(folder, exist_ok=True)

        stem, ext = os.path.splitext(os.path.basename(excel_path))
        ext = ext or ".xlsx"
        return os.path.abspath(os.path.normpath(os.path.join(folder, f"{stem}_{report_date}{ext}")))

    def _excel_file_format(self, file_path: str) -> int:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".xlsm":
            return 52
        if ext == ".xlsb":
            return 50
        if ext == ".xls":
            return 56
        return 51

    def _create_refreshed_excel_copy(self, excel_path: str, job: dict, output_folder: str) -> str:
        import pythoncom
        import win32com.client as win32

        excel_path = os.path.abspath(os.path.normpath(unquote(str(excel_path or ""))))
        copy_path = self._dated_excel_copy_path(excel_path, output_folder)
        report_date = str(self.start_date.get_date())[:10]

        pythoncom.CoInitialize()
        excel = None
        wb = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            excel.EnableEvents = False

            wb = excel.Workbooks.Open(
                excel_path,
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
            )

            for sheet in job.get("sheets", []):
                if not sheet.get("enabled", True):
                    continue
                try:
                    sheet_index = int(str(sheet.get("sheet", "1")).strip() or "1")
                    ws = wb.Worksheets(sheet_index)
                    ws.Range("B2").Value = report_date
                except Exception:
                    pass

            wb.RefreshAll()
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception as exc:
                self.write_log(f"Excel refresh warning: {exc}", level="WARN")

            deadline = time.time() + 180
            while time.time() < deadline:
                if self.stop_requested:
                    return ""
                try:
                    if excel.CalculateState == 0:
                        break
                except Exception:
                    break
                time.sleep(0.5)

            try:
                if os.path.exists(copy_path):
                    os.remove(copy_path)
            except Exception:
                pass

            wb.SaveCopyAs(copy_path)
            return copy_path
        finally:
            try:
                if wb:
                    wb.Close(False)
            except Exception:
                pass
            try:
                if excel:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    def run_feishu_chat(self):
        if run_create is None:
            raise RuntimeError(f"ไม่สามารถ import createpng.py ได้: {CREATEPNG_IMPORT_ERROR}")

        self._save_feishu_jobs_to_config()
        jobs = self._normalize_feishu_jobs(self.feishu_jobs)
        active_jobs = [job for job in jobs if job.get("enabled", True)]

        if not active_jobs:
            raise ValueError("ยังไม่มี Excel Job ที่เปิดใช้งานใน Feishu Chat")

        output_folder = self.get_setting("png_output_folder") or self.get_setting("jms_save_path")
        if output_folder:
            os.makedirs(output_folder, exist_ok=True)

        generated_images = []
        generated_excel_files = []
        created_count = 0
        excel_file_count = 0
        sent_count = 0
        sent_file_count = 0

        self.write_log("เริ่มเตรียมไฟล์ทั้งหมดก่อนส่ง Feishu")

        # STEP 1: Create all PNG and refreshed Excel copies first, keep import order
        for job in active_jobs:
            if self.stop_requested:
                self.write_log("ยกเลิกก่อนเตรียมไฟล์")
                return

            excel_path = self._resolve_excel_path(job.get("excel_file", ""))
            if not excel_path or not os.path.exists(excel_path):
                raise FileNotFoundError(f"ไม่พบ Excel File: {job.get('excel_file', '')}")

            self.write_log(f"Excel: {os.path.basename(excel_path)}")

            for sheet in job.get("sheets", []):
                if self.stop_requested:
                    self.write_log("ยกเลิกก่อนสร้าง PNG")
                    return

                if not sheet.get("enabled", True):
                    continue

                sheet_name = str(sheet.get("sheet", "1")).strip() or "1"
                cell_range = str(sheet.get("range", "A1:Z50")).strip() or "A1:Z50"

                png_filename = str(sheet.get("output", "report.png")).strip() or "report.png"
                if not png_filename.lower().endswith(".png"):
                    png_filename += ".png"

                job_output_folder = output_folder or os.path.dirname(excel_path)
                os.makedirs(job_output_folder, exist_ok=True)
                image_path = os.path.join(job_output_folder, png_filename)

                self.write_log(
                    f"สร้าง PNG: {os.path.basename(excel_path)} | Sheet {sheet_name} | {cell_range}"
                )

                run_create(
                    excel_path,
                    sheet_name,
                    cell_range,
                    image_path,
                    report_date=str(self.start_date.get_date()),
                    log=self.write_log,
                    stop_checker=lambda: self.stop_requested
                )

                if not os.path.exists(image_path):
                    raise FileNotFoundError(f"สร้าง PNG ไม่สำเร็จ: {image_path}")

                generated_images.append({
                    "path": image_path,
                    "filename": png_filename,
                    "send": bool(sheet.get("send", True)),
                })

                created_count += 1

            if job.get("send_excel", False):
                if self.stop_requested:
                    self.write_log("ยกเลิกก่อนสร้างสำเนา Excel")
                    return
                self.write_log(f"สร้างสำเนา Excel สำหรับส่ง: {os.path.basename(excel_path)}")
                excel_copy_path = self._create_refreshed_excel_copy(excel_path, job, output_folder)
                if excel_copy_path:
                    generated_excel_files.append({
                        "path": excel_copy_path,
                        "filename": os.path.basename(excel_copy_path),
                    })
                    excel_file_count += 1

        if self.stop_requested:
            self.write_log("ยกเลิกก่อน Upload Feishu")
            return

        if not generated_images and not generated_excel_files:
            raise ValueError("ไม่มีไฟล์ PNG หรือ Excel ที่ต้องส่งจาก Feishu Jobs")

        self.write_log(f"เตรียมไฟล์ครบแล้ว: PNG {created_count} รูป / Excel {excel_file_count} ไฟล์")

        # STEP 2: Upload + Send after every PNG is ready
        token = None

        for item in generated_images:
            if self.stop_requested:
                self.write_log("ยกเลิกก่อนส่ง Feishu")
                return

            if not item["send"]:
                continue

            if token is None:
                token = get_tenant_access_token(
                    self.get_setting("app_id"),
                    self.get_setting("app_secret"),
                    log=self.write_log
                )

            self.write_log(f"อัปโหลดรูปไป Feishu: {item['filename']}")
            image_key = upload_feishu_image(token, item["path"], log=self.write_log)

            self.write_log(f"ส่งรูปเข้า Feishu: {item['filename']}")
            send_feishu_image_by_chat_id(
                token,
                self.get_setting("chat_id"),
                image_key,
                log=self.write_log
            )

            sent_count += 1

        for item in generated_excel_files:
            if self.stop_requested:
                self.write_log("ยกเลิกก่อนส่งไฟล์ Excel")
                return

            if token is None:
                token = get_tenant_access_token(
                    self.get_setting("app_id"),
                    self.get_setting("app_secret"),
                    log=self.write_log
                )

            self.write_log(f"อัปโหลดไฟล์ Excel ไป Feishu: {item['filename']}")
            file_key = upload_feishu_file(token, item["path"], log=self.write_log)

            self.write_log(f"ส่งไฟล์ Excel เข้า Feishu: {item['filename']}")
            send_feishu_file_by_chat_id(
                token,
                self.get_setting("chat_id"),
                file_key,
                log=self.write_log
            )

            sent_file_count += 1

        self.write_log(
            f"Feishu Chat สำเร็จ: สร้าง {created_count} รูป / ส่งรูป {sent_count} รูป / ส่ง Excel {sent_file_count} ไฟล์"
        )


    def lock_ui(self):
        widgets = [
            self.run_hour_interval,
            self.run_minute_interval,
            self.start_hour,
            self.end_hour,
            self.start_date,
            self.end_date,
        ]

        for widget in widgets:
            try:
                widget.configure(state="disabled")
            except:
                pass

        self.tabview.set("  Home  ")

        for entry in self.widgets.values():
            try:
                entry.configure(state="disabled")
            except:
                pass

        try:
            self.tabview._segmented_button.configure(state="disabled")
        except:
            pass

    def unlock_ui(self):
        widgets = [
            self.run_hour_interval,
            self.run_minute_interval,
            self.start_hour,
            self.end_hour,
            self.start_date,
            self.end_date,
        ]

        for widget in widgets:
            try:
                widget.configure(state="readonly")
            except:
                pass

        self.tabview.set("  Home  ")

        for entry in self.widgets.values():
            try:
                entry.configure(state="normal")
            except:
                pass

        try:
            self.tabview._segmented_button.configure(state="normal")
        except:
            pass

        self.btn_start.configure(state="normal")
        self.btn_run.configure(state="normal")
        self.active_mode = None

    # ---------- Scheduler ----------
    def start_scheduler(self):
        if self.scheduler_running:
            return
        self.save_from_ui(silent=True)
        self.active_mode = "auto"
        self.lock_ui()

        self.btn_run.configure(
            state="disabled"
        )

        self.btn_start.configure(
            state="normal"
        )
        self.stop_requested = False
        self.scheduler_running = True
        self.after(0, self.toggle_auto_button)
        self.stop_requested = False
        self.write_log("เริ่ม Auto Scheduler")
        threading.Thread(target=self.scheduler_loop, daemon=True).start()

    def toggle_auto_button(self):

        if self.scheduler_running:

            self.btn_run.configure(
                state="disabled"
            )

            self.btn_start.configure(
                text="■ Stop",
                fg_color="#8a2d2d",
                hover_color="#a83232",
                command=self.stop_all
            )

        else:

            self.btn_run.configure(
                state="normal"
            )

            self.btn_start.configure(
                text="▶ Start Auto",
                fg_color="#1F6AA5",
                hover_color="#144870",
                command=self.start_scheduler
            )

    def toggle_run_button(self):

        if self.job_running:

            self.btn_start.configure(
                state="disabled"
            )

            self.btn_run.configure(
                text="■ Stop",
                fg_color="#8a2d2d",
                hover_color="#a83232",
                command=self.stop_all
            )

        else:

            self.btn_start.configure(
                state="normal"
            )

            self.btn_run.configure(
                text="⚡ Run Now",
                fg_color="#1F6AA5",
                hover_color="#144870",
                command=self.run_now
            )
    def stop_all(self):

        self.scheduler_running = False
        self.stop_requested = True

        self.set_status("Status: Stopped")
        self.write_log("รับคำสั่งหยุด")

        
    def scheduler_loop(self):

        while self.scheduler_running:

            try:
                hours = int(self.run_hour_interval.get() or 1)
                minutes = int(self.run_minute_interval.get() or 0)
            except ValueError:
                hours = 1
                minutes = 0

            now = datetime.now()

            # ตั้งนาทีตาม Combobox
            next_run = now.replace(
                minute=minutes,
                second=0,
                microsecond=0
            )

            # ถ้าเวลาปัจจุบันเลยจุดนั้นแล้ว
            # ให้ขยับทีละ X ชั่วโมง
            while next_run <= now:
                next_run += timedelta(hours=hours)

            while self.scheduler_running and datetime.now() < next_run:

                remain = int(
                    (next_run - datetime.now()).total_seconds()
                )

                m, s = divmod(max(remain, 0), 60)

                self.set_status(
                    f"Next run: {next_run.strftime('%H:%M:%S')} ({m:02d}:{s:02d})"
                )

                time.sleep(1)

            if not self.scheduler_running:
                break

            self.write_log(
                f"Auto run at {datetime.now().strftime('%H:%M:%S')}"
            )

            self.run_selected_jobs()

        self.scheduler_running = False

        self.after(
            0,
            self.toggle_auto_button
        )

        self.after(
            0,
            self.unlock_ui
        )

        self.set_status(
            "Status: Stopped"
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()

