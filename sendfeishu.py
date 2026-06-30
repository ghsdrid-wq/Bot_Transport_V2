import requests, time, hmac, hashlib, base64
import os

def request_with_retry(func, retries=3, delay=2, log=None, name=""):
    for i in range(retries):
        try:
            return func()
        except Exception as e:
            if log:
                log(f"  Retry [{i+1}/{retries}] {name} — {e}")
            if i < retries - 1:
                time.sleep(delay * (i + 1))
            else:
                raise

def gen_sign(timestamp, secret):
    string_to_sign = f"{timestamp}\n{secret}"
    hmac_code = hmac.new(string_to_sign.encode(), digestmod=hashlib.sha256).digest()
    return base64.b64encode(hmac_code).decode()

def get_token(app_id, app_secret, log=None):
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"

    def do_request():
        return requests.post(
            url,
            json={"app_id": app_id, "app_secret": app_secret},
            timeout=10
        ).json()

    res = request_with_retry(do_request, log=log, name="get_token")

    if "tenant_access_token" not in res:
        raise Exception(f"Get token failed: {res}")

    if log:
        log("  Feishu token OK")

    return res["tenant_access_token"]

def upload_image(token, path, log=None):
    url = "https://open.feishu.cn/open-apis/im/v1/images"

    headers = {
        "Authorization": f"Bearer {token}"
    }

    def do_request():
        with open(path, "rb") as f:
            files = {"image": f}
            return requests.post(
                url,
                headers=headers,
                files=files,
                data={"image_type": "message"},
                timeout=10
            ).json()

    res = request_with_retry(do_request, log=log, name="upload_image")

    if res.get("code") != 0:
        raise Exception(res)

    return res["data"]["image_key"]

def send_image(image_key, webhook, secret, log=None):
    timestamp = str(int(time.time()))
    sign = gen_sign(timestamp, secret)

    data = {
        "timestamp": timestamp,
        "sign": sign,
        "msg_type": "image",
        "content": {
            "image_key": image_key
        }
    }

    def do_request():
        return requests.post(webhook, json=data, timeout=10)

    request_with_retry(do_request, log=log, name="send_image")

def run_send(folder, webhook, secret, app_id, app_secret, log=None):
    def write(msg):
        if log:
            log(msg)
        else:
            print(msg)

    if not webhook:
        raise Exception("Missing WEBHOOK")

    if not app_id or not app_secret:
        raise Exception("Missing APP_ID / APP_SECRET")

    token = get_token(app_id, app_secret, log=write)

    # 🔥 หา png ทุกไฟล์ใน folder
    file_path = os.path.join(folder, "report.png")

    if not os.path.exists(file_path):
        raise Exception(f"File not found: {file_path}")

    write("── Feishu Delivery ─────────────────")
    write(f"  Uploading  : {os.path.basename(file_path)}")
    key = upload_image(token, file_path, log=write)

    write(f"  Sending    : {os.path.basename(file_path)}")
    send_image(key, webhook, secret, log=write)

    time.sleep(1)