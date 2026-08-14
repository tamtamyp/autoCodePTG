import os
import sys
import time
import json
import re
import csv
import threading
import logging
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs
import requests
import openpyxl

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

# Cấu hình logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

# ==============================================================================
# CẤU HÌNH MẶC ĐỊNH (Fallback nếu không đọc được từ Google Sheet)
# ==============================================================================
DEFAULT_CONFIG = {
    "BOT_TOKEN": "1281608028214813212:HNNRzFXskRXjEFtrHKXDqWeOauFfSkuPBVNQgoIYUlFFlikojFgxnopHrXicjZex",
    "SECRET_TOKEN": "abc-xyz-123",
    "CHAT_ID": "c40c79a60df0e4aebde1",
    "SERVER_ID": "2",
    "GAME_CODE": "661",
    "SHEET_URLS": ["https://docs.google.com/spreadsheets/d/1mdv1O31HGALyDTeZhmjn0aLNmjOmpR_3fO6RcudTerU/edit?usp=sharing"]
}

# Các biến cấu hình toàn cục
BOT_TOKEN = DEFAULT_CONFIG["BOT_TOKEN"]
SECRET_TOKEN = DEFAULT_CONFIG["SECRET_TOKEN"]
CHAT_ID = DEFAULT_CONFIG["CHAT_ID"]
SERVER_ID = DEFAULT_CONFIG["SERVER_ID"]
SHEET_URLS = DEFAULT_CONFIG["SHEET_URLS"]
GAME_CODE = DEFAULT_CONFIG["GAME_CODE"]

# Cấu hình mặc định của URL chứa file config
DEFAULT_CONFIG_SHEET_URL = "https://docs.google.com/spreadsheets/d/1uu_lE1QD5tB1HgJ5j3HtNeVdoMyEcAkvGeV3137264U/edit?usp=sharing"

def get_config_sheet_url():
    # 1. Ưu tiên đọc từ biến môi trường (dành cho Render)
    url = os.environ.get("CONFIG_SHEET_URL")
    if url:
        return url.strip()
            
    # 2. Mặc định
    return DEFAULT_CONFIG_SHEET_URL

def load_config_from_google_sheet(sheet_url):
    config = DEFAULT_CONFIG.copy()
    if not sheet_url:
        return config

    try:
        sheet_id_match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not sheet_id_match:
            logging.error(f"URL Google Sheet cấu hình không đúng định dạng: {sheet_url}")
            return config
        spreadsheet_id = sheet_id_match.group(1)
        
        gid = "0"
        gid_match = re.search(r"gid=([0-9]+)", sheet_url)
        if gid_match:
            gid = gid_match.group(1)
            
        csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"
        
        response = requests.get(csv_url, timeout=10)
        if response.status_code != 200:
            logging.error(f"Không thể tải Google Sheet cấu hình. Status code: {response.status_code}")
            return config
            
        content = response.content.decode('utf-8')
        csv_reader = csv.reader(content.splitlines())
        
        for row in csv_reader:
            if len(row) >= 2:
                key = str(row[0]).strip()
                val = str(row[1]).strip()
                if key in config:
                    if key == "SHEET_URLS":
                        config[key] = [url.strip() for url in val.split(",") if url.strip()]
                    else:
                        config[key] = val
        logging.info("Đã tải cấu hình thành công từ Google Sheet!")
    except Exception as e:
        logging.error(f"Lỗi khi đọc cấu hình từ Google Sheet: {e}")
        
    return config

def reload_config():
    global BOT_TOKEN, SECRET_TOKEN, CHAT_ID, SERVER_ID, GAME_CODE, SHEET_URLS
    url = get_config_sheet_url()
    logging.info(f"Đang tải cấu hình động từ Sheet URL: {url}...")
    config = load_config_from_google_sheet(url)
    
    BOT_TOKEN = config["BOT_TOKEN"]
    SECRET_TOKEN = config["SECRET_TOKEN"]
    CHAT_ID = config["CHAT_ID"]
    SERVER_ID = config["SERVER_ID"]
    GAME_CODE = config["GAME_CODE"]
    SHEET_URLS = config["SHEET_URLS"]

# Khởi động tải cấu hình lần đầu
try:
    reload_config()
except Exception as e:
    logging.error(f"Lỗi tải cấu hình khởi động: {e}")

PORT = int(os.environ.get("PORT", 5000))
# ==============================================================================


def get_google_sheet_roles(sheet_url):
    # Trích xuất Spreadsheet ID từ URL Google Sheet
    sheet_id_match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not sheet_id_match:
        raise ValueError("URL Google Sheet không hợp lệ. Vui lòng kiểm tra lại định dạng URL.")
    spreadsheet_id = sheet_id_match.group(1)
    
    # Trích xuất gid (mặc định là 0 nếu không có)
    gid = "0"
    gid_match = re.search(r"gid=([0-9]+)", sheet_url)
    if gid_match:
        gid = gid_match.group(1)
        
    csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"
    
    response = requests.get(csv_url, timeout=15)
    if response.status_code != 200:
        raise Exception(f"Không thể tải Google Sheet. Mã lỗi HTTP: {response.status_code}. Hãy đảm bảo sheet ở chế độ chia sẻ công khai ('Anyone with link can view').")
        
    content = response.content.decode('utf-8')
    csv_reader = csv.reader(content.splitlines())
    
    role_ids = []
    for row in csv_reader:
        if len(row) >= 2:
            val = row[1]  # Cột B
            if val is not None:
                val_str = val.strip()
                if val_str:
                    role_ids.append(val_str)
    return role_ids


def send_zalo_message(chat_id, text):
    """Gửi tin nhắn văn bản phản hồi lại người dùng thông qua API Zalo Bot Platform"""
    if not BOT_TOKEN or "ĐIỀN_TOKEN" in BOT_TOKEN:
        logging.info(f"[Zalo Bot MOCK -> Gửi tới {chat_id}]: {text}")
        return
        
    url = f"https://bot-api.zaloplatforms.com/bot{BOT_TOKEN}/sendMessage"
    headers = {
        "Content-Type": "application/json"
    }
    payload = {
        "chat_id": str(chat_id),
        "text": text
    }
    try:
        res = requests.post(url, headers=headers, json=payload, timeout=10)
        logging.info(f"[Zalo Bot] Phản hồi gửi tin nhắn: {res.status_code} - {res.text}")
    except Exception as e:
        logging.error(f"[Zalo Bot] Lỗi khi gửi tin nhắn Zalo: {e}")


def execute_redeem_logic(gift_code):
    """Thực hiện luồng quét sheet và nạp code cho tất cả role ID (trả về kết quả thô)"""
    role_ids = []
    for s_url in SHEET_URLS:
        s_url = s_url.strip()
        if not s_url:
            continue
        try:
            roles = get_google_sheet_roles(s_url)
            role_ids.extend(roles)
        except Exception as e:
            logging.error(f"Lỗi tải Google Sheet ({s_url}): {str(e)}")
            raise e
            
    role_ids = list(dict.fromkeys(role_ids))
    if not role_ids:
        return [], []
        
    api_url = "https://vgrapi-sea.vnggames.com/coordinator/api/v1/code/redeem"
    headers = {
        "x-client-region": "VN",
        "Content-Type": "application/json"
    }

    success_list = []
    failed_list = []

    for idx, role_id in enumerate(role_ids, 1):
        logging.info(f"[{idx}/{len(role_ids)}] Đang redeem cho {role_id}...")
        payload = {
            "serverId": str(SERVER_ID),
            "gameCode": str(GAME_CODE),
            "roleId": role_id,
            "roleName": role_id,
            "code": gift_code
        }

        try:
            response = requests.post(api_url, headers=headers, json=payload, timeout=10)
            status_code = response.status_code
            try:
                res_json = response.json()
                error_code = res_json.get("errorCode")
                error_msg = res_json.get("message") or res_json.get("description")
                
                if status_code == 200 and error_code == 1:
                    success_list.append({"roleId": role_id})
                    logging.info(f"    -> THÀNH CÔNG: {error_msg or 'Success'}")
                else:
                    msg = error_msg or f"Lỗi code {error_code}"
                    failed_list.append({"roleId": role_id, "error": msg})
                    logging.warning(f"    -> THẤT BẠI: {msg}")
            except Exception:
                err_msg = f"HTTP Status {status_code}: {response.text}"
                failed_list.append({"roleId": role_id, "error": err_msg})
                logging.warning(f"    -> THẤT BẠI: {err_msg}")
        except Exception as e:
            failed_list.append({"roleId": role_id, "error": f"Lỗi kết nối API: {str(e)}"})
            logging.error(f"    -> LỖI KẾT NỐI: {str(e)}")

        time.sleep(0.5)

    return success_list, failed_list


def run_redeem_in_background(chat_id, gift_code):
    """Luồng chạy ngầm dùng cho Zalo Webhook"""
    try:
        success_list, failed_list = execute_redeem_logic(gift_code)
    except Exception as e:
        send_zalo_message(chat_id, f"Lỗi trong quá trình redeem: {str(e)}")
        return

    total_success = len(success_list)
    total_failed = len(failed_list)
    
    if total_success == 0 and total_failed == 0:
        send_zalo_message(chat_id, "Không tìm thấy Role ID nào hợp lệ để redeem.")
        return

    summary_text = (
        f"KẾT QUẢ NẠP CODE: {gift_code}\n"
        f"Thành công: {total_success}\n"
        f"Thất bại: {total_failed}\n"
    )
    
    if failed_list:
        summary_text += "\nChi tiết lỗi:\n"
        for item in failed_list:
            summary_text += f"- {item['roleId']}: {item['error']}\n"
            
    send_zalo_message(chat_id, summary_text)


class CombinedRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        logging.info(f"[HTTP] {self.address_string()} - - {format%args}")

    def do_GET(self):
        try:
            reload_config()
        except Exception as e:
            logging.error(f"Lỗi tự động tải cấu hình: {e}")

        parsed_url = urlparse(self.path)
        
        # 1. Endpoint / phục vụ Health Check của Render
        if parsed_url.path == '/':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "alive"}).encode('utf-8'))
            
        # 2. Endpoint /redeem phục vụ gọi từ Postman qua GET
        elif parsed_url.path == '/redeem':
            query_params = parse_qs(parsed_url.query)
            gift_codes = query_params.get("code", [])
            gift_code = gift_codes[0].strip() if gift_codes else ""
            
            if not gift_code:
                self.send_error_response(400, "Thiếu tham số 'code' bắt buộc.")
                return
                
            self.process_redeem_sync(gift_code)
        else:
            self.send_error_response(404, "Endpoint not found.")

    def do_HEAD(self):
        parsed_url = urlparse(self.path)
        # Hỗ trợ HEAD request cho Health Check của Render
        if parsed_url.path == '/':
            self.send_response(200)
            self.end_headers()
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        try:
            reload_config()
        except Exception as e:
            logging.error(f"Lỗi tự động tải cấu hình: {e}")

        parsed_url = urlparse(self.path)
        
        # 3. Endpoint /webhook xử lý tin nhắn từ Zalo Bot
        if parsed_url.path == '/webhook':
            if SECRET_TOKEN and SECRET_TOKEN != "chuoi_mat_khau_tu_dat_cua_ban":
                req_secret = self.headers.get("X-Bot-Api-Secret-Token")
                if req_secret != SECRET_TOKEN:
                    self.send_response(403)
                    self.end_headers()
                    return

            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                self.send_response(400)
                self.end_headers()
                return

            logging.info(f"Nhận webhook Zalo: {json.dumps(data, ensure_ascii=False)}")

            event_name = data.get("event_name")
            
            # Hỗ trợ cả 2 event tin nhắn của Zalo
            if event_name in ["user_send_text", "message.text.received"]:
                message_text = data.get("message", {}).get("text", "").strip()
                
                # Cú pháp: !code <MÃ_CODE>
                gift_code = ""
                if message_text.lower().startswith("!code "):
                    gift_code = message_text[6:].strip()

                if gift_code:
                    # Ưu tiên lấy CHAT_ID cấu hình cứng, nếu không thì lấy ID người gửi tự động từ payload
                    target_chat_id = CHAT_ID
                    if not target_chat_id or "ĐIỀN_CHAT_ID" in target_chat_id:
                        target_chat_id = data.get("sender", {}).get("id") or data.get("message", {}).get("chat", {}).get("id")
                    
                    if target_chat_id:
                        # Gửi phản hồi ngay lập tức đến CHAT_ID để không bị timeout 2s
                        send_zalo_message(target_chat_id, f"Đang tiến hành redeem code '{gift_code}'. Vui lòng đợi kết quả...")
                        
                        # Chạy luồng ngầm xử lý nạp code
                        threading.Thread(
                            target=run_redeem_in_background,
                            args=(target_chat_id, gift_code),
                            daemon=True
                        ).start()
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
            
        # 4. Endpoint /redeem xử lý gọi Postman dạng POST (Đồng bộ)
        elif parsed_url.path == '/redeem':
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            try:
                params = json.loads(post_data.decode('utf-8'))
            except Exception:
                params = {}

            gift_code = params.get("code", "")
            if isinstance(gift_code, str):
                gift_code = gift_code.strip()
            else:
                gift_code = ""
            
            if not gift_code:
                self.send_error_response(400, "Thiếu tham số 'code' bắt buộc.")
                return
                
            self.process_redeem_sync(gift_code)
        else:
            self.send_error_response(404, "Endpoint not found.")

    def process_redeem_sync(self, gift_code):
        """Xử lý redeem đồng bộ để trả kết quả JSON về cho Postman"""
        try:
            success_list, failed_list = execute_redeem_logic(gift_code)
            response_payload = {
                "totalSuccess": len(success_list),
                "totalFailed": len(failed_list),
                "success": success_list,
                "failed": failed_list
            }
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps(response_payload, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_error_response(500, f"Lỗi hệ thống: {str(e)}")

    def send_error_response(self, status_code, message):
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        response = {"error": message}
        self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))


def run_server():
    server_address = ('', PORT)
    httpd = HTTPServer(server_address, CombinedRequestHandler)
    external_url = os.environ.get("RENDER_EXTERNAL_URL")
    
    webhook_url = f"{external_url}/webhook" if external_url else f"http://localhost:{PORT}/webhook"
    api_url = f"{external_url}/redeem" if external_url else f"http://localhost:{PORT}/redeem"
    
    logging.info("=" * 60)
    logging.info(f"Server chạy cổng: {PORT}")
    logging.info(f"Đường dẫn Webhook: {webhook_url}")
    logging.info(f"Đường dẫn gọi API: {api_url}")
    logging.info("=" * 60)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        logging.info("Đang tắt Server...")
        httpd.server_close()
        logging.info("Đã tắt Server.")

if __name__ == "__main__":
    run_server()
