import os
import sys
import time
import json
import re
import csv
import threading
import logging
from http.server import BaseHTTPRequestHandler, HTTPServer
import requests
import openpyxl

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

# Cấu hình logging để in ra console
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

# ==============================================================================
# CẤU HÌNH ZALO BOT & VNG REDEEM
# ==============================================================================
BOT_TOKEN = "ĐIỀN_TOKEN_ZALO_BOT_CỦA_BẠN"      # Token nhận từ bot.zaloplatforms.com
SECRET_TOKEN = "chuoi_mat_khau_tu_dat_cua_ban" # Secret token cấu hình khi setWebhook (tùy chọn)

SERVER_ID = "2"
GAME_CODE = "661"
SHEET_URLS = ["https://docs.google.com/spreadsheets/d/1mdv1O31HGALyDTeZhmjn0aLNmjOmpR_3fO6RcudTerU/edit?usp=sharing"]
# SHEET_URLS = [
#     "https://docs.google.com/spreadsheets/d/1wIL_pO9wdZjq5TX4S-e_zYgo0_Zc3O0_EKcVUWlDmz4/edit?usp=sharing"
#     ,"https://docs.google.com/spreadsheets/d/1s-K2MO92uzwkKSk7dZDo4vR6K7XicRvfIp63Vo7c-gA/edit?usp=sharing"
# ]
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


def run_redeem_in_background(chat_id, gift_code):
    """Luồng xử lý chạy ngầm tải Google Sheet, redeem code và trả kết quả về Zalo chat"""
    role_ids = []
    
    # 1. Tải danh sách roleId từ Google Sheet
    for s_url in SHEET_URLS:
        s_url = s_url.strip()
        if not s_url:
            continue
        try:
            roles = get_google_sheet_roles(s_url)
            role_ids.extend(roles)
        except Exception as e:
            send_zalo_message(chat_id, f"Lỗi tải Google Sheet ({s_url}): {str(e)}")
            return
            
    # Loại bỏ các role ID trùng lặp
    role_ids = list(dict.fromkeys(role_ids))
    if not role_ids:
        send_zalo_message(chat_id, "Không tìm thấy Role ID nào trong Google Sheets của bạn.")
        return
        
    send_zalo_message(chat_id, f"Đang xử lý nạp code '{gift_code}' cho {len(role_ids)} tài khoản. Vui lòng đợi giây lát...")
    
    # 2. Thực hiện gọi API Redeem VNG
    api_url = "https://vgrapi-sea.vnggames.com/coordinator/api/v1/code/redeem"
    headers = {
        "x-client-region": "VN",
        "Content-Type": "application/json"
    }

    success_list = []
    failed_list = []

    for idx, role_id in enumerate(role_ids, 1):
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
                    success_list.append(role_id)
                else:
                    msg = error_msg or f"Lỗi code {error_code}"
                    failed_list.append((role_id, msg))
            except Exception:
                if status_code == 200:
                    err_msg = f"Phản hồi không phải JSON: {response.text}"
                else:
                    err_msg = f"HTTP Status {status_code}: {response.text}"
                failed_list.append((role_id, err_msg))
        except Exception as e:
            failed_list.append((role_id, f"Lỗi kết nối API: {str(e)}"))

        # Tránh rate limiting của máy chủ
        time.sleep(0.5)

    # 3. Gửi tin nhắn thống kê kết quả về Zalo Chat
    summary_text = (
        f"KẾT QUẢ NẠP CODE: {gift_code}\n"
        f"Thành công: {len(success_list)}\n"
        f"Thất bại: {len(failed_list)}\n"
    )
    
    if failed_list:
        summary_text += "\nChi tiết lỗi:\n"
        for role_id, err in failed_list:
            summary_text += f"- {role_id}: {err}\n"
            
    send_zalo_message(chat_id, summary_text)


class ZaloWebhookRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # Log request thông qua logger thay vì print trực tiếp
        logging.info(f"[HTTP] {self.address_string()} - - {format%args}")

    def do_POST(self):
        if self.path == '/webhook':
            # Kiểm tra mã bí mật (SECRET_TOKEN) nếu được cấu hình
            if SECRET_TOKEN and SECRET_TOKEN != "chuoi_mat_khau_tu_dat_cua_ban":
                req_secret = self.headers.get("X-Bot-Api-Secret-Token")
                if req_secret != SECRET_TOKEN:
                    self.send_response(403)
                    self.end_headers()
                    self.wfile.write(b"Forbidden: Invalid X-Bot-Api-Secret-Token")
                    return

            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"Bad Request: Invalid JSON")
                return

            event_name = data.get("event_name")
            
            # Chỉ xử lý khi người dùng nhắn tin text vào Bot Zalo
            if event_name == "user_send_text":
                sender_id = data.get("sender", {}).get("id")
                message_text = data.get("message", {}).get("text", "").strip()
                
                # Cú pháp hỗ trợ: "!code <MÃ_CODE>"
                gift_code = ""
                if message_text.lower().startswith("!code "):
                    gift_code = message_text[6:].strip()

                if gift_code and sender_id:
                    # Chạy xử lý trong luồng riêng để trả về 200 OK cho Zalo ngay lập tức (< 2 giây)
                    threading.Thread(
                        target=run_redeem_in_background,
                        args=(sender_id, gift_code),
                        daemon=True
                    ).start()
            
            # Luôn trả về 200 OK thành công cho Zalo Server
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"Not Found")


def run_server():
    server_address = ('', PORT)
    httpd = HTTPServer(server_address, ZaloWebhookRequestHandler)
    external_url = os.environ.get("RENDER_EXTERNAL_URL")
    webhook_url = f"{external_url}/webhook" if external_url else f"http://localhost:{PORT}/webhook"
    
    logging.info("=" * 60)
    logging.info(f"Server Webhook Zalo Bot đang chạy tại cổng: {PORT}")
    logging.info(f"Đường dẫn Webhook: {webhook_url}")
    logging.info("=" * 60)
    logging.info("Nhấn Ctrl + C để tắt Server.\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        logging.info("Đang tắt Server...")
        httpd.server_close()
        logging.info("Đã tắt Server.")

if __name__ == "__main__":
    run_server()
