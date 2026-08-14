import os
import sys
import time
import json
import re
import csv
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs
import requests
import openpyxl

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

# ==========================================
# CẤU HÌNH CỐ ĐỊNH (HARDCODE)
# ==========================================
SERVER_ID = "2"
GAME_CODE = "661"
SHEET_URLS = ["https://docs.google.com/spreadsheets/d/1mdv1O31HGALyDTeZhmjn0aLNmjOmpR_3fO6RcudTerU/edit?usp=sharing"]
# SHEET_URLS = [
#     "https://docs.google.com/spreadsheets/d/1wIL_pO9wdZjq5TX4S-e_zYgo0_Zc3O0_EKcVUWlDmz4/edit?usp=sharing"
#     ,"https://docs.google.com/spreadsheets/d/1s-K2MO92uzwkKSk7dZDo4vR6K7XicRvfIp63Vo7c-gA/edit?usp=sharing"
# ]
# ==========================================

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

class RedeemRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # Override to log requests in a clean way
        print(f"[HTTP] {self.address_string()} - - [{self.log_date_time_string()}] {format%args}")

    def do_GET(self):
        parsed_url = urlparse(self.path)
        if parsed_url.path == '/redeem':
            # Lấy tham số qua Query String (Chỉ cần 'code')
            query_params = parse_qs(parsed_url.query)
            gift_codes = query_params.get("code", [])
            gift_code = gift_codes[0].strip() if gift_codes else ""
            
            if not gift_code:
                self.send_error_response(400, "Thiếu tham số 'code' bắt buộc.")
                return
            
            self.process_redeem(gift_code)
        else:
            self.send_error_response(404, "Endpoint not found. Use GET or POST to /redeem")

    def do_POST(self):
        parsed_url = urlparse(self.path)
        if parsed_url.path == '/redeem':
            # Đọc Body JSON của POST request
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
            
            self.process_redeem(gift_code)
        else:
            self.send_error_response(404, "Endpoint not found. Use GET or POST to /redeem")

    def process_redeem(self, gift_code):
        role_ids = []
        if SHEET_URLS:
            for s_url in SHEET_URLS:
                s_url = str(s_url).strip()
                if not s_url:
                    continue
                print(f"\n[+] Đang tải dữ liệu từ Google Sheet: {s_url}...")
                try:
                    roles_from_sheet = get_google_sheet_roles(s_url)
                    role_ids.extend(roles_from_sheet)
                except Exception as e:
                    self.send_error_response(500, f"Lỗi tải Google Sheet ({s_url}): {str(e)}")
                    return
        else:
            excel_path = "idacc.xlsx"
            if not os.path.exists(excel_path):
                self.send_error_response(404, f"Không tìm thấy file Excel: {excel_path} và không cấu hình SHEET_URLS")
                return

            print(f"\n[+] Đang đọc dữ liệu từ file Excel cục bộ: {excel_path}...")
            try:
                wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
                sheet = wb.active
                for row in sheet.iter_rows(values_only=True):
                    if len(row) >= 2:
                        val = row[1]  # Cột B
                        if val is not None:
                            val_str = str(val).strip()
                            if val_str:
                                role_ids.append(val_str)
            except Exception as e:
                self.send_error_response(500, f"Lỗi đọc file Excel: {str(e)}")
                return

        # Loại bỏ trùng lặp nếu trùng roleId giữa các sheet hoặc dòng
        role_ids = list(dict.fromkeys(role_ids))

        if not role_ids:
            self.send_error_response(400, "Không tìm thấy Role ID nào có dữ liệu.")
            return

        print(f"\n[+] Bắt đầu xử lý {len(role_ids)} account(s) cho Giftcode: '{gift_code}'")

        # 2. Gọi API coordination redeem
        api_url = "https://vgrapi-sea.vnggames.com/coordinator/api/v1/code/redeem"
        headers = {
            "x-client-region": "VN",
            "Content-Type": "application/json"
        }

        success_list = []
        failed_list = []

        for idx, role_id in enumerate(role_ids, 1):
            print(f"    [{idx}/{len(role_ids)}] Đang redeem cho {role_id}...")
            
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
                
                # Thử giải mã JSON phản hồi (cho cả 200 và các mã lỗi như 400)
                try:
                    res_json = response.json()
                    error_code = res_json.get("errorCode")
                    error_msg = res_json.get("message") or res_json.get("description")
                    
                    if status_code == 200 and error_code == 1:
                        success_list.append({
                            "roleId": role_id
                        })
                        print(f"        -> THÀNH CÔNG: {error_msg or 'Success'}")
                    else:
                        # Thất bại từ API VNG
                        msg = error_msg or f"Lỗi code {error_code}"
                        failed_list.append({
                            "roleId": role_id,
                            "error": msg
                        })
                        print(f"        -> THẤT BẠI: {msg}")
                except Exception:
                    # Trường hợp phản hồi không phải JSON
                    if status_code == 200:
                        err_msg = f"Phản hồi không phải JSON: {response.text}"
                    else:
                        err_msg = f"HTTP Status {status_code}: {response.text}"
                        
                    failed_list.append({
                        "roleId": role_id,
                        "error": err_msg
                    })
                    print(f"        -> THẤT BẠI: {err_msg}")

            except Exception as e:
                err_msg = f"Lỗi kết nối API: {str(e)}"
                failed_list.append({
                    "roleId": role_id,
                    "error": err_msg
                })
                print(f"        -> LỖI KẾT NỐI: {str(e)}")

            # Tránh rate-limiting
            time.sleep(0.5)

        # 3. Trả về kết quả JSON
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
        print(f"[+] Hoàn thành gọi API. Thành công: {len(success_list)} | Thất bại: {len(failed_list)}\n")

    def send_error_response(self, status_code, message):
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        response = {"error": message}
        self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))

def run_server(port=5000):
    server_address = ('', port)
    httpd = HTTPServer(server_address, RedeemRequestHandler)
    print("=" * 60)
    print(f"Server API đang chạy tại: http://localhost:{port}/redeem")
    print(f"Bạn có thể mở Postman và gọi GET hoặc POST tới endpoint này.")
    print("=" * 60)
    print("Nhấn Ctrl + C để dừng Server.\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[-] Đang tắt Server...")
        httpd.server_close()
        print("[+] Đã tắt Server.")

if __name__ == "__main__":
    run_server()
