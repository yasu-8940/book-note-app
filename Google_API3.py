import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime
import requests
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

# 🔹 Google Drive API のスコープ
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# =========================================================
# Google Drive サービス取得
# =========================================================
def get_gdrive_service():
    creds = None
    token_path = 'token.pickle'
    creds_path = 'credentials.json'

    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)

    return build('drive', 'v3', credentials=creds)

# =========================================================
# Excel ファイル作成（表紙画像付き）
# =========================================================
def create_excel_with_image(book, comment, filename="book_note.xlsx"):
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(['登録日', '書名', '著者', '出版社', '出版日', '概要', '感想', '表紙'])

    today = datetime.today().strftime("%Y-%m-%d")
    row = [
        today,
        book['title'],
        book['authors'],
        book['publisher'],
        book['publishedDate'],
        book['description'],
        comment,
        '',  # 画像用
    ]
    ws.append(row)

    if book['thumbnail']:
        response = requests.get(book['thumbnail'])
        img = Image.open(BytesIO(response.content))
        img_path = "cover_tmp.png"
        img.save(img_path)
        excel_img = XLImage(img_path)
        row_num = ws.max_row
        ws.add_image(excel_img, f'H{row_num}')
        os.remove(img_path)

    wb.save(filename)
    return filename

# =========================================================
# Google Drive にアップロード（存在すれば上書き）
# =========================================================
def upload_to_gdrive(local_path, drive_filename="book_note.xlsx"):
    service = get_gdrive_service()

    # 既存ファイルを探す
    query = f"name='{drive_filename}' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])

    file_metadata = {"name": drive_filename}
    media = MediaFileUpload(local_path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if items:
        # 既存ファイルを上書き
        file_id = items[0]['id']
        updated = service.files().update(fileId=file_id, media_body=media).execute()
        print(f"✅ 上書き保存OK: {updated['name']} ({file_id})")
        return file_id
    else:
        # 新規アップロード
        uploaded = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        print(f"✅ 新規アップロードOK: {uploaded['id']}")
        return uploaded['id']






