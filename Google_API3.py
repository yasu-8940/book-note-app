import streamlit as st
import os
import json
import pickle
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from datetime import datetime
from google.auth.transport.requests import Request
import requests
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import io, requests
# from __future__ import print_function
from pathlib import Path
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# 🔹 Google Drive API のスコープ
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# =========================================================
# Google Drive サービス取得
# =========================================================

def get_gdrive_service():
    """
    Google Drive API サービスを返す（サービスアカウント方式）
    Renderでは環境変数 GOOGLE_CREDENTIALS から読み込み、
    ローカルでは service_account.json を読み込む
    """
    creds = None

    # Render 環境
    if "GOOGLE_CREDENTIALS" in os.environ:
        service_account_info = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)

    # ローカル環境
    elif os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)

    else:
        raise FileNotFoundError("サービスアカウントの認証情報が見つかりません。")

    return build("drive", "v3", credentials=creds)

# =========================================================
# Excel ファイル作成（表紙画像付き）
# =========================================================
def create_excel_with_image(book, comment, base_xlsx_bytes=None, filename="book_note.xlsx"):

    # 既存ファイルのBytesが来ていればそれをベースに、無ければ新規
    if base_xlsx_bytes:
        wb = load_workbook(filename=BytesIO(base_xlsx_bytes))
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(['登録日','書名','著者','出版社','出版日','概要','感想','表紙'])

    today = datetime.today().strftime("%Y-%m-%d")
    ws.append([
        today,
        book.get('title',''),
        book.get('authors',''),
        book.get('publisher',''),
        book.get('publishedDate',''),
        book.get('description',''),
        comment,
        ''
    ])

    # 表紙（PIL→openpyxl 直接渡し）
    if book.get('thumbnail'):
        r = requests.get(book['thumbnail'], timeout=10)
        r.raise_for_status()
        img_pil = Image.open(BytesIO(r.content))
        excel_img = XLImage(img_pil)
        ws.add_image(excel_img, f"H{ws.max_row}")

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

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

# =========================================================
# GoogleAPPから本を探す
# =========================================================

def search_books_google_books(title):
    url = 'https://www.googleapis.com/books/v1/volumes'
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    params = {
        'q': title,
        'maxResults': 10,
        'printType': 'books',
        'langRestrict': 'ja',
        # 'key': 'AIzaSyA0gXAcX6_aShRD4eKPA6ag_4QBTQtvC0w'
    }
    try:
        response = requests.get(url, headers=headers, params=params)


        data = response.json()

        if 'items' not in data:
            st.warning("⚠️ 書籍が見つかりませんでした。")
            return []

        results = []

        for item in data['items']:
            info = item['volumeInfo']
            results.append({
                'title': info.get('title', ''),
                'authors': ', '.join(info.get('authors', [])),
                'publishedDate': info.get('publishedDate', ''),
                'description': info.get('description', ''),
                'thumbnail': info.get('imageLinks', {}).get('thumbnail', ''),
                'publisher': info.get('publisher', ''),
         })
        return results

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
        return []

# Google Drive 上書き保存関数
def upload_to_drive(excel_data, folder_id, filename="book_note.xlsx"):
    service = get_gdrive_service()

    # 既存ファイルがあるか検索
    query = f"'{folder_id}' in parents and name='{filename}' and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])

    media = MediaIoBaseUpload(
        io.BytesIO(excel_data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if items:
        # 更新（上書き）
        file_id = items[0]["id"]
        updated_file = service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        print(f"✅ 上書き保存OK: {filename} ({file_id})")
    else:
        # 新規作成
        file_metadata = {
            "name": filename,
            "parents": [folder_id]
        }
        new_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()
        print(f"✅ 新規作成OK: {filename} ({new_file['id']})")

# =========================================================
# Driveからファイルをダウンロード
# =========================================================

def download_from_drive(folder_id, filename="book_note.xlsx"):
    service = get_gdrive_service()

    # Drive上にファイルがあるか検索
    query = f"'{folder_id}' in parents and name='{filename}' and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    items = results.get("files", [])

    if not items:
        return None  # ファイルがまだ存在しない

    file_id = items[0]["id"]
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    return fh.getvalue()

# =========================================================
# Streamlit アプリ
# =========================================================
st.title("📚 読書ノート:シリーズ対応版（Google Books API）")

search_query = st.text_input("書名を入力してください（シリーズ名もOK）：")

if st.button("候補を検索"):
    try:
        results = search_books_google_books(search_query)
        st.session_state['search_results'] = results
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
    # results = search_books_google_books(search_query)
    # st.session_state['search_results'] = results

# 結果がある場合のみ処理
if 'search_results' in st.session_state and st.session_state['search_results']:
    results = st.session_state['search_results']
    options = [f"{book['title']} / {book['authors']}" for book in results]
    
    # 🔑 radioボタンは毎回再描画されるように
    selected = st.radio("候補から選んでください：", options, key="book_radio")
    selected_book = results[options.index(selected)]

    # 詳細表示
    st.subheader(selected_book['title'])
    st.write(f"著者: {selected_book['authors']}")
    st.write(f"出版社: {selected_book['publisher']}")
    st.write(f"出版日: {selected_book['publishedDate']}")
    st.write("概要:")
    st.write(selected_book['description'])

    if selected_book['thumbnail']:
        st.image(selected_book['thumbnail'], caption='表紙画像', width=150)

    # 感想入力
    st.markdown("---")
    comment = st.text_area("📖 感想を入力してください:")

    # ✅ Streamlit 側のダウンロードボタン（呼び出し例）
    if st.button("Excelでダウンロード（表紙付き）"):
        excel_data = create_excel_with_image(selected_book, comment)
        st.download_button(
            label="📥 Excelファイルをダウンロード",
            data=excel_data,
            file_name="book_note.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ✅ Streamlit Google Drive保存ボタン 
    if st.button("📤 Google Driveに保存（上書き）"):

        folder_id = "1CP9mzd7dOaPG9Fj88vY6OYSKwl7el1XT"

        # 1. Driveから既存Excelをダウンロード
        existing_bytes = download_from_drive(folder_id, "book_note.xlsx")

        # 2. 行を追加
        excel_data = create_excel_with_image(selected_book, comment, base_xlsx_bytes=existing_bytes)
        
        # 3. Drive へ保存（ここで folder_id を使う）

        upload_to_drive(excel_data, folder_id, filename="book_note.xlsx")
        st.success("✅ Google Driveに保存しました！")

