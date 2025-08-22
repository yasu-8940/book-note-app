import streamlit as st
import os
import json
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
import io
# from __future__ import print_function
from pathlib import Path
from googleapiclient.http import MediaIoBaseUpload

# 🔹 Google Drive API のスコープ
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# =========================================================
# Google Drive サービス取得
# =========================================================
def get_gdrive_service():
    creds = None
    token_path = 'token.pickle'

    # --- 1. token.pickle があれば再利用 ---
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)

    # --- 2. 認証が無効なら再認証 ---
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        
        else:
            creds_json = os.environ.get("GOOGLE_CREDENTIALS")
            if creds_json:
                # Renderなどクラウド環境：環境変数からロード
                creds_dict = json.loads(creds_json)
                flow = InstalledAppFlow.from_client_config(creds_dict, SCOPES)
            else:
                # ローカル環境：credentials.json を読む
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)

            creds = flow.run_local_server(port=0)

        # トークンを保存（ローカルでもRenderでも使える）
        with open(token_path, "wb") as token:
            pickle.dump(creds, token)

    return build('drive', 'v3', credentials=creds)

# =========================================================
# Excel ファイル作成（表紙画像付き）
# =========================================================
def create_excel_with_image(book, comment, filename="book_note.xlsx"):
    if os.path.exists(filename):
        try:
            wb = load_workbook(filename)
            ws = wb.active
        except Exception:
            # 壊れている場合は作り直す
            wb = Workbook()
            ws = wb.active
            ws.append(['登録日','書名','著者','出版社','出版日','概要','感想','表紙'])    
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
        ws.add_image(excel_img, f'H{ws.max_row}')

    # ✅ バイナリ化して返す（これが重要！）
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    # ✅ 保存後に削除
    if book['thumbnail'] and os.path.exists(img_path):
        os.remove(img_path)

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
def upload_to_gdrive(service, file_id, excel_data, filename="book_note.xlsx"):
    media = MediaIoBaseUpload(io.BytesIO(excel_data), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
    updated_file = service.files().update(
        fileId=file_id,
        media_body=media
    ).execute()
    return updated_file

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
        excel_data = create_excel_with_image(selected_book, comment)

        # 事前にファイルIDを保存しておく（初回だけアップロードしてID取得）
        file_id = "1a9jmvgdg1W9mnsdoCL8Qv6wAqkBHJFRp"  # あなたのDrive上の book_note.xlsx のID
        service = get_gdrive_service()
        updated_file = upload_to_gdrive(service, file_id, excel_data)

        st.success(f"✅ Google Driveに上書き保存しました！ ({updated_file['name']})")


