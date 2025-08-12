import streamlit as st
import requests
import csv
import os
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
from io import BytesIO
import os
from datetime import datetime
import pandas as pd
import tempfile

def create_excel_with_image(book, comment):
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
        ''
    ]
    ws.append(row)

    # 表紙画像の取得・挿入
    if book['thumbnail']:
        try:
            response = requests.get(book['thumbnail'])
            img = Image.open(BytesIO(response.content))

            # 一時ファイルパスを作成
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
                img_path = tmp_file.name
                img.save(img_path)

            excel_img = XLImage(img_path)
            row_num = ws.max_row
            ws.add_image(excel_img, f'H{row_num}')  # 画像列はH（8列目）

        except Exception as e:
            print(f"画像処理でエラー: {e}")
    
            # img_path = "cover_tmp.png"
            # img.save(img_path)

            # excel_img = XLImage(img_path)
            # ws.add_image(excel_img, f'H2')  # 表紙画像は2行目のH列に

            # 一時ファイル削除
            os.remove(img_path)
            
        except Exception as e:
            st.warning(f"画像挿入に失敗しました: {e}")

    # Excelバイナリとして保存
    excel_io = BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)

    return excel_io

# CSVファイルに書き出し
def write_to_csv(book, comment, filename=r"C:\Users\seki8\OneDrive\デスクトップ\python_lesson\読書ノート.csv"):
    file_exists = os.path.isfile(filename)
    with open(filename, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['書名', '著者', '出版社', '出版日', '概要', '感想'])
        writer.writerow([
            book['title'],
            book['authors'],
            book['publisher'],
            book['publishedDate'],
            book['description'],
            comment
        ])

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

# Streamlit アプリ
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

    # if st.button("CSVに保存"):
    #     write_to_csv(selected_book, comment)
    #     st.success("CSVに保存しました！")

    # ✅ Streamlit 側のダウンロードボタン（呼び出し例）
    if st.button("Excelでダウンロード（表紙付き）"):
        excel_data = create_excel_with_image(selected_book, comment)
        st.download_button(
            label="📥 Excelファイルをダウンロード",
            data=excel_data,
            file_name="book_note.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )





