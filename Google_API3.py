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

def write_to_excel_with_image(book, comment, filename=r"C:\Users\seki8\OneDrive\デスクトップ\python_lesson\読書ノート.xlsx"):
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(['登録日', '書名', '著者', '出版社', '出版日', '概要', '感想', '表紙'])

    # ✅ 今日の日付（登録日）
    today = datetime.today().strftime("%Y-%m-%d")

    # 書誌データ
    row = [
        today,
        book['title'],
        book['authors'],
        book['publisher'],
        book['publishedDate'],
        book['description'],
        comment,
        '',  # 画像用の列
    ]
    ws.append(row)

    # 表紙画像があるなら貼り付け
    if book['thumbnail']:
        response = requests.get(book['thumbnail'])
        img = Image.open(BytesIO(response.content))
        img_path = "cover_tmp.png"
        img.save(img_path)

        excel_img = XLImage(img_path)
        row_num = ws.max_row
        ws.add_image(excel_img, f'G{row_num}')

    wb.save(filename)

    # 一時画像ファイル削除
    if os.path.exists("cover_tmp.png"):
        os.remove("cover_tmp.png")

def search_books_google_books(title):
    params = {
        'q': title,
        'maxResults': 10,
        'printType': 'books',
        'langRestrict': 'ja',
    }
    response = requests.get('https://www.googleapis.com/books/v1/volumes', params=params)
    data = response.json()

    if 'items' not in data:
        return None
    
    books = []

    for item in data['items']:
        info = item['volumeInfo']
        books.append({
            'title': info.get('title', ''),
            'authors': ', '.join(info.get('authors', [])),
            'publishedDate': info.get('publishedDate', ''),
            'description': info.get('description', ''),
            'thumbnail': info.get('imageLinks', {}).get('thumbnail', ''),
            'publisher': info.get('publisher', ''),
        })
    return books

# Streamlit アプリ
st.title("📚 読書ノート:シリーズ対応版（Google Books API）")

search_query = st.text_input("書名を入力してください（シリーズ名もOK）：")

if st.button("候補を検索"):
    results = search_books_google_books(search_query)
    st.session_state['search_results'] = results

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

    if st.button("Excelに保存"):
        write_to_excel_with_image(selected_book, comment)
        st.success("Excelに保存しました（表紙付き）！")
