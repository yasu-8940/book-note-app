"""
Google Cloud Vision OCRを用いてpdfデータのOCRをかける
※pdfの場合は通常の画像データ（jpg/png）とは別の処理を行う
  導入に少し手間取る
※pythonの通常のライブラリを用いた場合に精度の限界があったためこれを試す
"""
from google.cloud import vision_v1
from google.cloud import storage
from google.cloud.vision_v1 import types
import os
import json
import csv
import glob

def get_files_by_extensions_glob(directory, extensions):
    """
    glob モジュールを使用して、指定されたディレクトリ内の特定の拡張子のファイルをリスト化します。
    
    Args:
      directory: 検索対象のディレクトリパス。
      extensions: 検索対象の拡張子を格納したリスト (例: ['.txt', '.py'])。
      
    Returns:
      指定された拡張子のファイル名のリスト。
    """
    print("directory =", directory)
    file_list = []
    for filename in os.listdir(directory):
        if any(filename.endswith(ext) for ext in extensions):
            full_path = os.path.join(directory, filename)
            print("✔ join:", full_path)
            file_list.append(full_path)
    return file_list

json_output_dir = r"C:\Users\seki8\OneDrive\デスクトップ\python_lesson\tmp_OCR"
csv_output_dir = r"C:\Users\seki8\OneDrive\デスクトップ\python_lesson\after_OCR"
os.makedirs(json_output_dir, exist_ok=True)  # なければ作成、あればスルー

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\seki8\lithe-catbird-446804-q2-441e0e8915c9.json"

directory_path = "C:/Users/seki8/OneDrive/デスクトップ/python_lesson/befor_OCR"  # ファイルが格納されているパス
target_extensions = [".pdf"]

# GCS設定
vision_client = vision_v1.ImageAnnotatorClient()

bucket_name = "my-pdf-ocr-output"  # アップロード用のバケット名
destination_uri_base = "gs://my-pdf-ocr-output/output/"  # OCR結果保存先

local_pdf_paths = get_files_by_extensions_glob(directory_path, target_extensions)
print("glob:", local_pdf_paths)

# GCSクライアント
storage_client = storage.Client()

# 各PDFを処理
for local_pdf_path in local_pdf_paths:
    # 1. PDFをGCSへアップロード
    pdf_name = os.path.basename(local_pdf_path)
    pdf_base = os.path.splitext(pdf_name)[0]  # 'receipt-01.pdf' → 'receipt-01'
    gcs_blob_path = f"input/{pdf_name}"
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(gcs_blob_path)
    blob.upload_from_filename(local_pdf_path)
    print(f"Uploaded {local_pdf_path} to gs://{bucket_name}/{gcs_blob_path}")

    # 2. GCS URI（OCR入力）
    gcs_source_uri = f"gs://{bucket_name}/{gcs_blob_path}"
    gcs_source = types.GcsSource(uri=gcs_source_uri)
    input_config = types.InputConfig(gcs_source=gcs_source, mime_type="application/pdf")

    # 3. OCR出力
    gcs_destination_uri = destination_uri_base + pdf_name.replace(".pdf", "/")
    gcs_destination = types.GcsDestination(uri=gcs_destination_uri)
    output_config = types.OutputConfig(gcs_destination=gcs_destination, batch_size=1) 

    async_request = types.AsyncAnnotateFileRequest(
        features=[types.Feature(type=vision_v1.Feature.Type.DOCUMENT_TEXT_DETECTION)],
        input_config=input_config,
        output_config=output_config,
    )

    # 4. 非同期OCRリクエスト送信
    operation = vision_client.async_batch_annotate_files(requests=[async_request])
    print(f"OCRリクエストを送信しました: {pdf_name} → {gcs_destination_uri}")

    # 5. 完了まで待つ
    operation.result(timeout=300)
    # print(f"OCR完了: {pdf_name}")

    # GCSにある結果ファイルを取得
    client = storage.Client()
    bucket_name = "my-pdf-ocr-output"
    # prefix = "output/"  # OCR結果のプレフィックス
    prefix = f"output/{pdf_base}/"

    blobs = client.list_blobs(bucket_name, prefix=prefix)
    # print("blobs=", blobs)

    for blob in blobs:
        print(blob.name)
        if blob.name.endswith(".json"):
            # base_name = os.path.splitext(os.path.basename(blob.name))[0]
            local_json_path = os.path.join(json_output_dir, f"{pdf_base}.json")
            blob.download_to_filename(local_json_path)
            print(f"保存しました: {local_json_path}")

            with open(local_json_path, "r", encoding="utf-8-sig") as f:
              data = json.load(f)

            csv_path = os.path.join(csv_output_dir, f"{pdf_base}.csv")
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
              writer = csv.writer(csvfile, delimiter=",")
              writer.writerow(["ファイル名", "OCR結果"])

              for response in data.get("responses", []):
                text = response.get("fullTextAnnotation", {}).get("text", "").strip()
                cleaned_text = text.replace("\n", " ").strip()
                writer.writerow([pdf_base, cleaned_text if cleaned_text else "（文字なし）"])
