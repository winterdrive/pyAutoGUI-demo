import io
import os
import logging
import traceback
from typing import Optional
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# load .env file "UPLOAD_FOLDER" by dontenv
load_dotenv()
upload_folder = os.getenv('UPLOAD_FOLDER')
download_url = os.getenv('DOWNLOAD_URL')
credentials = os.getenv('CLIENT_SECRET_FILE_PATH')

# 配置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 常量配置
UPLOAD_FOLDER = upload_folder
SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', credentials)
OUTPUT_IMAGE_PATH = "D:\\Users\\Downloads\\信仰曆\\2025嚕嚕米日曆\\lulumi圖\\"
EXCEL_FILE_PATH = "D:\\Users\\Downloads\\信仰曆\\2025嚕嚕米日曆\\2025嚕嚕米日曆｜語錄編輯表單 - 唐看這個.xlsx"


def get_drive_service():
    """建立並返回 Google Drive 服務客戶端"""
    try:
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        logger.error(f"Failed to create Google Drive service: {e}")
        raise


def upload_file(filename: str, folder_id: Optional[str] = UPLOAD_FOLDER):
    """上傳檔案到 Google Drive"""
    try:
        service = get_drive_service()
        media = MediaFileUpload(filename)
        file_metadata = {'name': filename, 'parents': [folder_id]}

        logger.info(f"開始上傳檔案 {filename}...")
        file = service.files().create(body=file_metadata, media_body=media).execute()
        logger.info(f"檔案上傳成功，file ID: {file.get('id')}")
    except HttpError as error:
        logger.error(f"An error occurred: {error}")
        raise
    except Exception as e:
        logger.error(f"Failed to upload file: {e}")
        raise


def download_file(file_id: str, output_path: str):
    """下載檔案從 Google Drive"""
    try:
        service = get_drive_service()
        request = service.files().get_media(fileId=file_id)
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)

        logger.info(f"開始下載檔案 ID: {file_id}...")
        done = False
        while not done:
            status, done = downloader.next_chunk()
            logger.info(f"Download progress: {int(status.progress() * 100)}%")

        # 將下載的內容寫入指定路徑
        with open(output_path, 'wb') as f:
            f.write(file.getvalue())
        logger.info(f"檔案下載成功，存儲至: {output_path}")

        return file.getvalue()
    except HttpError as error:
        logger.error(f"An error occurred: {error}")
        traceback.print_exc()
        raise
    except Exception as e:
        logger.error(f"Failed to download file: {e}")
        traceback.print_exc()
        raise


def url_parser(url: str):
    """
    從 Google Drive 共享連結中解析出檔案 ID
    """
    import re
    pattern = "([-\w]{25,})"
    fileId = re.findall(pattern, url)[0]
    return fileId


def load_excel_to_get_download_url_by_pandas(excel_path: str):
    """
    透過 pandas 讀取 Excel 檔案，獲取下載連結
    """
    import pandas as pd
    df = pd.read_excel(excel_path)
    download_url_list = df['圖片'].values
    return download_url_list


if __name__ == '__main__':
    target_results = load_excel_to_get_download_url_by_pandas(f"{EXCEL_FILE_PATH}")
    i = 1  # 設定初始編號
    for download_url in target_results:
        output_path = f"{OUTPUT_IMAGE_PATH}{str(i).zfill(3)}.jpg"

        if os.path.exists(output_path):
            print(f"{output_path} 已存在。")
            i += 1  # 跳過已存在的檔案
            continue

        try:
            file_id = url_parser(download_url)
            download_file(file_id, output_path)
            print(f"文件 {output_path} 下載成功。")
            i += 1  # 下載成功，增加編號
        except Exception as e:
            print(f"無法下載 {output_path}: {e}")
            i += 1  # 下載失敗，跳過編號
