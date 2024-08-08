import os

from dotenv import load_dotenv
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build, Resource


def create_session():
    # load .env file "UPLOAD_FOLDER" by dontenv
    load_dotenv()
    credentials = os.getenv('GOOGLE_PHOTO_FILE_PATH')
    scopes = ['https://www.googleapis.com/auth/photoslibrary.readonly']

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials, scopes)
            creds = flow.run_local_server()
        print(creds)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('photoslibrary', 'v1', credentials=creds, static_discovery=False)

    return service


def extract_album_id(url):
    import re

    match = re.search(r'photos\.app\.goo\.gl/(\w+)', url)
    return match.group(1) if match else None


def access_album(service: Resource):
    # Call the Photo v1 API，列出當前所有你已手動加入/建立的相簿
    results = service.albums().list(
        pageSize=10, fields="nextPageToken,albums(id,title)").execute()
    items = results.get('albums', [])
    if not items:
        print('No albums found.')
    else:
        print('Albums:')
        for item in items:
            print('{0} ({1})'.format(item['title'], item['id']))
            album_items = service.albums().get(albumId=item['id']).execute()
            print(album_items)
    # 列出當前所有你存取過的共享相簿
    shared_album = service.sharedAlbums().list().execute()
    print(shared_album)
    print(len(shared_album['sharedAlbums']))


def main():
    # Example Usage:
    service = create_session()
    access_album(service)


if __name__ == "__main__":
    main()
