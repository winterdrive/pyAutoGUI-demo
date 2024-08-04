import requests
import json
from config.base_config import UNSPLASH_ACCESS_KEY

def get_unsplash_image(photo_id: str, save_dir:str) -> None:
    """Downloads an image from Unsplash."""
    # check if file already exists
    import os
    if os.path.exists(f"{save_dir}/{photo_id}.jpg"):
        print('The image already exists.')
        return

    url = f"https://api.unsplash.com/photos/{photo_id}?client_id={UNSPLASH_ACCESS_KEY}"
    # print(url)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82 Safari/537'}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        # print(response.text)
        data = json.loads(response.text)
        # print(data)
        download_url = data['urls']['raw']
        with open(f"{save_dir}/{photo_id}.jpg", "wb") as file:
            res = requests.get(download_url, stream=True, headers=headers)
            file.write(res.content)
    else:
        print('Failed to download the image.')
        raise Exception('Failed to download the image.')

if __name__ == '__main__':
    # get_unsplash_image("3RIhxkIOBcE")
    get_unsplash_image("dAf7sFvPLQ4")
    