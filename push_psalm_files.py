import os
import requests
from dotenv import load_dotenv
from tqdm import tqdm
import fnmatch

# Load .env and get token
load_dotenv()
TOKEN = os.getenv("CHURCHTOOLS_TOKEN")

if not TOKEN:
    raise ValueError("CHURCHTOOLS_TOKEN not found in .env file")

BASE_URL = "https://feg-guemligen.church.tools/api"
DOMAIN_TYPE = "wiki_39"
DOMAIN_IDENTIFIER = "a879e45b-192e-494f-9724-7b83ac03deb3"

session = requests.Session()
session.headers.update({
    "Authorization": f"Login {TOKEN}",
    "accept": "application/json"
})


def get_files():
    url = f"{BASE_URL}/files/{DOMAIN_TYPE}/{DOMAIN_IDENTIFIER}"
    r = session.get(url)
    r.raise_for_status()
    return r.json()["data"]


def delete_file(file_id):
    url = f"{BASE_URL}/files/{file_id}"
    r = session.delete(url)
    r.raise_for_status()


def upload_file(local_path, filename):
    url = f"{BASE_URL}/files/{DOMAIN_TYPE}/{DOMAIN_IDENTIFIER}"
    with open(local_path, "rb") as f:
        files = {
            "files[]": (filename, f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        }
        data = {
            "image_options": "{}",
            "max_height": "",
            "max_width": ""
        }
        r = session.post(url, files=files, data=data)
        r.raise_for_status()


def main():
    files = get_files()
    psalm_files = [f for f in files if fnmatch.fnmatch(f["name"], "Psalm_*.pptx")]

    for file in tqdm(psalm_files, desc="Processing files"):
        local_path = os.path.join(os.getcwd(), file["name"])

        if not os.path.exists(local_path):
            print(f"⚠️ Local file not found, skipping: {file['name']}")
            continue

        try:
            delete_file(file["id"])
            upload_file(local_path, file["name"])
        except requests.HTTPError as e:
            print(f"❌ Error processing {file['name']}: {e}")


if __name__ == "__main__":
    main()
