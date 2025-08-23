import os
import sys

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


def request_with_retry(method, url, **kwargs):
    """Helper that retries once if 401 Unauthorized is returned."""
    resp = session.request(method, url, **kwargs)
    if resp.status_code == 401:
        # refresh session by forcing token header again
        session.headers.update({"Authorization": f"Login {TOKEN}"})
        resp = session.request(method, url, **kwargs)
    resp.raise_for_status()
    return resp


def get_files():
    url = f"{BASE_URL}/files/{DOMAIN_TYPE}/{DOMAIN_IDENTIFIER}"
    r = request_with_retry("GET", url)
    return r.json()["data"]


def delete_file(file_id):
    url = f"{BASE_URL}/files/{file_id}"
    request_with_retry("DELETE", url)


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
        request_with_retry("POST", url, files=files, data=data)


def main():
    try:
        files = get_files()
        print(f"✅ Successfully got list of {len(files)} files.")
        psalm_files = [f for f in files if fnmatch.fnmatch(f["name"], "Psalm_*.pptx")]
        psalm_files = sorted(psalm_files, key=lambda f: f["name"])
        print(f"  --> {len(psalm_files)} of them are Psalm PPTX files. Now traveling through them one by one.\n")
    except requests.HTTPError as e:
        print(f"❌ Error getting files, quit.")
        return

    #for file in tqdm(psalm_files, desc="Processing files"):
    for file in psalm_files:
        print(f"Handling file {file['name']}")

        local_path = os.path.join(os.getcwd(), file["name"])
        if not os.path.exists(local_path):
            print(f"  --> ⚠️ Local file not found, skipping: {file['name']}")
            continue

        try:
            delete_file(file["id"])
            print(f"✅ Deleted remote file {file['name']}")
        except requests.HTTPError as e:
            print(f"  --> ❌ Error deleting {file['name']}: {e}\n  Continuing.")
            continue
        try:
            upload_file(local_path, file["name"])
            print(f"  --> ✅ Uploaded new file {file['name']}")
        except requests.HTTPError as e:
            # Surface server message to help debugging (e.g., 401 details)
            msg = ""
            try:
                msg = f" | response: {e.response.text[:500]}"
            except Exception:
                pass
            print(f" --> ❌ Error uploading {file['name']}: {e}{msg}")


if __name__ == "__main__":
    main()
