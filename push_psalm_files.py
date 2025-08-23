import os
import fnmatch
import requests
from dotenv import load_dotenv

# Load .env and get token
load_dotenv()
USER, PASSWORD = os.getenv("CHURCHTOOLS_USER"), os.getenv("CHURCHTOOLS_PASSWORD")

BASE_URL = "https://feg-guemligen.church.tools/api"
DOMAIN_TYPE = "wiki_39"
DOMAIN_IDENTIFIER = "a879e45b-192e-494f-9724-7b83ac03deb3"

session = requests.Session()
session.headers.update({
    "accept": "application/json"
})

CSRF_TOKEN = None  # will be set after login


def login_with_user_pw():
    global CSRF_TOKEN
    url = f"{BASE_URL}/login"
    r = session.post(url, json={"username": USER, "password": PASSWORD})
    r.raise_for_status()
    # data = r.json()
    # CSRF_TOKEN = data.get("csrfToken")
    # if not CSRF_TOKEN:
    #     raise RuntimeError("No CSRF token returned from login. Response: %s" % data)
    # print("✅ Logged in and obtained CSRF token.")


def request_with_retry(method, url, **kwargs):
    # inject CSRF header for modifying requests
    if method.upper() in ("POST", "DELETE", "PUT", "PATCH") and CSRF_TOKEN:
        headers = kwargs.pop("headers", {})
        headers["X-CSRF-Token"] = CSRF_TOKEN
        kwargs["headers"] = headers

    resp = session.request(method, url, timeout=60, **kwargs)
    if resp.status_code == 401:
        # try re-login once
        login_with_user_pw()
        if method.upper() in ("POST", "DELETE", "PUT", "PATCH") and CSRF_TOKEN:
            headers = kwargs.pop("headers", {})
            headers["X-CSRF-Token"] = CSRF_TOKEN
            kwargs["headers"] = headers
        resp = session.request(method, url, timeout=60, **kwargs)
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
        login_with_user_pw()
        files = get_files()
        print(f"✅ Successfully got list of {len(files)} files.")
        psalm_files = [f for f in files if fnmatch.fnmatch(f["name"], "Psalm_*.pptx")]
        psalm_files = sorted(psalm_files, key=lambda f: f["name"])
        print(f"  --> {len(psalm_files)} Psalm PPTX files found. Processing...\n")
    except requests.HTTPError as e:
        print(f"❌ Error during login or file list: {e}")
        return

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
            print(f"  --> ❌ Error uploading {file['name']}: {e}{msg}")


if __name__ == "__main__":
    main()
