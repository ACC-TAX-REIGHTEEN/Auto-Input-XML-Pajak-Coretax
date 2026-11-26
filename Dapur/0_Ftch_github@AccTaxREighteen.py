import os
import hashlib
import configparser
from urllib.parse import quote
import requests

REPO_OWNER = "ACC-TAX-REIGHTEEN"
REPO_NAME = "Helper-For-Tax-Automation"
BRANCH = "main"
API_BASE_URL = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents"
REQUEST_TIMEOUT = (5, 15) 

FILES_TO_UPDATE = (
    "Helper_000000.txt",
    "Helper_120101.txt",
    "Helper_160100.txt",
    "Helper_200104.txt",
    "Helper_340300.txt",
    "Helper_400800.txt",
    "Helper_401100.txt",
    "Helper_401300.txt",
    "Helper_830900.txt",
    "Helper_842100.txt",
    "Helper_A.txt",
    "Helper_B.txt",
    "Helper_Del.txt",
    "KTP-OTH.txt",
    "KTP.txt",
)


def get_git_token(config_file="config.conf"):
    if not os.path.exists(config_file):
        return ""

    config = configparser.ConfigParser()
    try:
        config.read(config_file, encoding="utf-8-sig")
    except (configparser.Error, OSError):
        return ""

    if "GIT" in config and "git_token" in config["GIT"]:
        return config["GIT"]["git_token"].strip().strip('"').strip("'")

    return ""


def get_git_blob_sha(data):
    header = f"blob {len(data)}\0".encode("utf-8")
    return hashlib.sha1(header + data).hexdigest()


def get_local_blob_sha(filepath):
    if not os.path.exists(filepath):
        return None
    try:
        with open(filepath, "rb") as f:
            return get_git_blob_sha(f.read())
    except OSError:
        return None


def write_file_atomic(filepath, data):
    tmp_path = f"{filepath}.tmp"
    try:
        with open(tmp_path, "wb") as f:
            f.write(data)
        os.replace(tmp_path, filepath)
        return True
    except OSError:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        return False


def is_rate_limited(response):
    return (
        response.status_code == 429
        or (
            response.status_code == 403
            and (
                response.headers.get("X-RateLimit-Remaining") == "0"
                or "Retry-After" in response.headers
            )
        )
    )


def update_helpers():
    print("--> sedang menghubungkan")

    github_token = get_git_token()
    if not github_token:
        print("--> Dilewati: git_token tidak ditemukan di config.conf")
        return

    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {github_token}",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "AccTaxREighteen-HelperSync/1.0",
    })

    try:
        try:
            meta_resp = session.get(
                API_BASE_URL,
                params={"ref": BRANCH},
                headers={"Accept": "application/vnd.github+json"},
                timeout=REQUEST_TIMEOUT,
            )
        except (requests.ConnectionError, requests.Timeout):
            print("--> Mode offline / GitHub API tidak terjangkau, sinkronisasi dilewati")
            return
        except requests.RequestException:
            return

        if meta_resp.status_code == 401:
            print("--> Gagal autentikasi (401): git_token tidak valid atau kedaluwarsa")
            return
        if is_rate_limited(meta_resp):
            print("--> Batas kuota GitHub API tercapai, sinkronisasi dilewati")
            return
        if meta_resp.status_code in (403, 404):
            print(f"--> Akses repositori ditolak/tidak ditemukan ({meta_resp.status_code}), periksa izin token")
            return
        if meta_resp.status_code != 200:
            print(f"--> GitHub API mengembalikan status {meta_resp.status_code}, sinkronisasi dilewati")
            return

        try:
            repo_items = meta_resp.json()
        except ValueError:
            print("--> Respons bukan JSON valid (kemungkinan terhalang proxy/captive portal)")
            return

        if not isinstance(repo_items, list):
            return

        whitelist = set(FILES_TO_UPDATE)
        remote_files = {
            item["name"]: item.get("sha", "")
            for item in repo_items
            if isinstance(item, dict)
            and item.get("type") == "file"
            and item.get("name") in whitelist
        }

        print("--> Sedang memperbarui data")
        updated_count = 0

        for local_file in FILES_TO_UPDATE:
            remote_sha = remote_files.get(local_file)
            if not remote_sha:
                continue

            if get_local_blob_sha(local_file) == remote_sha:
                continue

            file_url = f"{API_BASE_URL}/{quote(local_file)}"
            try:
                file_resp = session.get(
                    file_url,
                    params={"ref": BRANCH},
                    headers={"Accept": "application/vnd.github.raw+json"},
                    timeout=REQUEST_TIMEOUT,
                )
            except requests.RequestException:
                continue

            if is_rate_limited(file_resp):
                print("--> Kuota API habis di tengah proses, sisa file dilewati")
                break
            if file_resp.status_code != 200:
                continue

            remote_content = file_resp.content

            if get_git_blob_sha(remote_content) != remote_sha:
                continue

            if write_file_atomic(local_file, remote_content):
                updated_count += 1

        if updated_count > 0:
            print(f"--> Data telah diperbarui ({updated_count} file diperbarui)")
        else:
            print("--> Data sudah versi terbaru")
    finally:
        session.close()


if __name__ == "__main__":
    update_helpers()