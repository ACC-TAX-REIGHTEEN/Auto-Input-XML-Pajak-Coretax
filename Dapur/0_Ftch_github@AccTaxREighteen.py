import os
import hashlib
import requests

FILES_TO_UPDATE = {
    "Helper_160100.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_160100.txt",
    "Helper_200104.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_200104.txt",
    "Helper_340300.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_340300.txt",
    "Helper_400800.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_400800.txt",
    "Helper_401100.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_401100.txt",
    "Helper_401300.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_401300.txt",
    "Helper_830900.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_830900.txt",
    "Helper_842100.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_842100.txt",
    "Helper_A.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_A.txt",
    "Helper_B.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_B.txt",
    "Helper_Del.txt": "https://raw.githubusercontent.com/ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation/refs/heads/main/Helper_Del.txt"
}

def check_internet():
    try:
        requests.get("https://www.google.com", timeout=5)
        return True
    except:
        return False

def get_bytes_md5(data):
    return hashlib.md5(data).hexdigest()

def update_helpers():
    if not check_internet():
        return

    for local_file, url in FILES_TO_UPDATE.items():
        try:
            response = requests.get(url, timeout=10)
            if response.status_code != 200:
                continue

            remote_content = response.content
            remote_md5 = get_bytes_md5(remote_content)

            if os.path.exists(local_file):
                with open(local_file, "rb") as f:
                    local_content = f.read()
                local_md5 = get_bytes_md5(local_content)

                if local_md5 != remote_md5:
                    with open(local_file, "wb") as f:
                        f.write(remote_content)
            else:
                with open(local_file, "wb") as f:
                    f.write(remote_content)

        except:
            continue

if __name__ == "__main__":
    update_helpers()