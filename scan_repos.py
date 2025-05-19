import os
import requests
import json
import tempfile
import subprocess
from openpyxl import Workbook
import shutil

# === CONFIG ===
org = "Anusreemv2004"
username = os.getenv('GITHUB_USERNAME')
token = os.getenv('GITHUB_TOKEN')

if not username or not token:
    raise Exception("Missing GITHUB_USERNAME or GITHUB_TOKEN environment variables")

headers = {
    "Authorization": f"token {token}",
    "Accept": "application/vnd.github+json"
}

# === Step 1: Fetch all repos in the org ===
print(f"📡 Fetching repositories for organization '{org}'...")
page = 1
all_repos = []

while True:
    url = f"https://api.github.com/orgs/{org}/repos?per_page=100&page={page}&type=all"
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"🚨 API Error {response.status_code}: {response.text}")
        break

    repos = response.json()
    if not repos:
        break

    all_repos.extend(repos)
    print(f"📥 Fetched {len(repos)} repos from page {page}")
    page += 1

print(f"🧮 Total repos fetched: {len(all_repos)}")

# === Step 2: Clone repos & extract custom.json metadata ===
repo_metadata_list = []
all_keys = set()

for repo in all_repos:
    print(f"🔍 Processing repo: {repo['full_name']}")
    clone_url = repo['clone_url']
    temp_dir = tempfile.mkdtemp()
    metadata = {}

    try:
        subprocess.run(
            ["git", "clone", "--depth=1", clone_url, temp_dir],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )

        custom_path = os.path.join(temp_dir, ".github", "custom.json")
        if os.path.exists(custom_path):
            with open(custom_path, "r") as f:
                metadata = json.load(f)
                all_keys.update(metadata.keys())

            repo_metadata_list.append({
                "name": repo["name"],
                "url": repo["html_url"],
                **metadata
            })
        else:
            print("   ⚠️ No .github/custom.json found")
    except Exception as e:
        print(f"   ⚠️ Failed to process repo '{repo['full_name']}': {e}")
    finally:
        shutil.rmtree(temp_dir)

# === Step 3: Write to Excel ===
wb = Workbook()
ws = wb.active
ws.title = "Repo Metadata"

sorted_keys = sorted(all_keys)
headers = ["Name", "URL"] + sorted_keys
ws.append(headers)

for entry in repo_metadata_list:
    row = [entry.get("name", ""), entry.get("url", "")]
    row.extend(entry.get(key, "") for key in sorted_keys)
    ws.append(row)

filename = f"{username}_org_repos_custom_metadata.xlsx"
wb.save(filename)
print(f"📁 Excel file saved as: {filename}")
