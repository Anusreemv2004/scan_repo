import os
import requests
import json
import tempfile
import subprocess
from openpyxl import Workbook
import shutil

# Step 1: Environment Variables
username = os.getenv('GITHUB_USERNAME')
token = os.getenv('GITHUB_TOKEN')

if not username or not token:
    raise Exception("Missing GITHUB_USERNAME or GITHUB_TOKEN")

# Step 2: API Setup
headers = {
    "Authorization": f"token {token}",
    "Accept": "application/vnd.github+json"
}

# Step 3: Fetch All Repos
print("📡 Fetching repositories...")
page = 1
all_repos = []

while True:
    url = f"https://api.github.com/user/repos?per_page=100&page={page}&visibility=all"
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"🚨 API Error: {response.status_code}")
        break

    repos = response.json()
    if not repos:
        break

    all_repos.extend(repos)
    page += 1

print(f"🧮 Total repositories fetched: {len(all_repos)}")

# Step 4: Collect metadata from custom.json files
repo_metadata_list = []
all_keys = set()

for repo in all_repos:
    clone_url = repo['clone_url']
    temp_dir = tempfile.mkdtemp()
    metadata = {}

    try:
        print(f"🔍 Cloning {repo['name']}...")
        subprocess.run(["git", "clone", "--depth=1", clone_url, temp_dir], check=True, stdout=subprocess.DEVNULL)

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

    except Exception as e:
        print(f"⚠️ Error reading {repo['name']}: {e}")
    finally:
        shutil.rmtree(temp_dir)

# Step 5: Write to Excel
wb = Workbook()
ws = wb.active
ws.title = "Repo Metadata"

# Ensure consistent column order
sorted_keys = sorted(all_keys)
headers = ["Name", "URL"] + sorted_keys
ws.append(headers)

for entry in repo_metadata_list:
    row = [
        entry.get("name", ""),
        entry.get("url", "")
    ] + [entry.get(key, "") for key in sorted_keys]
    ws.append(row)

filename = f"{username}_custom_metadata.xlsx"
wb.save(filename)
print(f"📁 Excel file saved as: {filename}")
