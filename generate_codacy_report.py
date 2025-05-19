from github import Github
import pandas as pd
import os

# Get GitHub token from environment variable
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
ORG_NAME = "Anusreemv2004"

g = Github(GITHUB_TOKEN)
org = g.get_organization(ORG_NAME)

repos = org.get_repos()

data = []

for repo in repos:
    try:
        hooks = repo.get_hooks()
        codacy_connected = False
        for hook in hooks:
            if "codacy" in hook.config.get("url", "").lower():
                codacy_connected = True
                break
        data.append({
            "repo_name": repo.name,
            "full_name": repo.full_name,
            "codacy_connected": codacy_connected,
            "repo_url": repo.html_url
        })
    except Exception as e:
        print(f"Error processing repo {repo.name}: {e}")

# Create DataFrame and export to Excel
df = pd.DataFrame(data)
df.to_excel("repos_with_codacy.xlsx", index=False)

print("Excel file generated: repos_with_codacy.xlsx")
