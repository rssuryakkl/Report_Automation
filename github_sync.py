import os
import base64
import logging
import requests

logger = logging.getLogger("GitHubSync")

def get_remote_filename(file_path: str) -> str:
    """Returns standard repository filename for the Excel report."""
    base_name = os.path.basename(file_path)
    # Standardize to Daily_report.xlsx matching the repo file
    if base_name.lower() == "daily_report.xlsx":
        return "Daily_report.xlsx"
    return base_name

def pull_excel_from_github(
    file_path: str,
    repo: str,
    token: str,
    branch: str = "main"
) -> tuple[bool, str]:
    """
    Downloads the latest Daily_report.xlsx from the GitHub cloud repository
    and saves it to the local file_path.
    """
    if not repo or not token:
        return False, "GitHub repository or token not provided."

    remote_name = get_remote_filename(file_path)
    url = f"https://api.github.com/repos/{repo}/contents/{remote_name}?ref={branch}"
    headers = {
        "Authorization": f"Bearer {token.strip()}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "DailyReportAutomation"
    }

    try:
        res = requests.get(url, headers=headers, timeout=15)
        if res.status_code == 200:
            data = res.json()
            content_b64 = data.get("content", "")
            file_bytes = base64.b64decode(content_b64)
            # Ensure parent directory exists
            os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
            with open(file_path, "wb") as f:
                f.write(file_bytes)
            return True, f"Successfully loaded latest `{remote_name}` from GitHub cloud!"
        elif res.status_code == 404:
            return False, f"Remote file `{remote_name}` not found in repository {repo} (branch: {branch})."
        elif res.status_code == 401:
            return False, "Invalid GitHub token (Unauthorized). Please check your token."
        elif res.status_code == 403:
            return False, "GitHub API permission denied. Ensure token has 'contents: read' access."
        else:
            return False, f"GitHub API error {res.status_code}: {res.text}"
    except Exception as e:
        return False, f"Network error connecting to GitHub: {e}"

def push_excel_to_github(
    file_path: str,
    repo: str,
    token: str,
    branch: str = "main",
    commit_message: str = None
) -> tuple[bool, str]:
    """
    Uploads/commits local Daily_report.xlsx directly to the GitHub cloud repository.
    This guarantees 100% persistence on Streamlit Cloud even when the container sleeps.
    """
    if not repo or not token:
        return False, "GitHub repository or token not configured."

    if not os.path.exists(file_path):
        return False, f"Local file not found at: {file_path}"

    remote_name = get_remote_filename(file_path)
    url = f"https://api.github.com/repos/{repo}/contents/{remote_name}"
    headers = {
        "Authorization": f"Bearer {token.strip()}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "DailyReportAutomation"
    }

    # Step 1: Get current file sha from GitHub if it exists
    sha = None
    try:
        get_res = requests.get(f"{url}?ref={branch}", headers=headers, timeout=15)
        if get_res.status_code == 200:
            sha = get_res.json().get("sha")
        elif get_res.status_code == 401:
            return False, "Invalid GitHub token (Unauthorized). Please check your token permissions."
        elif get_res.status_code == 403:
            return False, "GitHub API permission denied. Ensure token has 'contents: write' permissions."
    except Exception as e:
        return False, f"Failed checking current remote file: {e}"

    # Step 2: Read binary content and base64 encode
    try:
        with open(file_path, "rb") as f:
            file_bytes = f.read()
        content_b64 = base64.b64encode(file_bytes).decode("utf-8")
    except Exception as e:
        return False, f"Failed to read local Excel file: {e}"

    # Step 3: Put/Commit the updated file to GitHub
    payload = {
        "message": commit_message or f"chore: update {remote_name} from Cloud Web App [skip ci]",
        "content": content_b64,
        "branch": branch
    }
    if sha:
        payload["sha"] = sha

    try:
        put_res = requests.put(url, headers=headers, json=payload, timeout=20)
        if put_res.status_code in [200, 201]:
            return True, f"Successfully committed & preserved `{remote_name}` in GitHub Cloud!"
        else:
            return False, f"GitHub Commit failed ({put_res.status_code}): {put_res.text}"
    except Exception as e:
        return False, f"Failed uploading file to GitHub: {e}"

def test_github_connection(repo: str, token: str) -> tuple[bool, str]:
    """Tests if GitHub token is valid and has access to the target repository."""
    if not repo or not token:
        return False, "Repository or token is empty."
    url = f"https://api.github.com/repos/{repo.strip()}"
    headers = {
        "Authorization": f"Bearer {token.strip()}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "DailyReportAutomation"
    }
    try:
        res = requests.get(url, headers=headers, timeout=10)
        if res.status_code == 200:
            data = res.json()
            repo_full = data.get("full_name", repo)
            return True, f"Connected to `{repo_full}`"
        elif res.status_code == 401:
            return False, "Invalid GitHub Token (401 Unauthorized)."
        elif res.status_code == 404:
            return False, f"Repository `{repo}` not found or token has no access."
        else:
            return False, f"Error {res.status_code}: {res.text}"
    except Exception as e:
        return False, f"Connection failed: {e}"
