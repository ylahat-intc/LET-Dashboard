"""
LET Dashboard — Data Refresh Script
Pulls all records from the SharePoint list and saves a local cache.
Run this script whenever you want fresh data, then run generate_dashboard.py.

HOW TO USE:
  1. Run: py refresh_data.py        ← pulls latest from SharePoint
  2. Run: py generate_dashboard.py  ← generates HTML dashboard
  3. Open: LET_Dashboard.html       ← view in browser

Or run both together:
  py refresh_data.py && py generate_dashboard.py

NOTE: This script uses your existing M365 credentials (same as Copilot CLI).
      You must have access to the GLDC SharePoint site.
"""

import json
import os
import sys
import subprocess
import urllib.request
import urllib.error

# ── CONFIG ──────────────────────────────────────────────────────────────────
SITE_ID   = "intel.sharepoint.com,8c7b402a-80d5-4df7-9355-5ce084b0e63d,a9bcafb0-ecf2-4d91-87f0-f6b11664510c"
LIST_ID   = "781b2a30-b52a-4eda-b558-3f79056885a9"
OUTPUT    = os.path.join(os.path.dirname(__file__), "let_data_cache.json")
HTML_FILE = os.path.join(os.path.dirname(__file__), "LET_Dashboard.html")
# SharePoint path: Shared Documents/LET Dashboard/LET_Dashboard.html
SP_UPLOAD_PATH = "LET Dashboard/LET_Dashboard.html"

# Fields to fetch (maps internal names → readable)
FIELD_MAP = {
    "Title":    "RequestID",
    "field_2":  "DateRequested",
    "field_3":  "Status",
    "field_5":  "RequestType",
    "field_7":  "BU",
    "field_8":  "Group",
    "field_9":  "Permanent",
    "field_10": "CPA",
    "field_12": "TargetMoveIn",
    "field_13": "SqFtRequested",
    "field_14": "Benches",
    "field_15": "Racks",
    "field_16": "PowerKW",
    "field_17": "Touch",
    "field_18": "People",
    "field_19": "Platforms",
    "field_20": "SiteTo",
    "field_21": "SiteFrom",
    "field_26": "Owner",
    "field_27": "Title",
    "field_28": "StartDate",
    "field_29": "CloseDate",
    "field_30": "Workstations",
    "field_31": "SqFtAssigned",
    "field_32": "RoomAssigned",
    # Net Zero / Growth classification (added 2025-12, populated for newer requests only)
    "LETClassification": "LETClassification",
}
# ────────────────────────────────────────────────────────────────────────────

def get_token():
    """
    Get Microsoft Graph API token using Azure CLI or device code flow.
    Tries 'az account get-access-token' first (fastest if Azure CLI is installed).
    Falls back to MSAL device code flow.
    """
    # Try Azure CLI
    try:
        result = subprocess.run(
            ["az", "account", "get-access-token",
             "--resource", "https://graph.microsoft.com"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            token_data = json.loads(result.stdout)
            return token_data["accessToken"]
    except Exception:
        pass

    # Try MSAL with persistent token cache (login once, silent after)
    try:
        import msal
        CACHE_FILE = os.path.join(os.path.dirname(__file__), ".msal_token_cache.bin")
        cache = msal.SerializableTokenCache()
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, "r") as fh:
                cache.deserialize(fh.read())

        app = msal.PublicClientApplication(
            "14d82eec-204b-4c2f-b7e8-296a70dab67e",  # Microsoft Office client ID
            authority="https://login.microsoftonline.com/common",
            token_cache=cache
        )
        scopes = ["https://graph.microsoft.com/Sites.ReadWrite.All"]

        # Try silent auth from cache first
        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])

        # If silent fails → device flow (one-time browser login)
        if not result or "access_token" not in result:
            flow = app.initiate_device_flow(scopes=scopes)
            print(f"\n🔐 First-time login required (after this it's automatic).")
            print(f"   {flow['message']}\n")
            result = app.acquire_token_by_device_flow(flow)

        # Save updated cache
        if cache.has_state_changed:
            with open(CACHE_FILE, "w") as fh:
                fh.write(cache.serialize())

        if "access_token" in result:
            return result["access_token"]
        raise Exception(result.get("error_description", "Auth failed"))

    except ImportError:
        print("⚠️  MSAL not installed. Run: py -m pip install msal")
        sys.exit(1)

def fetch_all_items(token):
    """Fetch all list items via Graph API with paging."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    fields = ",".join(FIELD_MAP.keys())
    url = (f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}"
           f"/items?$expand=fields($select={fields})&$top=500")

    all_items = []
    page = 0
    while url:
        page += 1
        print(f"   Fetching page {page}... ", end="", flush=True)
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode())
        batch = data.get("value", [])
        all_items.extend(batch)
        print(f"{len(batch)} items (total: {len(all_items)})")
        url = data.get("@odata.nextLink")

    return all_items

def normalize_items(raw_items):
    """Extract and rename fields for the dashboard."""
    result = []
    for item in raw_items:
        f  = item.get("fields", {})
        created = item.get("createdDateTime", "")
        modified = item.get("lastModifiedDateTime", "")
        row = {
            "id":       item.get("id"),
            "webUrl":   item.get("webUrl"),
            "created":  created,
            "modified": modified,
            "createdBy":   item.get("createdBy", {}).get("displayName", ""),
            "modifiedBy":  item.get("lastModifiedBy", {}).get("displayName", ""),
            "fields":   f,   # keep raw fields for dashboard generator
        }
        result.append(row)
    return result

def upload_html_to_sharepoint(token):
    """Upload LET_Dashboard.html to SharePoint document library."""
    import urllib.parse
    if not os.path.exists(HTML_FILE):
        print("⚠️  HTML file not found — run generate_dashboard.py first")
        return False

    with open(HTML_FILE, "rb") as fh:
        content = fh.read()

    # URL-encode the path (spaces → %20)
    encoded_path = urllib.parse.quote(SP_UPLOAD_PATH)
    url = (f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}"
           f"/drive/root:/{encoded_path}:/content")
    req = urllib.request.Request(
        url,
        data=content,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "text/html",
        },
        method="PUT"
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode())
        web_url = result.get("webUrl", "")
        print(f"✅ Uploaded to SharePoint → {web_url}")
        return True
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"⚠️  Upload failed: {e.code} — {body[:200]}")
        return False


def upload_html_to_azure(connection_string=None):
    """Upload LET_Dashboard.html to Azure Blob Storage static website ($web container)."""
    if not os.path.exists(HTML_FILE):
        print("⚠️  HTML file not found — run generate_dashboard.py first")
        return False

    # Load connection string from config file if not provided
    cfg_file = os.path.join(os.path.dirname(__file__), "azure_config.txt")
    if not connection_string:
        if os.path.exists(cfg_file):
            with open(cfg_file) as f:
                for line in f:
                    if line.startswith("AZURE_CONNECTION_STRING="):
                        connection_string = line.split("=", 1)[1].strip()
                        break
    if not connection_string:
        print("⚠️  Azure connection string not configured.")
        print(f"   Add it to: {cfg_file}")
        print("   Line format:  AZURE_CONNECTION_STRING=DefaultEndpointsProtocol=https;...")
        return False

    try:
        from azure.storage.blob import BlobServiceClient, ContentSettings
        client = BlobServiceClient.from_connection_string(connection_string)
        blob = client.get_blob_client(container="$web", blob="LET_Dashboard.html")
        with open(HTML_FILE, "rb") as fh:
            blob.upload_blob(
                fh,
                overwrite=True,
                content_settings=ContentSettings(content_type="text/html; charset=utf-8")
            )
        # Derive the static website URL from connection string
        account = client.account_name
        print(f"✅ Uploaded to Azure → https://{account}.z13.web.core.windows.net/LET_Dashboard.html")
        print(f"   (Primary endpoint shown in Azure portal under Storage → Static website)")
        return True
    except Exception as e:
        print(f"⚠️  Azure upload failed: {e}")
        return False


def main():
    print("🔄 LET Dashboard — Data Refresh")
    print(f"   Target: SharePoint LET Lab Space Request/Release list")
    print(f"   Output: {OUTPUT}")
    print()

    print("🔐 Getting authentication token...")
    token = get_token()
    print("✅ Authenticated\n")

    print("📥 Fetching list items...")
    raw = fetch_all_items(token)
    print(f"\n✅ Fetched {len(raw)} total items")

    items = normalize_items(raw)

    # Save cache
    cache_data = {
        "refreshed_at": __import__("datetime").datetime.now().isoformat(),
        "count": len(items),
        "items": items
    }
    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, ensure_ascii=False, indent=2)

    print(f"✅ Cache saved → {OUTPUT}")
    print(f"\nNext: run 'py generate_dashboard.py' to build the HTML dashboard.")

def main_with_upload():
    """Full pipeline: refresh data → generate HTML → upload to SharePoint."""
    import subprocess
    main()
    print("\n🔨 Generating dashboard HTML...")
    result = subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), "generate_dashboard.py")],
                            capture_output=True, text=True)
    print(result.stdout.strip())
    if result.returncode != 0:
        print(f"⚠️  generate_dashboard.py failed: {result.stderr[:300]}")
        return

    print("\n☁️  Uploading to SharePoint...")
    token = get_token()
    upload_html_to_sharepoint(token)

    print("\n☁️  Uploading to Azure Static Website...")
    upload_html_to_azure()

    print("\n🌐 Copying to Startweb...")
    startweb_path = r"\\ger.corp.intel.com\ec\proj\ha\cs\360_Lab_EMEA\LET_Dashboard.html"
    try:
        import shutil
        shutil.copy2(HTML_FILE, startweb_path)
        print(f"✅ Copied to Startweb → http://startweb.intel.com/ec/proj/ha/cs/360_Lab_EMEA/LET_Dashboard.html")
    except Exception as e:
        print(f"⚠️  Startweb copy failed: {e}")

if __name__ == "__main__":
    if "--full" in sys.argv:
        main_with_upload()
    else:
        main()
