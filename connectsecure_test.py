#!/usr/bin/env python3
"""
ConnectSecure V4 API Test Script
"""

import requests
import json
import base64
import sys


# ──────────────────────────────────────────────
# CONFIGURATION - Fill these in
# ──────────────────────────────────────────────
TENANT_NAME   = "river-run"       # e.g. "riverrun"
HOSTNAME      = "pod104.myconnectsecure.com"  # Found in Profile > API Documentation
CLIENT_ID     = "c05d6ca5-8275-4857-bd93-42a193e7905f"
CLIENT_SECRET = "Z0FBQUFBQnBuSG81Nm1LS0h0YkRMWTY0YmlSSkR6Rml3SFhWZzRWTkpHdFVqQWxBaGhRX185NEZ5QXh3ZWNETHNOVGNGQlB5SmxQX28wRjNjWHVuM21mVFZ5ck5uNlVISHc3YWR6cGVRYng4SlJrVnp0V2drMUZDcXo4dGdyRExRdTZOZ2lsNXFfT2Q1Vm1XZmEzVU94d3Y5dFp6UHF4VUJTWGJGQzJyclhzOGNHaHFuTkdBQ3FZPQ=="
# ──────────────────────────────────────────────

BASE_URL = f"https://{HOSTNAME}"

session = requests.Session()
session.headers.update({
    "Content-Type": "application/json",
    "accept": "application/json"
})


def get_token():
    """Authenticate using base64 encoded Client-Auth-Token header."""
    print("[*] Authenticating...")

    raw = f"{TENANT_NAME}+{CLIENT_ID}:{CLIENT_SECRET}"
    encoded = base64.b64encode(raw.encode()).decode()

    headers = {
        "Client-Auth-Token": encoded,
        "Content-Type": "application/json",
        "accept": "application/json"
    }

    resp = requests.post(f"{BASE_URL}/w/authorize", headers=headers)
    resp.raise_for_status()
    data = resp.json()

    if not data.get("status"):
        print(f"[-] Auth failed: {json.dumps(data, indent=2)}")
        sys.exit(1)

    token = data["data"]["access_token"]
    user_id = data["data"]["user_id"]

    session.headers.update({
        "Authorization": f"Bearer {token}",
        "X-USER-ID": user_id
    })

    print(f"[+] Authentication successful!")
    print(f"    User ID: {user_id}\n")
    return token, user_id


def test_get_companies():
    """Retrieve list of companies."""
    print("[*] Fetching companies...")
    resp = session.get(f"{BASE_URL}/r/company/companies")
    resp.raise_for_status()
    data = resp.json()

    companies = data.get("data", [])
    total = data.get("total", len(companies))
    print(f"[+] Companies found: {total}")
    for c in companies[:5]:
        print(f"    - {c.get('name', 'N/A')} (ID: {c.get('id', 'N/A')}) | {c.get('description', '')}")
    if total > 5:
        print(f"    ... and {total - 5} more")
    print()
    return companies


def test_get_agents():
    """Retrieve agents."""
    print("[*] Fetching agents...")
    resp = session.get(f"{BASE_URL}/r/company/agents")
    resp.raise_for_status()
    data = resp.json()

    agents = data.get("data", [])
    total = data.get("total", len(agents))
    print(f"[+] Agents found: {total}")
    if agents:
        print(f"    First agent: {agents[0].get('hostname', 'N/A')}")
    print()


def test_get_assets():
    """Retrieve assets."""
    print("[*] Fetching assets...")
    resp = session.get(f"{BASE_URL}/r/asset/assets")
    resp.raise_for_status()
    data = resp.json()

    assets = data.get("data", [])
    total = data.get("total", len(assets))
    print(f"[+] Assets found: {total}")
    if assets:
        print(f"    First asset: {assets[0].get('hostname', 'N/A')}")
    print()


def test_get_company_stats():
    """Retrieve company stats (list)."""
    print("[*] Fetching company stats...")
    resp = session.get(f"{BASE_URL}/r/company/company_stats")
    resp.raise_for_status()
    data = resp.json()

    stats = data.get("data", [])
    total = data.get("total", len(stats))
    print(f"[+] Company stats records: {total}")
    print()
    return stats


def test_get_company_stats_by_id(company_id=737):
    """Retrieve company stat for a single company (GET /r/company/company_stats/{id})."""
    print(f"[*] Fetching company stats for company ID {company_id}...")
    resp = session.get(f"{BASE_URL}/r/company/company_stats/{company_id}")
    resp.raise_for_status()
    data = resp.json()

    if not data.get("status"):
        print(f"[-] API returned status=false: {data}")
        return None
    stats = data.get("data", {})
    print(f"[+] Company stats for ID {company_id}: company_id={stats.get('company_id')}, total_assets={stats.get('total_assets')}, date={stats.get('date')}")
    print()
    return stats


def test_get_users():
    """Retrieve users."""
    print("[*] Fetching users...")
    resp = session.get(f"{BASE_URL}/r/user/get_users")
    resp.raise_for_status()
    data = resp.json()

    users = data.get("data", [])
    total = data.get("total", len(users))
    print(f"[+] Users found: {total}")
    print()


def test_get_standard_reports(is_global=False, skip=None, limit=None):
    """List standard reports. GET /report_builder/standard_reports"""
    print("[*] Fetching standard reports...")
    params = {"isGlobal": is_global}
    if skip is not None:
        params["skip"] = skip
    if limit is not None:
        params["limit"] = limit

    resp = session.get(f"{BASE_URL}/report_builder/standard_reports", params=params)
    resp.raise_for_status()
    data = resp.json()

    if not data.get("status"):
        print(f"[-] API returned status=false: {data}")
        return None

    sections = data.get("message", [])
    total_reports = sum(len(s.get("Reports", [])) for s in sections)
    print(f"[+] Standard reports: {len(sections)} sections, {total_reports} report types")
    for sec in sections[:5]:
        section_name = sec.get("Section", "N/A")
        reports = sec.get("Reports", [])
        print(f"    - {section_name}: {len(reports)} reports")
    if len(sections) > 5:
        print(f"    ... and {len(sections) - 5} more sections")
    print()
    return sections


def test_get_report_jobs_view(condition="*", skip=0, limit=100, order_by=None):
    """Retrieve report jobs view. GET /r/company/report_jobs_view"""
    print("[*] Fetching report jobs view...")
    params = {"condition": condition, "skip": skip, "limit": limit}
    if order_by is not None:
        params["order_by"] = order_by

    resp = session.get(f"{BASE_URL}/r/company/report_jobs_view", params=params)
    resp.raise_for_status()
    data = resp.json()

    if not data.get("status"):
        print(f"[-] API returned status=false: {data}")
        return None

    jobs = data.get("data", [])
    total = data.get("total", len(jobs))
    print(f"[+] Report jobs: {len(jobs)} returned (total: {total})")
    for job in jobs[:3]:
        print(f"    - {job.get('job_id', 'N/A')}: {job.get('type', 'N/A')} | {job.get('status', 'N/A')} | {job.get('company_name', 'N/A')}")
    if len(jobs) > 3:
        print(f"    ... and {len(jobs) - 3} more")
    print()
    return jobs


def create_report_job(
    company_id,
    report_id,
    report_name,
    report_type,
    file_type=None,
    company_name="",
    is_filter=False,
    report_filter=None,
):
    """Create a report job. POST /report_builder/create_report_job

    Args:
        company_id: Company ID (from companies list)
        report_id: Report ID (from standard_reports, e.g. "34d96368304641a1b9d0b9a7cfaaf170")
        report_name: Display name for the report
        report_type: pdf, docx, xlsx, pptx
        file_type: Output format (defaults to report_type)
        company_name: Optional company name
        is_filter: Whether reportFilter is applied
        report_filter: Filter object when is_filter=True

    Returns:
        dict with status and message (contains job_id on success)
    """
    payload = {
        "company_id": company_id,
        "company_name": company_name,
        "reportId": report_id,
        "reportName": report_name,
        "reportType": report_type,
        "fileType": file_type or report_type,
        "isFilter": is_filter,
        "reportFilter": report_filter or {},
    }
    resp = session.post(f"{BASE_URL}/report_builder/create_report_job", json=payload)
    resp.raise_for_status()
    return resp.json()


def main():
    print("=" * 50)
    print("  ConnectSecure V4 API Test Script")
    print("=" * 50 + "\n")

    try:
        get_token()
        test_get_companies()
        test_get_agents()
        test_get_assets()
        test_get_company_stats()
        test_get_company_stats_by_id(737)  # single company: GET /r/company/company_stats/{id}
        test_get_users()
        test_get_standard_reports(is_global=False)  # GET /report_builder/standard_reports
        test_get_report_jobs_view()  # GET /r/company/report_jobs_view
        print("[+] All tests completed successfully.")

    except requests.exceptions.HTTPError as e:
        print(f"[-] HTTP Error: {e}")
        print(f"    Response: {e.response.text}")
        sys.exit(1)
    except requests.exceptions.ConnectionError as e:
        print(f"[-] Connection Error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[-] Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()