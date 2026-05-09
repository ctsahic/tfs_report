import argparse
import getpass
import json
import os
import re
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from tkcalendar import DateEntry
import pandas as pd
import requests

PAT_FILE = Path(".azure_devops_pat")
REPO_FILE = Path(".azure_devops_repos")
API_VERSION = "7.1-preview.1"


def parse_iso_datetime(value: str) -> datetime:
    formats = [
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(
        "Invalid date format. Use YYYY-MM-DD or YYYY-MM-DDTHH:MM or YYYY-MM-DD HH:MM"
    )


def load_pat() -> str:
    if PAT_FILE.exists():
        token = PAT_FILE.read_text().strip()
        if token:
            return token

    token = os.getenv("AZURE_DEVOPS_PAT")
    if token:
        return token.strip()

    token = getpass.getpass("Azure DevOps Personal Access Token: ").strip()
    if not token:
        raise SystemExit("A Personal Access Token is required.")

    save = input("Save PAT to local file '.azure_devops_pat' for reuse? (y/N): ").strip().lower()
    if save == "y":
        PAT_FILE.write_text(token + "\n")
        try:
            PAT_FILE.chmod(0o600)
        except OSError:
            pass
    return token


def parse_azure_devops_url(url: str) -> tuple[str, str]:
    normalized = url.rstrip("/")
    if normalized.startswith("https://"):
        normalized = normalized[len("https://") :]
    parts = normalized.split("/")
    if len(parts) < 3:
        raise ValueError("Azure DevOps URL must include organization and project.")
    if parts[0].lower() != "dev.azure.com":
        raise ValueError("URL must start with https://dev.azure.com.")
    return parts[1], parts[2]


def request_json(url: str, token: str, params: dict | None = None) -> dict:
    headers = {
        "Accept": "application/json",
    }
    response = requests.get(url, headers=headers, params=params, auth=("", token))
    response.raise_for_status()
    return response.json()


def request_text(url: str, token: str, params: dict | None = None) -> str:
    headers = {
        "Accept": "text/html,application/xhtml+xml",
    }
    response = requests.get(url, headers=headers, params=params, auth=("", token))
    response.raise_for_status()
    return response.text


def extract_work_item_title(payload: dict) -> str:
    title_candidates = [
        payload.get("title", ""),
        payload.get("name", ""),
        payload.get("resourceName", ""),
        payload.get("workItemTitle", ""),
    ]
    fields = payload.get("fields", {}) if isinstance(payload.get("fields", {}), dict) else {}
    title_candidates.extend(
        [
            fields.get("System.Title", ""),
            fields.get("System.Title ", ""),
            fields.get("title", ""),
        ]
    )
    for candidate in title_candidates:
        if candidate:
            return str(candidate)
    raw_text = str(payload)
    match = re.search(r"System\.Title['\"]?\s*[:=]\s*['\"]([^'\"]+)['\"]", raw_text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""


def get_repositories(organization: str, project: str, token: str) -> list[dict]:
    url = f"https://dev.azure.com/{organization}/{project}/_apis/git/repositories"
    return request_json(url, token, params={"api-version": API_VERSION}).get("value", [])


def get_pull_requests(organization: str, project: str, repo_id: str, token: str) -> list[dict]:
    url = (
        f"https://dev.azure.com/{organization}/{project}/_apis/git/repositories/{repo_id}/pullrequests"
    )
    return request_json(url, token, params={"api-version": API_VERSION, "searchCriteria.status": "all"}).get(
        "value", []
    )


def get_work_items_for_pr(organization: str, project: str, repo_id: str, pr_id: int, token: str) -> dict[int, str]:
    url = (
        f"https://dev.azure.com/{organization}/{project}/_apis/git/repositories/{repo_id}/pullRequests/{pr_id}/workitems"
    )
    work_items: dict[int, str] = {}
    for item in request_json(url, token, params={"api-version": API_VERSION}).get("value", []):
        work_item_id = item.get("id")
        if work_item_id is None:
            continue
        work_items[int(work_item_id)] = extract_work_item_title(item)
    return work_items


def get_work_item_details(organization: str, project: str, ids: list[int], token: str) -> dict[int, str]:
    if not ids:
        return {}

    result = {}
    for wid in ids:
        title = ""
        for item_url in [
            f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{wid}",
            f"https://dev.azure.com/{organization}/_apis/wit/workitems/{wid}",
        ]:
            try:
                response = request_json(
                    item_url,
                    token,
                    params={"api-version": API_VERSION},
                )
                title = extract_work_item_title(response)
                if title:
                    break
            except requests.HTTPError:
                continue

        if not title:
            for edit_url in [
                f"https://dev.azure.com/{organization}/{project}/_workitems/edit/{wid}/",
                f"https://dev.azure.com/{organization}/_workitems/edit/{wid}/",
            ]:
                try:
                    html = request_text(edit_url, token)
                    patterns = [
                        r"<title>(.*?)</title>",
                        r"meta[^>]+name=['\"]title['\"][^>]+content=['\"]([^'\"]+)['\"]",
                        r"meta[^>]+property=['\"]og:title['\"][^>]+content=['\"]([^'\"]+)['\"]",
                        r"data-title=['\"]([^'\"]+)['\"]",
                    ]
                    for pattern in patterns:
                        match = re.search(pattern, html, re.IGNORECASE | re.DOTALL)
                        if match:
                            page_title = match.group(1).strip()
                            page_title = re.sub(r"\s*-\s*Azure DevOps\s*$", "", page_title)
                            page_title = re.sub(r"^\s*#?\s*\d+\s*[-:–]\s*", "", page_title)
                            if page_title:
                                title = page_title
                                break
                    if title:
                        break
                except requests.HTTPError:
                    continue

        result[wid] = title
    return result


def build_report_rows(
    organization: str,
    project: str,
    repo_names: list[str],
    repositories: list[dict],
    from_dt: datetime,
    to_dt: datetime,
    token: str,
) -> list[dict]:
    repo_lookup = {repo["name"].lower(): repo for repo in repositories}
    selected_repos = []
    if repo_names:
        for name in repo_names:
            key = name.lower()
            if key not in repo_lookup:
                raise SystemExit(f"Repository not found: {name}")
            selected_repos.append(repo_lookup[key])
    else:
        selected_repos = list(repositories)

    all_rows = []
    work_item_cache: dict[int, str] = {}

    for repo in selected_repos:
        pr_list = get_pull_requests(organization, project, repo["id"], token)
        for pr in pr_list:
            created_date = datetime.fromisoformat(pr["creationDate"].replace("Z", "+00:00")).replace(tzinfo=None)
            if created_date < from_dt or created_date > to_dt:
                continue

            work_item_refs = get_work_items_for_pr(organization, project, repo["id"], pr["pullRequestId"], token)
            work_item_ids = list(work_item_refs.keys())
            for wid, title in work_item_refs.items():
                if title:
                    work_item_cache[wid] = title

            missing_ids = [wid for wid, title in work_item_refs.items() if not title and wid not in work_item_cache]
            if missing_ids:
                try:
                    work_item_cache.update(get_work_item_details(organization, project, missing_ids, token))
                except requests.HTTPError as exc:
                    print(f"Warning: failed to resolve work item details for IDs {missing_ids}: {exc}")

            work_item_titles = [work_item_cache.get(wid, "") for wid in work_item_ids]
            row = {
                "repository": repo["name"],
                "pr_id": pr["pullRequestId"],
                "title": pr.get("title", ""),
                "created_by": pr.get("createdBy", {}).get("displayName", ""),
                "creation_date": created_date.isoformat(sep=" ", timespec="seconds"),
                "status": pr.get("status", ""),
                "merge_status": pr.get("mergeStatus", ""),
                "source_ref": pr.get("sourceRefName", ""),
                "target_ref": pr.get("targetRefName", ""),
                "work_item_ids": ", ".join(str(x) for x in work_item_ids),
                "work_item_titles": "; ".join(work_item_titles),
                "description": pr.get("description", ""),
                "last_updated": pr.get("creationDate", ""),
            }
            all_rows.append(row)

    return all_rows


def save_outputs(rows: list[dict], output_prefix: str, formats: list[str]) -> list[str]:
    df = pd.DataFrame(rows)
    if df.empty:
        print("No pull requests found in the requested range.")
        return []

    saved_files = []
    if "excel" in formats:
        excel_path = f"{output_prefix}.xlsx"
        df.to_excel(excel_path, index=False, sheet_name="PR Report")
        workbook = load_workbook(excel_path)
        worksheet = workbook.active
        header_font = Font(bold=True)
        for col_idx, column_title in enumerate(df.columns, start=1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.font = header_font
            column_width = max(
                len(str(column_title)),
                max((len(str(worksheet.cell(row=row_idx, column=col_idx).value or "")) for row_idx in range(2, worksheet.max_row + 1)), default=0),
            ) + 2
            worksheet.column_dimensions[get_column_letter(col_idx)].width = min(column_width, 50)
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.freeze_panes = "A2"
        workbook.save(excel_path)
        saved_files.append(excel_path)
        print(f"Excel report written to: {excel_path}")
    if "html" in formats:
        html_path = f"{output_prefix}.html"
        df.to_html(html_path, index=False, escape=False)
        saved_files.append(html_path)
        print(f"HTML report written to: {html_path}")
    if "json" in formats:
        json_path = f"{output_prefix}.json"
        with open(json_path, "w", encoding="utf-8") as fp:
            json.dump(rows, fp, indent=2, ensure_ascii=False)
        saved_files.append(json_path)
        print(f"JSON report written to: {json_path}")
    return saved_files


def load_pat(token: str | None = None, ask_missing: bool = True, save: bool = False) -> str:
    if token:
        if save:
            PAT_FILE.write_text(token + "\n")
            try:
                PAT_FILE.chmod(0o600)
            except OSError:
                pass
        return token.strip()

    if PAT_FILE.exists():
        token = PAT_FILE.read_text().strip()
        if token:
            return token

    env_token = os.getenv("AZURE_DEVOPS_PAT")
    if env_token:
        return env_token.strip()

    if not ask_missing:
        raise ValueError("A Personal Access Token is required.")

    token = getpass.getpass("Azure DevOps Personal Access Token: ").strip()
    if not token:
        raise SystemExit("A Personal Access Token is required.")

    save_prompt = input("Save PAT to local file '.azure_devops_pat' for reuse? (y/N): ").strip().lower()
    if save_prompt == "y":
        PAT_FILE.write_text(token + "\n")
        try:
            PAT_FILE.chmod(0o600)
        except OSError:
            pass
    return token


def launch_gui() -> None:
    root = tk.Tk()
    root.title("Azure DevOps PR Report")

    def append_status(message: str) -> None:
        status_text.config(state="normal")
        status_text.insert("end", message + "\n")
        status_text.see("end")
        status_text.config(state="disabled")

    time_values = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 30)] + ["23:59"]

    def run_report() -> None:
        try:
            append_status("Starting report...")
            url = url_entry.get().strip()
            pat = pat_entry.get().strip() or None
            repos_raw = repos_text.get("1.0", "end").strip()
            repo_names = [name.strip() for name in re.split(r"[\n,]+", repos_raw) if name.strip()]
            # Persist repositories if requested
            if save_repos_var.get():
                try:
                    REPO_FILE.write_text(repos_raw.rstrip() + "\n")
                except Exception:
                    append_status("Warning: failed to save repositories to disk.")
            from_dt = parse_iso_datetime(f"{from_date_entry.get()} {from_time_var.get()}")
            to_dt = parse_iso_datetime(f"{to_date_entry.get()} {to_time_var.get()}")
            output_prefix = output_prefix_entry.get().strip() or "azure_pr_report"
            format_choice = format_var.get()
            output_formats = [format_choice]
            organization, project = parse_azure_devops_url(url)
            token = load_pat(pat, ask_missing=False, save=save_pat_var.get())
            repositories = get_repositories(organization, project, token)
            rows = build_report_rows(
                organization,
                project,
                repo_names,
                repositories,
                from_dt,
                to_dt,
                token,
            )
            saved_files = save_outputs(rows, output_prefix, output_formats)
            append_status("Report completed.")
            messagebox.showinfo("Done", "Report saved:\n" + "\n".join(saved_files) if saved_files else "No files created.")
        except Exception as exc:
            append_status(f"Error: {exc}")
            messagebox.showerror("Error", str(exc))

    frame = tk.Frame(root, padx=12, pady=12)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="Azure DevOps URL:").grid(row=0, column=0, sticky="w")
    url_entry = tk.Entry(frame, width=80)
    url_entry.insert(0, "https://dev.azure.com/cohentsahi/cohentzahi_agile")
    url_entry.grid(row=0, column=1, sticky="we", pady=2)

    tk.Label(frame, text="Personal Access Token:").grid(row=1, column=0, sticky="w")
    pat_entry = tk.Entry(frame, width=80, show="*")
    pat_entry.grid(row=1, column=1, sticky="we", pady=2)

    tk.Label(frame, text="Repositories (comma or newline separated):").grid(row=2, column=0, sticky="nw")
    repos_text = tk.Text(frame, width=80, height=4)
    repos_text.grid(row=2, column=1, sticky="we", pady=2)
    # Load saved repositories if present
    try:
        if REPO_FILE.exists():
            repos_text.insert("1.0", REPO_FILE.read_text())
    except Exception:
        pass

    tk.Label(frame, text="From date:").grid(row=3, column=0, sticky="w")
    from_frame = tk.Frame(frame)
    from_date_entry = DateEntry(from_frame, width=12, date_pattern="yyyy-MM-dd")
    from_date_entry.pack(side="left")
    from_time_var = tk.StringVar(value="00:00")
    from_time_combo = ttk.Combobox(from_frame, values=time_values, textvariable=from_time_var, width=6, state="readonly")
    from_time_combo.pack(side="left", padx=(8, 0))
    from_frame.grid(row=3, column=1, sticky="w", pady=2)

    tk.Label(frame, text="To date:").grid(row=4, column=0, sticky="w")
    to_frame = tk.Frame(frame)
    to_date_entry = DateEntry(to_frame, width=12, date_pattern="yyyy-MM-dd")
    to_date_entry.pack(side="left")
    to_time_var = tk.StringVar(value="23:59")
    to_time_combo = ttk.Combobox(to_frame, values=time_values, textvariable=to_time_var, width=6, state="readonly")
    to_time_combo.pack(side="left", padx=(8, 0))
    to_frame.grid(row=4, column=1, sticky="w", pady=2)

    tk.Label(frame, text="Output format:").grid(row=5, column=0, sticky="w")
    format_var = tk.StringVar(value="excel")
    tk.OptionMenu(frame, format_var, "excel", "html", "json").grid(row=5, column=1, sticky="w", pady=2)

    tk.Label(frame, text="Output prefix:").grid(row=6, column=0, sticky="w")
    output_prefix_entry = tk.Entry(frame, width=40)
    output_prefix_entry.insert(0, "azure_pr_report")
    output_prefix_entry.grid(row=6, column=1, sticky="w", pady=2)

    save_pat_var = tk.BooleanVar(value=False)
    tk.Checkbutton(frame, text="Save PAT locally (.azure_devops_pat)", variable=save_pat_var).grid(row=7, column=1, sticky="w", pady=4)
    save_repos_var = tk.BooleanVar(value=False)
    tk.Checkbutton(frame, text="Save repositories locally (.azure_devops_repos)", variable=save_repos_var).grid(row=8, column=1, sticky="w", pady=4)

    run_button = tk.Button(frame, text="Run Report", command=run_report, width=20)
    run_button.grid(row=9, column=1, sticky="w", pady=8)

    tk.Label(frame, text="Status:").grid(row=10, column=0, sticky="nw")
    status_text = tk.Text(frame, width=80, height=10, state="disabled")
    status_text.grid(row=10, column=1, sticky="we", pady=2)

    frame.columnconfigure(1, weight=1)
    root.mainloop()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Azure DevOps pull request report generator for date ranges, repos, and work item associations."
    )
    parser.add_argument(
        "--url",
        default="https://dev.azure.com/cohentsahi/cohentzahi_agile",
        help="Azure DevOps base URL including organization and project. Example: https://dev.azure.com/org/project",
    )
    parser.add_argument(
        "--repos",
        help="Comma-separated list of repository names. If omitted, all repos in the project are used.",
    )
    parser.add_argument(
        "--from",
        dest="from_date",
        type=parse_iso_datetime,
        help="Start date/time in YYYY-MM-DDTHH:MM or YYYY-MM-DD format.",
    )
    parser.add_argument(
        "--to",
        dest="to_date",
        type=parse_iso_datetime,
        help="End date/time in YYYY-MM-DDTHH:MM or YYYY-MM-DD format.",
    )
    parser.add_argument(
        "--format",
        choices=["excel", "html", "json", "both"],
        default="both",
        help="Output format for the report.",
    )
    parser.add_argument(
        "--output-prefix",
        default="azure_pr_report",
        help="Output file prefix (without extension).",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the GUI instead of using command-line mode.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    if args.gui or len(sys.argv) == 1:
        launch_gui()
        return

    if not args.from_date or not args.to_date:
        raise SystemExit("--from and --to are required in command-line mode.")

    organization, project = parse_azure_devops_url(args.url)
    token = load_pat()

    repo_names = [name.strip() for name in args.repos.split(",") if name.strip()] if args.repos else []
    repositories = get_repositories(organization, project, token)
    rows = build_report_rows(
        organization,
        project,
        repo_names,
        repositories,
        args.from_date,
        args.to_date,
        token,
    )

    output_formats = [args.format] if args.format != "both" else ["excel", "html"]
    save_outputs(rows, args.output_prefix, output_formats)


if __name__ == "__main__":
    main()
