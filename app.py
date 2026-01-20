import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from copy import copy as copy_style

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


SELECTED_FILES = []
OUTPUT_PATH = ""


def select_files(listbox):
    """Open file dialog to select multiple Excel files."""
    global SELECTED_FILES
    paths = filedialog.askopenfilenames(
        title="엑셀 파일 선택",
        filetypes=[("Excel files", "*.xlsx")],
    )
    if paths:
        SELECTED_FILES = list(paths)
        listbox.delete(0, tk.END)
        for path in SELECTED_FILES:
            listbox.insert(tk.END, path)


def select_output_path(path_var):
    """Open save dialog to select output Excel file path."""
    global OUTPUT_PATH
    path = filedialog.asksaveasfilename(
        title="저장 위치 선택",
        defaultextension=".xlsx",
        initialfile="Consolidated_Site_Assessment.xlsx",
        filetypes=[("Excel files", "*.xlsx")],
    )
    if path:
        OUTPUT_PATH = path
        path_var.set(path)


def _used_range(ws):
    """Return used range boundaries as (min_col, min_row, max_col, max_row)."""
    try:
        dimension = ws.calculate_dimension()
        if dimension == "A1" and ws["A1"].value is None:
            return None
        return range_boundaries(dimension)
    except Exception:
        return None


def extract_site_name(wb):
    """Extract site name from 요약 sheet based on fixed cell or label search."""
    if "요약" not in wb.sheetnames:
        return ""
    ws = wb["요약"]

    e1_value = ws["E1"].value
    if e1_value:
        return str(e1_value).strip()

    bounds = _used_range(ws)
    if not bounds:
        return ""
    min_col, min_row, max_col, max_row = bounds

    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if cell.value is None:
                continue
            cell_text = str(cell.value).strip()
            if "현장명" in cell_text:
                right_cell = ws.cell(row=cell.row, column=cell.column + 1)
                if right_cell.value:
                    return str(right_cell.value).strip()
    return ""


def extract_score_grade(wb):
    """Extract total score and grade from 요약 sheet with fallback label search."""
    if "요약" not in wb.sheetnames:
        return None, None
    ws = wb["요약"]

    total_score = ws["G32"].value
    grade = ws["G33"].value

    if total_score is not None and grade is not None:
        return total_score, grade

    bounds = _used_range(ws)
    if not bounds:
        return total_score, grade

    min_col, min_row, max_col, max_row = bounds
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if cell.value is None:
                continue
            cell_text = str(cell.value).replace(" ", "").strip()
            if total_score is None and cell_text in {"합계", "합계", "합계:"}:
                total_score = ws.cell(row=cell.row, column=7).value
            if grade is None and "평가등급" in cell_text:
                grade = ws.cell(row=cell.row, column=7).value
        if total_score is not None and grade is not None:
            break

    return total_score, grade


def copy_sheet(src_ws, dst_ws):
    """Copy values and basic formatting from source worksheet to destination worksheet."""
    bounds = _used_range(src_ws)
    if not bounds:
        return
    min_col, min_row, max_col, max_row = bounds

    for row in src_ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            dst_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                dst_cell._style = copy_style(cell._style)
            if cell.hyperlink:
                dst_cell.hyperlink = copy_style(cell.hyperlink)
            if cell.comment:
                dst_cell.comment = copy_style(cell.comment)

    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

    for col_idx in range(min_col, max_col + 1):
        letter = get_column_letter(col_idx)
        width = src_ws.column_dimensions[letter].width
        if width is not None:
            dst_ws.column_dimensions[letter].width = width

    for row_idx in range(min_row, max_row + 1):
        height = src_ws.row_dimensions[row_idx].height
        if height is not None:
            dst_ws.row_dimensions[row_idx].height = height


def safe_sheet_title(base_title, existing_titles):
    """Return a safe, unique sheet title within Excel's 31-char limit."""
    max_len = 31
    base_title = base_title[:max_len]
    title = base_title
    counter = 2
    while title in existing_titles:
        suffix = f"_{counter}"
        allowed_len = max_len - len(suffix)
        title = f"{base_title[:allowed_len]}{suffix}"
        counter += 1
    existing_titles.add(title)
    return title


def process_one_file(path, out_wb, summary_ws, existing_titles):
    """Process a single Excel file and append summary row."""
    file_name = os.path.basename(path)
    processed_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result = {
        "Index": None,
        "SiteName": "",
        "FileName": file_name,
        "TotalScore": None,
        "Grade": None,
        "ImportedSheets": 0,
        "Status": "OK",
        "ErrorMessage": "",
        "ProcessedAt": processed_at,
    }

    try:
        wb = load_workbook(path, data_only=True)
        result["SiteName"] = extract_site_name(wb)
        result["TotalScore"], result["Grade"] = extract_score_grade(wb)

        for sheet_name in wb.sheetnames:
            src_ws = wb[sheet_name]
            base_title = f"{result['SiteName']}_{sheet_name}" if result["SiteName"] else sheet_name
            safe_title = safe_sheet_title(base_title, existing_titles)
            dst_ws = out_wb.create_sheet(title=safe_title)
            copy_sheet(src_ws, dst_ws)
            result["ImportedSheets"] += 1
    except Exception as exc:
        result["Status"] = "FAIL"
        result["ErrorMessage"] = str(exc)

    return result


def _write_summary(summary_ws, results):
    headers = [
        "Index",
        "SiteName",
        "FileName",
        "TotalScore",
        "Grade",
        "ImportedSheets",
        "Status",
        "ErrorMessage",
        "ProcessedAt",
    ]
    summary_ws.append(headers)

    def parse_score(value):
        try:
            return float(value)
        except Exception:
            return None

    sorted_results = sorted(
        results,
        key=lambda r: (
            parse_score(r["TotalScore"]) is None,
            -(parse_score(r["TotalScore"]) or 0),
        ),
    )

    for idx, result in enumerate(sorted_results, start=1):
        result["Index"] = idx
        summary_ws.append([
            result["Index"],
            result["SiteName"],
            result["FileName"],
            result["TotalScore"],
            result["Grade"],
            result["ImportedSheets"],
            result["Status"],
            result["ErrorMessage"],
            result["ProcessedAt"],
        ])

    success_count = sum(1 for r in results if r["Status"] == "OK")
    fail_count = sum(1 for r in results if r["Status"] != "OK")

    summary_ws.append([])
    summary_ws.append(["Summary", "Success", success_count, "Fail", fail_count])


def main_run(root, listbox, progress, status_var):
    if not SELECTED_FILES:
        messagebox.showwarning("경고", "파일을 선택하세요.")
        return
    if not OUTPUT_PATH:
        messagebox.showwarning("경고", "저장 위치를 선택하세요.")
        return

    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    summary_ws = out_wb.create_sheet(title="집계")

    results = []
    existing_titles = set(out_wb.sheetnames)

    total_files = len(SELECTED_FILES)
    progress["maximum"] = total_files

    for idx, path in enumerate(SELECTED_FILES, start=1):
        status_var.set(f"처리 중: {os.path.basename(path)} ({idx}/{total_files})")
        root.update_idletasks()

        result = process_one_file(path, out_wb, summary_ws, existing_titles)
        results.append(result)

        progress["value"] = idx
        root.update_idletasks()

    _write_summary(summary_ws, results)

    out_wb.save(OUTPUT_PATH)

    success_count = sum(1 for r in results if r["Status"] == "OK")
    fail_count = sum(1 for r in results if r["Status"] != "OK")
    messagebox.showinfo(
        "완료",
        f"처리가 완료되었습니다.\n성공: {success_count} / 실패: {fail_count}\n저장 경로: {OUTPUT_PATH}",
    )
    status_var.set("완료")


def build_gui():
    root = tk.Tk()
    root.title("현장수준 등급평가 통합 도구")
    root.geometry("800x500")

    file_frame = ttk.Frame(root, padding=10)
    file_frame.pack(fill=tk.BOTH, expand=True)

    listbox = tk.Listbox(file_frame, height=10)
    listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

    scrollbar = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.config(yscrollcommand=scrollbar.set)

    button_frame = ttk.Frame(root, padding=10)
    button_frame.pack(fill=tk.X)

    ttk.Button(button_frame, text="파일 선택", command=lambda: select_files(listbox)).pack(
        side=tk.LEFT, padx=5
    )

    output_var = tk.StringVar()
    ttk.Button(
        button_frame,
        text="저장 위치 선택",
        command=lambda: select_output_path(output_var),
    ).pack(side=tk.LEFT, padx=5)

    ttk.Label(button_frame, textvariable=output_var).pack(side=tk.LEFT, padx=5)

    progress_frame = ttk.Frame(root, padding=10)
    progress_frame.pack(fill=tk.X)

    progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode="determinate")
    progress.pack(fill=tk.X, expand=True)

    status_var = tk.StringVar(value="대기 중")
    ttk.Label(progress_frame, textvariable=status_var).pack(anchor=tk.W)

    action_frame = ttk.Frame(root, padding=10)
    action_frame.pack(fill=tk.X)

    ttk.Button(
        action_frame,
        text="실행",
        command=lambda: main_run(root, listbox, progress, status_var),
    ).pack(side=tk.LEFT, padx=5)

    info_frame = ttk.Frame(root, padding=10)
    info_frame.pack(fill=tk.X)
    ttk.Label(
        info_frame,
        text=(
            "설치: pip install openpyxl\n"
            "실행: python app.py"
        ),
        justify=tk.LEFT,
    ).pack(anchor=tk.W)

    return root


if __name__ == "__main__":
    app_root = build_gui()
    app_root.mainloop()
