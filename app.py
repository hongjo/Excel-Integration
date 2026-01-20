import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from copy import copy


def select_files() -> Tuple[str, ...]:
    return filedialog.askopenfilenames(
        title="Select Excel Files",
        filetypes=[("Excel Files", "*.xlsx")],
    )


def select_output_path() -> str:
    return filedialog.asksaveasfilename(
        title="Save Consolidated File",
        defaultextension=".xlsx",
        initialfile="Consolidated_All_Sheets.xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
    )


def safe_sheet_title(base_title: str, existing_titles: set[str]) -> str:
    sanitized = base_title[:31]
    if sanitized not in existing_titles:
        return sanitized

    suffix = 2
    while True:
        suffix_str = f"_{suffix}"
        trimmed_base = sanitized[: 31 - len(suffix_str)]
        candidate = f"{trimmed_base}{suffix_str}"
        if candidate not in existing_titles:
            return candidate
        suffix += 1


def copy_sheet(
    src_ws: Worksheet,
    dst_wb: Workbook,
    new_title: str,
    values_ws: Optional[Worksheet] = None,
    value_only: bool = False,
) -> Worksheet:
    dst_ws = dst_wb.create_sheet(title=new_title)

    for col_letter, dimension in src_ws.column_dimensions.items():
        dst_dim = dst_ws.column_dimensions[col_letter]
        dst_dim.width = dimension.width
        dst_dim.hidden = dimension.hidden
        dst_dim.outline_level = dimension.outline_level
        dst_dim.collapsed = dimension.collapsed

    for row_idx, dimension in src_ws.row_dimensions.items():
        dst_dim = dst_ws.row_dimensions[row_idx]
        dst_dim.height = dimension.height
        dst_dim.hidden = dimension.hidden
        dst_dim.outline_level = dimension.outline_level
        dst_dim.collapsed = dimension.collapsed

    max_row = src_ws.max_row or 1
    max_col = src_ws.max_column or 1

    for row in src_ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell_value = cell.value
            if value_only and values_ws is not None:
                cell_value = values_ws.cell(row=cell.row, column=cell.column).value
            dst_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell_value)
            if cell.has_style:
                dst_cell.font = copy(cell.font)
                dst_cell.border = copy(cell.border)
                dst_cell.fill = copy(cell.fill)
                dst_cell.number_format = cell.number_format
                dst_cell.protection = copy(cell.protection)
                dst_cell.alignment = copy(cell.alignment)
            if cell.hyperlink:
                dst_cell._hyperlink = copy(cell.hyperlink)
            if cell.comment:
                dst_cell.comment = copy(cell.comment)

    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

    for cf in src_ws.conditional_formatting:
        dst_ws.conditional_formatting.add(cf, src_ws.conditional_formatting[cf])

    dst_ws.sheet_properties = copy(src_ws.sheet_properties)

    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes

    return dst_ws


def process_files(
    file_paths: List[str],
    output_path: str,
    gui_callbacks: Dict[str, callable],
) -> Dict[str, List[str]]:
    progress_callback = gui_callbacks.get("progress")
    status_callback = gui_callbacks.get("status")

    wb_out = Workbook()
    default_sheet = wb_out.active
    wb_out.remove(default_sheet)

    existing_titles: set[str] = set()
    success_files: List[str] = []
    failed_files: List[str] = []

    for file_idx, file_path in enumerate(file_paths, start=1):
        file_name = Path(file_path).name
        if file_name.startswith("~$"):
            continue

        if status_callback:
            status_callback(f"Loading: {file_name} ({file_idx}/{len(file_paths)})")

        try:
            wb_in = load_workbook(filename=file_path, data_only=False)
        except Exception as exc:  # noqa: BLE001
            failed_files.append(f"{file_name} (open error: {exc})")
            if progress_callback:
                progress_callback(file_idx, len(file_paths))
            continue

        wb_in_values: Optional[Workbook] = None

        try:
            for sheet_idx, sheet in enumerate(wb_in.worksheets):
                if sheet_idx == 0:
                    continue
                values_ws = None
                value_only = False
                if sheet_idx == 1:
                    if wb_in_values is None:
                        wb_in_values = load_workbook(
                            filename=file_path,
                            data_only=True,
                        )
                    values_ws = wb_in_values.worksheets[sheet_idx]
                    value_only = True
                base_title = f"{Path(file_name).stem}_{sheet.title}"
                new_title = safe_sheet_title(base_title, existing_titles)
                existing_titles.add(new_title)
                if status_callback:
                    status_callback(
                        f"Copying: {file_name} -> {sheet.title}"
                    )
                copy_sheet(
                    sheet,
                    wb_out,
                    new_title,
                    values_ws=values_ws,
                    value_only=value_only,
                )
            success_files.append(file_name)
        except Exception as exc:  # noqa: BLE001
            failed_files.append(f"{file_name} (sheet error: {exc})")
        finally:
            wb_in.close()
            if wb_in_values is not None:
                wb_in_values.close()

        if progress_callback:
            progress_callback(file_idx, len(file_paths))

    if status_callback:
        status_callback("Saving consolidated file...")

    wb_out.save(output_path)

    return {"success": success_files, "failed": failed_files}


def main() -> None:
    root = tk.Tk()
    root.title("Excel Sheet Consolidator")
    root.geometry("700x500")

    file_list: List[str] = []
    output_path: Optional[str] = None

    def update_file_listbox() -> None:
        listbox.delete(0, tk.END)
        for file_path in file_list:
            listbox.insert(tk.END, file_path)

    def on_select_files() -> None:
        nonlocal file_list
        selected = select_files()
        file_list = [path for path in selected if not Path(path).name.startswith("~$")]
        update_file_listbox()
        status_var.set(f"Selected {len(file_list)} files.")

    def on_select_output() -> None:
        nonlocal output_path
        selected_path = select_output_path()
        if selected_path:
            output_path = selected_path
            output_label_var.set(f"Output: {output_path}")

    def update_progress(current: int, total: int) -> None:
        progress_bar["maximum"] = max(total, 1)
        progress_bar["value"] = current
        root.update_idletasks()

    def update_status(message: str) -> None:
        status_var.set(message)
        root.update_idletasks()

    def on_run() -> None:
        if not file_list:
            messagebox.showerror("Error", "Please select Excel files.")
            return
        if not output_path:
            messagebox.showerror("Error", "Please select an output path.")
            return

        progress_bar["value"] = 0
        status_var.set("Starting...")

        try:
            results = process_files(
                file_list,
                output_path,
                {
                    "progress": update_progress,
                    "status": update_status,
                },
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", f"Unexpected error: {exc}")
            return

        success_count = len(results["success"])
        failed_count = len(results["failed"])
        summary_lines = [
            f"Success files: {success_count}",
            f"Failed files: {failed_count}",
        ]
        if failed_count:
            summary_lines.append("\nFailed list:")
            summary_lines.extend(results["failed"])

        summary = "\n".join(summary_lines)
        messagebox.showinfo("Completed", summary)
        status_var.set("Done")

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    select_button = ttk.Button(frame, text="파일 선택", command=on_select_files)
    select_button.pack(fill=tk.X)

    listbox = tk.Listbox(frame, height=10)
    listbox.pack(fill=tk.BOTH, expand=True, pady=5)

    output_button = ttk.Button(frame, text="저장 위치 선택", command=on_select_output)
    output_button.pack(fill=tk.X, pady=(5, 0))

    output_label_var = tk.StringVar(value="Output: (not selected)")
    output_label = ttk.Label(frame, textvariable=output_label_var)
    output_label.pack(fill=tk.X)

    run_button = ttk.Button(frame, text="실행", command=on_run)
    run_button.pack(fill=tk.X, pady=(10, 0))

    progress_bar = ttk.Progressbar(frame, mode="determinate")
    progress_bar.pack(fill=tk.X, pady=(10, 0))

    status_var = tk.StringVar(value="Ready")
    status_label = ttk.Label(frame, textvariable=status_var)
    status_label.pack(fill=tk.X, pady=(5, 0))

    instructions = ttk.Label(
        frame,
        text="Install: pip install openpyxl\nRun: python app.py",
        foreground="#666666",
    )
    instructions.pack(fill=tk.X, pady=(10, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
