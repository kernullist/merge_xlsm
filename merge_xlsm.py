import sys
import os
import win32com.client


def find_header_row(ws):
    last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    for r in range(1, last_row + 1):
        val = ws.Cells(r, 2).Value
        if val is not None and str(val).strip() == "Id":
            return r
    return None


def get_used_columns(ws):
    last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    return last_col


def read_data_rows(ws, header_row, num_cols):
    data_start = header_row + 1
    last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    rows = []
    for r in range(data_start, last_row + 1):
        key = ws.Cells(r, 2).Value
        if key is None or str(key).strip() == "":
            break
        row_data = []
        for c in range(1, num_cols + 1):
            row_data.append(ws.Cells(r, c).Value)
        rows.append((str(key).strip(), row_data))
    return rows


def merge_rows_by_key(main_rows, hotfix_rows, num_cols):
    from collections import defaultdict, OrderedDict

    main_groups = defaultdict(list)
    for key, row_data in main_rows:
        main_groups[key].append(row_data)

    hotfix_groups = defaultdict(list)
    for key, row_data in hotfix_rows:
        hotfix_groups[key].append(row_data)

    merged = []
    processed_keys = set()

    all_keys = OrderedDict()
    for key, _ in hotfix_rows:
        if key not in all_keys:
            all_keys[key] = None
    for key, _ in main_rows:
        if key not in all_keys:
            all_keys[key] = None

    for key in all_keys:
        if key in processed_keys:
            continue
        processed_keys.add(key)

        m_list = main_groups.get(key, [])
        h_list = hotfix_groups.get(key, [])
        max_count = max(len(m_list), len(h_list))

        for i in range(max_count):
            m_data = m_list[i] if i < len(m_list) else None
            h_data = h_list[i] if i < len(h_list) else None

            if h_data is not None and m_data is not None:
                row = []
                for c in range(num_cols):
                    hval = h_data[c] if c < len(h_data) else None
                    mval = m_data[c] if c < len(m_data) else None
                    row.append(hval if hval is not None else mval)
                merged.append(row)
            elif h_data is not None:
                merged.append(list(h_data))
            elif m_data is not None:
                merged.append(list(m_data))

    return merged


def merge_sheet_data(ws1, ws2, sheet_name):
    header1 = find_header_row(ws1)
    header2 = find_header_row(ws2)

    if header1 is None or header2 is None:
        print(f"  Skipping sheet '{sheet_name}': no 'Id' header found")
        return False

    num_cols1 = get_used_columns(ws1)
    num_cols2 = get_used_columns(ws2)
    num_cols = max(num_cols1, num_cols2)

    main_rows = read_data_rows(ws1, header1, num_cols)
    hotfix_rows = read_data_rows(ws2, header2, num_cols)

    merged_rows = merge_rows_by_key(main_rows, hotfix_rows, num_cols)

    data_start = header1 + 1
    old_last_row = ws1.UsedRange.Row + ws1.UsedRange.Rows.Count - 1

    clear_end = max(data_start + len(merged_rows), old_last_row)
    for r in range(data_start, clear_end + 1):
        for c in range(1, num_cols + 1):
            ws1.Cells(r, c).ClearContents()

    for r_idx, row_data in enumerate(merged_rows):
        for c_idx, val in enumerate(row_data):
            if val is not None:
                ws1.Cells(data_start + r_idx, c_idx + 1).Value = val

    print(f"  Merged '{sheet_name}': {len(main_rows)} main + {len(hotfix_rows)} hotfix -> {len(merged_rows)} rows")
    return True


def merge_xlsm(file1, file2, output_file):
    file1 = os.path.abspath(file1)
    file2 = os.path.abspath(file2)
    output_file = os.path.abspath(output_file)

    if not os.path.exists(file1):
        print(f"Error: {file1} not found.")
        return
    if not os.path.exists(file2):
        print(f"Error: {file2} not found.")
        return

    try:
        excel = win32com.client.Dispatch("Excel.Application")
    except Exception as e:
        print(f"Error starting Excel. Is Excel installed? {e}")
        return

    excel.Visible = False
    excel.DisplayAlerts = False

    wb1 = None
    wb2 = None

    try:
        print(f"Opening '{os.path.basename(file1)}'...")
        wb1 = excel.Workbooks.Open(file1)

        print(f"Opening '{os.path.basename(file2)}'...")
        wb2 = excel.Workbooks.Open(file2)

        wb1_sheet_names = set()
        for i in range(1, wb1.Sheets.Count + 1):
            wb1_sheet_names.add(wb1.Sheets(i).Name)

        print("Merging sheets...")

        for i in range(1, wb2.Sheets.Count + 1):
            sheet2 = wb2.Sheets(i)
            sheet_name = sheet2.Name

            if sheet_name in wb1_sheet_names:
                sheet1 = wb1.Sheets(sheet_name)
                print(f"  Merging data in sheet '{sheet_name}'...")
                merge_sheet_data(sheet1, sheet2, sheet_name)
            else:
                sheet2.Copy(After=wb1.Sheets(wb1.Sheets.Count))
                print(f"  Copied new sheet '{sheet_name}' from hotfix")

        print("Attempting to merge VBA modules...")
        try:
            temp_dir = os.environ.get('TEMP', 'C:\\Temp')
            for comp in wb2.VBProject.VBComponents:
                if comp.Type in [1, 2, 3]:
                    if comp.Type == 1:
                        ext = ".bas"
                    elif comp.Type == 2:
                        ext = ".cls"
                    else:
                        ext = ".frm"

                    temp_path = os.path.join(temp_dir, comp.Name + ext)
                    print(f"  Exporting VBA module: {comp.Name}")
                    comp.Export(temp_path)

                    wb1.VBProject.VBComponents.Import(temp_path)
                    os.remove(temp_path)

                    if comp.Type == 3:
                        frx_path = os.path.join(temp_dir, comp.Name + ".frx")
                        if os.path.exists(frx_path):
                            os.remove(frx_path)

        except Exception as vba_e:
            print("\nWARNING: Could not merge VBA modules.")
            print("To allow VBA merging, go to Excel -> Options -> Trust Center -> Trust Center Settings -> Macro Settings")
            print("And check 'Trust access to the VBA project object model'.")
            print(f"Error details: {vba_e}\n")

        print(f"Saving merged workbook as '{os.path.basename(output_file)}'...")
        wb1.SaveAs(output_file, FileFormat=52)
        print("Done!")

    except Exception as e:
        print(f"Error during merge operation: {e}")
    finally:
        if wb2:
            wb2.Close(False)
        if wb1:
            wb1.Close(False)
        excel.Quit()


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python merge_xlsm.py <main.xlsm> <hotfix.xlsm> <output.xlsm>")
        sys.exit(1)

    merge_xlsm(sys.argv[1], sys.argv[2], sys.argv[3])