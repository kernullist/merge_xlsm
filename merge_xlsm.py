import sys
import os
import win32com.client


XL_CALCULATION_AUTOMATIC = -4105
XL_CALCULATION_MANUAL = -4135
VBA_COMPONENT_EXTENSIONS = {
    1: ".bas",
    2: ".cls",
    3: ".frm",
}


def normalize_2d(values):
    if values is None:
        return []
    if not isinstance(values, tuple):
        return [[values]]

    if not values:
        return []

    first = values[0]
    if isinstance(first, tuple):
        return [list(row) for row in values]

    return [list(values)]


def find_header_row(ws):
    last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    if last_row < 1:
        return None

    col_values = normalize_2d(ws.Range(ws.Cells(1, 2), ws.Cells(last_row, 2)).Value)
    for idx, row in enumerate(col_values, start=1):
        val = row[0] if row else None
        if val is not None and str(val).strip() == "Id":
            return idx
    return None


def get_used_columns(ws):
    last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    return last_col


def read_data_rows(ws, header_row, num_cols):
    data_start = header_row + 1
    last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    if data_start > last_row:
        return []

    values = normalize_2d(ws.Range(ws.Cells(data_start, 1), ws.Cells(last_row, num_cols)).Value)
    rows = []
    for row_data in values:
        key = row_data[1] if len(row_data) > 1 else None
        if key is None or str(key).strip() == "":
            break
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

    clear_end = max(data_start + len(merged_rows) - 1, old_last_row)
    if clear_end >= data_start:
        ws1.Range(ws1.Cells(data_start, 1), ws1.Cells(clear_end, num_cols)).ClearContents()

    if merged_rows:
        write_values = tuple(tuple(row) for row in merged_rows)
        ws1.Range(
            ws1.Cells(data_start, 1),
            ws1.Cells(data_start + len(merged_rows) - 1, num_cols),
        ).Value = write_values

    print(f"  Merged '{sheet_name}': {len(main_rows)} main + {len(hotfix_rows)} hotfix -> {len(merged_rows)} rows")
    return True


def get_vba_component(vb_project, component_name):
    try:
        return vb_project.VBComponents(component_name)
    except Exception:
        return None


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
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    wb1 = None
    wb2 = None
    previous_calculation = None
    calculation_switched = False

    try:
        print(f"Opening '{os.path.basename(file1)}'...")
        wb1 = excel.Workbooks.Open(
            file1,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        print(f"Opening '{os.path.basename(file2)}'...")
        wb2 = excel.Workbooks.Open(
            file2,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        try:
            previous_calculation = excel.Calculation
            excel.Calculation = XL_CALCULATION_MANUAL
            calculation_switched = True
        except Exception as calc_e:
            print(f"WARNING: Could not switch Excel to manual calculation mode: {calc_e}")

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
            temp_dir = os.environ.get("TEMP", "C:\\Temp")
            for comp in wb2.VBProject.VBComponents:
                ext = VBA_COMPONENT_EXTENSIONS.get(comp.Type)
                if ext:
                    temp_path = os.path.join(temp_dir, comp.Name + ext)
                    frx_path = os.path.join(temp_dir, comp.Name + ".frx")
                    print(f"  Exporting VBA module: {comp.Name}")
                    comp.Export(temp_path)

                    existing_comp = get_vba_component(wb1.VBProject, comp.Name)
                    if existing_comp is not None and existing_comp.Type in VBA_COMPONENT_EXTENSIONS:
                        print(f"  Replacing existing VBA module: {comp.Name}")
                        wb1.VBProject.VBComponents.Remove(existing_comp)

                    wb1.VBProject.VBComponents.Import(temp_path)

                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                    if comp.Type == 3 and os.path.exists(frx_path):
                        os.remove(frx_path)

        except Exception as vba_e:
            print("\nWARNING: Could not merge VBA modules.")
            print("To allow VBA merging, go to Excel -> Options -> Trust Center -> Trust Center Settings -> Macro Settings")
            print("And check 'Trust access to the VBA project object model'.")
            print(f"Error details: {vba_e}\n")

        if calculation_switched:
            print("Recalculating workbook...")
            excel.Calculation = XL_CALCULATION_AUTOMATIC
            excel.CalculateFullRebuild()

        print(f"Saving merged workbook as '{os.path.basename(output_file)}'...")
        wb1.SaveAs(output_file, FileFormat=52)
        print("Done!")

    except Exception as e:
        print(f"Error during merge operation: {e}")
    finally:
        if calculation_switched and previous_calculation is not None:
            try:
                excel.Calculation = previous_calculation
            except Exception:
                pass
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
