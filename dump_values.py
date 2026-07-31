import openpyxl

wb = openpyxl.load_workbook('실손보험_세대별_자기부담계산기.xlsx', data_only=True)

def dump_sheet_values(ws):
    print("="*60)
    print(f"Sheet: {ws.title} (CALCULATED VALUES)")
    print("="*60)
    for row in ws.iter_rows(values_only=False):
        row_vals = []
        for cell in row:
            if cell.value is not None:
                row_vals.append(f"{cell.coordinate}:{cell.value}")
        if row_vals:
            print(" | ".join(row_vals))

dump_sheet_values(wb['실손보험금계산기'])
dump_sheet_values(wb['검증_사례비교'])
