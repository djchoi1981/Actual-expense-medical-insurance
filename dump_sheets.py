import openpyxl

wb = openpyxl.load_workbook('/Users/djchoi81/Downloads/실손보험_세대별_자기부담계산기_5.xlsx', data_only=False)

with open('sheets_dump.txt', 'w', encoding='utf-8') as f:
    def dump_sheet(ws):
        f.write("="*80 + "\n")
        f.write(f"Sheet: {ws.title} ({ws.dimensions})\n")
        f.write("="*80 + "\n")
        for row in ws.iter_rows(values_only=False):
            row_vals = []
            for cell in row:
                if cell.value is not None:
                    row_vals.append(f"{cell.coordinate}:{cell.value}")
            if row_vals:
                f.write(" | ".join(row_vals) + "\n")

    dump_sheet(wb['실손보험금계산기'])
    dump_sheet(wb['검증_사례비교'])
    dump_sheet(wb['세대별공제산식기준'])
