import os
import sys
import shutil
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from .config import (
    PLANT_MAPPINGS,
    FONT_MAIN_TITLE,
    FONT_SUBTITLE,
    FONT_SEC_TITLE,
    FONT_TBL_HEADER,
    FONT_DATA,
    FONT_DATA_BOLD,
    FONT_RESULT,
    FILL_NAVY,
    FILL_SOFT_HEADER,
    FILL_SUB_BANNER,
    FILL_ZEBRA,
    FILL_TOTAL,
    FILL_HIGHLIGHT,
    CELL_BORDER,
    TOTAL_BORDER
)
from .classifier import classify_ticket, get_plant_location
from .processor import parse_datetime_value


def check_file_writable(file_path):
    """
    Returns True if file is writable or doesn't exist yet;
    Returns False if locked by an active process (such as Microsoft Excel).
    """
    if not os.path.exists(file_path):
        return True
    try:
        with open(file_path, 'a'):
            pass
        return True
    except (PermissionError, OSError):
        return False


def ensure_file_writable(file_path):
    """
    Prevents PermissionError crashes by verifying write availability.
    If the file is locked in Excel, prompts the user to close it.
    """
    while not check_file_writable(file_path):
        print("\n" + "!" * 75)
        print("  UYARI: HEDEF EXCEL DOSYASI SU ANDA BASKA BIR PROGRAMDA ACIK!")
        print("!" * 75)
        print(f"  Kilitli Dosya : {file_path}")
        print("  Bu dosya su anda Microsoft Excel veya baska bir uygulama tarafindan kullaniliyor.")
        print("  Lutfen acik olan Excel dosyasini KAPATIP devam etmek icin [Enter]'a basin.")
        print("  (Islemi iptal etmek icin 'Q' tuslayip Enter'a basabilirsiniz)")
        try:
            choice = input("\nSeciminiz [Enter: Yeniden Dene, Q: Cikis]: ").strip()
        except KeyboardInterrupt:
            print("\n\n[*] Islem kullanici tarafindan iptal edildi. Cikis yapiliyor.\n")
            sys.exit(0)
            
        if choice.upper() == 'Q':
            print("\n[*] Islem kullanici tarafindan iptal edildi. Cikis yapildi.\n")
            sys.exit(0)


def build_sla_report_workbook(all_records, header_cols, input_path, output_path, total_source_count):
    """
    Creates the output Excel report by copying the full original workbook
    (preserving all quarterly sheets Tickets-Q1..Q4) and populating 'Year-2026'
    and 'Summary' sheets.
    """
    # Verify file is not open/locked before copying
    ensure_file_writable(output_path)
    
    # Copy original file to output path so input file remains 100% untouched
    shutil.copy2(input_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    
    # -------------------------------------------------------------------------
    # 1. Populate 'Year-2026' Consolidated Tab
    # -------------------------------------------------------------------------
    if 'Year-2026' in wb.sheetnames:
        ws_year = wb['Year-2026']
        ws_year.delete_rows(1, ws_year.max_row + 10)
        ws_year.delete_cols(1, ws_year.max_column + 10)
    else:
        ws_year = wb.create_sheet('Year-2026')
        
    ws_year.views.sheetView[0].showGridLines = True
    
    # Headers: 33 original + 6 helper columns
    extended_headers = list(header_cols) + [
        'In_Scope',
        'Resolution_Hours',
        'SLA_Status',
        'Scope_Category',
        'Plant_Site',
        'Quarter'
    ]
    ws_year.append(extended_headers)
    
    # Style Header Row
    for col_idx in range(1, len(extended_headers) + 1):
        cell = ws_year.cell(row=1, column=col_idx)
        cell.font = FONT_TBL_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_year.row_dimensions[1].height = 28
    
    date_col_indices = [3, 10, 11]  # Col C (Creation), Col J (Solution), Col K (Closure)
    
    in_scope_count = 0
    for idx, rec in enumerate(all_records, start=2):
        in_scope, category = classify_ticket(rec)
        plant_site = get_plant_location(rec)
        if in_scope:
            in_scope_count += 1
            
        row_values = []
        for c_idx, h in enumerate(header_cols, start=1):
            val = rec.get(h)
            if c_idx in date_col_indices:
                val = parse_datetime_value(val)
            row_values.append(val)
            
        # Helper Columns:
        # Col 34 (AH): In_Scope
        row_values.append(in_scope)
        # Col 35 (AI): Resolution_Hours: =(J{idx}-C{idx})*24
        row_values.append(f"=(J{idx}-C{idx})*24")
        # Col 36 (AJ): SLA_Status: =IF(AI{idx}<=4, "PASSED", "FAILED")
        row_values.append(f'=IF(AI{idx}<=4, "PASSED", "FAILED")')
        # Col 37 (AK): Scope_Category
        row_values.append(category)
        # Col 38 (AL): Plant_Site
        row_values.append(plant_site)
        # Col 39 (AM): Quarter
        row_values.append(rec.get('_quarter', 'Q1'))
        
        ws_year.append(row_values)
        
        # Date & format adjustments
        for d_col in date_col_indices:
            c_val = ws_year.cell(row=idx, column=d_col).value
            if isinstance(c_val, datetime):
                ws_year.cell(row=idx, column=d_col).number_format = 'yyyy-mm-dd hh:mm:ss'
                
        ws_year.cell(row=idx, column=35).number_format = '0.00'
        ws_year.cell(row=idx, column=34).alignment = Alignment(horizontal='center')
        ws_year.cell(row=idx, column=35).alignment = Alignment(horizontal='right')
        ws_year.cell(row=idx, column=36).alignment = Alignment(horizontal='center')
        
    max_row = len(all_records) + 1
    
    # Auto-adjust column widths for Year-2026
    for col in ws_year.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col[:100])
        ws_year.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    ws_year.column_dimensions['AH'].width = 14
    ws_year.column_dimensions['AI'].width = 18
    ws_year.column_dimensions['AJ'].width = 14
    ws_year.column_dimensions['AK'].width = 24
    ws_year.column_dimensions['AL'].width = 16
    ws_year.column_dimensions['AM'].width = 12

    # -------------------------------------------------------------------------
    # 2. Build 'Summary' Executive Dashboard Tab
    # -------------------------------------------------------------------------
    if 'Summary' in wb.sheetnames:
        ws_sum = wb['Summary']
        ws_sum.delete_rows(1, ws_sum.max_row + 10)
        ws_sum.delete_cols(1, ws_sum.max_column + 10)
    else:
        ws_sum = wb.create_sheet('Summary', 0)
        
    summary_idx = wb.sheetnames.index('Summary')
    if summary_idx != 0:
        wb._sheets.insert(0, wb._sheets.pop(summary_idx))
        
    ws_sum.views.sheetView[0].showGridLines = True
    
    # Title Banner (Row 2 & 3)
    ws_sum.merge_cells('B2:G2')
    cell_title = ws_sum['B2']
    cell_title.value = "PROMETEON TYRE GROUP - IT INFRASTRUCTURE SLA PERFORMANCE REPORT (2026)"
    cell_title.font = FONT_MAIN_TITLE
    cell_title.fill = FILL_NAVY
    cell_title.alignment = Alignment(horizontal='center', vertical='center')
    ws_sum.row_dimensions[2].height = 32
    
    ws_sum.merge_cells('B3:G3')
    cell_sub = ws_sum['B3']
    cell_sub.value = f"Scope: BusinessLine = 'Infrastructure Services' | Severity = '1 - Molto alta' & '2 - Alta' | SLA Target: <= 4.0h | Total Analyzed: {total_source_count:,}"
    cell_sub.font = FONT_SUBTITLE
    cell_sub.fill = FILL_SUB_BANNER
    cell_sub.alignment = Alignment(horizontal='center', vertical='center')
    ws_sum.row_dimensions[3].height = 20

    # -------------------------------------------------------------------------
    # Table 1: Manager's Exact Summary Table (Rows 5 to 10)
    # -------------------------------------------------------------------------
    headers_t1 = ['Metric', 'Count / Value', 'Ratio / Rate', 'Status']
    for col_idx, h_text in enumerate(headers_t1, start=2):
        c = ws_sum.cell(row=5, column=col_idx)
        c.value = h_text
        c.font = FONT_TBL_HEADER
        c.fill = FILL_NAVY
        c.alignment = Alignment(horizontal='center' if col_idx > 2 else 'left', vertical='center')
        c.border = CELL_BORDER
    ws_sum.row_dimensions[5].height = 26
    
    # Row 6: Total Critical Tickets
    ws_sum['B6'] = "Total Critical Tickets (Sev 1 + Sev 2)"
    ws_sum['C6'] = f"=COUNTIF('Year-2026'!$AH$2:$AH${max_row}, TRUE)"
    ws_sum['D6'] = 1.0
    ws_sum['E6'] = "-"
    
    # Row 7: Closed Within 4 Hours (SLA Passed)
    ws_sum['B7'] = "Closed Within 4 Hours (SLA Passed)"
    ws_sum['C7'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AJ$2:$AJ${max_row}, "PASSED")'
    ws_sum['D7'] = "=IF(C6>0, C7/C6, 0)"
    ws_sum['E7'] = "-"
    
    # Row 8: Exceeding 4 Hours / SLA Breach (Failed)
    ws_sum['B8'] = "Exceeding 4 Hours / SLA Breach (Failed)"
    ws_sum['C8'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AJ$2:$AJ${max_row}, "FAILED")'
    ws_sum['D8'] = "=IF(C6>0, C8/C6, 0)"
    ws_sum['E8'] = "-"
    
    # Row 9: KPI Target
    ws_sum['B9'] = "KPI Target"
    ws_sum['C9'] = "-"
    ws_sum['D9'] = 0.90
    ws_sum['E9'] = "-"
    
    # Row 10: Result
    ws_sum['B10'] = "Result"
    ws_sum['C10'] = "-"
    ws_sum['D10'] = "-"
    ws_sum['E10'] = '=IF(D7>=D9, "PASSED", "FAILED")'
    
    # Format Table 1
    for r in range(6, 11):
        ws_sum.row_dimensions[r].height = 22
        for col_idx in range(2, 6):
            cell = ws_sum.cell(row=r, column=col_idx)
            cell.font = FONT_DATA_BOLD if r in [6, 10] else FONT_DATA
            cell.border = CELL_BORDER
            if r == 10:
                cell.fill = FILL_HIGHLIGHT
                if col_idx == 5:
                    cell.font = FONT_RESULT
            elif r % 2 == 1:
                cell.fill = FILL_ZEBRA
                
            if col_idx == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.number_format = '#,##0'
            elif col_idx == 4:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if r in [6, 7, 8, 9] and cell.value != "-":
                    cell.number_format = '0.0%'
            elif col_idx == 5:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # -------------------------------------------------------------------------
    # Table 2: Domain Breakdown (Rows 13 to 17)
    # -------------------------------------------------------------------------
    ws_sum.cell(row=12, column=2, value="1. DOMAIN & INFRASTRUCTURE CATEGORY BREAKDOWN").font = FONT_SEC_TITLE
    
    headers_t2 = ['Category', 'In-Scope Incidents', 'SLA Met (Passed)', 'SLA Breached (Failed)', 'SLA Compliance (%)', 'Avg Resolution (Hours)']
    for col_idx, h_text in enumerate(headers_t2, start=2):
        c = ws_sum.cell(row=13, column=col_idx)
        c.value = h_text
        c.font = FONT_TBL_HEADER
        c.fill = FILL_SOFT_HEADER
        c.alignment = Alignment(horizontal='center' if col_idx > 2 else 'left', vertical='center')
        c.border = CELL_BORDER
    ws_sum.row_dimensions[13].height = 24
    
    # Row 14: Network
    ws_sum['B14'] = "Managed Network Services"
    ws_sum['C14'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Network")'
    ws_sum['D14'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Network", \'Year-2026\'!$AJ$2:$AJ${max_row}, "PASSED")'
    ws_sum['E14'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Network", \'Year-2026\'!$AJ$2:$AJ${max_row}, "FAILED")'
    ws_sum['F14'] = "=IF(C14>0, D14/C14, 0)"
    ws_sum['G14'] = f'=AVERAGEIFS(\'Year-2026\'!$AI$2:$AI${max_row}, \'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Network")'
    
    # Row 15: Industrial OT
    ws_sum['B15'] = "Industrial OT & Plant Infrastructure"
    ws_sum['C15'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Industrial OT")'
    ws_sum['D15'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Industrial OT", \'Year-2026\'!$AJ$2:$AJ${max_row}, "PASSED")'
    ws_sum['E15'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Industrial OT", \'Year-2026\'!$AJ$2:$AJ${max_row}, "FAILED")'
    ws_sum['F15'] = "=IF(C15>0, D15/C15, 0)"
    ws_sum['G15'] = f'=AVERAGEIFS(\'Year-2026\'!$AI$2:$AI${max_row}, \'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AK$2:$AK${max_row}, "Industrial OT")'
    
    # Row 16: Total
    ws_sum['B16'] = "Total In-Scope Infrastructure"
    ws_sum['C16'] = "=SUM(C14:C15)"
    ws_sum['D16'] = "=SUM(D14:D15)"
    ws_sum['E16'] = "=SUM(E14:E15)"
    ws_sum['F16'] = "=IF(C16>0, D16/C16, 0)"
    ws_sum['G16'] = f'=AVERAGEIFS(\'Year-2026\'!$AI$2:$AI${max_row}, \'Year-2026\'!$AH$2:$AH${max_row}, TRUE)'
    
    for r in range(14, 17):
        ws_sum.row_dimensions[r].height = 20
        is_tot = (r == 16)
        for col_idx in range(2, 8):
            cell = ws_sum.cell(row=r, column=col_idx)
            cell.font = FONT_DATA_BOLD if is_tot else FONT_DATA
            cell.border = TOTAL_BORDER if is_tot else CELL_BORDER
            if is_tot:
                cell.fill = FILL_TOTAL
            elif r % 2 == 1:
                cell.fill = FILL_ZEBRA
                
            if col_idx == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif col_idx in [3, 4, 5]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0'
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '0.0%'
            elif col_idx == 7:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '0.00'

    # -------------------------------------------------------------------------
    # Table 3: Severity Breakdown (Rows 19 to 23)
    # -------------------------------------------------------------------------
    ws_sum.cell(row=18, column=2, value="2. BREAKDOWN BY SEVERITY LEVEL").font = FONT_SEC_TITLE
    
    headers_t3 = ['Severity Level', 'In-Scope Incidents', 'SLA Met (Passed)', 'SLA Breached (Failed)', 'SLA Compliance (%)']
    for col_idx, h_text in enumerate(headers_t3, start=2):
        c = ws_sum.cell(row=19, column=col_idx)
        c.value = h_text
        c.font = FONT_TBL_HEADER
        c.fill = FILL_SOFT_HEADER
        c.alignment = Alignment(horizontal='center' if col_idx > 2 else 'left', vertical='center')
        c.border = CELL_BORDER
    ws_sum.row_dimensions[19].height = 24
    
    # Row 20: Severity 1
    ws_sum['B20'] = "Severity 1 - Molto alta (Critical)"
    ws_sum['C20'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$I$2:$I${max_row}, "1 - Molto alta")'
    ws_sum['D20'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$I$2:$I${max_row}, "1 - Molto alta", \'Year-2026\'!$AJ$2:$AJ${max_row}, "PASSED")'
    ws_sum['E20'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$I$2:$I${max_row}, "1 - Molto alta", \'Year-2026\'!$AJ$2:$AJ${max_row}, "FAILED")'
    ws_sum['F20'] = "=IF(C20>0, D20/C20, 0)"
    
    # Row 21: Severity 2
    ws_sum['B21'] = "Severity 2 - Alta (High)"
    ws_sum['C21'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$I$2:$I${max_row}, "2 - Alta")'
    ws_sum['D21'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$I$2:$I${max_row}, "2 - Alta", \'Year-2026\'!$AJ$2:$AJ${max_row}, "PASSED")'
    ws_sum['E21'] = f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$I$2:$I${max_row}, "2 - Alta", \'Year-2026\'!$AJ$2:$AJ${max_row}, "FAILED")'
    ws_sum['F21'] = "=IF(C21>0, D21/C21, 0)"
    
    # Row 22: Total
    ws_sum['B22'] = "Total Severity 1 & 2"
    ws_sum['C22'] = "=SUM(C20:C21)"
    ws_sum['D22'] = "=SUM(D20:D21)"
    ws_sum['E22'] = "=SUM(E20:E21)"
    ws_sum['F22'] = "=IF(C22>0, D22/C22, 0)"
    
    for r in range(20, 23):
        ws_sum.row_dimensions[r].height = 20
        is_tot = (r == 22)
        for col_idx in range(2, 7):
            cell = ws_sum.cell(row=r, column=col_idx)
            cell.font = FONT_DATA_BOLD if is_tot else FONT_DATA
            cell.border = TOTAL_BORDER if is_tot else CELL_BORDER
            if is_tot:
                cell.fill = FILL_TOTAL
            elif r % 2 == 1:
                cell.fill = FILL_ZEBRA
                
            if col_idx == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif col_idx in [3, 4, 5]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0'
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '0.0%'

    # -------------------------------------------------------------------------
    # Table 4: Plant Breakdown (Rows 25 to 32)
    # -------------------------------------------------------------------------
    ws_sum.cell(row=24, column=2, value="3. BREAKDOWN BY MANUFACTURING PLANT & LOCATION").font = FONT_SEC_TITLE
    
    headers_t4 = ['Manufacturing Plant / Site', 'In-Scope Incidents', 'SLA Met (Passed)', 'SLA Breached (Failed)', 'SLA Compliance (%)']
    for col_idx, h_text in enumerate(headers_t4, start=2):
        c = ws_sum.cell(row=25, column=col_idx)
        c.value = h_text
        c.font = FONT_TBL_HEADER
        c.fill = FILL_SOFT_HEADER
        c.alignment = Alignment(horizontal='center' if col_idx > 2 else 'left', vertical='center')
        c.border = CELL_BORDER
    ws_sum.row_dimensions[25].height = 24
    
    start_r = 26
    for idx, (p_label, p_code) in enumerate(PLANT_MAPPINGS):
        curr_r = start_r + idx
        ws_sum.cell(row=curr_r, column=2, value=p_label)
        ws_sum.cell(row=curr_r, column=3, value=f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AL$2:$AL${max_row}, "{p_code}")')
        ws_sum.cell(row=curr_r, column=4, value=f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AL$2:$AL${max_row}, "{p_code}", \'Year-2026\'!$AJ$2:$AJ${max_row}, "PASSED")')
        ws_sum.cell(row=curr_r, column=5, value=f'=COUNTIFS(\'Year-2026\'!$AH$2:$AH${max_row}, TRUE, \'Year-2026\'!$AL$2:$AL${max_row}, "{p_code}", \'Year-2026\'!$AJ$2:$AJ${max_row}, "FAILED")')
        ws_sum.cell(row=curr_r, column=6, value=f'=IF(C{curr_r}>0, D{curr_r}/C{curr_r}, 0)')
        
    tot_plant_r = start_r + len(PLANT_MAPPINGS)
    ws_sum.cell(row=tot_plant_r, column=2, value="Total Plants & Sites")
    ws_sum.cell(row=tot_plant_r, column=3, value=f"=SUM(C{start_r}:C{tot_plant_r-1})")
    ws_sum.cell(row=tot_plant_r, column=4, value=f"=SUM(D{start_r}:D{tot_plant_r-1})")
    ws_sum.cell(row=tot_plant_r, column=5, value=f"=SUM(E{start_r}:E{tot_plant_r-1})")
    ws_sum.cell(row=tot_plant_r, column=6, value=f"=IF(C{tot_plant_r}>0, D{tot_plant_r}/C{tot_plant_r}, 0)")
    
    for r in range(start_r, tot_plant_r + 1):
        ws_sum.row_dimensions[r].height = 20
        is_tot = (r == tot_plant_r)
        for col_idx in range(2, 7):
            cell = ws_sum.cell(row=r, column=col_idx)
            cell.font = FONT_DATA_BOLD if is_tot else FONT_DATA
            cell.border = TOTAL_BORDER if is_tot else CELL_BORDER
            if is_tot:
                cell.fill = FILL_TOTAL
            elif r % 2 == 1:
                cell.fill = FILL_ZEBRA
                
            if col_idx == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif col_idx in [3, 4, 5]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0'
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '0.0%'

    # Set Column Widths for Summary
    ws_sum.column_dimensions['A'].width = 4
    ws_sum.column_dimensions['B'].width = 44
    ws_sum.column_dimensions['C'].width = 22
    ws_sum.column_dimensions['D'].width = 22
    ws_sum.column_dimensions['E'].width = 22
    ws_sum.column_dimensions['F'].width = 22
    ws_sum.column_dimensions['G'].width = 22
    while True:
        try:
            ensure_file_writable(output_path)
            wb.save(output_path)
            wb.close()
            break
        except (PermissionError, OSError):
            ensure_file_writable(output_path)
            
    return in_scope_count, max_row
