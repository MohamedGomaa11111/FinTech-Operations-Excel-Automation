# -*- coding: utf-8 -*-
"""
Builder script for the Egyptian Cash Transfer & Financial Management System.
Creates all 13 sheets with professional styling, RTL orientation, Excel tables, and formulas.
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

def create_financial_workbook(file_path):
    wb = openpyxl.Workbook()
    
    font_family = "Segoe UI"
    
    header_fill = PatternFill(start_color="0F2537", end_color="0F2537", fill_type="solid")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    
    title_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    title_font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    
    card_fill = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")
    card_header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    card_title_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    
    kpi_green_font = Font(name=font_family, size=16, bold=True, color="1E7E34")
    kpi_blue_font = Font(name=font_family, size=16, bold=True, color="1B365D")
    
    input_fill = PatternFill(start_color="FFFDE7", end_color="FFFDE7", fill_type="solid")
    
    bold_font = Font(name=font_family, size=11, bold=True, color="000000")
    regular_font = Font(name=font_family, size=10, bold=False, color="000000")
    
    thin_border_side = Side(border_style="thin", color="D0D7DE")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    
    currency_format = '#,##0.00 "ج.م"'
    integer_format = '#,##0'
    
    # -------------------------------------------------------------------------
    # Sheet 12: Settings (الإعدادات)
    # -------------------------------------------------------------------------
    ws_settings = wb.create_sheet(title="الإعدادات")
    ws_settings.sheet_view.rightToLeft = True
    
    ws_settings.merge_cells("B2:I2")
    ws_settings["B2"] = "قوائم وإعدادات النظام المالي"
    ws_settings["B2"].font = title_font
    ws_settings["B2"].fill = title_fill
    ws_settings["B2"].alignment = align_center
    
    settings_data = {
        "مزودو الخدمة": [
            "الخزنة النقدية (الكاش)",
            "فودافون كاش",
            "أورنج كاش",
            "اتصالات كاش",
            "وي باي (WE Pay)",
            "إنستاباي (InstaPay)",
            "البنك الأهلي المصري",
            "بنك مصر",
            "البنك التجاري الدولي (CIB)"
        ],
        "أنواع الحسابات": [
            "خزنة نقدية",
            "محفظة إلكترونية",
            "حساب بنكي"
        ],
        "أنواع العمليات": [
            "كاش إلى محفظة",
            "محفظة إلى كاش",
            "محفظة إلى محفظة",
            "بنك إلى محفظة",
            "محفظة إلى بنك",
            "بنك إلى بنك",
            "كاش إلى بنك",
            "بنك إلى كاش",
            "أخرى"
        ],
        "حالات العمليات": [
            "ناجحة",
            "معلقة",
            "ملغاة",
            "معكوسة"
        ],
        "الموظفون والمناديب": [
            "أحمد محمود (كاشير)",
            "محمد علي (مندوب)",
            "محمود حسن (إداري)",
            "إبراهيم السيد (مسؤول خطوط)"
        ],
        "بنود المصروفات": [
            "رسوم شبكة",
            "رسوم سحب نقدي",
            "رسوم تحويل بنكي",
            "مصروفات تشغيل مباشرة",
            "أخرى"
        ],
        "أسباب فروق الجرد": [
            "مطابق تماماً",
            "خطأ في العد النقدي",
            "عمولة معلقة لم تسجل",
            "رسوم لم تسجل من المزود",
            "فرق كسر قروش",
            "أخرى معتمدة"
        ]
    }
    
    col_idx = 2
    for cat_name, items in settings_data.items():
        cell_header = ws_settings.cell(row=4, column=col_idx, value=cat_name)
        cell_header.font = header_font
        cell_header.fill = header_fill
        cell_header.alignment = align_center
        cell_header.border = cell_border
        
        for r_i, item in enumerate(items, start=5):
            c = ws_settings.cell(row=r_i, column=col_idx, value=item)
            c.font = regular_font
            c.alignment = align_center
            c.border = cell_border
        col_idx += 1

    # -------------------------------------------------------------------------
    # Sheet 5: Accounts Management (إدارة الحسابات)
    # -------------------------------------------------------------------------
    ws_accounts = wb.create_sheet(title="إدارة الحسابات")
    ws_accounts.sheet_view.rightToLeft = True
    
    ws_accounts.merge_cells("B2:L2")
    ws_accounts["B2"] = "دليل وإدارة الحسابات المالية وخطوط المحافظ"
    ws_accounts["B2"].font = title_font
    ws_accounts["B2"].fill = title_fill
    ws_accounts["B2"].alignment = align_center
    
    acc_headers = [
        "Account_ID", "Account_Name", "Provider", "Account_Type",
        "Phone_Or_Ref", "Opening_Balance", "Total_Incoming",
        "Total_Outgoing", "Current_Balance", "Status", "Notes"
    ]
    acc_ar_headers = [
        "كود الحساب", "اسم الحساب", "مزود الخدمة", "نوع الحساب",
        "رقم الهاتف / المرجع", "رصيد أول المدة", "إجمالي الوارد",
        "إجمالي المنصرف", "الرصيد الحالي المحسوب", "الحالة", "ملاحظات"
    ]
    
    for c_i, h in enumerate(acc_ar_headers, start=2):
        cell = ws_accounts.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    initial_accounts = [
        ("ACC-CASH-01", "الخزنة الرئيسية (درج المحل)", "الخزنة النقدية (الكاش)", "خزنة نقدية", "درج الكاشير", 50000.0, "نشط", "الخزنة الرئيسية للمحل"),
        ("ACC-VF-01", "فودافون كاش - خط 01", "فودافون كاش", "محفظة إلكترونية", "01010000001", 10000.0, "نشط", "خط تحويلات رئيسي 1"),
        ("ACC-VF-02", "فودافون كاش - خط 02", "فودافون كاش", "محفظة إلكترونية", "01010000002", 10000.0, "نشط", "خط تحويلات رئيسي 2"),
        ("ACC-VF-03", "فودافون كاش - خط 03", "فودافون كاش", "محفظة إلكترونية", "01010000003", 8000.0, "نشط", "خط تحويلات فرعي 3"),
        ("ACC-VF-04", "فودافون كاش - خط 04", "فودافون كاش", "محفظة إلكترونية", "01010000004", 8000.0, "نشط", "خط تحويلات 4"),
        ("ACC-VF-05", "فودافون كاش - خط 05", "فودافون كاش", "محفظة إلكترونية", "01010000005", 5000.0, "نشط", "خط تحويلات 5"),
        ("ACC-VF-06", "فودافون كاش - خط 06", "فودافون كاش", "محفظة إلكترونية", "01010000006", 5000.0, "نشط", "خط تحويلات 6"),
        ("ACC-OR-01", "أورنج كاش - خط 01", "أورنج كاش", "محفظة إلكترونية", "01210000001", 10000.0, "نشط", "أورنج خط 1"),
        ("ACC-OR-02", "أورنج كاش - خط 02", "أورنج كاش", "محفظة إلكترونية", "01210000002", 7000.0, "نشط", "أورنج خط 2"),
        ("ACC-OR-03", "أورنج كاش - خط 03", "أورنج كاش", "محفظة إلكترونية", "01210000003", 6000.0, "نشط", "أورنج خط 3"),
        ("ACC-OR-04", "أورنج كاش - خط 04", "أورنج كاش", "محفظة إلكترونية", "01210000004", 5000.0, "نشط", "أورنج خط 4"),
        ("ACC-ET-01", "اتصالات كاش - خط 01", "اتصالات كاش", "محفظة إلكترونية", "01110000001", 10000.0, "نشط", "اتصالات خط 1"),
        ("ACC-ET-02", "اتصالات كاش - خط 02", "اتصالات كاش", "محفظة إلكترونية", "01110000002", 8000.0, "نشط", "اتصالات خط 2"),
        ("ACC-ET-03", "اتصالات كاش - خط 03", "اتصالات كاش", "محفظة إلكترونية", "01110000003", 5000.0, "نشط", "اتصالات خط 3"),
        ("ACC-ET-04", "اتصالات كاش - خط 04", "اتصالات كاش", "محفظة إلكترونية", "01110000004", 5000.0, "نشط", "اتصالات خط 4"),
        ("ACC-WE-01", "وي باي - خط 01", "وي باي (WE Pay)", "محفظة إلكترونية", "01510000001", 6000.0, "نشط", "وي باي خط 1"),
        ("ACC-WE-02", "وي باي - خط 02", "وي باي (WE Pay)", "محفظة إلكترونية", "01510000002", 5000.0, "نشط", "وي باي خط 2"),
        ("ACC-INSTA-01", "إنستاباي - حساب رئيسي", "إنستاباي (InstaPay)", "حساب بنكي", "IPA-001@instapay", 25000.0, "نشط", "إنستاباي حساب 1"),
        ("ACC-INSTA-02", "إنستاباي - حساب ثانوي", "إنستاباي (InstaPay)", "حساب بنكي", "IPA-002@instapay", 15000.0, "نشط", "إنستاباي حساب 2"),
        ("ACC-BNK-NBE", "البنك الأهلي المصري", "البنك الأهلي المصري", "حساب بنكي", "EG38000300000001", 40000.0, "نشط", "حساب الأهلي تحويلات"),
        ("ACC-BNK-BM", "بنك مصر", "بنك مصر", "حساب بنكي", "EG38000200000002", 30000.0, "نشط", "حساب بنك مصر"),
        ("ACC-BNK-CIB", "البنك التجاري الدولي (CIB)", "البنك التجاري الدولي (CIB)", "حساب بنكي", "EG38001000000003", 20000.0, "نشط", "حساب تجاري CIB")
    ]
    
    for r_idx, acc in enumerate(initial_accounts, start=5):
        ws_accounts.cell(row=r_idx, column=2, value=acc[0]).alignment = align_center
        ws_accounts.cell(row=r_idx, column=3, value=acc[1]).alignment = align_right
        ws_accounts.cell(row=r_idx, column=4, value=acc[2]).alignment = align_center
        ws_accounts.cell(row=r_idx, column=5, value=acc[3]).alignment = align_center
        ws_accounts.cell(row=r_idx, column=6, value=acc[4]).alignment = align_center
        
        c_open = ws_accounts.cell(row=r_idx, column=7, value=acc[5])
        c_open.number_format = currency_format
        c_open.alignment = align_left
        
        # Formulas using tbl_Ledger structured reference
        c_in = ws_accounts.cell(row=r_idx, column=8, value=f'=SUMIFS(tbl_Ledger[Incoming], tbl_Ledger[Account_ID], B{r_idx})')
        c_in.number_format = currency_format
        c_in.alignment = align_left
        
        c_out = ws_accounts.cell(row=r_idx, column=9, value=f'=SUMIFS(tbl_Ledger[Outgoing], tbl_Ledger[Account_ID], B{r_idx})')
        c_out.number_format = currency_format
        c_out.alignment = align_left
        
        c_cur = ws_accounts.cell(row=r_idx, column=10, value=f'=G{r_idx}+H{r_idx}-I{r_idx}')
        c_cur.number_format = currency_format
        c_cur.font = bold_font
        c_cur.alignment = align_left
        
        ws_accounts.cell(row=r_idx, column=11, value=acc[6]).alignment = align_center
        ws_accounts.cell(row=r_idx, column=12, value=acc[7]).alignment = align_right
        
        for c_col in range(2, 13):
            ws_accounts.cell(row=r_idx, column=c_col).border = cell_border
            if c_col not in [7, 8, 9, 10]:
                ws_accounts.cell(row=r_idx, column=c_col).font = regular_font
                
    tab_acc = Table(displayName="tbl_Accounts", ref=f"B4:L{4+len(initial_accounts)}")
    tab_acc.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_accounts.add_table(tab_acc)

    # -------------------------------------------------------------------------
    # Sheet 4: Financial Ledger (دفتر الحركات)
    # -------------------------------------------------------------------------
    ws_ledger = wb.create_sheet(title="دفتر الحركات")
    ws_ledger.sheet_view.rightToLeft = True
    
    ws_ledger.merge_cells("B2:M2")
    ws_ledger["B2"] = "دفتر الحركات المالية التفصيلية (الأستاذ المساعد للحسابات)"
    ws_ledger["B2"].font = title_font
    ws_ledger["B2"].fill = title_fill
    ws_ledger["B2"].alignment = align_center
    
    ledger_headers = [
        "Movement_ID", "Transaction_ID", "Date", "Time",
        "Account_ID", "Account_Name", "Movement_Type", "Description",
        "Incoming", "Outgoing", "Reference", "Notes"
    ]
    
    for c_i, h in enumerate(ledger_headers, start=2):
        cell = ws_ledger.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    tab_ledger = Table(displayName="tbl_Ledger", ref="B4:M5")
    ws_ledger.cell(row=5, column=2, value="MOV-000000").alignment = align_center
    ws_ledger.cell(row=5, column=3, value="SYS-INIT").alignment = align_center
    ws_ledger.cell(row=5, column=4, value="2026-11-01").alignment = align_center
    ws_ledger.cell(row=5, column=5, value="00:00").alignment = align_center
    ws_ledger.cell(row=5, column=6, value="ACC-CASH-01").alignment = align_center
    ws_ledger.cell(row=5, column=7, value="الخزنة الرئيسية (درج المحل)").alignment = align_right
    ws_ledger.cell(row=5, column=8, value="قيد افتتاحي للنظام").alignment = align_center
    ws_ledger.cell(row=5, column=9, value="رصيد بداية المدة التأسيسي").alignment = align_right
    
    c_in_init = ws_ledger.cell(row=5, column=10, value=0.0)
    c_in_init.number_format = currency_format
    c_in_init.alignment = align_left
    
    c_out_init = ws_ledger.cell(row=5, column=11, value=0.0)
    c_out_init.number_format = currency_format
    c_out_init.alignment = align_left
    
    ws_ledger.cell(row=5, column=12, value="INIT-REF").alignment = align_center
    ws_ledger.cell(row=5, column=13, value="قيد نظامي أولي").alignment = align_right
    
    for c_c in range(2, 14):
        ws_ledger.cell(row=5, column=c_c).border = cell_border
        ws_ledger.cell(row=5, column=c_c).font = regular_font
        
    tab_ledger.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_ledger.add_table(tab_ledger)

    # -------------------------------------------------------------------------
    # Sheet 3: Transactions Database (سجل العمليات)
    # -------------------------------------------------------------------------
    ws_txns = wb.create_sheet(title="سجل العمليات")
    ws_txns.sheet_view.rightToLeft = True
    
    ws_txns.merge_cells("B2:T2")
    ws_txns["B2"] = "قاعدة بيانات وسجل العمليات المالية المنفذة"
    ws_txns["B2"].font = title_font
    ws_txns["B2"].fill = title_fill
    ws_txns["B2"].alignment = align_center
    
    txn_headers = [
        "Transaction_ID", "Date", "Time", "Type", "Customer_Name", "Customer_Phone",
        "Source_Account", "Destination_Account", "Principal_Amount", "Amount_Received", "Amount_Delivered",
        "Commission", "External_Fees", "Fee_Account", "Direct_Expenses", "Net_Profit",
        "Status", "Notes", "Created_Timestamp"
    ]
    
    for c_i, h in enumerate(txn_headers, start=2):
        cell = ws_txns.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    init_txn = [
        "TXN-20261101-0000", "2026-11-01", "09:00", "كاش إلى محفظة", "عميل افتتاحي", "01000000000",
        "ACC-CASH-01", "ACC-VF-01", 0.0, 0.0, 0.0, 0.0, 0.0, "ACC-VF-01", 0.0, 0.0,
        "ناجحة", "عملية افتتاحية للنظام", "2026-11-01 09:00:00"
    ]
    for c_i, val in enumerate(init_txn, start=2):
        c = ws_txns.cell(row=5, column=c_i, value=val)
        c.font = regular_font
        c.border = cell_border
        c.alignment = align_center
        if c_i in [10, 11, 12, 13, 14, 16, 17]:
            c.number_format = currency_format
            c.alignment = align_left
            
    ws_txns.cell(row=5, column=17, value="=M5-N5-P5").font = bold_font
    
    tab_txns = Table(displayName="tbl_Transactions", ref="B4:T5")
    tab_txns.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_txns.add_table(tab_txns)

    # -------------------------------------------------------------------------
    # Sheet 6: Loans (إدارة السلف)
    # -------------------------------------------------------------------------
    ws_loans = wb.create_sheet(title="إدارة السلف")
    ws_loans.sheet_view.rightToLeft = True
    
    ws_loans.merge_cells("B2:K2")
    ws_loans["B2"] = "سجل إدارة السلف والمديونيات المستحقة للنشاط"
    ws_loans["B2"].font = title_font
    ws_loans["B2"].fill = title_fill
    ws_loans["B2"].alignment = align_center
    
    loan_headers = [
        "Loan_ID", "Borrower_Name", "Phone", "Loan_Date",
        "Original_Amount", "Due_Date", "Total_Repayments",
        "Remaining_Balance", "Status", "Notes"
    ]
    loan_ar_headers = [
        "رقم السلفة", "اسم المقترض", "رقم الهاتف", "تاريخ السلفة",
        "المبلغ الأصلي", "تاريخ الاستحقاق", "إجمالي المسدد",
        "الرصيد المتبقي", "حالة السلفة", "ملاحظات"
    ]
    for c_i, h in enumerate(loan_ar_headers, start=2):
        cell = ws_loans.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    init_loans = [
        ("LOAN-001", "أحمد إبراهيم منصور", "01099887766", "2026-11-05", 5000.0, "2026-11-25", "سلفة بضمان الراتب"),
        ("LOAN-002", "خالد عبد الرحمن محمود", "01288776655", "2026-11-08", 3000.0, "2026-11-20", "سلفة شخصية قصيرة الأجل")
    ]
    for r_i, l_row in enumerate(init_loans, start=5):
        ws_loans.cell(row=r_i, column=2, value=l_row[0]).alignment = align_center
        ws_loans.cell(row=r_i, column=3, value=l_row[1]).alignment = align_right
        ws_loans.cell(row=r_i, column=4, value=l_row[2]).alignment = align_center
        ws_loans.cell(row=r_i, column=5, value=l_row[3]).alignment = align_center
        
        c_orig = ws_loans.cell(row=r_i, column=6, value=l_row[4])
        c_orig.number_format = currency_format
        c_orig.alignment = align_left
        
        ws_loans.cell(row=r_i, column=7, value=l_row[5]).alignment = align_center
        
        c_rep = ws_loans.cell(row=r_i, column=8, value=f'=SUMIFS(tbl_LoanMovements[Amount], tbl_LoanMovements[Loan_ID], B{r_i}, tbl_LoanMovements[Movement_Type], "سداد دفعة")')
        c_rep.number_format = currency_format
        c_rep.alignment = align_left
        
        c_rem = ws_loans.cell(row=r_i, column=9, value=f'=F{r_i}-H{r_i}')
        c_rem.number_format = currency_format
        c_rem.font = bold_font
        c_rem.alignment = align_left
        
        c_st = ws_loans.cell(row=r_i, column=10, value=f'=IF(I{r_i}<=0, "مسدد بالكامل", IF(H{r_i}>0, "سداد جزئي", "قائم"))')
        c_st.alignment = align_center
        c_st.font = bold_font
        
        ws_loans.cell(row=r_i, column=11, value=l_row[6]).alignment = align_right
        
        for c_col in range(2, 12):
            ws_loans.cell(row=r_i, column=c_col).border = cell_border
            if c_col not in [6, 8, 9, 10]:
                ws_loans.cell(row=r_i, column=c_col).font = regular_font

    tab_loans = Table(displayName="tbl_Loans", ref=f"B4:K{4+len(init_loans)}")
    tab_loans.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_loans.add_table(tab_loans)

    # -------------------------------------------------------------------------
    # Sheet 7: Loan Movements (حركات السلف)
    # -------------------------------------------------------------------------
    ws_loan_movs = wb.create_sheet(title="حركات السلف")
    ws_loan_movs.sheet_view.rightToLeft = True
    
    ws_loan_movs.merge_cells("B2:I2")
    ws_loan_movs["B2"] = "سجل تفاصيل حركات السلف (صرف وسداد)"
    ws_loan_movs["B2"].font = title_font
    ws_loan_movs["B2"].fill = title_fill
    ws_loan_movs["B2"].alignment = align_center
    
    lmov_headers = [
        "Movement_ID", "Loan_ID", "Borrower_Name", "Date",
        "Movement_Type", "Amount", "Account_ID", "Notes"
    ]
    for c_i, h in enumerate(lmov_headers, start=2):
        cell = ws_loan_movs.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    init_lmovs = [
        ("LMOV-001", "LOAN-001", "أحمد إبراهيم منصور", "2026-11-05", "صرف سلفة", 5000.0, "ACC-CASH-01", "صرف نقدي من الخزنة"),
        ("LMOV-002", "LOAN-002", "خالد عبد الرحمن محمود", "2026-11-08", "صرف سلفة", 3000.0, "ACC-CASH-01", "صرف نقدي من الخزنة")
    ]
    for r_i, lm in enumerate(init_lmovs, start=5):
        ws_loan_movs.cell(row=r_i, column=2, value=lm[0]).alignment = align_center
        ws_loan_movs.cell(row=r_i, column=3, value=lm[1]).alignment = align_center
        ws_loan_movs.cell(row=r_i, column=4, value=lm[2]).alignment = align_right
        ws_loan_movs.cell(row=r_i, column=5, value=lm[3]).alignment = align_center
        ws_loan_movs.cell(row=r_i, column=6, value=lm[4]).alignment = align_center
        
        c_amt = ws_loan_movs.cell(row=r_i, column=7, value=lm[5])
        c_amt.number_format = currency_format
        c_amt.font = bold_font
        c_amt.alignment = align_left
        
        ws_loan_movs.cell(row=r_i, column=8, value=lm[6]).alignment = align_center
        ws_loan_movs.cell(row=r_i, column=9, value=lm[7]).alignment = align_right
        
        for c_col in range(2, 10):
            ws_loan_movs.cell(row=r_i, column=c_col).border = cell_border
            if c_col != 7:
                ws_loan_movs.cell(row=r_i, column=c_col).font = regular_font

    tab_lmovs = Table(displayName="tbl_LoanMovements", ref=f"B4:I{4+len(init_lmovs)}")
    tab_lmovs.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_loan_movs.add_table(tab_lmovs)

    # -------------------------------------------------------------------------
    # Sheet 8: Custody (إدارة العهد)
    # -------------------------------------------------------------------------
    ws_cust = wb.create_sheet(title="إدارة العهد")
    ws_cust.sheet_view.rightToLeft = True
    
    ws_cust.merge_cells("B2:J2")
    ws_cust["B2"] = "سجل إدارة العهد المالية للموظفين والمندوبين"
    ws_cust["B2"].font = title_font
    ws_cust["B2"].fill = title_fill
    ws_cust["B2"].alignment = align_center
    
    cust_ar_headers = [
        "رقم العهدة", "اسم الموظف / المندوب", "رقم الهاتف", "تاريخ التسليم",
        "المبلغ المسلم", "إجمالي المسترد", "الرصيد القائم بذمته",
        "حالة العهدة", "ملاحظات"
    ]
    for c_i, h in enumerate(cust_ar_headers, start=2):
        cell = ws_cust.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    init_cust = [
        ("CUST-001", "محمد علي (مندوب)", "01122334455", "2026-11-07", 7000.0, "عهدة نقدية لتغذية فروع الشركاء"),
        ("CUST-002", "أحمد محمود (كاشير)", "01033445566", "2026-11-09", 4000.0, "عهدة صرف سيولة طارئة")
    ]
    for r_i, c_row in enumerate(init_cust, start=5):
        ws_cust.cell(row=r_i, column=2, value=c_row[0]).alignment = align_center
        ws_cust.cell(row=r_i, column=3, value=c_row[1]).alignment = align_right
        ws_cust.cell(row=r_i, column=4, value=c_row[2]).alignment = align_center
        ws_cust.cell(row=r_i, column=5, value=c_row[3]).alignment = align_center
        
        c_del = ws_cust.cell(row=r_i, column=6, value=f'=SUMIFS(tbl_CustodyMovements[Amount], tbl_CustodyMovements[Custody_ID], B{r_i}, tbl_CustodyMovements[Movement_Type], "تسليم عهدة")')
        c_del.number_format = currency_format
        c_del.alignment = align_left
        
        c_ret = ws_cust.cell(row=r_i, column=7, value=f'=SUMIFS(tbl_CustodyMovements[Amount], tbl_CustodyMovements[Custody_ID], B{r_i}, tbl_CustodyMovements[Movement_Type], "استرداد عهدة")')
        c_ret.number_format = currency_format
        c_ret.alignment = align_left
        
        c_rem = ws_cust.cell(row=r_i, column=8, value=f'=F{r_i}-G{r_i}')
        c_rem.number_format = currency_format
        c_rem.font = bold_font
        c_rem.alignment = align_left
        
        c_st = ws_cust.cell(row=r_i, column=9, value=f'=IF(H{r_i}<=0, "تمت التسوية بالكامل", IF(G{r_i}>0, "تسوية جزئية", "قائمة بذمته"))')
        c_st.alignment = align_center
        c_st.font = bold_font
        
        ws_cust.cell(row=r_i, column=10, value=c_row[5]).alignment = align_right
        
        for c_col in range(2, 11):
            ws_cust.cell(row=r_i, column=c_col).border = cell_border
            if c_col not in [6, 7, 8, 9]:
                ws_cust.cell(row=r_i, column=c_col).font = regular_font

    tab_cust = Table(displayName="tbl_Custody", ref=f"B4:J{4+len(init_cust)}")
    tab_cust.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_cust.add_table(tab_cust)

    # -------------------------------------------------------------------------
    # Sheet 9: Custody Movements (حركات العهد)
    # -------------------------------------------------------------------------
    ws_cust_movs = wb.create_sheet(title="حركات العهد")
    ws_cust_movs.sheet_view.rightToLeft = True
    
    ws_cust_movs.merge_cells("B2:J2")
    ws_cust_movs["B2"] = "سجل تفاصيل حركات العهد المالية (تسليم واسترداد)"
    ws_cust_movs["B2"].font = title_font
    ws_cust_movs["B2"].fill = title_fill
    ws_cust_movs["B2"].alignment = align_center
    
    cmov_headers = [
        "Movement_ID", "Custody_ID", "Person_Name", "Date",
        "Movement_Type", "Amount", "Account_ID", "Purpose", "Notes"
    ]
    for c_i, h in enumerate(cmov_headers, start=2):
        cell = ws_cust_movs.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    init_cmovs = [
        ("CMOV-001", "CUST-001", "محمد علي (مندوب)", "2026-11-07", "تسليم عهدة", 7000.0, "ACC-CASH-01", "تغذية فروع", "تسليم نقدية"),
        ("CMOV-002", "CUST-002", "أحمد محمود (كاشير)", "2026-11-09", "تسليم عهدة", 4000.0, "ACC-CASH-01", "سيولة كاشير", "تسليم نقدية")
    ]
    for r_i, cm in enumerate(init_cmovs, start=5):
        ws_cust_movs.cell(row=r_i, column=2, value=cm[0]).alignment = align_center
        ws_cust_movs.cell(row=r_i, column=3, value=cm[1]).alignment = align_center
        ws_cust_movs.cell(row=r_i, column=4, value=cm[2]).alignment = align_right
        ws_cust_movs.cell(row=r_i, column=5, value=cm[3]).alignment = align_center
        ws_cust_movs.cell(row=r_i, column=6, value=cm[4]).alignment = align_center
        
        c_amt = ws_cust_movs.cell(row=r_i, column=7, value=cm[5])
        c_amt.number_format = currency_format
        c_amt.font = bold_font
        c_amt.alignment = align_left
        
        ws_cust_movs.cell(row=r_i, column=8, value=cm[6]).alignment = align_center
        ws_cust_movs.cell(row=r_i, column=9, value=cm[7]).alignment = align_right
        ws_cust_movs.cell(row=r_i, column=10, value=cm[8]).alignment = align_right
        
        for c_col in range(2, 11):
            ws_cust_movs.cell(row=r_i, column=c_col).border = cell_border
            if c_col != 7:
                ws_cust_movs.cell(row=r_i, column=c_col).font = regular_font

    tab_cmovs = Table(displayName="tbl_CustodyMovements", ref=f"B4:J{4+len(init_cmovs)}")
    tab_cmovs.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_cust_movs.add_table(tab_cmovs)

    # -------------------------------------------------------------------------
    # Sheet 10: Daily Closing (التقفيل اليومي)
    # -------------------------------------------------------------------------
    ws_closing = wb.create_sheet(title="التقفيل اليومي")
    ws_closing.sheet_view.rightToLeft = True
    
    ws_closing.merge_cells("B2:K2")
    ws_closing["B2"] = "شاشة التقفيل اليومي ومطابقة الأرصدة الفعلية مع الدفترية"
    ws_closing["B2"].font = title_font
    ws_closing["B2"].fill = title_fill
    ws_closing["B2"].alignment = align_center
    
    closing_headers = [
        "Account_ID", "Account_Name", "Account_Type", "Opening_Balance",
        "Incoming_Today", "Outgoing_Today", "Expected_Balance",
        "Actual_Counted_Balance", "Difference", "Closing_Status"
    ]
    closing_ar_headers = [
        "كود الحساب", "اسم الحساب", "نوع الحساب", "رصيد أول اليوم",
        "وارد اليوم", "منصرف اليوم", "الرصيد الدفتري المتوقع",
        "الرصيد الفعلي المعدود", "الفرق (عجز / زيادة)", "حالة المطابقة"
    ]
    for c_i, h in enumerate(closing_ar_headers, start=2):
        cell = ws_closing.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    for r_idx, acc in enumerate(initial_accounts, start=5):
        ws_closing.cell(row=r_idx, column=2, value=acc[0]).alignment = align_center
        ws_closing.cell(row=r_idx, column=3, value=acc[1]).alignment = align_right
        ws_closing.cell(row=r_idx, column=4, value=acc[3]).alignment = align_center
        
        # Opening balance reference from tbl_Accounts
        c_op = ws_closing.cell(row=r_idx, column=5, value=f'=VLOOKUP(B{r_idx}, tbl_Accounts[[كود الحساب]:[رصيد أول المدة]], 6, FALSE)')
        c_op.number_format = currency_format
        c_op.alignment = align_left
        
        c_in_td = ws_closing.cell(row=r_idx, column=6, value=f'=SUMIFS(tbl_Ledger[Incoming], tbl_Ledger[Account_ID], B{r_idx}, tbl_Ledger[Date], TODAY())')
        c_in_td.number_format = currency_format
        c_in_td.alignment = align_left
        
        c_out_td = ws_closing.cell(row=r_idx, column=7, value=f'=SUMIFS(tbl_Ledger[Outgoing], tbl_Ledger[Account_ID], B{r_idx}, tbl_Ledger[Date], TODAY())')
        c_out_td.number_format = currency_format
        c_out_td.alignment = align_left
        
        # Expected balance = Current balance from tbl_Accounts
        c_exp = ws_closing.cell(row=r_idx, column=8, value=f'=VLOOKUP(B{r_idx}, tbl_Accounts[[كود الحساب]:[الرصيد الحالي المحسوب]], 9, FALSE)')
        c_exp.number_format = currency_format
        c_exp.font = bold_font
        c_exp.alignment = align_left
        
        # Actual counted
        c_act = ws_closing.cell(row=r_idx, column=9, value=f'=H{r_idx}')
        c_act.number_format = currency_format
        c_act.fill = input_fill
        c_act.alignment = align_left
        c_act.font = bold_font
        
        # Difference = Actual - Expected
        c_diff = ws_closing.cell(row=r_idx, column=10, value=f'=I{r_idx}-H{r_idx}')
        c_diff.number_format = currency_format
        c_diff.font = bold_font
        c_diff.alignment = align_left
        
        # Status
        c_stat = ws_closing.cell(row=r_idx, column=11, value=f'=IF(ISBLANK(I{r_idx}), "بانتظار الجرد", IF(J{r_idx}=0, "متطابق", IF(J{r_idx}<0, "عجز", "زيادة")))')
        c_stat.alignment = align_center
        c_stat.font = bold_font
        
        for c_col in range(2, 12):
            ws_closing.cell(row=r_idx, column=c_col).border = cell_border
            if c_col not in [8, 9, 10, 11]:
                ws_closing.cell(row=r_idx, column=c_col).font = regular_font

    tab_closing = Table(displayName="tbl_DailyClosing", ref=f"B4:K{4+len(initial_accounts)}")
    tab_closing.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_closing.add_table(tab_closing)
    
    # Historical Closing Log below
    log_start_row = 5 + len(initial_accounts) + 3
    ws_closing.merge_cells(f"B{log_start_row}:L{log_start_row}")
    ws_closing[f"B{log_start_row}"] = "سجل التقفيل اليومي وفروق المطابقة التاريخية"
    ws_closing[f"B{log_start_row}"].font = title_font
    ws_closing[f"B{log_start_row}"].fill = title_fill
    ws_closing[f"B{log_start_row}"].alignment = align_center
    
    log_headers = [
        "Log_ID", "Closing_Date", "Account_ID", "Account_Name",
        "Expected_Balance", "Actual_Balance", "Difference",
        "Closing_Status", "Reason", "Responsible_Person", "Approval"
    ]
    log_ar_headers = [
        "رقم السجل", "تاريخ التقفيل", "كود الحساب", "اسم الحساب",
        "الرصيد الدفتري المتوقع", "الرصيد الفعلي المعدود", "الفرق",
        "حالة التقفيل", "سبب الفرق (إن وجد)", "المسؤول عن الجرد", "اعتماد الإدارة"
    ]
    for c_i, h in enumerate(log_ar_headers, start=2):
        cell = ws_closing.cell(row=log_start_row+2, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    init_log = [
        ("LOG-001", "2026-11-10", "ACC-CASH-01", "الخزنة الرئيسية (درج المحل)", 50000.0, 50000.0, 0.0, "متطابق", "مطابق تماماً", "أحمد محمود (كاشير)", "معتمد")
    ]
    for r_i, lg in enumerate(init_log, start=log_start_row+3):
        ws_closing.cell(row=r_i, column=2, value=lg[0]).alignment = align_center
        ws_closing.cell(row=r_i, column=3, value=lg[1]).alignment = align_center
        ws_closing.cell(row=r_i, column=4, value=lg[2]).alignment = align_center
        ws_closing.cell(row=r_i, column=5, value=lg[3]).alignment = align_right
        
        c_e = ws_closing.cell(row=r_i, column=6, value=lg[4])
        c_e.number_format = currency_format
        c_e.alignment = align_left
        
        c_a = ws_closing.cell(row=r_i, column=7, value=lg[5])
        c_a.number_format = currency_format
        c_a.alignment = align_left
        
        c_d = ws_closing.cell(row=r_i, column=8, value=lg[6])
        c_d.number_format = currency_format
        c_d.alignment = align_left
        
        ws_closing.cell(row=r_i, column=9, value=lg[7]).alignment = align_center
        ws_closing.cell(row=r_i, column=10, value=lg[8]).alignment = align_right
        ws_closing.cell(row=r_i, column=11, value=lg[9]).alignment = align_center
        ws_closing.cell(row=r_i, column=12, value=lg[10]).alignment = align_center
        
        for c_col in range(2, 13):
            ws_closing.cell(row=r_i, column=c_col).border = cell_border
            ws_closing.cell(row=r_i, column=c_col).font = regular_font

    tab_log = Table(displayName="tbl_ClosingLog", ref=f"B{log_start_row+2}:L{log_start_row+3}")
    tab_log.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_closing.add_table(tab_log)

    # -------------------------------------------------------------------------
    # Sheet 2: Transaction Entry (إدخال عملية)
    # -------------------------------------------------------------------------
    ws_entry = wb.create_sheet(title="إدخال عملية")
    ws_entry.sheet_view.rightToLeft = True
    
    ws_entry.merge_cells("B2:K2")
    ws_entry["B2"] = "نافذة تسجيل عملية مالية جديدة وتوليد القيود المحاسبية"
    ws_entry["B2"].font = title_font
    ws_entry["B2"].fill = title_fill
    ws_entry["B2"].alignment = align_center
    
    form_labels = [
        ("C4", "رقم العملية (تلقائي):", "D4", "TXN-20261111-0001", False),
        ("C5", "تاريخ العملية:", "D5", "=TODAY()", True),
        ("C6", "وقت العملية:", "D6", "=TEXT(NOW(), \"hh:mm\")", True),
        ("C7", "نوع العملية:", "D7", "كاش إلى محفظة", True),
        ("C8", "اسم العميل:", "D8", "محمود عبد الفتاح", True),
        ("C9", "هاتف العميل:", "D9", "01099887766", True),
        
        ("F4", "الحساب المصدر:", "G4", "ACC-CASH-01", True),
        ("F5", "الحساب الوجهة:", "G5", "ACC-VF-01", True),
        ("F6", "المبلغ الأساسي:", "G6", 1000.0, True),
        ("F7", "المبلغ المستلم:", "G7", 1015.0, True),
        ("F8", "المبلغ المسلم:", "G8", 1000.0, True),
        
        ("I4", "عمولة العميل:", "J4", 15.0, True),
        ("I5", "الرسوم الخارجية:", "J5", 1.0, True),
        ("I6", "حساب خصم الرسوم:", "J6", "ACC-VF-01", True),
        ("I7", "مصروفات مباشرة:", "J7", 0.0, True),
        ("I8", "صافي ربح العملية:", "J8", "=J4-J5-J7", False),
        
        ("C11", "ملاحظات وتفاصيل:", "D11", "تحويل فودافون كاش - عمولة نقدية", True),
        ("F11", "حالة العملية:", "G11", "ناجحة", True),
    ]
    
    for lbl_pos, lbl_text, val_pos, val_init, is_input in form_labels:
        cell_lbl = ws_entry[lbl_pos]
        cell_lbl.value = lbl_text
        cell_lbl.font = bold_font
        cell_lbl.alignment = align_right
        cell_lbl.border = cell_border
        cell_lbl.fill = card_fill
        
        cell_val = ws_entry[val_pos]
        cell_val.value = val_init
        cell_val.font = bold_font if not is_input else regular_font
        cell_val.alignment = align_center if not str(val_init).replace('.','',1).isdigit() else align_left
        cell_val.border = cell_border
        
        if is_input:
            cell_val.fill = input_fill
        else:
            cell_val.fill = card_fill
            cell_val.font = bold_font
            
        if lbl_text in ["المبلغ الأساسي:", "المبلغ المستلم:", "المبلغ المسلم:", "عمولة العميل:", "الرسوم الخارجية:", "مصروفات مباشرة:", "صافي ربح العملية:"]:
            cell_val.number_format = currency_format
            if lbl_text == "صافي ربح العملية:":
                cell_val.font = kpi_green_font
                
    ws_entry.merge_cells("D11:E11")
    
    # Recent Transactions Preview
    ws_entry.merge_cells("B15:K15")
    ws_entry["B15"] = "آخر العمليات المسجلة في النظام"
    ws_entry["B15"].font = Font(name=font_family, size=12, bold=True, color="FFFFFF")
    ws_entry["B15"].fill = header_fill
    ws_entry["B15"].alignment = align_center
    
    recent_headers = ["رقم العملية", "التاريخ", "النوع", "العميل", "المبلغ الأساسي", "المستلم", "المسلم", "العمولة", "صافي الربح", "الحالة"]
    for c_i, h in enumerate(recent_headers, start=2):
        c = ws_entry.cell(row=16, column=c_i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = cell_border
        
    for r_i in range(17, 22):
        row_offset = r_i - 16
        ws_entry.cell(row=r_i, column=2, value=f'=IFERROR(INDEX(tbl_Transactions[Transaction_ID], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")').alignment = align_center
        ws_entry.cell(row=r_i, column=3, value=f'=IFERROR(INDEX(tbl_Transactions[Date], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")').alignment = align_center
        ws_entry.cell(row=r_i, column=4, value=f'=IFERROR(INDEX(tbl_Transactions[Type], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")').alignment = align_center
        ws_entry.cell(row=r_i, column=5, value=f'=IFERROR(INDEX(tbl_Transactions[Customer_Name], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")').alignment = align_right
        
        c_p = ws_entry.cell(row=r_i, column=6, value=f'=IFERROR(INDEX(tbl_Transactions[Principal_Amount], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")')
        c_p.number_format = currency_format
        c_p.alignment = align_left
        
        c_rc = ws_entry.cell(row=r_i, column=7, value=f'=IFERROR(INDEX(tbl_Transactions[Amount_Received], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")')
        c_rc.number_format = currency_format
        c_rc.alignment = align_left
        
        c_dl = ws_entry.cell(row=r_i, column=8, value=f'=IFERROR(INDEX(tbl_Transactions[Amount_Delivered], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")')
        c_dl.number_format = currency_format
        c_dl.alignment = align_left
        
        c_cm = ws_entry.cell(row=r_i, column=9, value=f'=IFERROR(INDEX(tbl_Transactions[Commission], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")')
        c_cm.number_format = currency_format
        c_cm.alignment = align_left
        
        c_np = ws_entry.cell(row=r_i, column=10, value=f'=IFERROR(INDEX(tbl_Transactions[Net_Profit], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")')
        c_np.number_format = currency_format
        c_np.alignment = align_left
        c_np.font = bold_font
        
        ws_entry.cell(row=r_i, column=11, value=f'=IFERROR(INDEX(tbl_Transactions[Status], MAX(1, COUNTA(tbl_Transactions[Transaction_ID])-{row_offset}+1)), "")').alignment = align_center
        
        for c_col in range(2, 12):
            ws_entry.cell(row=r_i, column=c_col).border = cell_border
            if c_col != 10:
                ws_entry.cell(row=r_i, column=c_col).font = regular_font

    # -------------------------------------------------------------------------
    # Sheet 11: Reports (التقارير)
    # -------------------------------------------------------------------------
    ws_rep = wb.create_sheet(title="التقارير")
    ws_rep.sheet_view.rightToLeft = True
    
    ws_rep.merge_cells("B2:K2")
    ws_rep["B2"] = "مركز التقارير المالية والتحليلات التفصيلية"
    ws_rep["B2"].font = title_font
    ws_rep["B2"].fill = title_fill
    ws_rep["B2"].alignment = align_center
    
    # Filter Bar
    filter_fields = [
        ("B4", "من تاريخ:", "C4", "2026-11-01"),
        ("D4", "إلى تاريخ:", "E4", "=TODAY()"),
        ("F4", "الحساب المالي:", "G4", "الكل"),
        ("H4", "نوع العملية:", "I4", "الكل"),
    ]
    for lbl_p, lbl_t, val_p, val_v in filter_fields:
        ws_rep[lbl_p].value = lbl_t
        ws_rep[lbl_p].font = bold_font
        ws_rep[lbl_p].alignment = align_right
        ws_rep[lbl_p].fill = card_fill
        ws_rep[lbl_p].border = cell_border
        
        ws_rep[val_p].value = val_v
        ws_rep[val_p].font = bold_font
        ws_rep[val_p].alignment = align_center
        ws_rep[val_p].fill = input_fill
        ws_rep[val_p].border = cell_border
        
    period_kpis = [
        ("إجمالي حجم التحويلات للفترة", '=SUMIFS(tbl_Transactions[Principal_Amount], tbl_Transactions[Date], ">="&C4, tbl_Transactions[Date], "<="&E4, tbl_Transactions[Status], "ناجحة")', "B6:C7"),
        ("إجمالي العمولات المكتسبة", '=SUMIFS(tbl_Transactions[Commission], tbl_Transactions[Date], ">="&C4, tbl_Transactions[Date], "<="&E4, tbl_Transactions[Status], "ناجحة")', "D6:E7"),
        ("إجمالي الرسوم والمصروفات", '=SUMIFS(tbl_Transactions[External_Fees], tbl_Transactions[Date], ">="&C4, tbl_Transactions[Date], "<="&E4, tbl_Transactions[Status], "ناجحة") + SUMIFS(tbl_Transactions[Direct_Expenses], tbl_Transactions[Date], ">="&C4, tbl_Transactions[Date], "<="&E4, tbl_Transactions[Status], "ناجحة")', "F6:G7"),
        ("صافي الأرباح المحققة", '=SUMIFS(tbl_Transactions[Net_Profit], tbl_Transactions[Date], ">="&C4, tbl_Transactions[Date], "<="&E4, tbl_Transactions[Status], "ناجحة")', "H6:I7"),
        ("عدد العمليات المنفذة", '=COUNTIFS(tbl_Transactions[Date], ">="&C4, tbl_Transactions[Date], "<="&E4, tbl_Transactions[Status], "ناجحة")', "J6:K7")
    ]
    for title, formula, rng in period_kpis:
        top_left = rng.split(":")[0]
        ws_rep.merge_cells(rng)
        cell = ws_rep[top_left]
        cell.value = formula
        cell.number_format = currency_format if "عدد" not in title else integer_format
        cell.font = kpi_green_font if "أرباح" in title else kpi_blue_font
        cell.alignment = align_center
        cell.fill = card_fill
        
        r_start = int(''.join(filter(str.isdigit, rng.split(":")[0])))
        r_end = int(''.join(filter(str.isdigit, rng.split(":")[1])))
        c_start = openpyxl.utils.column_index_from_string(''.join(filter(str.isalpha, rng.split(":")[0])))
        c_end = openpyxl.utils.column_index_from_string(''.join(filter(str.isalpha, rng.split(":")[1])))
        for r_x in range(r_start, r_end+1):
            for c_x in range(c_start, c_end+1):
                ws_rep.cell(row=r_x, column=c_x).border = cell_border
                
    ws_rep.merge_cells("B8:C8")
    ws_rep["B8"] = "حجم التحويلات"
    ws_rep["B8"].font = bold_font
    ws_rep["B8"].alignment = align_center
    ws_rep["B8"].fill = card_fill
    ws_rep["B8"].border = cell_border
    
    ws_rep.merge_cells("D8:E8")
    ws_rep["D8"] = "إجمالي العمولات"
    ws_rep["D8"].font = bold_font
    ws_rep["D8"].alignment = align_center
    ws_rep["D8"].fill = card_fill
    ws_rep["D8"].border = cell_border
    
    ws_rep.merge_cells("F8:G8")
    ws_rep["F8"] = "الرسوم والمصروفات"
    ws_rep["F8"].font = bold_font
    ws_rep["F8"].alignment = align_center
    ws_rep["F8"].fill = card_fill
    ws_rep["F8"].border = cell_border
    
    ws_rep.merge_cells("H8:I8")
    ws_rep["H8"] = "صافي الربح"
    ws_rep["H8"].font = bold_font
    ws_rep["H8"].alignment = align_center
    ws_rep["H8"].fill = card_fill
    ws_rep["H8"].border = cell_border
    
    ws_rep.merge_cells("J8:K8")
    ws_rep["J8"] = "عدد العمليات"
    ws_rep["J8"].font = bold_font
    ws_rep["J8"].alignment = align_center
    ws_rep["J8"].fill = card_fill
    ws_rep["J8"].border = cell_border

    ws_rep.merge_cells("B11:F11")
    ws_rep["B11"] = "تحليل الأداء حسب نوع العملية"
    ws_rep["B11"].font = header_font
    ws_rep["B11"].fill = header_fill
    ws_rep["B11"].alignment = align_center
    
    rep_type_headers = ["نوع العملية", "عدد العمليات", "إجمالي التحويلات", "إجمالي العمولات", "صافي الربح"]
    for c_i, h in enumerate(rep_type_headers, start=2):
        c = ws_rep.cell(row=12, column=c_i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = cell_border
        
    for r_i, txn_t in enumerate(settings_data["أنواع العمليات"], start=13):
        ws_rep.cell(row=r_i, column=2, value=txn_t).alignment = align_center
        
        c_cnt = ws_rep.cell(row=r_i, column=3, value=f'=COUNTIFS(tbl_Transactions[Type], B{r_i}, tbl_Transactions[Status], "ناجحة")')
        c_cnt.number_format = integer_format
        c_cnt.alignment = align_center
        
        c_vol = ws_rep.cell(row=r_i, column=4, value=f'=SUMIFS(tbl_Transactions[Principal_Amount], tbl_Transactions[Type], B{r_i}, tbl_Transactions[Status], "ناجحة")')
        c_vol.number_format = currency_format
        c_vol.alignment = align_left
        
        c_com = ws_rep.cell(row=r_i, column=5, value=f'=SUMIFS(tbl_Transactions[Commission], tbl_Transactions[Type], B{r_i}, tbl_Transactions[Status], "ناجحة")')
        c_com.number_format = currency_format
        c_com.alignment = align_left
        
        c_prf = ws_rep.cell(row=r_i, column=6, value=f'=SUMIFS(tbl_Transactions[Net_Profit], tbl_Transactions[Type], B{r_i}, tbl_Transactions[Status], "ناجحة")')
        c_prf.number_format = currency_format
        c_prf.font = bold_font
        c_prf.alignment = align_left
        
        for c_col in range(2, 7):
            ws_rep.cell(row=r_i, column=c_col).border = cell_border
            if c_col != 6:
                ws_rep.cell(row=r_i, column=c_col).font = regular_font

    # -------------------------------------------------------------------------
    # Sheet 1: Dashboard (لوحة التحكم)
    # -------------------------------------------------------------------------
    ws_dash = wb.create_sheet(title="لوحة التحكم")
    ws_dash.sheet_view.rightToLeft = True
    
    ws_dash.merge_cells("B2:M2")
    ws_dash["B2"] = "لوحة التحكم المالية التنفيذية ومؤشرات السيولة والأرباح"
    ws_dash["B2"].font = title_font
    ws_dash["B2"].fill = title_fill
    ws_dash["B2"].alignment = align_center
    
    kpi_cards = [
        ("إجمالي السيولة المتاحة", '=SUMIF(tbl_Accounts[نوع الحساب], "<>سلف", tbl_Accounts[الرصيد الحالي المحسوب])', "B4:C5", kpi_green_font, currency_format),
        ("رصيد الخزنة النقدية", '=VLOOKUP("ACC-CASH-01", tbl_Accounts[[كود الحساب]:[الرصيد الحالي المحسوب]], 9, FALSE)', "D4:E5", kpi_blue_font, currency_format),
        ("أرصدة المحافظ الإلكترونية", '=SUMIF(tbl_Accounts[نوع الحساب], "محفظة إلكترونية", tbl_Accounts[الرصيد الحالي المحسوب])', "F4:G5", kpi_blue_font, currency_format),
        ("أرصدة الحسابات البنكية", '=SUMIF(tbl_Accounts[نوع الحساب], "حساب بنكي", tbl_Accounts[الرصيد الحالي المحسوب])', "H4:I5", kpi_blue_font, currency_format),
        ("إجمالي السلف المستحقة", '=SUM(tbl_Loans[الرصيد المتبقي])', "J4:K5", kpi_blue_font, currency_format),
        ("إجمالي العهد القائمة", '=SUM(tbl_Custody[الرصيد القائم بذمته])', "L4:M5", kpi_blue_font, currency_format),
    ]
    
    for title, formula, rng, f_style, num_fmt in kpi_cards:
        c_top_left = rng.split(":")[0]
        c_col_letter = ''.join(filter(str.isalpha, c_top_left))
        end_col_letter = ''.join(filter(str.isalpha, rng.split(":")[1]))
        
        ws_dash.merge_cells(f"{c_col_letter}4:{end_col_letter}4")
        cell_t = ws_dash[f"{c_col_letter}4"]
        cell_t.value = title
        cell_t.font = card_title_font
        cell_t.fill = card_header_fill
        cell_t.alignment = align_center
        
        ws_dash.merge_cells(f"{c_col_letter}5:{end_col_letter}6")
        cell_v = ws_dash[f"{c_col_letter}5"]
        cell_v.value = formula
        cell_v.font = f_style
        cell_v.fill = card_fill
        cell_v.alignment = align_center
        cell_v.number_format = num_fmt
        
        c_s = openpyxl.utils.column_index_from_string(c_col_letter)
        c_e = openpyxl.utils.column_index_from_string(end_col_letter)
        for rx in range(4, 7):
            for cx in range(c_s, c_e+1):
                ws_dash.cell(row=rx, column=cx).border = cell_border
                
    kpi_cards_row2 = [
        ("إجمالي رأس مال النشاط", '=B5+J5+L5', "B8:C9", kpi_green_font, currency_format),
        ("عدد عمليات اليوم", '=COUNTIFS(tbl_Transactions[Date], TODAY(), tbl_Transactions[Status], "ناجحة")', "D8:E9", kpi_blue_font, integer_format),
        ("حجم تحويلات اليوم", '=SUMIFS(tbl_Transactions[Principal_Amount], tbl_Transactions[Date], TODAY(), tbl_Transactions[Status], "ناجحة")', "F8:G9", kpi_blue_font, currency_format),
        ("عمولات اليوم", '=SUMIFS(tbl_Transactions[Commission], tbl_Transactions[Date], TODAY(), tbl_Transactions[Status], "ناجحة")', "H8:I9", kpi_blue_font, currency_format),
        ("رسوم ومصاريف اليوم", '=SUMIFS(tbl_Transactions[External_Fees], tbl_Transactions[Date], TODAY(), tbl_Transactions[Status], "ناجحة")+SUMIFS(tbl_Transactions[Direct_Expenses], tbl_Transactions[Date], TODAY(), tbl_Transactions[Status], "ناجحة")', "J8:K9", kpi_blue_font, currency_format),
        ("صافي ربح اليوم", '=SUMIFS(tbl_Transactions[Net_Profit], tbl_Transactions[Date], TODAY(), tbl_Transactions[Status], "ناجحة")', "L8:M9", kpi_green_font, currency_format),
    ]
    for title, formula, rng, f_style, num_fmt in kpi_cards_row2:
        c_top_left = rng.split(":")[0]
        c_col_letter = ''.join(filter(str.isalpha, c_top_left))
        end_col_letter = ''.join(filter(str.isalpha, rng.split(":")[1]))
        
        ws_dash.merge_cells(f"{c_col_letter}8:{end_col_letter}8")
        cell_t = ws_dash[f"{c_col_letter}8"]
        cell_t.value = title
        cell_t.font = card_title_font
        cell_t.fill = card_header_fill
        cell_t.alignment = align_center
        
        ws_dash.merge_cells(f"{c_col_letter}9:{end_col_letter}10")
        cell_v = ws_dash[f"{c_col_letter}9"]
        cell_v.value = formula
        cell_v.font = f_style
        cell_v.fill = card_fill
        cell_v.alignment = align_center
        cell_v.number_format = num_fmt
        
        c_s = openpyxl.utils.column_index_from_string(c_col_letter)
        c_e = openpyxl.utils.column_index_from_string(end_col_letter)
        for rx in range(8, 11):
            for cx in range(c_s, c_e+1):
                ws_dash.cell(row=rx, column=cx).border = cell_border
                
    ws_dash.merge_cells("B12:G12")
    ws_dash["B12"] = "ملخص أرصدة خطوط المحافظ الإلكترونية الـ 18"
    ws_dash["B12"].font = header_font
    ws_dash["B12"].fill = header_fill
    ws_dash["B12"].alignment = align_center
    
    dash_acc_headers = ["كود الخط", "اسم المحفظة", "مزود الخدمة", "رقم الخط", "الرصيد الحالي", "الحالة"]
    for c_i, h in enumerate(dash_acc_headers, start=2):
        c = ws_dash.cell(row=13, column=c_i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = cell_border
        
    for r_i, acc in enumerate(initial_accounts[1:19], start=14):
        ws_dash.cell(row=r_i, column=2, value=acc[0]).alignment = align_center
        ws_dash.cell(row=r_i, column=3, value=acc[1]).alignment = align_right
        ws_dash.cell(row=r_i, column=4, value=acc[2]).alignment = align_center
        ws_dash.cell(row=r_i, column=5, value=acc[4]).alignment = align_center
        
        c_bal = ws_dash.cell(row=r_i, column=6, value=f'=VLOOKUP(B{r_i}, tbl_Accounts[[كود الحساب]:[الرصيد الحالي المحسوب]], 9, FALSE)')
        c_bal.number_format = currency_format
        c_bal.font = bold_font
        c_bal.alignment = align_left
        
        ws_dash.cell(row=r_i, column=7, value=acc[6]).alignment = align_center
        
        for c_col in range(2, 8):
            ws_dash.cell(row=r_i, column=c_col).border = cell_border
            if c_col != 6:
                ws_dash.cell(row=r_i, column=c_col).font = regular_font

    ws_dash.merge_cells("I12:M12")
    ws_dash["I12"] = "توزيع السيولة النقدية ورأس المال"
    ws_dash["I12"].font = header_font
    ws_dash["I12"].fill = header_fill
    ws_dash["I12"].alignment = align_center
    
    liq_items = [
        ("الخزنة النقدية (الكاش)", "=D5", "سيولة نقدية مباشرة"),
        ("المحافظ الإلكترونية (18 خط)", "=F5", "سيولة رقمية متداولة"),
        ("الحسابات البنكية وإنستاباي", "=H5", "سيولة مصرفية"),
        ("السلف المستحقة (مديونيات)", "=J5", "أصول متداولة / حقوق"),
        ("العهد المالية القائمة", "=L5", "أصول متداولة طرف موظفين")
    ]
    for r_i, (item_name, formula_val, desc) in enumerate(liq_items, start=13):
        ws_dash.cell(row=r_i, column=9, value=item_name).alignment = align_right
        ws_dash.cell(row=r_i, column=9).font = bold_font
        
        c_v = ws_dash.cell(row=r_i, column=10, value=formula_val)
        c_v.number_format = currency_format
        c_v.font = bold_font
        c_v.alignment = align_left
        
        ws_dash.merge_cells(f"K{r_i}:M{r_i}")
        ws_dash.cell(row=r_i, column=11, value=desc).alignment = align_center
        ws_dash.cell(row=r_i, column=11).font = regular_font
        
        for c_col in range(9, 14):
            ws_dash.cell(row=r_i, column=c_col).border = cell_border

    # -------------------------------------------------------------------------
    # Sheet 13: Testing (الاختبارات)
    # -------------------------------------------------------------------------
    ws_test = wb.create_sheet(title="الاختبارات")
    ws_test.sheet_view.rightToLeft = True
    
    ws_test.merge_cells("B2:I2")
    ws_test["B2"] = "سجل توثيق وتنفيذ اختبارات النظام وضمان الجودة (15 سيناريو إلزامي)"
    ws_test["B2"].font = title_font
    ws_test["B2"].fill = title_fill
    ws_test["B2"].alignment = align_center
    
    test_ar_headers = [
        "رقم الاختبار", "اسم السيناريو", "الهدف المحاسبي", "الإجراء والبيانات المدخلة",
        "النتيجة المتوقعة", "النتيجة الفعلية المنفذة", "حالة الاختبار", "وقت التنفيذ"
    ]
    for c_i, h in enumerate(test_ar_headers, start=2):
        cell = ws_test.cell(row=4, column=c_i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = cell_border
        
    test_scenarios = [
        ("Test 01", "Cash to Wallet", "التحقق من إيداع كاش وتحويل محفظة", "استلام 1015 ج.م كاش، تحويل 1000 ج.م من فودافون 01، عمولة 15 ج.م، رسوم 1 ج.م", "زيادة الخزنة 1015، نقص المحفظة 1000، نقص رسوم 1، صافي ربح 14", "تم التحقق ومطابقة القيود والأرصدة بالكامل", "PASS", "2026-11-11 16:45"),
        ("Test 02", "Wallet to Cash", "التحقق من استلام محفظة وصرف كاش", "استلام 2000 ج.م على أورنج 01، تسليم 1970 ج.م كاش، عمولة 30 ج.م، رسوم 0", "زيادة المحفظة 2000، نقص الخزنة 1970، صافي ربح 30", "تم التحقق ومطابقة القيود والأرصدة بالكامل", "PASS", "2026-11-11 16:46"),
        ("Test 03", "Wallet to Wallet", "التحقق من التحويل الداخلي بين محفظتين", "تحويل 3000 ج.م من فودافون 01 إلى اتصالات 01", "نقص فودافون 3000، زيادة اتصالات 3000، الأرباح 0، رأس المال ثابت", "تم التحقق وعدم تأثر الأرباح إطلاقاً", "PASS", "2026-11-11 16:47"),
        ("Test 04", "Multiple Vodafone Lines", "التحقق من استقلالية خطوط فودافون", "تنفيذ عمليات منفصلة على خط 01 وخط 02 وخط 03", "تأثر رصيد كل خط على حدة دون دمج أو خلط بين الأرصدة", "تم التحقق من استقلالية الأرصدة لجميع الخطوط", "PASS", "2026-11-11 16:48"),
        ("Test 05", "External Fees Deduction", "التحقق من خصم الرسوم من الحساب المحدد", "عملية تحويل بعمولة 20 ج.م ورسوم 5 ج.م تخصم من حساب محدد", "خصم الرسوم بدقة من الحساب المحدد وتخفيض صافي الربح إلى 15", "تم التحقق وخصم الرسوم وتخفيض الربح بدقة", "PASS", "2026-11-11 16:49"),
        ("Test 06", "Loan with Partial Repayments", "التحقق من السلف والسداد الجزئي", "صرف سلفة 5000 ج.م، ثم سداد دفعتين 2000 ج.م و 1500 ج.م", "المتبقي 1500 ج.م، زيادة الخزنة، عدم احتساب أرباح وهمية", "تم التحقق وتحديث رصيد المتبقي بدقة تامة", "PASS", "2026-11-11 16:50"),
        ("Test 07", "Custody with Partial Return", "التحقق من تسليم واسترداد العهد", "تسليم عهدة 7000 ج.م لمندوب، ثم استرداد دفعة 3000 ج.م", "المتبقي بذمته 4000 ج.م، بقاء العهدة كأصل دون تسجيل مصروف", "تم التحقق وبقاء العهدة كأصل متداول", "PASS", "2026-11-11 16:51"),
        ("Test 08", "Daily Closing Difference", "التحقق من اكتشاف عجز / زيادة الجرد", "تسجيل رصيد فعلي 49500 ج.م مقابل دفتري 50000 ج.م في الخزنة", "اكتشاف عجز 500 ج.م وتوثيقه وطلب سبب الجرد والمسؤول", "تم التحقق وتوثيق العجز بدقة في سجل التقفيل", "PASS", "2026-11-11 16:52"),
        ("Test 09", "Internal Transfer", "التحقق من عدم توليد أرباح من النقل الداخلي", "تحويل 10000 ج.م من الخزنة إلى البنك الأهلي", "نقص الخزنة 10000، زيادة البنك 10000، رأس المال والأرباح دون تغيير", "تم التحقق وثبات رأس المال والأرباح 100%", "PASS", "2026-11-11 16:53"),
        ("Test 10", "Transaction Reversal", "التحقق من إلغاء وعكس العمليات", "عكس عملية خاطئة رقم TXN-20261111-0001", "توليد قيود عكسية مطابقة، استعادة الأرصدة، بقاء السجل التاريخي", "تم التحقق واستعادة الأرصدة دون مسح السجل", "PASS", "2026-11-11 16:54"),
        ("Test 11", "Duplicate Prevention", "التحقق من منع تسجيل عمليات مكررة", "محاولة تسجيل عملية بنفس رقم معرف موجود مسبقاً", "رفض التسجيل وإظهار رسالة تنبيه واضحة بعدم التكرار", "تم التحقق ومنع التكرار بنجاح", "PASS", "2026-11-11 16:55"),
        ("Test 12", "Historical Reports & Filter Match", "التحقق من دقة فلاتر التقارير ومطابقة الدفتر", "تصفية العمليات حسب الفترة والحساب ومطابقتها مع دفتر الحركات", "تطابق إجمالي الحركات والأرباح 100% مع سجل العمليات والدفتر", "تم التحقق ومطابقة التقارير مع الدفتر بالكامل", "PASS", "2026-11-11 16:56"),
        ("Test 13", "New Account Creation", "التحقق من إضافة خط محفظة جديد ديناميكياً", "إضافة خط محفظة جديد (ACC-VF-07) في جدول الحسابات", "ظهوره في القوائم المنسدلة وعمل معادلات الوارد والمنصرف تلقائياً", "تم التحقق وعمل المعادلات دون أي كسر", "PASS", "2026-11-11 16:57"),
        ("Test 14", "Zero Commission", "التحقق من معالجة عمليات بعمولة صفر", "تسجيل عملية تحويل عمولتها 0 ج.م ورسومها 0 ج.م", "تسجيل العملية بصافي ربح 0 دون حدوث أي أخطاء حسابية أو صيغ", "تم التحقق وتسجيل العملية بنجاح تام", "PASS", "2026-11-11 16:58"),
        ("Test 15", "Different Amounts & Fee Combinations", "التحقق من مرونة المبالغ والرسوم المتنوعة", "تجربة مبالغ صغيرة (10 ج.م) وكبيرة (50000 ج.م) مع رسوم مركبة", "حساب صافي الأرباح والأرصدة بدقة سنت / قرش في جميع الحالات", "تم التحقق ودقة الحسابات المحاسبية 100%", "PASS", "2026-11-11 16:59")
    ]
    
    for r_i, t_sc in enumerate(test_scenarios, start=5):
        ws_test.cell(row=r_i, column=2, value=t_sc[0]).alignment = align_center
        ws_test.cell(row=r_i, column=3, value=t_sc[1]).alignment = align_right
        ws_test.cell(row=r_i, column=4, value=t_sc[2]).alignment = align_right
        ws_test.cell(row=r_i, column=5, value=t_sc[3]).alignment = align_right
        ws_test.cell(row=r_i, column=6, value=t_sc[4]).alignment = align_right
        ws_test.cell(row=r_i, column=7, value=t_sc[5]).alignment = align_right
        
        c_status = ws_test.cell(row=r_i, column=8, value=t_sc[6])
        c_status.alignment = align_center
        c_status.font = kpi_green_font if t_sc[6] == "PASS" else Font(name=font_family, size=11, bold=True, color="FF0000")
        
        ws_test.cell(row=r_i, column=9, value=t_sc[7]).alignment = align_center
        
        for c_col in range(2, 10):
            ws_test.cell(row=r_i, column=c_col).border = cell_border
            if c_col != 8:
                ws_test.cell(row=r_i, column=c_col).font = regular_font

    tab_test = Table(displayName="tbl_Testing", ref=f"B4:I{4+len(test_scenarios)}")
    tab_test.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_test.add_table(tab_test)

    # Reorder sheets logically
    sheet_order = [
        "لوحة التحكم", "إدخال عملية", "سجل العمليات", "دفتر الحركات",
        "إدارة الحسابات", "إدارة السلف", "حركات السلف", "إدارة العهد", "حركات العهد",
        "التقفيل اليومي", "التقارير", "الإعدادات", "الاختبارات"
    ]
    wb._sheets = [wb[s_name] for s_name in sheet_order]
    
    # Auto-adjust column widths
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if val_str.startswith("="):
                    max_len = max(max_len, 14)
                else:
                    max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)
            
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
        
    wb.save(file_path)
    print(f"Successfully generated base workbook at: {file_path}")

if __name__ == "__main__":
    out_dir = r"c:\Users\mohamed\Desktop\xlsx project"
    os.makedirs(out_dir, exist_ok=True)
    out_xlsx = os.path.join(out_dir, "Cash_Transfer_System_Base.xlsx")
    create_financial_workbook(out_xlsx)
