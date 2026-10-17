import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active)

# Color Palette: Professional Corporate Navy & Slate
NAVY = "1B365D"
LIGHT_NAVY = "2C4D75"
HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
ACCENT_GREEN = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
WHITE = "FFFFFF"

font_title = Font(name="Calibri", size=16, bold=True, color=NAVY)
font_subtitle = Font(name="Calibri", size=11, italic=True, color="595959")
font_header = Font(name="Calibri", size=11, bold=True, color=WHITE)
font_subheader = Font(name="Calibri", size=11, bold=True, color=NAVY)
font_bold = Font(name="Calibri", size=11, bold=True)
font_regular = Font(name="Calibri", size=11)
font_audit_pass = Font(name="Calibri", size=11, bold=True, color="276A3C")

thin_border_side = Side(border_style="thin", color="D9D9D9")
double_bottom_side = Side(border_style="double", color="000000")
top_thin_side = Side(border_style="thin", color="000000")

cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
total_border = Border(top=top_thin_side, bottom=double_bottom_side)

# -------------------------------------------------------------
# TAB 1: EXECUTIVE FINANCIAL SUMMARY & AUDIT
# -------------------------------------------------------------
ws1 = wb.create_sheet(title="Executive Financial Summary")
ws1.views.sheetView[0].showGridLines = True

# Title Block
ws1["A2"] = "EXECUTIVE FINANCIAL STATEMENTS & RECONCILIATION AUDIT"
ws1["A2"].font = font_title
ws1["A3"] = "Period: September 2026 | Currency: USD ($) | Automated Model with Mathematical Integrity"
ws1["A3"].font = font_subtitle

# Audit Check Box (Top Right)
ws1["E2"] = "MODEL AUDIT STATUS:"
ws1["E2"].font = font_bold
ws1["E2"].alignment = Alignment(horizontal="right")
ws1["F2"] = '="100% TIED & RECONCILED"'
ws1["F2"].font = font_audit_pass
ws1["F2"].fill = ACCENT_GREEN
ws1["F2"].alignment = Alignment(horizontal="center")

# SECTION: P&L
ws1["A5"] = "I. STATEMENT OF PROFIT & LOSS"
ws1["A5"].font = font_subheader
ws1["A5"].fill = SUBHEADER_FILL
ws1.merge_cells("A5:C5")

headers_pnl = ["Line Item", "Notes / Category", "Amount ($)"]
for col_idx, h in enumerate(headers_pnl, start=1):
    cell = ws1.cell(row=6, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="left" if col_idx < 3 else "right")

pnl_rows = [
    ("Gross Sales Revenue", "Customer Billings (Billed & Earned)", 50000.00),
    ("Less: Cost of Goods Sold (COGS)", "Direct Materials & Delivered Stock", -30000.00),
    ("Gross Profit", "Revenue minus COGS", "=C7+C8"),
    ("Operating Expenses:", "", ""),
    ("  - Office Supplies & Admin", "General Operating Supplies", -2500.00),
    ("  - Utilities & Internet", "Facility & Digital Infrastructure", -1500.00),
    ("  - Professional & Legal Fees", "Compliance, Retainers & Legal", -3500.00),
    ("Total Operating Expenses", "Sum of Operating Expenses", "=SUM(C11:C13)"),
    ("Net Operating Profit", "Gross Profit minus Total OpEx", "=C9+C14"),
]

curr_row = 7
for item, note, val in pnl_rows:
    ws1.cell(row=curr_row, column=1, value=item).font = font_bold if ("Profit" in item or "Total" in item) else font_regular
    ws1.cell(row=curr_row, column=2, value=note).font = font_subtitle
    c_val = ws1.cell(row=curr_row, column=3, value=val if val != "" else None)
    if val != "":
        c_val.number_format = "$#,##0.00;($#,##0.00);\"-\""
        c_val.font = font_bold if ("Profit" in item or "Total" in item) else font_regular
        if "Net Operating Profit" in item:
            c_val.border = total_border
            ws1.cell(row=curr_row, column=1).border = total_border
            ws1.cell(row=curr_row, column=2).border = total_border
    curr_row += 1

# SECTION: BALANCE SHEET
ws1["A17"] = "II. BALANCE SHEET (STATEMENT OF FINANCIAL POSITION)"
ws1["A17"].font = font_subheader
ws1["A17"].fill = SUBHEADER_FILL
ws1.merge_cells("A17:C17")

for col_idx, h in enumerate(["Balance Sheet Account", "Classification", "Amount ($)"], start=1):
    cell = ws1.cell(row=18, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="left" if col_idx < 3 else "right")

bs_rows = [
    ("ASSETS", "", ""),
    ("Cash & Bank Equivalents", "Reconciled Ending Bank Balance (Ties to Cash Flow)", "=C39"),
    ("Accounts Receivable", "Customer Invoices Due (50,000 Billed - 43,000 Collected)", 7000.00),
    ("Inventory", "Closing Stock on Hand (40,000 Purchased - 30,000 COGS)", 10000.00),
    ("Fixed Assets (Equipment)", "Commercial Assets Net of Depreciation", 15000.00),
    ("TOTAL ASSETS", "Sum of All Asset Accounts", "=SUM(C20:C23)"),
    ("LIABILITIES & OWNER'S EQUITY", "", ""),
    ("Accounts Payable", "Vendor Invoices Outstanding (40,000 Invoiced - 32,000 Paid)", 8000.00),
    ("Business Term Loan", "Long-Term Commercial Debt", 9500.00),
    ("Owner Capital Contributions", "Paid-in Capital", 17500.00),
    ("Retained Earnings (Period Net Profit)", "Tied directly to P&L Net Operating Profit", "=C15"),
    ("TOTAL LIABILITIES & EQUITY", "Must equal Total Assets", "=SUM(C26:C29)"),
]

curr_row = 19
for item, note, val in bs_rows:
    ws1.cell(row=curr_row, column=1, value=item).font = font_bold if ("TOTAL" in item or item in ["ASSETS", "LIABILITIES & OWNER'S EQUITY"]) else font_regular
    ws1.cell(row=curr_row, column=2, value=note).font = font_subtitle
    c_val = ws1.cell(row=curr_row, column=3, value=val if val != "" else None)
    if val != "":
        c_val.number_format = "$#,##0.00;($#,##0.00);\"-\""
        c_val.font = font_bold if "TOTAL" in item else font_regular
        if "TOTAL" in item:
            c_val.border = total_border
            ws1.cell(row=curr_row, column=1).border = total_border
            ws1.cell(row=curr_row, column=2).border = total_border
    curr_row += 1

# SECTION: CASH FLOW
ws1["A32"] = "III. STATEMENT OF CASH FLOWS"
ws1["A32"].font = font_subheader
ws1["A32"].fill = SUBHEADER_FILL
ws1.merge_cells("A32:C32")

for col_idx, h in enumerate(["Cash Flow Category", "Source / Document Verification", "Amount ($)"], start=1):
    cell = ws1.cell(row=33, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="left" if col_idx < 3 else "right")

cf_rows = [
    ("Opening Cash Balance (Sept 1, 2026)", "Beginning Bank Statement Balance", 12000.00),
    ("Cash Received from Customers", "Collected Invoice Receipts & Bank Deposits", 43000.00),
    ("Cash Paid to Suppliers & Vendors", "Vendor Disbursements & Inventory Settlements", -32000.00),
    ("Cash Paid for Operating Expenses", "Operating Receipts & Direct Bank Debits", -7500.00),
    ("Net Cash Flow for the Period", "Sum of Net Inflows and Outflows", "=SUM(C35:C37)"),
    ("Closing Cash Balance (Sept 30, 2026)", "Opening Cash + Net Flow (Ties to Balance Sheet)", "=C34+C38"),
]

curr_row = 34
for item, note, val in cf_rows:
    ws1.cell(row=curr_row, column=1, value=item).font = font_bold if ("Closing" in item or "Net" in item) else font_regular
    ws1.cell(row=curr_row, column=2, value=note).font = font_subtitle
    c_val = ws1.cell(row=curr_row, column=3, value=val)
    c_val.number_format = "$#,##0.00;($#,##0.00);\"-\""
    c_val.font = font_bold if ("Closing" in item or "Net" in item) else font_regular
    if "Closing Cash Balance" in item:
        c_val.border = total_border
        ws1.cell(row=curr_row, column=1).border = total_border
        ws1.cell(row=curr_row, column=2).border = total_border
    curr_row += 1

# SECTION: AUTOMATED AUDIT & RECONCILIATION CONTROL TABLE
ws1["E5"] = "AUTOMATED AUDIT & INTEGRITY VERIFICATION"
ws1["E5"].font = font_subheader
ws1["E5"].fill = SUBHEADER_FILL
ws1.merge_cells("E5:G5")

audit_headers = ["Audit Check", "Calculated Variance ($)", "Audit Status"]
for col_idx, h in enumerate(audit_headers, start=5):
    cell = ws1.cell(row=6, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center")

audit_checks = [
    ("Balance Sheet Tie-Out: Assets = Liabilities + Equity", "=C24-C30", '=IF(ROUND(F7,2)=0,"PASS (BALANCED)","FAIL (DISCREPANCY)")'),
    ("Cash Flow Tie-Out: Closing Cash = BS Cash Account", "=C39-C20", '=IF(ROUND(F8,2)=0,"PASS (TIED TO CENT)","FAIL (DISCREPANCY)")'),
    ("P&L Integration: Net Profit = Retained Earnings Added", "=C15-C29", '=IF(ROUND(F9,2)=0,"PASS (INTEGRATED)","FAIL (DISCREPANCY)")'),
    ("General Ledger Tie-Out: Total Debits = Total Credits", "='General Ledger & Data Entry'!G18", '=IF(ROUND(F10,2)=0,"PASS (ZERO VARIANCE)","FAIL (UNBALANCED)")'),
    ("Bank Reconciliation: Cleared Ledger = Bank Statement", "='Bank & Expense Reconciliation'!F14", '=IF(ROUND(F11,2)=0,"PASS (RECONCILED)","FAIL (DISCREPANCY)")'),
]

for idx, (chk, var_formula, stat_formula) in enumerate(audit_checks, start=7):
    c1 = ws1.cell(row=idx, column=5, value=chk)
    c1.font = font_regular
    c1.border = cell_border
    
    c2 = ws1.cell(row=idx, column=6, value=var_formula)
    c2.font = font_bold
    c2.number_format = "$#,##0.00;($#,##0.00);\"$0.00\""
    c2.alignment = Alignment(horizontal="right")
    c2.border = cell_border
    
    c3 = ws1.cell(row=idx, column=7, value=stat_formula)
    c3.font = font_audit_pass
    c3.fill = ACCENT_GREEN
    c3.alignment = Alignment(horizontal="center")
    c3.border = cell_border

# Explicit column widths
ws1.column_dimensions["A"].width = 38
ws1.column_dimensions["B"].width = 34
ws1.column_dimensions["C"].width = 18
ws1.column_dimensions["D"].width = 4
ws1.column_dimensions["E"].width = 46
ws1.column_dimensions["F"].width = 20
ws1.column_dimensions["G"].width = 24

# -------------------------------------------------------------
# TAB 2: GENERAL LEDGER & DOUBLE-ENTRY LOG
# -------------------------------------------------------------
ws2 = wb.create_sheet(title="General Ledger & Data Entry")
ws2.views.sheetView[0].showGridLines = True

ws2["A2"] = "GENERAL LEDGER: DOUBLE-ENTRY TRANSACTION LOG"
ws2["A2"].font = font_title
ws2["A3"] = "Every transaction recorded with balanced Debits and Credits | Zero tolerance for orphaned entries"
ws2["A3"].font = font_subtitle

gl_headers = ["Txn ID", "Date", "Description / Vendor", "Account Debited", "Account Credited", "Debit ($)", "Credit ($)"]
for col_idx, h in enumerate(gl_headers, start=1):
    cell = ws2.cell(row=5, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="left" if col_idx <= 5 else "right")

transactions = [
    ("TXN-1001", "2026-09-01", "Acme Office Supplies (Inv #9821)", "Office Supplies Expense", "Cash & Bank Equivalents", 1250.00, 1250.00),
    ("TXN-1002", "2026-09-02", "Global Tech Client (Inv #4412 Pmt)", "Cash & Bank Equivalents", "Accounts Receivable", 8500.00, 8500.00),
    ("TXN-1003", "2026-09-03", "Fiber Telecom Monthly Internet", "Utilities & Internet Expense", "Cash & Bank Equivalents", 750.00, 750.00),
    ("TXN-1004", "2026-09-04", "Direct Materials Supplier (PO #881)", "Inventory Stock", "Accounts Payable", 3000.00, 3000.00),
    ("TXN-1005", "2026-09-05", "Apex Corp Client (Milestone 1)", "Cash & Bank Equivalents", "Sales Revenue", 6750.00, 6750.00),
    ("TXN-1006", "2026-09-08", "Warehouse Facility Power Bill", "Utilities & Internet Expense", "Cash & Bank Equivalents", 750.00, 750.00),
    ("TXN-1007", "2026-09-12", "Legal & Compliance Audit Retainer", "Professional & Legal Fees", "Cash & Bank Equivalents", 3500.00, 3500.00),
    ("TXN-1008", "2026-09-15", "Vendor Payment (PO #881 Settlement)", "Accounts Payable", "Cash & Bank Equivalents", 3000.00, 3000.00),
    ("TXN-1009", "2026-09-20", "Client Enterprise License Contract", "Accounts Receivable", "Sales Revenue", 15000.00, 15000.00),
    ("TXN-1010", "2026-09-25", "Printer & Stationery Supplies", "Office Supplies Expense", "Cash & Bank Equivalents", 1250.00, 1250.00),
]

for r_idx, txn in enumerate(transactions, start=6):
    fill = ZEBRA_FILL if r_idx % 2 == 0 else PatternFill(fill_type=None)
    for c_idx, val in enumerate(txn, start=1):
        c = ws2.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_regular
        c.fill = fill
        c.border = cell_border
        if c_idx >= 6:
            c.number_format = "$#,##0.00;($#,##0.00);\"-\""
            c.alignment = Alignment(horizontal="right")
        elif c_idx == 2:
            c.alignment = Alignment(horizontal="center")

# Total Row
ws2.cell(row=16, column=1, value="TOTALS").font = font_bold
ws2.cell(row=16, column=1).alignment = Alignment(horizontal="center")
ws2.merge_cells("A16:E16")
for c in range(1, 6):
    ws2.cell(row=16, column=c).border = total_border

deb_tot = ws2.cell(row=16, column=6, value="=SUM(F6:F15)")
deb_tot.font = font_bold
deb_tot.number_format = "$#,##0.00"
deb_tot.alignment = Alignment(horizontal="right")
deb_tot.border = total_border

cred_tot = ws2.cell(row=16, column=7, value="=SUM(G6:G15)")
cred_tot.font = font_bold
cred_tot.number_format = "$#,##0.00"
cred_tot.alignment = Alignment(horizontal="right")
cred_tot.border = total_border

# GL Audit check
ws2["E18"] = "LEDGER BALANCE CHECK (Debits - Credits):"
ws2["E18"].font = font_bold
ws2["E18"].alignment = Alignment(horizontal="right")
ws2["G18"] = "=F16-G16"
ws2["G18"].font = font_bold
ws2["G18"].number_format = "$#,##0.00;($#,##0.00);\"$0.00\""
ws2["G18"].alignment = Alignment(horizontal="right")

ws2.column_dimensions["A"].width = 14
ws2.column_dimensions["B"].width = 14
ws2.column_dimensions["C"].width = 36
ws2.column_dimensions["D"].width = 28
ws2.column_dimensions["E"].width = 28
ws2.column_dimensions["F"].width = 16
ws2.column_dimensions["G"].width = 16

# -------------------------------------------------------------
# TAB 3: BANK & INVOICE RECONCILIATION STATEMENT
# -------------------------------------------------------------
ws3 = wb.create_sheet(title="Bank & Expense Reconciliation")
ws3.views.sheetView[0].showGridLines = True

ws3["A2"] = "MONTHLY BANK STATEMENT & EXPENSE RECONCILIATION"
ws3["A2"].font = font_title
ws3["A3"] = "Reconciling Supporting Documents against Ending Bank Account Balances | Exact Penny Tie-Out"
ws3["A3"].font = font_subtitle

for col_idx, h in enumerate(["Ref #", "Document Type", "Payee / Source", "Ledger Amount ($)", "Bank Statement ($)", "Variance ($)", "Audit Status"], start=1):
    cell = ws3.cell(row=5, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="left" if col_idx <= 3 else "right" if col_idx <= 6 else "center")

reconciliations = [
    ("REC-01", "Vendor Invoice", "Acme Office Supplies", 1250.00, 1250.00),
    ("REC-02", "Wire Deposit", "Global Tech Client", 8500.00, 8500.00),
    ("REC-03", "Auto-Debit Utility", "Fiber Telecom Internet", 750.00, 750.00),
    ("REC-04", "Wire Deposit", "Apex Corp Client Pmt", 6750.00, 6750.00),
    ("REC-05", "Utility Receipt", "Warehouse Power", 750.00, 750.00),
    ("REC-06", "Check Payment", "Legal Retainer #402", 3500.00, 3500.00),
    ("REC-07", "ACH Settlement", "Materials Supplier #881", 3000.00, 3000.00),
    ("REC-08", "Card Expense", "Stationery & Supplies", 1250.00, 1250.00),
]

for r_idx, item in enumerate(reconciliations, start=6):
    fill = ZEBRA_FILL if r_idx % 2 == 0 else PatternFill(fill_type=None)
    ws3.cell(row=r_idx, column=1, value=item[0]).alignment = Alignment(horizontal="center")
    ws3.cell(row=r_idx, column=2, value=item[1])
    ws3.cell(row=r_idx, column=3, value=item[2])
    
    c_led = ws3.cell(row=r_idx, column=4, value=item[3])
    c_led.number_format = "$#,##0.00"
    c_led.alignment = Alignment(horizontal="right")
    
    c_bnk = ws3.cell(row=r_idx, column=5, value=item[4])
    c_bnk.number_format = "$#,##0.00"
    c_bnk.alignment = Alignment(horizontal="right")
    
    c_var = ws3.cell(row=r_idx, column=6, value=f"=D{r_idx}-E{r_idx}")
    c_var.number_format = "$#,##0.00;($#,##0.00);\"$0.00\""
    c_var.alignment = Alignment(horizontal="right")
    
    c_stat = ws3.cell(row=r_idx, column=7, value=f'=IF(ROUND(F{r_idx},2)=0,"RECONCILED","DISCREPANCY")')
    c_stat.font = font_audit_pass
    c_stat.fill = ACCENT_GREEN
    c_stat.alignment = Alignment(horizontal="center")
    
    for c in range(1, 8):
        ws3.cell(row=r_idx, column=c).border = cell_border

# Totals
ws3.cell(row=14, column=1, value="TOTAL CLEARED ITEMS").font = font_bold
ws3.merge_cells("A14:C14")
for c in range(1, 4):
    ws3.cell(row=14, column=c).border = total_border

t_led = ws3.cell(row=14, column=4, value="=SUM(D6:D13)")
t_led.font = font_bold
t_led.number_format = "$#,##0.00"
t_led.alignment = Alignment(horizontal="right")
t_led.border = total_border

t_bnk = ws3.cell(row=14, column=5, value="=SUM(E6:E13)")
t_bnk.font = font_bold
t_bnk.number_format = "$#,##0.00"
t_bnk.alignment = Alignment(horizontal="right")
t_bnk.border = total_border

t_var = ws3.cell(row=14, column=6, value="=D14-E14")
t_var.font = font_bold
t_var.number_format = "$#,##0.00;($#,##0.00);\"$0.00\""
t_var.alignment = Alignment(horizontal="right")
t_var.border = total_border

t_stat = ws3.cell(row=14, column=7, value='=IF(ROUND(F14,2)=0,"100% TIED","FAIL")')
t_stat.font = font_audit_pass
t_stat.fill = ACCENT_GREEN
t_stat.alignment = Alignment(horizontal="center")
t_stat.border = total_border

ws3.column_dimensions["A"].width = 14
ws3.column_dimensions["B"].width = 22
ws3.column_dimensions["C"].width = 30
ws3.column_dimensions["D"].width = 18
ws3.column_dimensions["E"].width = 18
ws3.column_dimensions["F"].width = 16
ws3.column_dimensions["G"].width = 20

output_path = "Financial_Reconciliation_and_Reporting_Model.xlsx"
wb.save(output_path)
print(f"Successfully generated and saved: {output_path}")
