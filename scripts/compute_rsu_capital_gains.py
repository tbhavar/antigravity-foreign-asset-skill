"""
compute_rsu_capital_gains.py
Automates calculation of Short-Term and Long-Term Capital Gains for foreign shares / RSUs.
Preserves dynamic Excel formulas and generates ICAI-compliant statutory computation notes.
"""

import os
import datetime
import pandas as pd
import openpyxl
from copy import copy
from sbi_forex_rates import SBIForexResolver

def compute_and_export_rsu_cg(
    sales_records,
    output_filepath="RSU Gain.xlsx",
    template_filepath="resources/RSU-CG-Calculations_template.xlsx",
    forex_csv="resources/SBI_REFERENCE_RATES_USD.csv",
    financial_year="2025-26",
    share_name="Equity Shares - RSU"
):
    """
    sales_records: list of dicts containing:
        - date_acq: datetime or str ('YYYY-MM-DD' or 'MM/DD/YYYY')
        - date_sold: datetime or str ('YYYY-MM-DD' or 'MM/DD/YYYY')
        - qty: float
        - vest_fmv: float (USD)
        - sale_price: float (USD)
        - selling_expenses_usd: float (optional, default 0.0)
    """
    forex = SBIForexResolver(forex_csv)

    # Use template if available, otherwise fallback to creating workbook
    if os.path.exists(template_filepath):
        wb = openpyxl.load_workbook(template_filepath)
        ws = wb.active
        start_row = 11
        is_extended_template = True
    elif os.path.exists(output_filepath):
        wb = openpyxl.load_workbook(output_filepath)
        ws = wb.active
        start_row = 2
        is_extended_template = False
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        start_row = 2
        is_extended_template = False

    # Handle NIL sales scenario
    if not sales_records or len(sales_records) == 0:
        print(f"[RSU CG] No sales recorded for FY {financial_year}. Generating NIL Statement.")
        if is_extended_template:
            ws['D7'].value = 0.0
            ws['D8'].value = 0.0
            ws.cell(row=11, column=1, value=1)
            ws.cell(row=11, column=2, value="N/A")
            ws.cell(row=11, column=3, value="N/A")
            ws.cell(row=11, column=4, value=financial_year)
            ws.cell(row=11, column=6, value=share_name)
            ws.cell(row=11, column=7, value=0)
            ws.cell(row=11, column=18, value=0.0)
            ws.cell(row=11, column=21, value=0.0)
        else:
            # Minimal template row
            ws.cell(row=2, column=1, value=1)
            ws.cell(row=2, column=2, value="N/A")
            ws.cell(row=2, column=3, value="N/A")
            ws.cell(row=2, column=4, value=financial_year)
            ws.cell(row=2, column=5, value=0)
            ws.cell(row=2, column=6, value=share_name)
            ws.cell(row=2, column=7, value=0)
            ws.cell(row=2, column=8, value=0)
            ws.cell(row=2, column=9, value="=G2*H2")
            ws.cell(row=2, column=10, value=0)
            ws.cell(row=2, column=11, value="=I2*J2")
            ws.cell(row=2, column=12, value=0)
            ws.cell(row=2, column=13, value="=G2*L2")
            ws.cell(row=2, column=14, value=0)
            ws.cell(row=2, column=15, value=0)
            ws.cell(row=2, column=16, value="=M2*N2")
            ws.cell(row=2, column=17, value="=P2-(O2*N2)")
            ws.cell(row=2, column=18, value=0)
            ws.cell(row=2, column=19, value="=Q2-K2")
            
            ws.cell(row=3, column=1, value="Total / Net Capital Gain (Loss)")
            ws.cell(row=3, column=7, value="=SUM(G2:G2)")
            ws.cell(row=3, column=9, value="=SUM(I2:I2)")
            ws.cell(row=3, column=11, value="=SUM(K2:K2)")
            ws.cell(row=3, column=13, value="=SUM(M2:M2)")
            ws.cell(row=3, column=16, value="=SUM(P2:P2)")
            ws.cell(row=3, column=17, value="=SUM(Q2:Q2)")
            ws.cell(row=3, column=18, value="=SUM(R2:R2)")
            ws.cell(row=3, column=19, value="=SUM(S2:S2)")

        wb.save(output_filepath)
        print(f"[RSU CG] Successfully saved to {output_filepath}")
        return output_filepath

    # Cache row styling from template
    ref_row = start_row
    styles = []
    max_cols = 22 if is_extended_template else 19
    for c in range(1, max_cols + 1):
        cell = ws.cell(row=ref_row, column=c)
        styles.append({
            'font': copy(cell.font),
            'fill': copy(cell.fill),
            'border': copy(cell.border),
            'alignment': copy(cell.alignment),
            'number_format': cell.number_format
        })

    end_row = start_row + len(sales_records) - 1

    for idx, rec in enumerate(sales_records):
        curr_row = start_row + idx
        date_acq = pd.to_datetime(rec['date_acq'])
        date_sold = pd.to_datetime(rec['date_sold'])
        days_held = (date_sold - date_acq).days
        is_long_term = days_held > 730 # 24 months for foreign unlisted shares

        qty = float(rec['qty'])
        vest_fmv = float(rec['vest_fmv'])
        sale_price = float(rec['sale_price'])
        selling_exp_usd = float(rec.get('selling_expenses_usd', 0.0))

        # Rule 115 TT Buying Rates: Last day of month preceding acquisition & sale
        _, acq_tt_rate = forex.get_preceding_month_end_rate(date_acq)
        _, sale_tt_rate = forex.get_preceding_month_end_rate(date_sold)

        if is_extended_template:
            # Extended template mapping (Row 11 onwards)
            ws.cell(row=curr_row, column=1, value=idx + 1)
            ws.cell(row=curr_row, column=2, value=date_acq.strftime('%m-%d-%y'))
            ws.cell(row=curr_row, column=3, value=date_sold.strftime('%m-%d-%y'))
            ws.cell(row=curr_row, column=4, value=financial_year)
            ws.cell(row=curr_row, column=5, value=days_held)
            ws.cell(row=curr_row, column=6, value=share_name)
            ws.cell(row=curr_row, column=7, value=qty)
            ws.cell(row=curr_row, column=8, value=vest_fmv)
            ws.cell(row=curr_row, column=9, value=f"=ROUND(G{curr_row}*H{curr_row}, 2)")
            ws.cell(row=curr_row, column=10, value=acq_tt_rate)
            ws.cell(row=curr_row, column=11, value=f"=ROUND(I{curr_row}*J{curr_row}, 2)")
            ws.cell(row=curr_row, column=12, value=sale_price)
            ws.cell(row=curr_row, column=13, value=f"=ROUND(G{curr_row}*L{curr_row}, 2)")
            ws.cell(row=curr_row, column=14, value=sale_tt_rate)
            ws.cell(row=curr_row, column=15, value=f"=ROUND(M{curr_row}*N{curr_row}, 2)")
            ws.cell(row=curr_row, column=16, value=selling_exp_usd)
            ws.cell(row=curr_row, column=17, value=f"=ROUND(O{curr_row}-(P{curr_row}*N{curr_row}), 2)")
            
            # STCG vs LTCG
            if not is_long_term:
                ws.cell(row=curr_row, column=18, value=f"=ROUND(Q{curr_row}-K{curr_row}, 2)")
                ws.cell(row=curr_row, column=19, value=1)
                ws.cell(row=curr_row, column=20, value=0)
                ws.cell(row=curr_row, column=21, value=0.0)
            else:
                ws.cell(row=curr_row, column=18, value=0.0)
                ws.cell(row=curr_row, column=19, value=1)
                ws.cell(row=curr_row, column=20, value=f"=K{curr_row}")
                ws.cell(row=curr_row, column=21, value=f"=ROUND(Q{curr_row}-K{curr_row}, 2)")

        else:
            # Arun Nair style template
            ws.cell(row=curr_row, column=1, value=idx + 1)
            ws.cell(row=curr_row, column=2, value=date_acq.strftime('%m-%d-%y'))
            ws.cell(row=curr_row, column=3, value=date_sold.strftime('%m-%d-%y'))
            ws.cell(row=curr_row, column=4, value=financial_year)
            ws.cell(row=curr_row, column=5, value=days_held)
            ws.cell(row=curr_row, column=6, value=share_name)
            ws.cell(row=curr_row, column=7, value=qty)
            ws.cell(row=curr_row, column=8, value=vest_fmv)
            ws.cell(row=curr_row, column=9, value=f"=ROUND(G{curr_row}*H{curr_row}, 2)")
            ws.cell(row=curr_row, column=10, value=acq_tt_rate)
            ws.cell(row=curr_row, column=11, value=f"=ROUND(I{curr_row}*J{curr_row}, 2)")
            ws.cell(row=curr_row, column=12, value=sale_price)
            ws.cell(row=curr_row, column=13, value=f"=ROUND(G{curr_row}*L{curr_row}, 2)")
            ws.cell(row=curr_row, column=14, value=sale_tt_rate)
            ws.cell(row=curr_row, column=15, value=selling_exp_usd)
            ws.cell(row=curr_row, column=16, value=f"=ROUND(M{curr_row}*N{curr_row}, 2)")
            ws.cell(row=curr_row, column=17, value=f"=ROUND(P{curr_row}-(O{curr_row}*N{curr_row}), 2)")
            
            if not is_long_term:
                ws.cell(row=curr_row, column=18, value=f"=ROUND(Q{curr_row}-K{curr_row}, 2)")
                ws.cell(row=curr_row, column=19, value=0.0)
            else:
                ws.cell(row=curr_row, column=18, value=0.0)
                ws.cell(row=curr_row, column=19, value=f"=ROUND(Q{curr_row}-K{curr_row}, 2)")

        # Apply cell styles
        for c in range(1, max_cols + 1):
            st = styles[c - 1]
            cell = ws.cell(row=curr_row, column=c)
            if st['font']: cell.font = st['font']
            if st['fill']: cell.fill = st['fill']
            if st['border']: cell.border = st['border']
            if st['alignment']: cell.alignment = st['alignment']
            if st['number_format']: cell.number_format = st['number_format']

    # Update summary formulas
    if is_extended_template:
        ws['D7'].value = f"=SUM(R{start_row}:R{end_row})"
        ws['D8'].value = f"=SUM(U{start_row}:U{end_row})"
    else:
        tot_row = end_row + 1
        ws.cell(row=tot_row, column=1, value="Total / Net Capital Gain (Loss)")
        ws.cell(row=tot_row, column=7, value=f"=SUM(G{start_row}:G{end_row})")
        ws.cell(row=tot_row, column=9, value=f"=SUM(I{start_row}:I{end_row})")
        ws.cell(row=tot_row, column=11, value=f"=SUM(K{start_row}:K{end_row})")
        ws.cell(row=tot_row, column=13, value=f"=SUM(M{start_row}:M{end_row})")
        ws.cell(row=tot_row, column=16, value=f"=SUM(P{start_row}:P{end_row})")
        ws.cell(row=tot_row, column=17, value=f"=SUM(Q{start_row}:Q{end_row})")
        ws.cell(row=tot_row, column=18, value=f"=SUM(R{start_row}:R{end_row})")
        ws.cell(row=tot_row, column=19, value=f"=SUM(S{start_row}:S{end_row})")

    wb.save(output_filepath)
    print(f"[RSU CG] Successfully populated {output_filepath} with {len(sales_records)} rows.")
    return output_filepath

if __name__ == "__main__":
    # Test execution
    dummy_sales = [
        {"date_acq": "2025-03-01", "date_sold": "2025-05-27", "qty": 4.0, "vest_fmv": 910.58, "sale_price": 902.31}
    ]
    compute_and_export_rsu_cg(dummy_sales, output_filepath="test_rsu_gain.xlsx")
    if os.path.exists("test_rsu_gain.xlsx"):
        os.remove("test_rsu_gain.xlsx")
