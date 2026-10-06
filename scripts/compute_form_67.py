"""
compute_form_67.py
Prepares Form 67 statement for claiming Double Taxation Relief u/s 90 under DTAA (Article 10 for Dividend).
Retains live Excel formulas and includes statutory explanatory notes for tax filing.
"""

import os
import datetime
import pandas as pd
import openpyxl
from sbi_forex_rates import SBIForexResolver

def compute_and_export_form_67(
    dividend_records,
    output_filepath="Form-67.xlsx",
    template_filepath="resources/Form-67_template.xlsx",
    forex_csv="resources/SBI_REFERENCE_RATES_USD.csv",
    country_name="United States of America",
    taxpayer_tin="XXX-XX-0000",
    indian_slab_rate=0.312, # 30% slab + 4% cess under New Tax Regime
    dtaa_article="Article 10",
    dtaa_rate=0.25
):
    """
    dividend_records: list of dicts:
        - date_received: 'YYYY-MM-DD' or datetime
        - gross_div_usd: float
        - tax_withheld_usd: float
    """
    forex = SBIForexResolver(forex_csv)

    total_gross_inr = 0.0
    total_tax_paid_inr = 0.0
    reconciliation_details = []

    for item in dividend_records:
        dt = pd.to_datetime(item['date_received'])
        g_usd = float(item['gross_div_usd'])
        t_usd = float(item.get('tax_withheld_usd', g_usd * 0.25))

        # Rule 115 TT Buying rate on preceding month end
        rate_date, tt_rate = forex.get_preceding_month_end_rate(dt)
        g_inr = round(g_usd * tt_rate, 2)
        t_inr = round(t_usd * tt_rate, 2)

        total_gross_inr += g_inr
        total_tax_paid_inr += t_inr
        reconciliation_details.append({
            'date': dt.strftime('%Y-%m-%d'),
            'rate_date': rate_date,
            'tt_rate': tt_rate,
            'gross_usd': g_usd,
            'gross_inr': g_inr,
            'tax_usd': t_usd,
            'tax_inr': t_inr
        })

    # Prepare Workbook
    if os.path.exists(template_filepath):
        wb = openpyxl.load_workbook(template_filepath)
        ws = wb.active
    elif os.path.exists(output_filepath):
        wb = openpyxl.load_workbook(output_filepath)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        # Setup headers
        ws.cell(row=1, column=3, value="Tax paid")
        ws.cell(row=1, column=4, value="Tax payable")
        ws.cell(row=1, column=5, value="Relief")
        ws.cell(row=1, column=6, value="Taxpayer")
        ws.cell(row=1, column=7, value="Heads of")
        ws.cell(row=1, column=8, value="Article of")
        ws.cell(row=1, column=9, value="Section of")
        ws.cell(row=1, column=10, value="DTAA")
        ws.cell(row=1, column=11, value="Tax under")
        ws.cell(row=1, column=12, value="Tax u/s 115JC")
        ws.cell(row=1, column=13, value="Tax rate")
        ws.cell(row=1, column=14, value="Nature of")

        headers_r2 = [
            'Country Name', 'Income outside India', 'outside India', 'in India', 'claimed',
            'Identification No.', 'income', 'DTAA', 'relief', 'rate',
            'normal provisions - IT Act', 'Tax u/s 115JC', 'outside India', 'income'
        ]
        for c_idx, h in enumerate(headers_r2, 1):
            ws.cell(row=2, column=c_idx, value=h)

    # Set row 3 computation with live Excel formulas
    ws.cell(row=3, column=1, value=country_name)
    ws.cell(row=3, column=2, value=round(total_gross_inr, 2))
    ws.cell(row=3, column=3, value=round(total_tax_paid_inr, 2))
    ws.cell(row=3, column=4, value=f"=ROUND(B3*{indian_slab_rate}, 2)")
    ws.cell(row=3, column=5, value="=MIN(C3, D3)")
    ws.cell(row=3, column=6, value=taxpayer_tin)
    ws.cell(row=3, column=7, value="Income from Other Sources")
    ws.cell(row=3, column=8, value=dtaa_article)
    ws.cell(row=3, column=9, value="Section 90")
    ws.cell(row=3, column=10, value=dtaa_rate)
    ws.cell(row=3, column=11, value="=D3")
    ws.cell(row=3, column=12, value=0)
    ws.cell(row=3, column=13, value="=ROUND(C3/B3, 4)")
    ws.cell(row=3, column=14, value="Dividend Income")

    wb.save(output_filepath)
    print(f"[Form 67] Successfully exported to {output_filepath} (Gross: ₹{total_gross_inr:.2f}, Tax: ₹{total_tax_paid_inr:.2f})")
    return output_filepath, reconciliation_details

if __name__ == "__main__":
    dummy_divs = [
        {"date_received": "2025-06-15", "gross_div_usd": 15.0, "tax_withheld_usd": 3.75},
        {"date_received": "2025-09-15", "gross_div_usd": 15.0, "tax_withheld_usd": 3.75}
    ]
    compute_and_export_form_67(dummy_divs, output_filepath="test_form67.xlsx")
    if os.path.exists("test_form67.xlsx"):
        os.remove("test_form67.xlsx")
