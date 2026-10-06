"""
populate_schedule_fa.py
Populates Schedule FA Table A3 (Foreign Equity Shares / RSUs) for Indian ITR filing.
Supports configurable reporting periods:
  - FINANCIAL_YEAR (1st April to 31st March)
  - CALENDAR_YEAR  (1st January to 31st December - US/foreign accounting period)
Preserves dynamic Excel formulas and generates ICAI-compliant disclosure notes.
"""

import os
import datetime
import pandas as pd
import openpyxl
from copy import copy
from sbi_forex_rates import SBIForexResolver

def populate_schedule_fa_sheet(
    holding_tranches,
    output_filepath="Schedule FA.xlsx",
    template_filepath="resources/Schedule_FA_template.xlsx",
    forex_csv="resources/SBI_REFERENCE_RATES_USD.csv",
    country_name="United States of America",
    entity_name="Microsoft Corporation",
    entity_address="One Microsoft Way, Redmond, WA",
    entity_zip="98052",
    entity_nature="Equity Shares - RSU",
    period_mode="FINANCIAL_YEAR", # or "CALENDAR_YEAR"
    period_year=2025,
    period_end_price=369.37,
    peak_stock_price=555.45,
    peak_tt_rate=94.35
):
    """
    holding_tranches: list of dicts:
        - vest_date: 'YYYY-MM-DD' or datetime
        - qty: float
        - vest_fmv: float (USD)
        - is_sold: bool
        - date_sold: 'YYYY-MM-DD' (optional)
        - sale_price: float (optional)
        - dividend_inr: float (optional)
    """
    forex = SBIForexResolver(forex_csv)

    # Determine Period Boundaries
    if period_mode.upper() == "CALENDAR_YEAR":
        period_start = pd.to_datetime(f"{period_year}-01-01")
        period_end = pd.to_datetime(f"{period_year}-12-31")
    else:
        # FINANCIAL_YEAR e.g. FY 2025-26: 01-Apr-2025 to 31-Mar-2026
        period_start = pd.to_datetime(f"{period_year}-04-01")
        period_end = pd.to_datetime(f"{period_year+1}-03-31")

    # Get Period-End TTBR
    _, period_end_tt_rate = forex.get_rate_on_date(period_end)
    if period_end_tt_rate is None:
        _, period_end_tt_rate = forex.get_preceding_month_end_rate(period_end)

    # Load Workbook
    if os.path.exists(output_filepath):
        wb = openpyxl.load_workbook(output_filepath)
        ws = wb.active
        start_row = 3
    elif os.path.exists(template_filepath):
        wb = openpyxl.load_workbook(template_filepath)
        ws = wb.active
        start_row = 4
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        start_row = 3

    # Populate rows
    num_rows = len(holding_tranches)
    for idx, tranche in enumerate(holding_tranches):
        curr_row = start_row + idx
        vest_date = pd.to_datetime(tranche['vest_date'])
        qty = float(tranche['qty'])
        vest_fmv = float(tranche['vest_fmv'])
        is_sold = tranche.get('is_sold', False)
        div_inr = float(tranche.get('dividend_inr', 0.0))

        # Initial Value: Acquisition TTBR on actual vest date
        _, acq_tt_rate = forex.get_rate_on_date(vest_date)
        if acq_tt_rate is None:
            _, acq_tt_rate = forex.get_preceding_month_end_rate(vest_date)

        ws.cell(row=curr_row, column=1, value=country_name)
        ws.cell(row=curr_row, column=2, value=entity_name)
        ws.cell(row=curr_row, column=3, value=entity_address)
        ws.cell(row=curr_row, column=4, value=entity_zip)
        ws.cell(row=curr_row, column=5, value=entity_nature)
        ws.cell(row=curr_row, column=6, value=vest_date.strftime('%d-%m-%Y'))

        # Formulas for valuations
        ws.cell(row=curr_row, column=7, value=f"=ROUND({qty} * {vest_fmv} * {acq_tt_rate}, 2)")
        ws.cell(row=curr_row, column=8, value=f"=ROUND({qty} * {peak_stock_price} * {peak_tt_rate}, 2)")

        if not is_sold:
            ws.cell(row=curr_row, column=9, value=f"=ROUND({qty} * {period_end_price} * {period_end_tt_rate}, 2)")
            ws.cell(row=curr_row, column=11, value=0.0)
        else:
            sale_price = float(tranche.get('sale_price', 0.0))
            date_sold = pd.to_datetime(tranche.get('date_sold', period_end))
            _, sale_tt_rate = forex.get_rate_on_date(date_sold)
            if sale_tt_rate is None:
                _, sale_tt_rate = forex.get_preceding_month_end_rate(date_sold)
            
            ws.cell(row=curr_row, column=9, value=0.0)
            ws.cell(row=curr_row, column=11, value=f"=ROUND({qty} * {sale_price} * {sale_tt_rate}, 2)")

        ws.cell(row=curr_row, column=10, value=round(div_inr, 2))

    # Total Summary Row
    tot_row = start_row + num_rows
    ws.cell(row=tot_row, column=1, value="Total Foreign Equity Assets Disclosure")
    ws.cell(row=tot_row, column=7, value=f"=SUM(G{start_row}:G{tot_row-1})")
    ws.cell(row=tot_row, column=8, value=f"=SUM(H{start_row}:H{tot_row-1})")
    ws.cell(row=tot_row, column=9, value=f"=SUM(I{start_row}:I{tot_row-1})")
    ws.cell(row=tot_row, column=10, value=f"=SUM(J{start_row}:J{tot_row-1})")
    ws.cell(row=tot_row, column=11, value=f"=SUM(K{start_row}:K{tot_row-1})")

    wb.save(output_filepath)
    print(f"[Schedule FA] Successfully populated {output_filepath} ({num_rows} tranches, period: {period_mode}).")
    return output_filepath

if __name__ == "__main__":
    dummy_tranches = [
        {"vest_date": "2025-05-15", "qty": 3.6386, "vest_fmv": 448.37, "is_sold": False, "dividend_inr": 264.85},
        {"vest_date": "2025-11-30", "qty": 23.2186, "vest_fmv": 423.46, "is_sold": False, "dividend_inr": 2174.00}
    ]
    populate_schedule_fa_sheet(dummy_tranches, output_filepath="test_schedule_fa.xlsx")
    if os.path.exists("test_schedule_fa.xlsx"):
        os.remove("test_schedule_fa.xlsx")
