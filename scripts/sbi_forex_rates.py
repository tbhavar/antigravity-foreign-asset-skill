"""
sbi_forex_rates.py
Utility module to parse historical State Bank of India (SBI) Telegraphic Transfer (TT) rates
and resolve foreign exchange rates under Rule 115 of the Indian Income-tax Rules, 1962.
"""

import os
import pandas as pd
import datetime

class SBIForexResolver:
    def __init__(self, csv_path="resources/SBI_REFERENCE_RATES_USD.csv"):
        self.csv_path = csv_path
        self.df = None
        self._load_rates()

    def _load_rates(self):
        if not os.path.exists(self.csv_path):
            # Try alternate fallback locations
            fallbacks = [
                os.path.basename(self.csv_path),
                os.path.join("..", self.csv_path),
                os.path.join("resources", os.path.basename(self.csv_path)),
                "SBI_REFERENCE_RATES_USD (1).csv",
                "SBI_REFERENCE_RATES_USD.csv"
            ]
            for fb in fallbacks:
                if os.path.exists(fb):
                    self.csv_path = fb
                    break

        if not os.path.exists(self.csv_path):
            raise FileNotFoundError(f"SBI Rates file not found at: {self.csv_path}")

        df = pd.read_csv(self.csv_path)
        
        # Determine date column name
        date_col = next((c for c in df.columns if 'DATE' in c.upper()), df.columns[0])
        buy_col = next((c for c in df.columns if 'TT BUY' in c.upper() or 'TT_BUY' in c.upper()), None)
        sell_col = next((c for c in df.columns if 'TT SELL' in c.upper() or 'TT_SELL' in c.upper()), None)

        if not buy_col:
            raise ValueError(f"Could not find 'TT BUY' column in {self.csv_path}")

        # Parse dates flexibly
        df['DateClean'] = pd.to_datetime(df[date_col].astype(str).str.split().str[0], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['DateClean'])
        df['TT_BUY_NUM'] = pd.to_numeric(df[buy_col], errors='coerce')
        if sell_col:
            df['TT_SELL_NUM'] = pd.to_numeric(df[sell_col], errors='coerce')
        else:
            df['TT_SELL_NUM'] = df['TT_BUY_NUM']

        self.df = df.sort_values('DateClean').reset_index(drop=True)

    def get_rate_on_date(self, date_input, rate_type='BUY', lookback_days=10):
        """
        Retrieves SBI TT rate for a specific date. If the bank was closed (holiday/weekend),
        looks back up to `lookback_days` to find the latest published rate.
        """
        target_date = pd.to_datetime(date_input)
        col = 'TT_BUY_NUM' if rate_type.upper() == 'BUY' else 'TT_SELL_NUM'

        curr_date = target_date
        for _ in range(lookback_days):
            match = self.df[self.df['DateClean'] == curr_date]
            if not match.empty:
                val = match.iloc[0][col]
                if pd.notna(val) and val > 0:
                    return curr_date.strftime('%Y-%m-%d'), float(val)
            curr_date = curr_date - datetime.timedelta(days=1)

        return None, None

    def get_preceding_month_end_rate(self, date_input, rate_type='BUY'):
        """
        Rule 115 Compliance:
        Returns the SBI TT Buying Rate on the LAST DAY OF THE MONTH PRECEDING
        the month in which the income arose or asset was transferred.
        """
        d = pd.to_datetime(date_input)
        first_day_of_month = d.replace(day=1)
        last_day_preceding_month = first_day_of_month - datetime.timedelta(days=1)
        return self.get_rate_on_date(last_day_preceding_month, rate_type=rate_type)

if __name__ == "__main__":
    resolver = SBIForexResolver()
    sample_date = "2025-05-15"
    actual_date, rate = resolver.get_preceding_month_end_rate(sample_date)
    print(f"Sample: For transaction on {sample_date}, Rule 115 Preceding Month-End ({actual_date}) SBI TTBR = ₹{rate}")
