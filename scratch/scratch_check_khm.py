import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

CREDENTIALS_PATH = r'c:\Users\lap4all\Documents\Auto report\credentials.json'
SPREADSHEET_ID   = '1XaziS_8UB2lCwL01bxam106EdIUHtlMISDGJX1PU5A8'
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(SPREADSHEET_ID)
    
    print("Reading f30...")
    ws = ss.worksheet("f30")
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    
    print(f"Total rows in f30: {len(df)}")
    print("Columns:", df.columns.tolist())
    
    # Let's inspect some date and revenue columns
    date_col = "Ngày LTC đầu tiên"
    print("\nRaw date values (first 10):")
    print(df[date_col].head(10).tolist())
    print("Raw date types:")
    print(df[date_col].head(10).apply(type).tolist())
    
    # Try parsing using the parse_vn_date function from the original code
    _MONTH_MAP = {f"thg {i}": i for i in range(1, 13)}
    import re
    def parse_vn_date(s):
        if isinstance(s, (int, float)):
            return pd.to_datetime(s, unit='D', origin='1899-12-30').to_pydatetime()
        s = str(s).strip()
        if s.isdigit():
            return pd.to_datetime(int(s), unit='D', origin='1899-12-30').to_pydatetime()
        for m, num in _MONTH_MAP.items():
            if m in s:
                parts = re.findall(r"\d+", s)
                if len(parts) >= 2:
                    return datetime(int(parts[-1]), num, int(parts[0]))
        return None

    from datetime import datetime
    df['parsed_date'] = df[date_col].apply(parse_vn_date)
    print("\nParsed date values (first 10):")
    print(df['parsed_date'].head(10).tolist())

    df[date_col] = pd.to_datetime(df['parsed_date'], errors='coerce')
    
    # Revenue column
    rev_col = "DoanhThu_NoVAT"
    df[rev_col] = pd.to_numeric(df[rev_col], errors='coerce').fillna(0)
    
    # Group by month of date_col
    df['Month'] = df[date_col].dt.to_period('M')
    summary = df.groupby('Month').agg(
        Count=('Mã KH', 'count'),
        Total_Revenue=(rev_col, 'sum'),
        Mean_Revenue=(rev_col, 'mean')
    )
    print("\nSummary by Month of First LTC Date:")
    print(summary)
    
    # Analysis of May vs June shops
    df_may = df[df['Month'] == '2026-05']
    df_june = df[df['Month'] == '2026-06']
    
    print(f"\nMay shops: total={len(df_may)}, with revenue>0={len(df_may[df_may[rev_col]>0])}, zero revenue={len(df_may[df_may[rev_col]==0])}")
    print(f"June shops: total={len(df_june)}, with revenue>0={len(df_june[df_june[rev_col]>0])}, zero revenue={len(df_june[df_june[rev_col]==0])}")
    
    print("\nTop 15 May shops by revenue:")
    print(df_may.sort_values(by=rev_col, ascending=False)[['Mã KH', 'Tên KH', date_col, rev_col]].head(15))
    
    print("\nTop 15 June shops by revenue:")
    print(df_june.sort_values(by=rev_col, ascending=False)[['Mã KH', 'Tên KH', date_col, rev_col]].head(15))
    
    # Print details for outlier shop 5099749
    outlier = df[df['Mã KH'] == '5099749']
    print("\nOutlier Shop Details:")
    print(outlier.to_string())
    
    # Check if there are other outliers in June
    print("\nJune shops with revenue > 10M:")
    print(df_june[df_june[rev_col] > 10000000][['Mã KH', 'Tên KH', 'Tinh', 'AM', rev_col]])

if __name__ == "__main__":
    main()
