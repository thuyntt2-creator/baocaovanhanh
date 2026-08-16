import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import numpy as np

gc = gspread.authorize(Credentials.from_service_account_file('credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets']))
sh = gc.open_by_key('14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4')
ws = sh.worksheet('TTS')
vals = ws.get_all_values()
df = pd.DataFrame(vals[1:], columns=vals[0])

latest_date = df['ltc_date'].max()
print("Latest date:", latest_date)
tts_latest = df[df['ltc_date'] == latest_date].copy()
tts_latest['vol_ltc'] = pd.to_numeric(tts_latest['vol_ltc'], errors='coerce').fillna(0)
tts_latest['ontime_xuat_first_mile'] = pd.to_numeric(tts_latest['ontime_xuat_first_mile'], errors='coerce').fillna(0)

# Group by bc_lay
tts_grouped = tts_latest.groupby('bc_lay', as_index=False).agg({
    'vol_ltc': 'sum',
    'ontime_xuat_first_mile': 'sum'
})
tts_grouped['rot_count'] = tts_grouped['vol_ltc'] - tts_grouped['ontime_xuat_first_mile']
tts_grouped['rate'] = np.where(tts_grouped['vol_ltc'] > 0, tts_grouped['rot_count'] / tts_grouped['vol_ltc'], 0.0)

top10 = tts_grouped.sort_values(by='rot_count', ascending=False).head(10).reset_index(drop=True)
print('--- TOP 10 AFTER COMBINING BY BC_LAY ---')
for idx, r in top10.iterrows():
    print(f"{idx+1}. {r['bc_lay']} | Vol Can: {r['vol_ltc']:.0f} | Vol Rot: {r['rot_count']:.0f} | Ty le: {r['rate']*100:.2f}%")
