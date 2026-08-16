import os
import io
import sys
import pandas as pd

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
csv_path = os.path.join(SCRIPT_DIR, 'gtc_inspect.csv')

df = pd.read_csv(csv_path)

print(f"Number of columns: {len(df.columns)}")
print(f"Number of rows: {len(df)}")

# Print cell values of first few rows for all columns
for col_idx in range(len(df.columns)):
    vals = [str(df.iloc[row_idx, col_idx]) for row_idx in range(min(5, len(df)))]
    print(f"Col {col_idx:2d}: " + " | ".join(vals))
