import sys, os
sys.path.append(r"c:\Users\lap4all\Documents\Auto report")
from update_aging_assignments import run_calculations

try:
    run_calculations()
except Exception as e:
    print("MAIN EXCEPTION:", e)
