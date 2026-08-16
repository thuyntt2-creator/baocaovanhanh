import sys
try:
    import win32com.client
    print("win32com.client is available!")
except ImportError:
    print("win32com.client is NOT available.")
