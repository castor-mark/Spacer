# config.py
import os

# Folder where Excel files are located
EXCEL_FOLDER = "excel_files"

# Output folder for reports
OUTPUT_FOLDER = "reports"

# Create folders if they don't exist
os.makedirs(EXCEL_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)