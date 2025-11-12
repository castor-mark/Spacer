# Enhanced Excel Validator

A powerful Python tool for validating and cleaning data in Excel files. This tool checks for spacing issues, time format problems, and file extension case errors, then generates detailed reports with highlighted problematic cells.

## Features

- 🔍 **Multiple Validation Types**:

  - 📝 Spacing issues and special characters
  - ⏰ 24-hour time format validation (e.g., `05:11:20` not `5:11:20`)
  - 📁 File extension case checking (e.g., `.pdf` not `.PDF`)
- 🎯 **Flexible Column Selection**:

  - Single column validation
  - Multiple columns (comma-separated)
  - All columns validation
- 🔴 **Strict Mode**:

  - Flag ANY space for columns that should never contain spaces (filenames, IDs, SKUs)
  - Normal mode for typical text columns
- 📊 **Comprehensive Reporting**:

  - Excel reports with highlighted problematic cells
  - CSV reports for easy data review
  - Timestamped folders for historical records
  - "Latest" folder for quick access to recent results
- 🔄 **Session Management**:

  - Accumulate results when analyzing multiple columns
  - Generate comprehensive reports for entire session

## Installation

### Prerequisites

- Python 3.7 or higher (3.8+ recommended)
- pip package manager

### Quick Setup

```bash
# Clone or download the project
git clone <repository-url>  # or download and extract

# Navigate to project directory
cd excel-validator

# Create virtual environment (recommended)
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On Mac/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the script
python excel_validator.py
```
