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

## Project Structure

```
.
├── config.py               # Configuration settings for the validator
├── excel_validator.py      # Main script for running Excel validations
├── README.md               # Project documentation
├── requirements.txt        # Python dependencies
├── excel_files/            # Directory where input Excel files should be placed
└── reports/                # Directory where validation reports will be generated
```

## Installation

### Prerequisites

- Python 3.7 or higher (3.8+ recommended)
- pip package manager

### Quick Setup

```bash
# Clone or download the project
git clone https://github.com/castor-mark/Spacer.git  # or download and extract

# Navigate to project directory
cd Spacer

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

## Usage

1.  **Place Input Files**: Put all Excel files you wish to validate into the `excel_files/` directory.
2.  **Run the Validator**: Execute the `excel_validator.py` script. The script will guide you through the validation options.
3.  **View Reports**: After validation, detailed reports (Excel and CSV) will be generated in the `reports/` directory. Each run creates a timestamped subfolder, and a `latest/` symlink/shortcut will point to the most recent results.