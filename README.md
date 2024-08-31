# Kaspersky Report Generation Script

This script generates a report in Excel format based on data from various sources. It consists of two main components:

1. `kasperskyreportgen.py`: This script reads data from CSV and Excel files, performs data manipulation and merges the data into a single DataFrame. It then writes the merged data to an Excel file and appends EDR details to the same file. Additionally, it counts and prints the number of hosts with EDR requirements, EDR deployed, and EDR pending.

2. `xls_to_xlsx.sh`: This script converts XLS files to XLSX format using LibreOffice. It searches for files matching a specific pattern in a given directory and converts them to XLSX.

## Prerequisites

- Python 3.x
- LibreOffice (for running `xls_to_xlsx.sh`)
- The following Python packages:
  - pandas
  - openpyxl

## Usage

1. Clone the repository.
2. Install the required Python packages by running `pip install -r requirements.txt`.