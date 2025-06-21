# 🧠 Report-Generator-Bot

**Report-Generator-Bot** is a Python-based automation tool that streamlines the creation of both **Power BI** and **Excel reports**. It dynamically binds user-provided CSV datasets with pre-designed `.pbit` (Power BI template) and `.xlsx` (Excel template) files, enabling quick and consistent report generation without manual rework.

This tool is ideal for automating business intelligence workflows and generating recurring reports across multiple departments and industries.

---

## ✨ Key Features

- **Excel Report Automation**
  - Automatically populates a styled Excel template (`.xlsx`) with selected CSV data.
  - Generates reports with updated tables and chart placeholders.
  - One-click GUI interface to generate and save reports to the `/output` folder.

- **Power BI Report Automation**
  - Automatically replaces the dataset in a `.pbit` template with the selected CSV file.
  - Opens the report in Power BI Desktop, ready for refresh and saving as `.pbix`.
  - Eliminates repetitive tasks in business reporting.

- **Dual GUI Interfaces**
  - Two dedicated GUIs: one for Power BI, one for Excel.
  - Built using `customtkinter` for a modern desktop experience.
  - Real-time logging of actions and errors in the app window.

- **No Command Line Required**
  - Designed for users with zero coding experience.
  - Simple point-and-click operation for loading data and generating reports.

- **Lightweight & Configurable**
  - Just update the template files to fit your brand or report design.
  - Automatically saves reports in timestamped format under `/output`.

---

## 📁 Folder Structure
```plaintext
Report-Generator-Bot/
├── data/
│   └── input.csv                 # Updated dynamically via the GUI
├── output/
│   └── [Generated reports]
├── template/
│   ├── generate_dash.pbit       # Power BI Template
│   └── sales_template.xlsx      # Excel Report Template
├── generate_powerbi_report.py   # Power BI Report Generator GUI
├── generate_excel_report.py     # Excel Report Generator GUI
└── README.md

```
---

## 🚀 How to Use
### Generate Power BI Report
 ```bash
 python generate_powerbi_report.py
```


### Generate Excel Report
```bash
- python generate_excel_report.py
```
### Prerequisites
- Python 3.8 or higher
- Power BI Desktop installed (for Power BI automation)
- Required Python libraries:
  ```bash
  pip install pandas openpyxl customtkinter

##  Use Cases
- Automate weekly/monthly reports without rebuilding visuals
- Sales & Marketing dashboards with different regional data
- Finance & Accounting month-end closures
- Healthcare or academic reporting with standardized layouts
- Inventory or operations reports with minimal human effort



## Notes
- Ensure your `.pbit` file references the correct path: `data/input.csv`
- Excel templates must have headers matching the structure of your CSV files
- Reports are automatically saved with a timestamp to avoid overwriting
- Close Excel/Power BI reports before re-generating to prevent save conflicts



