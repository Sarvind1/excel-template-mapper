# Excel Template Mapper

Automates the population of shipping and procurement Excel templates (SCI and PL formats) from inbound product list source data. Maps fields from source data sheets, generates Excel formulas for automatic population, and identifies unmappable fields.

## Features

- **Automated Field Mapping**: Maps columns from source "YW1 Inbound PL" and "Container" sheets to template fields
- **Formula Generation**: Creates Excel formulas to automatically populate template fields with source data
- **Unmappable Field Detection**: Identifies and reports fields that cannot be automatically mapped
- **Template Analysis**: Analyzes template structure and validates mapping coverage
- **Summary Reports**: Generates comprehensive reports of mapping results and formula applications

## Tech Stack

- Python 3.7+
- openpyxl (Excel file manipulation)
- pandas (data analysis)
- Standard library: json, re, collections

## Setup

1. **Create a virtual environment**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

2. **Install dependencies**:
   ```bash
   pip install openpyxl pandas
   ```

3. **Prepare input files**:
   - Place your "YW1 Inbound PL.xlsx" workbook in the project root
   - Ensure it contains sheets: "YW1 Inbound PL", "Container", "SCI Template - Single Batch", "PL Template - Single Batch"

## Usage

Run the mapping and formula generation workflow:

```bash
# 1. Analyze templates and map source columns
python map_and_generate_formulas.py

# 2. Add formulas to templates
python add_formulas_to_templates.py

# 3. Generate summary report
python create_summary_report.py
```

Each script outputs progress and analysis results to the console. Generated Excel files will include formulas automatically populating template fields from source data.

## Project Structure

- `map_and_generate_formulas.py` - Maps source data columns to template fields
- `add_formulas_to_templates.py` - Injects Excel formulas into templates
- `create_summary_report.py` - Generates comprehensive mapping reports
- `analyze_templates.py` - Analyzes template structure and fields
- `detailed_template_review.py` - Detailed structural analysis
- `explore_excel.py` - Explores workbook layout and formula discovery