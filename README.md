# Excel Template Mapper

An automated Excel data mapping and formula generation tool that transforms raw product and container data into standardized supply chain templates. Maps data from YW1 Inbound PL sheets into SCI (Shipping Confirmation Item) and PL (Packing List) templates with intelligent field matching and formula population.

## Features

- **Automated Field Mapping**: Analyzes source data headers and intelligently maps to template fields
- **Formula Generation**: Automatically populates template cells with Excel formulas that reference source data
- **Unmappable Field Detection**: Identifies fields that require manual intervention
- **Comprehensive Reporting**: Generates detailed analysis of mapping coverage and formula completeness
- **Multi-Sheet Support**: Works with multiple source sheets (YW1 Inbound PL, Container data)

## Tech Stack

- **Python 3.x**
- **openpyxl** - Excel file manipulation
- **pandas** - Data analysis and processing

## Setup

1. Create a virtual environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

2. Install dependencies:
   ```bash
   pip install openpyxl pandas
   ```

3. Prepare your source Excel file:
   - Place `YW1 Inbound PL.xlsx` in the project root
   - Ensure source sheets exist: "YW1 Inbound PL", "Container"
   - Ensure template sheets exist: "SCI Template - Single Batch", "PL Template - Single Batch"

## Usage

### 1. Initial Analysis
Explore your Excel structure:
```bash
python explore_excel.py
```

### 2. Analyze and Build Mappings
Map source columns to template fields:
```bash
python map_and_generate_formulas.py
```

### 3. Add Formulas to Templates
Populate templates with formulas:
```bash
python add_formulas_to_templates.py
```

### 4. Generate Summary Report
Create a comprehensive review:
```bash
python create_summary_report.py
```

The tool generates:
- `analysis_output.txt` - Detailed field-by-field analysis
- `detailed_review.txt` - Template structure review
- Modified Excel files with formulas populated
- `template_analysis.json` - Machine-readable mapping data

## Workflow

1. Run `explore_excel.py` to understand your workbook structure
2. Run `map_and_generate_formulas.py` to identify available mappings
3. Run `add_formulas_to_templates.py` to populate templates with formulas
4. Review `analysis_output.txt` and `detailed_review.txt` for unmappable fields
5. Manually handle unmappable fields as needed

## Output Files

- `YW1 Inbound PL_with_unmapped.xlsx` - Initial mapping with unmapped fields identified
- `YW1 Inbound PL_with_formulas.xlsx` - Final templates with all formulas populated
- Analysis reports for verification and debugging