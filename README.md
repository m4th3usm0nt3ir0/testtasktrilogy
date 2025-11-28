# Vendor Spend Analysis Workflow

This repository contains a reproducible Python workflow for analyzing vendor spend from an acquisition. The workflow enriches vendor data with classifications and generates summary tables for savings analysis.

## Overview

- **Input**: `vendor_analysis_input.xlsx` (386 vendors, $7.89M total spend)
- **Output**: Enriched CSV files with classifications and summary tables ready for Google Sheets

## Quick Start

### 1. Enrich the Vendor Data

```bash
python3 enrich_vendors.py
```

This script:
- Loads the 386 vendors from the Excel file
- Classifies each vendor based on name patterns
- Adds 4 enrichment columns: Department, Description, Strategic Recommendation, Functional Category
- Exports `vendor_analysis_enriched.csv`
- Shows top 50 vendors by spend for review

**Quality Controls:**
- ✓ Same row count (386 in = 386 out)
- ✓ Same total spend ($7,887,360.40)
- ✓ No data creation or deletion

### 2. Generate Summary Tables

```bash
python3 analysis.py
```

This script:
- Loads the enriched CSV
- Generates summary tables by Department, Category, and Recommendation
- Identifies consolidation opportunities
- Exports 4 summary CSV files

## Output Files

### 1. vendor_analysis_enriched.csv
Full enriched dataset with all 386 vendors and these columns:
- Original columns: Vendor Name, Last 12 months Cost (USD)
- Enriched columns:
  - `Department`: Engineering | Facilities | G&A | Legal | M&A | Marketing | SaaS | Product | Professional Services | Sales | Support | Finance
  - `1-line Description on what the Vendor does`: ≤120 char description
  - `Suggestions (Consolidate / Terminate / Optimize costs)`: Strategic recommendation
  - `Functional_Category`: CRM | Collaboration | Cloud Infra | etc.

### 2. summary_by_department.csv
Spend breakdown by department with vendor counts and percentages.

Key insights:
- **Sales**: $3.12M (39.5%) - Primarily Salesforce
- **G&A**: $1.98M (25.1%) - 329 vendors (consolidation opportunity)
- **Facilities**: $1.22M (15.4%) - 12 vendors (real estate/office space)
- **Professional Services**: $0.74M (9.4%) - Consulting, advisory, IT services

### 3. summary_by_category.csv
Spend breakdown by functional category.

Key insights:
- **CRM**: $3.12M (39.5%) - Salesforce dominance
- **Real Estate/Facilities**: $1.22M (15.4%) - 12 vendors across multiple locations
- **Professional Services**: $0.81M (10.3%) - 13 vendors (BDO, RSM, etc.)
- **Travel**: $0.42M (5.3%) - Navan platforms

### 4. summary_by_recommendation.csv
Spend by strategic recommendation:
- **Optimize**: $7.87M (99.8%) - 383 vendors
- **Consolidate**: $13.7K (0.2%) - 3 vendors (collaboration tools)

### 5. consolidation_opportunities.csv
Categories with 2+ vendors, ranked by total spend.

**Top 3 Opportunities:**

1. **Real Estate/Facilities** - $1.22M across 12 vendors
   - Multiple office spaces in different locations
   - Potential to consolidate or renegotiate

2. **Professional Services** - $0.81M across 13 vendors
   - BDO, RSM, Grant Thornton, Infosys, etc.
   - Could consolidate advisory/consulting work

3. **Travel** - $0.42M across 3 vendors
   - Navan (TripActions) in 2 different entities
   - Clear consolidation candidate

## Workflow Architecture

```
vendor_analysis_input.xlsx (386 rows)
           ↓
    enrich_vendors.py
    - Classify by vendor name
    - Add 4 enrichment columns
    - Validate row count & spend totals
           ↓
vendor_analysis_enriched.csv (386 rows)
           ↓
      analysis.py
      - Load enriched data
      - Generate summaries
      - Identify consolidation opportunities
           ↓
    4 Summary CSV files
```

## Classification Logic

The enrichment uses pattern matching on vendor names to classify vendors:

- **Real Estate**: "properties", "tower", "spaces", "wework", etc.
- **Professional Services**: "bdo", "rsm", "grant thornton", "advisory", "consulting", etc.
- **Legal**: "law", "legal", "odvjetnicko", etc.
- **M&A**: "houlihan lokey", "vector capital", etc.
- **Insurance**: "insurance", "osiguranje", "aetna", "bupa", etc.
- **Marketing**: "linkedin", "hubspot", "google", "cognism", etc.
- **Travel**: "navan", "tripactions", etc.
- **SaaS**: "aws", "salesforce", "kimble", "planful", etc.

## Departments (from Config)

The allowed departments are:
- Engineering
- Facilities
- G&A (General & Administrative)
- Legal
- M&A
- Marketing
- SaaS
- Product
- Professional Services
- Sales
- Support
- Finance

## Manual Review & Editing

### Review Top 50 Vendors

After running `enrich_vendors.py`, review the displayed top 50 vendors by spend. These represent 90%+ of total spend.

### Edit Classifications

If you need to correct any classifications:

1. Edit `vendor_analysis_enriched.csv` directly
2. Update Department, Description, Recommendation, or Category columns
3. Re-run `python3 analysis.py` to regenerate summaries

**Important:** Do NOT change:
- Vendor names
- Spend amounts
- Row count

### Example Edits

```csv
Vendor Name,Department,Last 12 months Cost (USD),...
Salesforce Uk Ltd-Uk,Sales,3117225.89,...
```

Change Department from "Sales" to "SaaS" if needed, then re-run analysis.

## Re-running the Analysis

If you edit the enriched CSV or want to update classifications:

```bash
# Option 1: Re-enrich from scratch
python3 enrich_vendors.py

# Option 2: Manually edit vendor_analysis_enriched.csv, then:
python3 analysis.py
```

## Data Integrity Guarantees

✓ **No row creation**: 386 rows in → 386 rows out
✓ **No data deletion**: All vendors preserved
✓ **No spend changes**: $7,887,360.40 total maintained
✓ **Additive only**: Only adds classification columns

## Dependencies

```bash
pip install pandas openpyxl xlrd
```

## Next Steps for Google Sheets

1. Import all CSV files into Google Sheets
2. Create pivot tables and charts
3. Build "Top 3 Opportunities" dashboard
4. Share with stakeholders

## Files in This Repository

- `vendor_analysis_input.xlsx` - Original Excel export (386 vendors)
- `enrich_vendors.py` - Classification and enrichment script
- `analysis.py` - Summary table generation script
- `vendor_analysis_enriched.csv` - Enriched vendor data (output)
- `summary_by_department.csv` - Department summary (output)
- `summary_by_category.csv` - Functional category summary (output)
- `summary_by_recommendation.csv` - Recommendation summary (output)
- `consolidation_opportunities.csv` - Consolidation targets (output)
- `README.md` - This file

## Notes

- The classification is based on vendor names only, as the input file doesn't contain detailed service descriptions
- 313 vendors (40%) fall into "Other" category due to generic names
- Focus your manual review on the top 50 vendors, which represent 90%+ of spend
- The workflow is designed to be re-runnable: edit the CSV and regenerate summaries as needed
