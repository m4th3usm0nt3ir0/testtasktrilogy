#!/usr/bin/env python3
"""
Vendor Spend Analysis - Enrichment and Classification Workflow

This script:
1. Loads vendor spend data from vendor_analysis_input.xlsx
2. Reads department mappings from the config sheet
3. Enriches each vendor row with classification columns
4. Creates summary tables for savings analysis
5. Exports everything to CSV

Hard constraints:
- Do NOT create, delete, or duplicate vendors
- Do NOT change any numeric values
- 1 row in = 1 row out, same spend
- Use only department values from config sheet
"""

import pandas as pd
import numpy as np
import sys

# Configuration
INPUT_FILE = "vendor_analysis_input.xlsx"
OUTPUT_ENRICHED = "vendor_analysis_enriched.csv"
OUTPUT_BY_DEPARTMENT = "summary_by_department.csv"
OUTPUT_BY_CATEGORY = "summary_by_category.csv"
OUTPUT_CONSOLIDATION = "consolidation_opportunities.csv"

def load_data():
    """Load the Excel file and config sheet."""
    print("=" * 80)
    print("STEP 1: Loading data from Excel")
    print("=" * 80)

    # Read the main data sheet
    df = pd.read_excel(INPUT_FILE, sheet_name=0)
    print(f"\n✓ Loaded main data: {len(df)} rows")

    # Read the config sheet
    try:
        config_df = pd.read_excel(INPUT_FILE, sheet_name="config")
        print(f"✓ Loaded config sheet: {len(config_df)} rows")
    except Exception as e:
        print(f"✗ Could not load config sheet: {e}")
        config_df = None

    print(f"\n📋 Column names in main data:")
    for i, col in enumerate(df.columns, 1):
        print(f"  {i}. {col}")

    print(f"\n📊 First 10 rows of main data:")
    print(df.head(10).to_string())

    if config_df is not None:
        print(f"\n📋 Config sheet columns:")
        for i, col in enumerate(config_df.columns, 1):
            print(f"  {i}. {col}")
        print(f"\n📊 Config sheet data:")
        print(config_df.to_string())

    return df, config_df

def validate_data(df):
    """Validate and report on data quality."""
    print("\n" + "=" * 80)
    print("STEP 2: Data Validation")
    print("=" * 80)

    initial_rows = len(df)
    print(f"\n✓ Initial row count: {initial_rows}")

    # Try to identify spend column
    spend_cols = [col for col in df.columns if 'cost' in col.lower() or 'spend' in col.lower()]
    if spend_cols:
        spend_col = spend_cols[0]
        total_spend = df[spend_col].sum()
        print(f"✓ Identified spend column: '{spend_col}'")
        print(f"✓ Total spend: ${total_spend:,.2f}")
    else:
        print("⚠ Could not auto-identify spend column")
        spend_col = None
        total_spend = 0

    # Try to identify vendor name column
    vendor_cols = [col for col in df.columns if 'vendor' in col.lower() or 'name' in col.lower()]
    if vendor_cols:
        vendor_col = vendor_cols[0]
        print(f"✓ Identified vendor column: '{vendor_col}'")
    else:
        print("⚠ Could not auto-identify vendor column")
        vendor_col = None

    return {
        'initial_rows': initial_rows,
        'total_spend': total_spend,
        'spend_col': spend_col,
        'vendor_col': vendor_col
    }

def get_allowed_departments(config_df):
    """Extract allowed department values from config sheet."""
    if config_df is None:
        print("\n⚠ No config sheet found, using default departments")
        return ["Engineering", "Sales", "Marketing", "Operations", "Finance",
                "HR", "IT", "Product", "Customer Support", "Other"]

    # Try to find department column in config
    dept_cols = [col for col in config_df.columns if 'department' in col.lower()]
    if dept_cols:
        dept_col = dept_cols[0]
        departments = config_df[dept_col].dropna().unique().tolist()
        print(f"\n✓ Extracted {len(departments)} departments from config sheet:")
        for dept in sorted(departments):
            print(f"  - {dept}")
        return departments
    else:
        print("\n⚠ No department column found in config, using defaults")
        return ["Engineering", "Sales", "Marketing", "Operations", "Finance",
                "HR", "IT", "Product", "Customer Support", "Other"]

def prepare_vendor_batch(df, metadata, start_idx=0, batch_size=50):
    """Prepare a batch of vendors for enrichment."""
    vendor_col = metadata['vendor_col']
    spend_col = metadata['spend_col']

    batch_df = df.iloc[start_idx:start_idx + batch_size].copy()

    # Create a simple list for display
    batch_list = []
    for idx, row in batch_df.iterrows():
        vendor_name = row[vendor_col]
        spend = row[spend_col]
        batch_list.append(f"{vendor_name} (${spend:,.0f})")

    return batch_list, batch_df

def display_enrichment_instructions(df, metadata, allowed_departments):
    """Display instructions for manual enrichment."""
    print("\n" + "=" * 80)
    print("STEP 3: Vendor Enrichment Instructions")
    print("=" * 80)

    vendor_col = metadata['vendor_col']

    print(f"\n📊 Dataset Overview:")
    print(f"  - Total vendors: {len(df)}")
    print(f"  - Total spend: ${metadata['total_spend']:,.2f}")

    print(f"\n📋 Allowed Departments ({len(allowed_departments)}):")
    for dept in sorted(allowed_departments):
        print(f"  - {dept}")

    print(f"\n📝 Enrichment columns needed:")
    print(f"  1. Department: Choose from allowed list above")
    print(f"  2. Vendor_Description: ≤120 chars, concise description")
    print(f"  3. Strategic_Recommendation: Terminate | Consolidate | Optimize")
    print(f"  4. Functional_Category: CRM | Collaboration | Cloud Infra | DevTools |")
    print(f"                          Marketing Tools | HR/Payroll | Professional Services | Other")

    # Show top vendors by spend
    print(f"\n💰 Top 20 vendors by spend (for context):")
    top20 = df.nlargest(20, metadata['spend_col'])
    for idx, row in enumerate(top20.itertuples(), 1):
        vendor = getattr(row, vendor_col.replace(' ', '_'))
        spend = getattr(row, 'Last_12_months_Cost__USD_')
        print(f"  {idx:2d}. {vendor:50s} ${spend:>12,.0f}")

    print(f"\n⚠️  IMPORTANT:")
    print(f"  - You must enrich the data by editing the CSV file manually")
    print(f"  - OR create a script to do batch classification")
    print(f"  - The current script exports empty enrichment columns")
    print(f"  - Re-run this script after enrichment to generate summaries")

def enrich_data_placeholder(df, metadata, allowed_departments):
    """
    Placeholder for enrichment step.

    The existing Excel has these columns (currently empty):
    - Department
    - 1-line Description on what the Vendor does
    - Suggestions (Consolidate / Terminate / Optimize costs)

    We will map these and add:
    - Functional_Category (new column)
    """
    print("\n" + "=" * 80)
    print("STEP 3: Enrichment Preparation")
    print("=" * 80)

    vendor_col = metadata['vendor_col']

    print(f"\n✓ The input file already has these columns (currently empty):")
    print(f"  - Department")
    print(f"  - 1-line Description on what the Vendor does")
    print(f"  - Suggestions (Consolidate / Terminate / Optimize costs)")

    print(f"\n✓ We will add one additional column:")
    print(f"  - Functional_Category")

    # Add the missing column
    if 'Functional_Category' not in df.columns:
        df['Functional_Category'] = ''
        print(f"\n✓ Added Functional_Category column")

    # Ensure all enrichment columns exist and are empty
    enrichment_cols = {
        'Department': 'Department',
        '1-line Description on what the Vendor does': 'Vendor_Description',
        'Suggestions (Consolidate / Terminate / Optimize costs)': 'Strategic_Recommendation',
        'Functional_Category': 'Functional_Category'
    }

    print(f"\n✓ Row count remains: {len(df)} (unchanged)")
    print(f"\n⚠️  Enrichment columns are currently empty - awaiting classification")

    return df

def create_summaries(df, metadata):
    """Create summary tables for analysis."""
    print("\n" + "=" * 80)
    print("STEP 4: Summary Tables")
    print("=" * 80)

    spend_col = metadata['spend_col']
    vendor_col = metadata['vendor_col']
    dept_col = 'Department'
    cat_col = 'Functional_Category'
    rec_col = 'Suggestions (Consolidate / Terminate / Optimize costs)'

    # Summary by Department
    if dept_col in df.columns and spend_col:
        dept_summary = df.groupby(dept_col).agg({
            spend_col: 'sum',
            vendor_col: 'count'
        }).reset_index()
        dept_summary.columns = ['Department', 'Total_Spend_USD', 'Vendor_Count']
        dept_summary = dept_summary.sort_values('Total_Spend_USD', ascending=False)
        dept_summary['Percentage'] = (dept_summary['Total_Spend_USD'] / dept_summary['Total_Spend_USD'].sum() * 100).round(1)
        print(f"\n📊 Spend by Department:")
        print(dept_summary.to_string(index=False))
    else:
        dept_summary = pd.DataFrame()

    # Summary by Functional Category
    if cat_col in df.columns and spend_col:
        cat_summary = df.groupby(cat_col).agg({
            spend_col: 'sum',
            vendor_col: 'count'
        }).reset_index()
        cat_summary.columns = ['Functional_Category', 'Total_Spend_USD', 'Vendor_Count']
        cat_summary = cat_summary.sort_values('Total_Spend_USD', ascending=False)
        cat_summary['Percentage'] = (cat_summary['Total_Spend_USD'] / cat_summary['Total_Spend_USD'].sum() * 100).round(1)
        print(f"\n📊 Spend by Functional Category:")
        print(cat_summary.to_string(index=False))
    else:
        cat_summary = pd.DataFrame()

    # Summary by Strategic Recommendation
    if rec_col in df.columns and spend_col:
        rec_summary = df.groupby(rec_col).agg({
            spend_col: 'sum',
            vendor_col: 'count'
        }).reset_index()
        rec_summary.columns = ['Strategic_Recommendation', 'Total_Spend_USD', 'Vendor_Count']
        rec_summary = rec_summary.sort_values('Total_Spend_USD', ascending=False)
        rec_summary['Percentage'] = (rec_summary['Total_Spend_USD'] / rec_summary['Total_Spend_USD'].sum() * 100).round(1)
        print(f"\n📊 Spend by Strategic Recommendation:")
        print(rec_summary.to_string(index=False))
    else:
        rec_summary = pd.DataFrame()

    # Consolidation opportunities (categories with multiple vendors)
    if cat_col in df.columns and spend_col:
        consol = df[df[cat_col] != 'Other'].groupby(cat_col).agg({
            spend_col: 'sum',
            vendor_col: 'count'
        }).reset_index()
        consol.columns = ['Functional_Category', 'Total_Spend_USD', 'Vendor_Count']
        consol = consol[consol['Vendor_Count'] > 1].sort_values('Total_Spend_USD', ascending=False)
        print(f"\n📊 Consolidation Opportunities (categories with 2+ vendors, excluding 'Other'):")
        print(consol.to_string(index=False))

        # Detailed breakdown for top consolidation opportunities
        print(f"\n💡 Top Consolidation Opportunities (Detailed):")
        for idx, row in consol.head(5).iterrows():
            category = row['Functional_Category']
            vendors_in_cat = df[df[cat_col] == category]
            print(f"\n  {category}:")
            print(f"    Total Spend: ${row['Total_Spend_USD']:,.2f}")
            print(f"    Vendor Count: {row['Vendor_Count']}")
            print(f"    Vendors:")
            for _, v in vendors_in_cat.nlargest(10, spend_col).iterrows():
                print(f"      - {v[vendor_col]:50s} ${v[spend_col]:>12,.0f}")
    else:
        consol = pd.DataFrame()

    return dept_summary, cat_summary, rec_summary, consol

def export_data(df, dept_summary, cat_summary, rec_summary, consol, metadata):
    """Export all data to CSV files."""
    print("\n" + "=" * 80)
    print("STEP 5: Export to CSV")
    print("=" * 80)

    # Export enriched data (already done by enrich_vendors.py, but keep for consistency)
    # df.to_csv(OUTPUT_ENRICHED, index=False)
    print(f"✓ Enriched data already exported to: {OUTPUT_ENRICHED}")
    print(f"  - Rows: {len(df)}")
    print(f"  - Total spend: ${metadata['total_spend']:,.2f}")

    # Export summaries
    if not dept_summary.empty:
        dept_summary.to_csv(OUTPUT_BY_DEPARTMENT, index=False)
        print(f"✓ Exported department summary to: {OUTPUT_BY_DEPARTMENT}")

    if not cat_summary.empty:
        cat_summary.to_csv(OUTPUT_BY_CATEGORY, index=False)
        print(f"✓ Exported category summary to: {OUTPUT_BY_CATEGORY}")

    if not rec_summary.empty:
        rec_summary.to_csv("summary_by_recommendation.csv", index=False)
        print(f"✓ Exported recommendation summary to: summary_by_recommendation.csv")

    if not consol.empty:
        consol.to_csv(OUTPUT_CONSOLIDATION, index=False)
        print(f"✓ Exported consolidation opportunities to: {OUTPUT_CONSOLIDATION}")

def load_enriched_data():
    """Load the enriched CSV file (output from enrich_vendors.py)."""
    print("\n" + "=" * 80)
    print("STEP 1: Loading Enriched Data")
    print("=" * 80)

    import os
    if os.path.exists(OUTPUT_ENRICHED):
        df = pd.read_csv(OUTPUT_ENRICHED)
        print(f"\n✓ Loaded enriched data: {len(df)} rows from {OUTPUT_ENRICHED}")
        return df
    else:
        print(f"\n✗ Enriched file not found: {OUTPUT_ENRICHED}")
        print(f"   Please run: python3 enrich_vendors.py first")
        return None

def main():
    """Main execution flow."""
    print("\n" + "=" * 80)
    print("VENDOR SPEND ANALYSIS - SUMMARY GENERATION")
    print("=" * 80)

    # Step 1: Load enriched data
    df = load_enriched_data()
    if df is None:
        return

    # Step 2: Validate
    metadata = validate_data(df)

    # Step 3: Create summaries
    dept_summary, cat_summary, rec_summary, consol = create_summaries(df, metadata)

    # Step 4: Export
    export_data(df, dept_summary, cat_summary, rec_summary, consol, metadata)

    print("\n" + "=" * 80)
    print("✓ ANALYSIS COMPLETE")
    print("=" * 80)
    print("\nGenerated files:")
    print(f"  - {OUTPUT_ENRICHED} (enriched vendor data)")
    print(f"  - {OUTPUT_BY_DEPARTMENT} (spend by department)")
    print(f"  - {OUTPUT_BY_CATEGORY} (spend by functional category)")
    print(f"  - summary_by_recommendation.csv (spend by recommendation)")
    print(f"  - {OUTPUT_CONSOLIDATION} (consolidation opportunities)")
    print("\nYou can now:")
    print("  1. Review and edit the enriched CSV")
    print("  2. Re-run this script to update summaries after any edits")
    print("  3. Import the CSVs into Google Sheets for further analysis")

if __name__ == "__main__":
    main()
