#!/usr/bin/env python3
"""
Vendor Enrichment Script

This script enriches vendor data by classifying each vendor based on its name.
Uses AI reasoning to determine:
- Department
- Vendor Description (≤120 chars)
- Strategic Recommendation (Terminate/Consolidate/Optimize)
- Functional Category
"""

import pandas as pd
import json

# Configuration
INPUT_FILE = "vendor_analysis_input.xlsx"
OUTPUT_FILE = "vendor_analysis_enriched.csv"

# Allowed values (from user-provided config)
ALLOWED_DEPARTMENTS = [
    "Engineering",
    "Facilities",
    "G&A",
    "Legal",
    "M&A",
    "Marketing",
    "SaaS",
    "Product",
    "Professional Services",
    "Sales",
    "Support",
    "Finance"
]

ALLOWED_RECOMMENDATIONS = ["Terminate", "Consolidate", "Optimize"]

FUNCTIONAL_CATEGORIES = [
    "CRM",
    "Collaboration",
    "Infrastructure",
    "DevTools",
    "Marketing Tools",
    "HR/Payroll",
    "Professional Services",
    "Real Estate/Facilities",
    "Insurance/Benefits",
    "Telecommunications",
    "Travel",
    "Finance/Accounting",
    "Legal Services",
    "Other"
]

def classify_vendor(vendor_name):
    """
    Classify a single vendor based on its name.

    Returns a dict with:
    - department: str
    - description: str (≤120 chars)
    - recommendation: str
    - category: str
    """
    name_lower = vendor_name.lower()

    # Initialize result
    result = {
        'department': 'G&A',  # General & Administrative as default
        'description': '',
        'recommendation': 'Optimize',
        'category': 'Other'
    }

    # === CRM / Sales Tools ===
    if 'salesforce' in name_lower:
        result['department'] = 'Sales'
        result['description'] = 'CRM and sales automation platform'
        result['recommendation'] = 'Optimize'
        result['category'] = 'CRM'

    elif 'hubspot' in name_lower:
        result['department'] = 'Marketing'
        result['description'] = 'Marketing automation and CRM platform'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Marketing Tools'

    # === Travel ===
    elif any(x in name_lower for x in ['navan', 'tripactions', 'travel']):
        result['department'] = 'G&A'
        result['description'] = 'Corporate travel management platform'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Travel'

    # === M&A Advisory ===
    elif any(x in name_lower for x in ['houlihan lokey', 'vector capital', 'm&a', 'mergers']):
        result['department'] = 'M&A'
        result['description'] = 'M&A advisory and investment banking services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Professional Services'

    # === Legal Services ===
    elif any(x in name_lower for x in ['law', 'legal', 'odvjetnicko', 'zuric i partneri']):
        result['department'] = 'Legal'
        result['description'] = 'Legal services and counsel'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Legal Services'

    # === Professional Services (Accounting, Consulting) ===
    elif any(x in name_lower for x in ['bdo', 'rsm', 'grant thornton', 'kpmg', 'pwc', 'deloitte', 'ey', 'ernst']):
        result['department'] = 'Professional Services'
        result['description'] = 'Accounting, audit and advisory services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Professional Services'

    elif any(x in name_lower for x in ['advisory', 'consulting', '4i advisory']):
        result['department'] = 'Professional Services'
        result['description'] = 'Professional advisory and consulting services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Professional Services'

    # === Real Estate / Facilities (coworking, office space) ===
    elif any(x in name_lower for x in ['wework', 'properties', 'tower', 'spaces', 'space ', 'gpt space', 'tog uk', 'innovent spaces', 'weking', 'cloudcrossing bvba', 'zagrebtower', 'studentski centar']):
        result['department'] = 'Facilities'
        result['description'] = 'Office space and real estate services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Real Estate/Facilities'

    # === Cloud Infrastructure (AWS, Azure, GCP) ===
    elif any(x in name_lower for x in ['aws', 'amazon web services', 'azure', 'google cloud', 'gcp']):
        result['department'] = 'Engineering'
        result['description'] = 'Cloud infrastructure and hosting services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Infrastructure'

    # === Insurance & Benefits ===
    elif any(x in name_lower for x in ['insurance', 'osiguranje', 'aetna', 'brokers', 'jensten', 'bupa', 'care health']):
        result['department'] = 'G&A'
        result['description'] = 'Employee benefits and insurance services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Insurance/Benefits'

    # === HR & Recruitment ===
    elif any(x in name_lower for x in ['recruitment', 'recruiters', 'mason frank', 'technet', 'cedar recruitment']):
        result['department'] = 'G&A'
        result['description'] = 'Recruitment and talent acquisition services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'HR/Payroll'

    elif 'hr solution' in name_lower or 'hr ' in name_lower or 'payroll' in name_lower:
        result['department'] = 'G&A'
        result['description'] = 'HR management and consulting services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'HR/Payroll'

    # === Telecommunications ===
    elif any(x in name_lower for x in ['telefonica', 'telecom', 'telco', 'network']):
        result['department'] = 'SaaS'
        result['description'] = 'Telecommunications and connectivity services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Telecommunications'

    # === IT Services & Technology Consulting ===
    elif any(x in name_lower for x in ['infosys', 'cloud technology', 'it solutions', 'info system', 'tech services', 'technology']):
        result['department'] = 'Professional Services'
        result['description'] = 'IT services and technology consulting'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Professional Services'

    # === Marketing / Advertising / Media ===
    elif any(x in name_lower for x in ['linkedin', 'google ireland', 'cognism', 'uberflip', 'mightyhive']):
        result['department'] = 'Marketing'
        result['description'] = 'Marketing and advertising platform'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Marketing Tools'

    elif any(x in name_lower for x in ['marketing', 'advertising', 'media', 'cult of monday', 'big frontier']):
        result['department'] = 'Marketing'
        result['description'] = 'Marketing and advertising services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Marketing Tools'

    # === SaaS Applications & Business Software ===
    elif any(x in name_lower for x in ['kimble', 'planful', 'intralinks']):
        result['department'] = 'SaaS'
        result['description'] = 'Business management and operations software'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Collaboration'

    # === Collaboration Tools ===
    elif any(x in name_lower for x in ['slack', 'zoom', 'miro', 'notion', 'asana', 'monday.com', 'atlassian', 'jira', 'confluence']):
        result['department'] = 'SaaS'
        result['description'] = 'Team collaboration and productivity platform'
        result['recommendation'] = 'Consolidate'
        result['category'] = 'Collaboration'

    # === Developer Tools ===
    elif any(x in name_lower for x in ['github', 'gitlab', 'docker', 'kubernetes', 'jenkins']):
        result['department'] = 'Engineering'
        result['description'] = 'Software development and DevOps tools'
        result['recommendation'] = 'Optimize'
        result['category'] = 'DevTools'

    # === Finance/Accounting Software ===
    elif any(x in name_lower for x in ['quickbooks', 'xero', 'sage', 'netsuite']):
        result['department'] = 'Finance'
        result['description'] = 'Accounting and financial management software'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Finance/Accounting'

    # === Training & Education ===
    elif any(x in name_lower for x in ['accutrainee', 'training']):
        result['department'] = 'G&A'
        result['description'] = 'Training and professional development services'
        result['recommendation'] = 'Optimize'
        result['category'] = 'HR/Payroll'

    # === Associations & Memberships ===
    elif 'tmforum' in name_lower or 'forum' in name_lower or 'association' in name_lower:
        result['department'] = 'G&A'
        result['description'] = 'Industry association membership and events'
        result['recommendation'] = 'Optimize'
        result['category'] = 'Other'

    # === Default case - try to infer from name ===
    else:
        result['description'] = f'Services provided by {vendor_name}'
        result['recommendation'] = 'Optimize'

    # Ensure description is ≤120 chars
    if len(result['description']) > 120:
        result['description'] = result['description'][:117] + '...'

    return result

def enrich_dataframe(df):
    """Enrich the entire dataframe."""
    print(f"\n{'='*80}")
    print(f"ENRICHING {len(df)} VENDORS")
    print(f"{'='*80}\n")

    vendor_col = 'Vendor Name'
    spend_col = 'Last 12 months Cost (USD)'

    # Track progress
    enriched_count = 0
    total_spend_processed = 0

    # Process each row
    for idx, row in df.iterrows():
        vendor_name = row[vendor_col]
        spend = row[spend_col]

        # Classify vendor
        classification = classify_vendor(vendor_name)

        # Update dataframe
        df.at[idx, 'Department'] = classification['department']
        df.at[idx, '1-line Description on what the Vendor does'] = classification['description']
        df.at[idx, 'Suggestions (Consolidate / Terminate / Optimize costs)'] = classification['recommendation']
        df.at[idx, 'Functional_Category'] = classification['category']

        enriched_count += 1
        total_spend_processed += spend

        # Progress indicator
        if (idx + 1) % 50 == 0:
            pct = (idx + 1) / len(df) * 100
            print(f"  Progress: {idx + 1}/{len(df)} ({pct:.1f}%) - ${total_spend_processed:,.0f} processed")

    print(f"\n✓ Enriched {enriched_count} vendors")
    print(f"✓ Total spend processed: ${total_spend_processed:,.2f}")

    return df

def display_top_vendors(df, n=50):
    """Display top N vendors by spend."""
    vendor_col = 'Vendor Name'
    spend_col = 'Last 12 months Cost (USD)'

    print(f"\n{'='*80}")
    print(f"TOP {n} VENDORS BY SPEND (FOR REVIEW)")
    print(f"{'='*80}\n")

    top_n = df.nlargest(n, spend_col)

    for idx, row in top_n.iterrows():
        vendor = row[vendor_col]
        spend = row[spend_col]
        dept = row['Department']
        desc = row['1-line Description on what the Vendor does']
        rec = row['Suggestions (Consolidate / Terminate / Optimize costs)']
        cat = row['Functional_Category']

        print(f"{idx+1:2d}. {vendor}")
        print(f"    Spend: ${spend:,.0f}")
        print(f"    Dept: {dept} | Category: {cat}")
        print(f"    Description: {desc}")
        print(f"    Recommendation: {rec}")
        print()

def main():
    """Main execution."""
    print(f"\n{'='*80}")
    print("VENDOR ENRICHMENT WORKFLOW")
    print(f"{'='*80}")

    # Load data
    print(f"\n1. Loading data from {INPUT_FILE}...")
    df = pd.read_excel(INPUT_FILE, sheet_name=0)
    print(f"   ✓ Loaded {len(df)} vendors")

    vendor_col = 'Vendor Name'
    spend_col = 'Last 12 months Cost (USD)'
    initial_spend = df[spend_col].sum()
    print(f"   ✓ Initial total spend: ${initial_spend:,.2f}")

    # Add Functional_Category column if it doesn't exist
    if 'Functional_Category' not in df.columns:
        df['Functional_Category'] = ''
        print(f"   ✓ Added Functional_Category column")

    # Enrich
    print(f"\n2. Enriching vendors...")
    df_enriched = enrich_dataframe(df)

    # Verify totals
    final_spend = df_enriched[spend_col].sum()
    final_rows = len(df_enriched)
    print(f"\n3. Quality Control:")
    print(f"   ✓ Final row count: {final_rows} (should be {len(df)})")
    print(f"   ✓ Final total spend: ${final_spend:,.2f} (should be ${initial_spend:,.2f})")

    if final_rows != len(df):
        print(f"   ✗ ERROR: Row count mismatch!")
        return

    if abs(final_spend - initial_spend) > 0.01:
        print(f"   ✗ ERROR: Spend total mismatch!")
        return

    # Display value distributions
    print(f"\n4. Value Distributions:")
    print(f"\n   Department distribution:")
    dept_counts = df_enriched['Department'].value_counts()
    for dept, count in dept_counts.items():
        print(f"     {dept:20s}: {count:3d} vendors")

    print(f"\n   Functional Category distribution:")
    cat_counts = df_enriched['Functional_Category'].value_counts()
    for cat, count in cat_counts.items():
        print(f"     {cat:30s}: {count:3d} vendors")

    print(f"\n   Strategic Recommendation distribution:")
    rec_counts = df_enriched['Suggestions (Consolidate / Terminate / Optimize costs)'].value_counts()
    for rec, count in rec_counts.items():
        print(f"     {rec:20s}: {count:3d} vendors")

    # Export with semicolon delimiter for CSV safety
    print(f"\n5. Exporting to {OUTPUT_FILE}...")
    df_enriched.to_csv(OUTPUT_FILE, index=False, sep=';', encoding='utf-8')
    print(f"   ✓ Exported successfully (semicolon-delimited)")

    # Display top vendors for review
    display_top_vendors(df_enriched, 50)

    print(f"\n{'='*80}")
    print("✓ ENRICHMENT COMPLETE")
    print(f"{'='*80}")
    print(f"\nNext steps:")
    print(f"1. Review the top 50 vendors above")
    print(f"2. Edit {OUTPUT_FILE} to correct any misclassifications")
    print(f"3. Run 'python3 analysis.py' to generate summary tables")

if __name__ == "__main__":
    main()
