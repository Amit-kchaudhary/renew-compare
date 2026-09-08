import os
import glob
import pandas as pd

def main():
    source_dir = r"F:\targetpend\source"
    output_dir = r"F:\targetpend\output"

    # Find daily CSV files in the source directory (excluding temp and renewals files)
    csv_files = [f for f in glob.glob(os.path.join(source_dir, "*.csv"))
                 if "recent_renewal" not in os.path.basename(f) and "temp" not in f]
                 
    if not csv_files:
        print("Error: No daily CSV status files found in the source directory.")
        return

    # Pick the latest CSV file by modification time
    latest_file = max(csv_files, key=os.path.getmtime)
    print(f"Reading source file: {latest_file}")

    # Load the CSV
    df = pd.read_csv(latest_file)

    # Clean column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Map required columns case-insensitively
    cols_map = {c.lower(): c for c in df.columns}
    required_cols_lower = ['region', 'partner', 'outlet', 'onu status', 'user id']
    missing_cols = [col for col in required_cols_lower if col not in cols_map]
    
    if missing_cols:
        print(f"Error: Missing columns in CSV (case-insensitive search): {missing_cols}")
        return

    region_col = cols_map['region']
    partner_col = cols_map['partner']
    outlet_col = cols_map['outlet']
    onu_status_col = cols_map['onu status']
    user_id_col = cols_map['user id']

    # Prepare online/offline flag counts
    df['ONU Status Cleaned'] = df[onu_status_col].astype(str).str.strip().str.lower()
    df['online_count'] = (df['ONU Status Cleaned'] == 'online').astype(int)
    df['offline_count'] = (df['ONU Status Cleaned'] == 'offline').astype(int)

    # Group by Region, Partner, and Outlet (keeping NaN/missing outlets in output)
    grouped = df.groupby([region_col, partner_col, outlet_col], dropna=False).agg(
        Online_users=('online_count', 'sum'),
        Offline_users=('offline_count', 'sum'),
        Total_users=(user_id_col, 'count')
    ).reset_index()

    # Calculate Target: Total users count divided by 3 and rounded off to nearest integer
    grouped['Target'] = (grouped['Total_users'] / 3.0).round().astype(int)

    # Rename columns to match the exact schema requested by the user:
    grouped = grouped.rename(columns={
        region_col: 'region',
        partner_col: 'Partner',
        outlet_col: 'Outlet',
        'Online_users': 'Online users',
        'Offline_users': 'Offlince Users',  # Matches user's exact spelling 'Offlince Users'
        'Total_users': 'Total users',
        'Target': 'Target'
    })

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Save output to Excel
    output_filename = "user_status_summary.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    
    # Write to Excel
    grouped.to_excel(output_path, index=False)
    print(f"Successfully created summary Excel file: {output_path}")
    print(f"Total rows written: {len(grouped)}")

if __name__ == "__main__":
    main()
