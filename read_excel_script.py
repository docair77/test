import pandas as pd

def main():
    file_path = "your_file.xlsx"
    output_file_path = "comparison_results.xlsx"
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names[:2]

        if len(sheet_names) < 2:
            print(f"Error: The Excel file '{file_path}' must contain at least two sheets.")
            return

        original_sheet1_name = sheet_names[0]
        original_sheet2_name = sheet_names[1]
        print(f"Reading sheets: {original_sheet1_name}, {original_sheet2_name}")

        df1 = pd.read_excel(xls, sheet_name=original_sheet1_name)
        df2 = pd.read_excel(xls, sheet_name=original_sheet2_name)

        if df1.empty or df2.empty:
            print(f"Error: One or both sheets ('{original_sheet1_name}', '{original_sheet2_name}') are empty.")
            return

        if 'ID' not in df1.columns:
            print(f"Error: 'ID' column not found in sheet '{original_sheet1_name}'.")
            return
        if 'ID' not in df2.columns:
            print(f"Error: 'ID' column not found in sheet '{original_sheet2_name}'.")
            return

        ids1 = set(df1['ID'])
        ids2 = set(df2['ID'])

        common_ids = list(ids1.intersection(ids2))
        unmatched_sheet1_ids = list(ids1.difference(ids2))
        unmatched_sheet2_ids = list(ids2.difference(ids1))

        print(f"\nAnalysis of IDs:")
        print(f"Number of common IDs: {len(common_ids)}")
        print(f"Number of IDs in '{original_sheet1_name}' but not in '{original_sheet2_name}': {len(unmatched_sheet1_ids)}")
        print(f"Number of IDs in '{original_sheet2_name}' but not in '{original_sheet1_name}': {len(unmatched_sheet2_ids)}")

        # --- Data for "sheet 1" (Common IDs) ---
        df1_common = df1[df1['ID'].isin(common_ids)].copy()
        df2_common = df2[df2['ID'].isin(common_ids)].copy()
        merged_common_df = pd.merge(df1_common, df2_common, on='ID', suffixes=('', '_sheet2'))

        # --- Data for "sheet 2" (Unmatched IDs) ---
        df1_unmatched = df1[df1['ID'].isin(unmatched_sheet1_ids)].copy()
        df1_unmatched['SourceSheet'] = original_sheet1_name

        df2_unmatched = df2[df2['ID'].isin(unmatched_sheet2_ids)].copy()
        df2_unmatched['SourceSheet'] = original_sheet2_name

        # Concatenate the unmatched data
        # Ensure columns are aligned; pandas.concat handles this if column names are the same.
        # If column names differ (e.g. 'Product' vs 'Customer'), unmatched rows will have NaNs for columns not present in their original sheet.
        # This is generally desired for this kind of report.
        all_unmatched_df = pd.concat([df1_unmatched, df2_unmatched], ignore_index=True, sort=False)


        # Create/Overwrite the Excel workbook and write both sheets
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            merged_common_df.to_excel(writer, sheet_name='sheet 1', index=False)
            print(f"\nSuccessfully wrote merged data for common IDs to 'sheet 1' in '{output_file_path}'.")

            all_unmatched_df.to_excel(writer, sheet_name='sheet 2', index=False)
            print(f"Successfully wrote all unmatched ID data to 'sheet 2' in '{output_file_path}'.")

    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
