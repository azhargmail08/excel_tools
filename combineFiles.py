import pandas as pd
import os
import glob

def combine_excel_files():
    print("=== Excel Files Combiner (.xlsx only) ===")
    
    # Get directory path containing Excel files
    directory_path = input("Enter the directory path containing .xlsx files: ").strip().strip('"').strip("'")
    
    # Validate directory
    if not os.path.isdir(directory_path):
        print(f"Error: '{directory_path}' is not a valid directory")
        return
    
    # Find all .xlsx files in the directory
    excel_files = glob.glob(os.path.join(directory_path, '*.xlsx'))
    
    if not excel_files:
        print(f"No .xlsx files found in {directory_path}")
        return
        
    print(f"Found {len(excel_files)} .xlsx files:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {os.path.basename(file)}")
    
    # Ask if a specific column should be added to identify source files
    add_source_column = input("Add a column to identify source files? (y/n): ").lower() == 'y'
    
    # Combine all files
    all_dfs = []
    for file_path in excel_files:
        try:
            # Read the file
            df = pd.read_excel(file_path)
            
            # Add source file information if requested
            if add_source_column:
                df['Source_File'] = os.path.basename(file_path)
            
            all_dfs.append(df)
            print(f"✓ Successfully read {os.path.basename(file_path)} - {len(df)} rows")
            
        except Exception as e:
            print(f"× Error reading {os.path.basename(file_path)}: {str(e)}")
    
    if not all_dfs:
        print("No data could be read from the Excel files")
        return
    
    # Combine all dataframes
    combined_df = pd.concat(all_dfs, ignore_index=True)
    
    # Save combined data
    output_file = 'combined_files.xlsx'
    combined_df.to_excel(output_file, index=False)
    
    print(f"\nSuccessfully combined {len(all_dfs)} files into {output_file}")
    print(f"Total rows: {len(combined_df)}")

if __name__ == "__main__":
    combine_excel_files()