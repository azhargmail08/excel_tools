import pandas as pd
import os

def separate_csv():
    # Get user input
    csv_path = input("Enter the path to your CSV file: ").strip().strip('"').strip("'")
    column_name = input("Enter the column name to separate by: ")
    output_dir = input("Enter the path where you want to save the Excel files: ").strip().strip('"').strip("'")

    try:
        # Read CSV file
        df = pd.read_csv(csv_path)
        
        # Check if column exists
        if column_name not in df.columns:
            print(f"Error: Column '{column_name}' not found in the CSV file")
            return

        # Check if directory exists, if not ask to create it
        if not os.path.exists(output_dir):
            create_dir = input(f"Directory '{output_dir}' doesn't exist. Create it? (y/n): ")
            if create_dir.lower() == 'y':
                os.makedirs(output_dir)
            else:
                print("Operation cancelled.")
                return

        # Group by specified column and save to separate files
        for value, group in df.groupby(column_name):
            # Clean filename by removing invalid characters
            clean_value = str(value).replace('/', '_').replace('\\', '_')
            output_path = os.path.join(output_dir, f"{clean_value}.xlsx")
            group.to_excel(output_path, index=False)
            print(f"Created: {output_path}")

        print("Separation completed successfully!")

    except FileNotFoundError:
        print("Error: CSV file not found")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    separate_csv()
