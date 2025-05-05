import pandas as pd
import os
import warnings

def merge_specific_excel_files(directory_path, file1, column1, file2, column2):
    # Construct full file paths
    file_path1 = os.path.join(directory_path, file1)
    file_path2 = os.path.join(directory_path, file2)

    try:
        # Read the specified Excel files into DataFrames
        df1 = pd.read_excel(file_path1)
        df2 = pd.read_excel(file_path2)
        print(f"\n================================\nSuccessfully read files: {file1} and {file2}")

        # Merge the DataFrames on the specified columns
        merged_data = pd.merge(df1, df2, left_on=column1, right_on=column2, how='outer')
        print(f"\n================================\nMerged {file1} and {file2} on columns: {column1} and {column2}")

        return merged_data, df1, df2

    except Exception as e:
        print(f"Error processing files: {e}")
        return None, None, None

def main():
    # Suppress OpenPyXL warnings
    warnings.simplefilter(action='ignore', category=SyntaxWarning)
    
    # Get the directory path from the user
    directory_path = input("Enter the path to the directory containing Excel files (e.g., D:\\Desktop\\Folder\\): ")

    # Get the first file name and column name from the user
    file1 = input("\n\nEnter the name of the first Excel file to merge (e.g., GW Master.xlsx): ")
    column1 = input(f"\nEnter the column name in {file1} to merge on: ")

    # Get the second file name and column name from the user
    file2 = input("\n\nEnter the name of the second Excel file to merge (e.g., loading.xlsx): ")
    column2 = input(f"\nEnter the column name in {file2} to merge on: ")

    # Merge the specified files
    merged_data, df1, df2 = merge_specific_excel_files(directory_path, file1, column1, file2, column2)

    # Check if merging was successful
    if merged_data is not None and not merged_data.empty:
        # Ask user for columns to include in the output
        print("\nAvailable columns in the merged data:")
        print(merged_data.columns.tolist())

        column_selection = input("\nEnter the columns you want in the output file (comma-separated) or type 'All' to include all columns: ")

        if column_selection.strip().lower() == 'all':
            selected_columns = merged_data.columns.tolist()
        else:
            selected_columns = [col.strip() for col in column_selection.split(',') if col.strip() in merged_data.columns.tolist()]

        # Filter the merged data to include only the selected columns
        merged_data = merged_data[selected_columns]

        # Ask for output file format
        output_format = input("\nEnter the desired output file format (.xlsx or .csv): ").strip().lower()
        output_file = input("\nEnter the name for the output file (without extension): ")

        if output_format == '.xlsx':
            output_file += '.xlsx'
            merged_data.to_excel(output_file, index=False)
            print(f"\nMerged data saved to {output_file}\n================================\n")
        elif output_format == '.csv':
            output_file += '.csv'
            merged_data.to_csv(output_file, index=False)
            print(f"\nMerged data saved to {output_file}\n================================\n")
        else:
            print("Invalid output format specified. Please use .xlsx or .csv.")
    else:
        print("\nNo data was merged or an error occurred.\n================================\n")

if __name__ == "__main__":
    main()
