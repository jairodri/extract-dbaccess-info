import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


def dump_db_info_to_csv(db_name: str, table_dataframes: dict, output_dir: str, sep: str=','):
    """
    Saves each DataFrame in the provided dictionary to a CSV file, organizing the files within a directory 
    named after the database.

    The CSV files are named after the corresponding table names. Each DataFrame is saved to a separate CSV file 
    in the specified output directory, which will include a subdirectory named after the database if not already included.

    Parameters:
    -----------
    db_name : str
        The name of the database, which will be used to create a subdirectory within the output directory.

    table_dataframes : dict of pandas.DataFrame
        A dictionary where each key is a table name and each value is a DataFrame containing the table's column metadata.
    
    output_dir : str
        The directory where the CSV files will be saved. If the directory does not contain a subdirectory with the 
        name of the database, one will be added.

    sep : str, optional
        Field delimiter for the output file. The default is a comma.

    Returns:
    --------
    None
    """
    # Ensure the output directory includes the database name
    if not output_dir.endswith(db_name):
        output_dir = os.path.join(output_dir, db_name)


    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Iterate over the dictionary to save each DataFrame to a CSV file
    for table_name, dataframe in table_dataframes.items():
        # Create the CSV file path
        file_path = os.path.join(output_dir, f"{table_name}.csv")
        
        # Save the DataFrame to a CSV file
        dataframe.to_csv(file_path, sep=sep, index=False)



def dump_db_info_to_excel(db_name: str, table_dataframes: dict, output_dir: str):

    # Ensure the output directory includes the database name
    if not output_dir.endswith(db_name):
        output_dir = os.path.join(output_dir, db_name)

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Create the Excel workbook
    workbook = Workbook()

    # Remove the default sheet created by openpyxl
    workbook.remove(workbook.active)
    
    # Default Excel font size if not specified
    standard_font_size = 11  # Default Excel font size if not specified

    for table_name, dataframe in table_dataframes.items():
        # Create a new sheet with the table name
        sheet = workbook.create_sheet(title=table_name)
        
        # Add the DataFrame to the sheet
        for r_idx, row in enumerate(dataframe_to_rows(dataframe, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                # Ensure the value is converted to a string if it's not a basic data type
                cell_value = str(value) if not isinstance(value, (int, float, type(None))) else value
                cell = sheet.cell(row=r_idx, column=c_idx, value=cell_value)
                if r_idx == 1:  # Apply formatting to header row
                    cell.font = Font(bold=True, size=standard_font_size+1)
                    cell.fill = PatternFill(end_color='00D9D9D9', start_color='00D9D9D9', fill_type='solid')
        
        # Auto-size columns based on the maximum width of the data and the header
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name

            # Calculate the width required by the header (considering formatting)
            header_length = len(str(col[0].value))
            adjusted_header_length = header_length * 1.5  # Factor to account for bold and larger font size

            # Compare the header length with the lengths of the data values
            for cell in col:
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            
            # Use the greater of the header length or data length for column width
            max_length = max(max_length, adjusted_header_length)

            # Adjust the column width
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width


        # Apply a filter to all columns
        sheet.auto_filter.ref = sheet.dimensions


    # Save the workbook to the output directory with the name of the database
    excel_file_path = os.path.join(output_dir, f"{db_name}.xlsx")
    workbook.save(excel_file_path)


