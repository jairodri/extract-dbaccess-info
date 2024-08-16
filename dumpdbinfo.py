import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, colors
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.hyperlink import Hyperlink


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


def create_hyperlink(ws, at_cell, sheet_name, cell_ref='A1', display_name=None):
    """
    Creates a hyperlink in a specified cell that links to another cell within the same workbook.

    This function adds a hyperlink to a specified cell in an Excel worksheet (`ws`). The hyperlink points 
    to a cell within another sheet (or the same sheet) within the same workbook. The cell containing the 
    hyperlink is formatted with a blue, underlined font to resemble a standard hyperlink.

    Parameters:
    -----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        The worksheet where the hyperlink will be created.
    
    at_cell : str
        The cell reference (e.g., 'B2') where the hyperlink will be placed in the `ws` worksheet.
    
    sheet_name : str
        The name of the sheet to which the hyperlink will point.
    
    cell_ref : str, optional
        The cell reference within the `sheet_name` sheet that the hyperlink will point to. Default is 'A1'.
    
    display_name : str, optional
        The text to be displayed in the cell containing the hyperlink. If not provided, defaults to the `sheet_name`.

    Returns:
    --------
    None
    """
    if display_name is None:
        display_name = sheet_name
    to_location = "'{0}'!{1}".format(sheet_name, cell_ref)
    ws[at_cell].hyperlink = Hyperlink(display=display_name, ref=at_cell, location=to_location)
    ws[at_cell].value = display_name
    ws[at_cell].font = Font(u='single', color=colors.BLUE)


def dump_db_info_to_excel(db_name: str, table_dataframes: dict, output_dir: str):
    """
    Exports metadata from a database to an Excel workbook with a separate sheet for each table's metadata.

    This function generates an Excel workbook where each table's metadata is stored in a separate sheet. 
    The first sheet, titled "Tables," serves as an index listing all table names, with hyperlinks to 
    their respective sheets for easy navigation.

    Each table sheet will include a hyperlink in the header row to return to the index sheet.

    Parameters:
    -----------
    db_name : str
        The name of the database, which will be used to name the output Excel file.
    
    table_dataframes : dict of pandas.DataFrame
        A dictionary where each key is a table name and each value is a DataFrame containing the table's column metadata.
    
    output_dir : str
        The directory where the Excel file will be saved. The function will ensure the directory structure 
        includes a folder named after the database.

    Returns:
    --------
    None
    """

    # Ensure the output directory includes the database name
    if not output_dir.endswith(db_name):
        output_dir = os.path.join(output_dir, db_name)

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Create the Excel workbook
    workbook = Workbook()

    # Default Excel font size if not specified
    standard_font_size = 11  # Default Excel font size if not specified

    # Use the default sheet as the index sheet
    index_sheet = workbook.active
    index_sheet.title = "Tables"
    
    # Add header to the index sheet
    index_sheet.cell(row=1, column=1, value="Table").font = Font(color="FFFFFF", bold=True, size=standard_font_size+1)
    index_sheet.cell(row=1, column=1).fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    # Populate the index sheet with table names
    for i, table_name in enumerate(table_dataframes.keys(), start=2):
        index_sheet.cell(row=i, column=1, value=table_name)
    
    # Adjust column width for the index sheet
    max_length = max(len(table_name) for table_name in table_dataframes.keys()) + 2
    index_sheet.column_dimensions['A'].width = max_length

    # Apply a filter to all columns
    index_sheet.auto_filter.ref = index_sheet.dimensions

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
                    cell.font = Font(color="FFFFFF", bold=True, size=standard_font_size+1)
                    cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        
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

        # Add a hyperlink to return to the "Tables" sheet in the last cell of the header row
        last_col_idx = len(dataframe.columns) + 1
        return_cell = sheet.cell(row=1, column=last_col_idx)
        create_hyperlink(sheet, at_cell=return_cell.coordinate, sheet_name="Tables", cell_ref='A1', display_name="Return to Tables")

        # Ajustar el ancho de la columna para adaptarse al mensaje "Return to Tables"
        sheet.column_dimensions[return_cell.column_letter].width = len("Return to Tables") + 2


    # Links are created for each table in the list for easy access to its sheet.
    for i in range(2, index_sheet.max_row + 1):
        create_hyperlink(index_sheet, 'A' + str(i), index_sheet['A' + str(i)].value, cell_ref='A1')

    # Save the workbook to the output directory with the name of the database
    excel_file_path = os.path.join(output_dir, f"{db_name}.xlsx")
    workbook.save(excel_file_path)


