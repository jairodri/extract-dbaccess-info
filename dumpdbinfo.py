import pandas as pd
import os


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

