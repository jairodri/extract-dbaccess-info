import pandas as pd
import os


def dump_db_info_to_csv(table_dataframes: dict, output_dir: str):
    """
    Saves each DataFrame in the provided dictionary to a CSV file.
    
    The CSV files are named after the corresponding table names. 
    Each DataFrame is saved to a separate CSV file in the specified output directory.

    Parameters:
    -----------
    table_dataframes : dict of pandas.DataFrame
        A dictionary where each key is a table name and each value is a DataFrame containing the table's column metadata.
    
    output_dir : str
        The directory where the CSV files will be saved.
    
    Returns:
    --------
    None
    """
    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Iterate over the dictionary to save each DataFrame to a CSV file
    for table_name, dataframe in table_dataframes.items():
        # Create the CSV file path
        file_path = os.path.join(output_dir, f"{table_name}.csv")
        
        # Save the DataFrame to a CSV file
        dataframe.to_csv(file_path, index=False)
        
        print(f"Saved {table_name} to {file_path}")
