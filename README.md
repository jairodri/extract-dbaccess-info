
# Extract DB Access Info

This repository provides tools to extract metadata and data from Microsoft Access databases using Python, Pandas, and SQLAlchemy. The extracted information is organized into pandas DataFrames, which can then be processed to generate CSV or Excel files.

## Features

- Extracts metadata about tables and their columns from a Microsoft Access database.
- Retrieves all data from tables in the database.
- Exports metadata and data to CSV or Excel formats.

## Main Functions

### `getdbinfo.py`

This file contains functions to extract metadata and data from a Microsoft Access database.

#### `get_db_info_metadata(db_path: str) -> dict`

Extracts metadata from a Microsoft Access database and returns it in a dictionary format.

- **Parameters:**
  - `db_path` (str): The file path to the Microsoft Access database.
  
- **Returns:**
  - A dictionary with the database name as the key and another dictionary as the value. The inner dictionary contains table names as keys and pandas DataFrames as values, where each DataFrame includes detailed metadata about the table's columns (e.g., column name, data type, primary key, etc.).

#### `get_db_info_data(db_path: str) -> dict`

Extracts all data from the tables of a Microsoft Access database and returns it in a dictionary format.

- **Parameters:**
  - `db_path` (str): The file path to the Microsoft Access database.

- **Returns:**
  - A dictionary with the database name as the key and another dictionary as the value. The inner dictionary contains table names as keys and pandas DataFrames as values, where each DataFrame includes the data from the corresponding table.

### `dumpdbinfo.py`

This file contains functions to export the extracted data and metadata to CSV and Excel files.

#### `dump_db_info_to_csv(db_info: dict, output_dir: str, sep: str = ',')`

Saves each DataFrame from the extracted database information into a CSV file, organized in a directory named after the database.

- **Parameters:**
  - `db_info` (dict): A dictionary containing the database name as the key and another dictionary as the value. The inner dictionary includes table names as keys and pandas DataFrames with the table's data.
  - `output_dir` (str): The directory where the CSV files will be saved.
  - `sep` (str, optional): Field delimiter for the output file. The default is a comma (`,`).

- **Returns:**
  - None. This function creates CSV files in the specified output directory.

#### `dump_db_info_to_excel(db_info: dict, output_dir: str, include_record_count: bool = False, max_records_per_table: int = 50000)`

Exports the extracted database information into an Excel workbook, with each table's data on a separate sheet.

- **Parameters:**
  - `db_info` (dict): A dictionary containing the database name as the key and another dictionary as the value. The inner dictionary includes table names as keys and pandas DataFrames with the table's data.
  - `output_dir` (str): The directory where the Excel file will be saved.
  - `include_record_count` (bool, optional): If `True`, adds a column to the index sheet showing the number of records in each table. Default is `False`.
  - `max_records_per_table` (int, optional): The maximum number of records to include per table in the Excel sheet. Default is 50,000.

- **Returns:**
  - None. This function creates an Excel file in the specified output directory.

## License

This project is licensed under the MIT License.  