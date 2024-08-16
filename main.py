from getdbinfo import get_db_info_metadata, get_db_info_data
from dumpdbinfo import dump_db_info_to_csv, dump_db_info_to_excel

import os
from dotenv import load_dotenv


# Load variables from the .env file
load_dotenv()

# Get the database path from the environment variable
access_db = os.getenv('ACCESS_DB_PATH')

# Get the directory where the CSV files will be saved
output_dir_metadata = os.getenv('OUTPUT_DIR_METADATA')
output_dir_data = os.getenv('OUTPUT_DIR_DATA')

if __name__ == '__main__':
    db_name, table_df_metadata = get_db_info_metadata(access_db)
    dump_db_info_to_csv(db_name, table_df_metadata, output_dir_metadata, sep='|')
    dump_db_info_to_excel(db_name, table_df_metadata, output_dir_metadata)
    db_name, table_df_data = get_db_info_data(access_db)
    dump_db_info_to_csv(db_name, table_df_data, output_dir_data, sep='|')

