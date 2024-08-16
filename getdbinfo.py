import sqlalchemy as sa
import sqlalchemy_access as sa_a
import pandas as pd


def get_db_info_metadata(db_path:str):

    """
    Extracts metadata information from a Microsoft Access database and returns it as a dictionary of pandas DataFrames.

    This function creates an SQLAlchemy engine using an ODBC connection string to connect to a Microsoft Access 
    database specified by `db_path`. It then reflects the database schema to retrieve metadata about the tables 
    in the database. For each table, it gathers detailed information about its columns, including name, data type, 
    nullability, primary key status, default value, uniqueness, index presence, and any comments. 

    Each table's column information is stored in a pandas DataFrame, and the function returns a dictionary where 
    the keys are table names and the values are the corresponding DataFrames containing the column details.

    Parameters:
    -----------
    db_path : str
        The file path to the Microsoft Access database.

    Returns:
    --------
    dict of pandas.DataFrame
        A dictionary where each key is a table name from the database and each value is a DataFrame with the columns' 
        metadata for that table.
    """

    #
    # Create the SQLAlchemy engine with an ODBC connection string
    #  
    driver = "{Microsoft Access Driver (*.mdb, *.accdb)}"
    connection_string = (
        f"DRIVER={driver};"
        f"DBQ={db_path};"
        f"ExtendedAnsiSQL=1;"
        )
    connection_url = sa.engine.URL.create(
        "access+pyodbc",
        query={"odbc_connect": connection_string}
        )
    engine = sa.create_engine(connection_url)
    #

    # Reflect the database schema (retrieve information about the tables)
    metadata = sa.MetaData()
    metadata.reflect(bind=engine)

    # Create a DataFrame for each table
    table_dataframes = {}

    # Iterate through the tables and retrieve their columns
    for table_name in metadata.tables:
        table = metadata.tables[table_name]
        
        # List to store the columns of the table
        columns_data = []

        for column in table.columns:
            # Get the characteristics of the column
            column_info = {
                "Column Name": column.name,
                "Data Type": column.type,
                "Nullable": column.nullable,
                "Primary Key": column.primary_key,
                "Default": column.default,
                "Unique": column.unique,
                "Index": column.index,
                "Comment": column.comment
            }
            # Add additional information if applicable
            if hasattr(column.type, 'length'):
                column_info["Length"] = column.type.length
            else:
                column_info["Length"] = None

            # Add the column information to the list
            columns_data.append(column_info)

        # Create a DataFrame for the table
        df = pd.DataFrame(columns_data)
        table_dataframes[table_name] = df

    return table_dataframes    

