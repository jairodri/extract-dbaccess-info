import sqlalchemy as sa
import sqlalchemy_access as sa_a
import pandas as pd
import os


def get_db_info_metadata(db_path:str):
    """
    Extracts metadata information from a Microsoft Access database and returns it as a dictionary
    where the key is the database name and the value is another dictionary containing table names 
    and their corresponding DataFrame with columns' metadata.

    This function creates an SQLAlchemy engine using an ODBC connection string to connect to a Microsoft Access 
    database specified by `db_path`. It then reflects the database schema to retrieve metadata about the tables 
    in the database. For each table, it gathers detailed information about its columns, including name, data type, 
    nullability, primary key status, default value, uniqueness, index presence, and any comments. 

    Parameters:
    -----------
    db_path : str
        The file path to the Microsoft Access database.

    Returns:
    --------
    dict: A dictionary where the key is the name of the database and the value is another dictionary 
    where each key is a table name from the database and each value is a DataFrame with the columns' 
    metadata for that table.
    """

    # Extract the database name from the db_path
    db_name = os.path.splitext(os.path.basename(db_path))[0]

    #
    # Create the SQLAlchemy engine with an ODBC connection string
    # Ordinary unprotected Access database
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

    # Wrap the table_dataframes in a new dictionary with db_name as the key
    db_info_dict = {db_name: table_dataframes}

    return db_info_dict


def get_db_info_data(db_path:str):
    """
    Extracts data from a Microsoft Access database and returns it as a dictionary
    where the key is the database name and the value is another dictionary containing table names 
    and their corresponding DataFrame with columns' data.

    This function creates an SQLAlchemy engine using an ODBC connection string to connect to a Microsoft Access 
    database specified by `db_path`. It then reflects the database schema to retrieve metadata about the tables 
    in the database. For each table, it retrieves all data. 

    Parameters:
    -----------
    db_path : str
        The file path to the Microsoft Access database.

    Returns:
    --------
    dict: A dictionary where the key is the name of the database and the value is another dictionary 
    where each key is a table name from the database and each value is a DataFrame with the columns' 
    data for that table.
    """

    # Extract the database name from the db_path
    db_name = os.path.splitext(os.path.basename(db_path))[0]

    #
    # Create the SQLAlchemy engine with an ODBC connection string
    # Ordinary unprotected Access database
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

    # Create a dictionary to store DataFrames
    table_data = {}

    # Iterate over each table and load its data into a DataFrame
    for table_name in metadata.tables:
        # Query all data from the table
        query = sa.select(metadata.tables[table_name])
        df = pd.read_sql(query, engine)
        
        # Store the DataFrame in the dictionary
        table_data[table_name] = df

    # Return the dictionary of DataFrames
    # Wrap the table_dataframes in a new dictionary with db_name as the key
    db_data_dict = {db_name: table_data}

    return db_data_dict

