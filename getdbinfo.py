import sqlalchemy as sa
import sqlalchemy_access as sa_a
import pandas as pd


def get_db_info(db_path:str):

    #
    # Create the SQLAlchemy engine
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
    dataframes = {}

    # Iterate through the tables and retrieve their columns
    for table_name in metadata.tables:
        table = metadata.tables[table_name]

        # List to store the columns of the table
        columns_data = []

        for column in table.columns:
            # Get the characteristics of the column
            column_info = {
                "Column Name": column.name,
                "Data Type": str(column.type),
                "Nullable": column.nullable,
                "Primary Key": column.primary_key,
                "Default": column.default,
                "Unique": column.unique
            }
            # Add additional information if applicable
            if hasattr(column.type, 'length'):
                column_info["Length"] = column.type.length

            # Add the column information to the list
            columns_data.append(column_info)

        # Create a DataFrame for the table
        df = pd.DataFrame(columns_data)
        dataframes[table_name] = df

        # Print the DataFrame for the current table
        print(f"DataFrame for table '{table_name}':\n", df, "\n")

