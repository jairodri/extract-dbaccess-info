import sqlalchemy as sa
import sqlalchemy_access as sa_a
import pandas as pd


def get_db_info(db_path:str):

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

    # Crear un objeto MetaData
    metadata = sa.MetaData()

    # Reflejar la base de datos (obtener información sobre las tablas)
    metadata.reflect(bind=engine)

    # Crear un DataFrame para cada tabla
    dataframes = {}

    # Obtener y listar las tablas y sus columnas
    for table_name in metadata.tables:
    
        table = metadata.tables[table_name]
        # Lista para almacenar las columnas de la tabla
        columns_data = []

        for column in table.columns:
            # Obtener las características de la columna
            column_info = {
                "Column Name": column.name,
                "Data Type": str(column.type),
                "Nullable": column.nullable,
                "Primary Key": column.primary_key,
                "Default": column.default,
                "Unique": column.unique
            }
            # Añadir información adicional si es aplicable
            if hasattr(column.type, 'length'):
                column_info["Length"] = column.type.length

            # Añadir la información de la columna a la lista
            columns_data.append(column_info)

        # Crear un DataFrame de la tabla
        df = pd.DataFrame(columns_data)
        dataframes[table_name] = df

        # Crear un DataFrame de la tabla
        df = pd.DataFrame(columns_data)
        dataframes[table_name] = df

        # Mostrar el DataFrame de la tabla actual
        print(f"DataFrame for table '{table_name}':\n", df, "\n")

