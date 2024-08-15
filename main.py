from getdbinfo import get_db_info
import os
from dotenv import load_dotenv


# Cargar variables desde el archivo .env
load_dotenv()

access_db = os.getenv('ACCESS_DB_PATH')
if __name__ == '__main__':
    get_db_info(access_db)
