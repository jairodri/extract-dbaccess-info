from getdbinfo import get_db_info
import os
from dotenv import load_dotenv


# Load variables from the .env file
load_dotenv()

# Get the database path from the environment variable
access_db = os.getenv('ACCESS_DB_PATH')

if __name__ == '__main__':
    get_db_info(access_db)
