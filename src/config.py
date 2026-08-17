import os
from dotenv import load_dotenv

load_dotenv(".env.local")

DATABASE_URL = os.getenv("DATABASE_URL")
DATABASE_USER = os.getenv("DATABASE_USERNAME")
DATABASE_PASSWORD = os.getenv("DATABASE_PASSWORD")
DATABASE_NAME = os.getenv("DATABASE_NAME", "ad2l_production")
DATABASE_PORT = int(os.getenv("DATABASE_PORT", 3306))

OPEN_DOTA_API_KEY = os.getenv("OPEN_DOTA_API_KEY")
