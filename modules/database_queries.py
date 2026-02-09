import os
import pandas as pd
import mysql.connector
from tkinter import Tk, Label, Entry, Button
from dotenv import load_dotenv

load_dotenv()

# Variables globales para la conexión
database="bi"

def execute_query(query):
    """Ejecuta una consulta SQL y devuelve un DataFrame."""
    global host, user, passwordß
    db_config = {
        "host": os.getenv("DB_HOST"),
        "user": os.getenv("DB_USER"),
        "password": os.getenv("DB_PASSWORD"),
        "database": database
    }
    connection = mysql.connector.connect(**db_config)
    data = pd.read_sql(query, connection)
    connection.close()
    return data


