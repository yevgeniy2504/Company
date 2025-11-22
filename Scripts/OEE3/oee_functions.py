import pandas as pd
import os
from sqlalchemy import create_engine
import logging


class CsvSaveError(Exception):
    pass


class DatabaseLoadError(Exception):
    pass


class ExcelLoadError(Exception):
    pass


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    filename="app.log",
    filemode="a"
)
logger = logging.getLogger(__name__)


def save_df_to_csv(df, filename, folder='output', index=False, encoding='utf-8-sig', decimal=','):
    try:
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, filename)
        df.to_csv(path, index=index, encoding=encoding, decimal=decimal)
        return path
    except Exception as e:
        logger.error(f"Ошибка сохранения CSV: {e}")
        raise CsvSaveError(f"Ошибка сохранения CSV файла '{filename}': {e}")


def load_from_db_sqlalchemy(driver, server, database, query):
    """
    Загружает данные из SQL Server; пробрасывает DatabaseLoadError.
    """
    connection_string = (
        f"mssql+pyodbc://@{server}/{database}"
        f"?driver={driver.replace(' ', '+')}"
        f"&trusted_connection=yes"
    )

    try:
        engine = create_engine(connection_string)
        with engine.connect() as connection:
            df = pd.read_sql(query, connection)
        return df
    except Exception as e:
        logger.error(f"Ошибка загрузки из базы: {e}")
        raise DatabaseLoadError(f"Ошибка загрузки данных из базы '{database}': {e}")


def load_excel_sheet(file_path, sheet_name):
    """
    Загружает лист Excel; пробрасывает ExcelLoadError.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        logger.error(f"Ошибка загрузки Excel: {e}")
        raise ExcelLoadError(f"Ошибка загрузки листа '{sheet_name}' из файла '{file_path}': {e}")
