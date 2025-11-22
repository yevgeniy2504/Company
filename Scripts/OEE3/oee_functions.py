import pandas as pd
import os
from sqlalchemy import create_engine, text
import logging


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    filename="app.log",
    filemode="a"
)
logger = logging.getLogger(__name__)


def save_df_to_csv(df, filename, folder='output', index=False, encoding='utf-8-sig', decimal=','):
    """
    Сохраняет DataFrame в CSV. Ошибки пробрасываются вверх.
    """
    try:
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, filename)
        df.to_csv(path, index=index, encoding=encoding, decimal=decimal)
        return path
    except Exception as e:
        logger.error(f"Ошибка сохранения CSV: {e}")
        raise


def load_from_db_sqlalchemy(driver, server, database, query, start_date=None, end_date=None):
    """
    Загружает данные из SQL Server. Ошибки пробрасываются вверх.
    Поддерживает параметры :start_date и :end_date.
    """

    connection_string = (
        f"mssql+pyodbc://@{server}/{database}"
        f"?driver={driver.replace(' ', '+')}"
        f"&trusted_connection=yes"
    )

    params = {}
    if start_date:
        params["start_date"] = start_date
    if end_date:
        params["end_date"] = end_date

    try:
        engine = create_engine(connection_string)
        with engine.connect() as connection:
            df = pd.read_sql(text(query), connection, params=params)
        return df
    except Exception as e:
        logger.error(f"Ошибка загрузки SQL: {e}")
        raise


def load_excel_sheet(file_path, sheet_name):
    """
    Загружает лист Excel. Ошибки пробрасываются вверх.
    """
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        logger.error(f"Ошибка загрузки Excel: {e}")
        raise
