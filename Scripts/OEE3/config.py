# ============================
# CONFIGURATION
# ============================

from datetime import datetime, timedelta

# === Подключение к SQL ===
DRIVER = "ODBC Driver 17 for SQL Server"
SERVER = "BKZTKDSDB41,21433"

DATABASE_1 = "PIFD"
DATABASE_2 = "acqkzdem1"


# ============================
# ДАТЫ ФИЛЬТРАЦИИ
# ============================

# Конец прошлого месяца 20:00:00 в ISO-формате
today = datetime.today()
first_day_current_month = today.replace(day=1)
last_day_prev_month = first_day_current_month - timedelta(days=1)

END_DATE = last_day_prev_month.replace(
    hour=20, minute=0, second=0
).strftime("%Y-%m-%d %H:%M:%S")

# Начало периода: задаём дату и время вручную через datetime()
START_DATE = datetime(2022, 1, 1, 8, 0, 0).strftime("%Y-%m-%d %H:%M:%S")


# ============================
# EXCEL файлы
# ============================

PATH_TO_EXCEL_FILE = (
    r"k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\2 - ON GOING\2 - OE\2502 DIGITAL PROJECTS YEVGENIY\TRS\TRS_relations.xlsx"
)

EXELL_SHEET_NAME_1 = "Drilling_Operations"
EXELL_SHEET_NAME_2 = "Rigs-Name"
EXELL_SHEET_NAME_3 = "Standart avarage meters"


# ========================================================================
# QUERY 1 — события PI (PIFD) c PI FILETIME → datetime (UTC+6)
# ========================================================================

QUERY_1 = """
SELECT 
    AFE.name COLLATE DATABASE_DEFAULT AS equipment,
    AFET.name COLLATE DATABASE_DEFAULT AS event,

    DATEADD(
        SECOND,
        (AVEvF.starttime / 10000000) - 63240134400 + 6 * 3600,
        '2005-01-01'
    ) AS starttime,

    CASE 
        WHEN AVEvF.endtime > 2638317188000000000 THEN NULL
        ELSE DATEADD(
            SECOND,
            (AVEvF.endtime / 10000000) - 63240134400 + 6 * 3600,
            '2005-01-01'
        )
    END AS endtime

FROM PIFD.dbo.AFEventFrame AS AVEvF
LEFT JOIN PIFD.dbo.AFElementTemplate AS AFET 
    ON AVEvF.fktemplateid = AFET.rid
LEFT JOIN PIFD.dbo.AFElement AS AFE 
    ON AVEvF.fkprimaryreferencedelement = AFE.rid

WHERE AFE.name COLLATE DATABASE_DEFAULT IN (
    SELECT DISTINCT equipment COLLATE DATABASE_DEFAULT 
    FROM PI_EXTRA.dbo.DR_REL_EQUIPMENT_EVENT
)

AND AFE.id IN (
    SELECT fkelementid 
    FROM PIFD.dbo.AFElementVersion 
    WHERE fktemplateid = (
        SELECT id 
        FROM PIFD.dbo.AFElementTemplate 
        WHERE name = 'RIG'
          AND fkdatabaseid = (
              SELECT id FROM PIFD.dbo.AFDatabase WHERE name = 'DEM_DR'
          )
    )
)

AND DATEADD(
        SECOND,
        (AVEvF.starttime / 10000000) - 63240134400 + 6 * 3600,
        '2005-01-01'
    ) >= :start_date

AND DATEADD(
        SECOND,
        (AVEvF.starttime / 10000000) - 63240134400 + 6 * 3600,
        '2005-01-01'
    ) <= :end_date
"""


# ========================================================================
# QUERY 2 — данные о скважинах (ACQKZDEM1)
# Простая фильтрация дат (ISO-формат из Python), без усложнений
# ========================================================================

QUERY_2 = """
SELECT
    HOLEID,
    DrillCompany,
    DrillRig,
    HolePurpose,
    HoleStatus,
    ENDDATE,
    CASE 
        WHEN ISNUMERIC(InvoiceDepth) = 1 AND InvoiceDepth <> ''
            THEN CAST(InvoiceDepth AS FLOAT)
        ELSE DEPTH
    END AS DEPTH
FROM ACQKZDEM1.dbo.AV_AREVA_V_DRILLING_SPREADSHEET_DOP
WHERE 
    Sourse_DB = 'DEM'
    AND HoleStatus NOT IN ('PROJECT', 'CONSTRUCTION')

    -- Проверка что ENDDATE преобразуем
    AND ISDATE(ENDDATE) = 1

    -- Фильтр по датам (через старый CONVERT)
    AND CONVERT(datetime, ENDDATE, 120) >= :start_date
    AND CONVERT(datetime, ENDDATE, 120) <= :end_date
"""
