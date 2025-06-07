import pandas as pd
import pyodbc
import calendar

df_operations = pd.read_excel(
    "k:/DOP/OED/METHOD&TOOLS/3 - PROJECTS/2 - ON GOING/2 - OE/2502 DIGITAL PROJECTS YEVGENIY/TRS/TRS_relations.xlsx",
    sheet_name='Operations')

df_operations = df_operations[['event', 'Category']]
mask = df_operations['Category'].notna()
df_operations_u = df_operations[mask].reset_index(drop=True)


driver = 'ODBC Driver 17 for SQL Server'
server = 'BKZTKDSDB41,21433'
database = 'PIFD'

# Строка подключения
conn_str = f'DRIVER={{{driver}}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'

# SQL-запрос
query = """
SELECT 
    AFE.name COLLATE DATABASE_DEFAULT AS [equipment],
    AFET.name COLLATE DATABASE_DEFAULT AS [event],
    DATEADD(SECOND, (AVEvF.starttime / 10000000) - 63240134400 + 6 * 3600, '2005-01-01') AS starttime,
    IIF(AVEvF.endtime > 2638317188000000000, NULL, 
        DATEADD(SECOND, (AVEvF.endtime / 10000000) - 63240134400 + 6 * 3600, '2005-01-01')) AS endtime
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
"""

# Загрузка в DataFrame
try:
    with pyodbc.connect(conn_str) as conn:
        df_osi_operations = pd.read_sql(query, conn) # pyright: ignore[reportArgumentType]
        print(df_osi_operations.head())  # Проверка результата
except Exception as e:
    print("Ошибка при подключении или выполнении запроса:", e)


df_osi_operations['duration'] = df_osi_operations['endtime'] - df_osi_operations['starttime']

df_osi_operations['duration'] = df_osi_operations['duration'].dt.total_seconds() / 3600



df_rigs = pd.read_excel(
    "k:/DOP/OED/METHOD&TOOLS/3 - PROJECTS/2 - ON GOING/2 - OE/2502 DIGITAL PROJECTS YEVGENIY/TRS/TRS_relations.xlsx",
    sheet_name='Rigs-Name')


df_circ_data = pd.read_excel(
    "k:/DOP/OED/METHOD&TOOLS/3 - PROJECTS/2 - ON GOING/2 - OE/2502 DIGITAL PROJECTS YEVGENIY/TRS/TRS_relations.xlsx",
    sheet_name='Standart avarage meters')


df_osi_operations['endtime'].max()


database_a = 'acqkzdem1'
# SQL-запрос
query_2 = """
SELECT
    HOLEID,
    DrillCompany,
    DrillRig,
    HolePurpose,
    HoleStatus,
    ENDDATE,
    CASE 
        WHEN TRY_CAST([InvoiceDepth] AS FLOAT) IS NOT NULL AND TRY_CAST([InvoiceDepth] AS FLOAT) > 0 
            THEN CAST([InvoiceDepth] AS FLOAT)
        ELSE [DEPTH]
    END AS DEPTH
    FROM [ACQKZDEM1].[dbo].[AV_AREVA_V_DRILLING_SPREADSHEET_DOP]
    where [Sourse_DB]='DEM' 
      and NOT ([HoleStatus] = 'PROJECT' or [HoleStatus]='CONSTRUCTION')
      AND ENDDATE >= '2020-01-01'
"""

# Строка подключения
conn_str = f'DRIVER={{{driver}}};SERVER={server};DATABASE={database_a};Trusted_Connection=yes;'

# Загрузка данных в DataFrame
try:
    with pyodbc.connect(conn_str) as conn:
        df_meters = pd.read_sql(query_2, conn) # pyright: ignore[reportArgumentType]
        print(df_meters.head())  # Проверка
except Exception as e:
    print("Ошибка при подключении или выполнении запроса:", e)


mask = df_rigs['Rig-AcQuire'].notna()
df_rigs = df_rigs[mask].reset_index(drop=True)
df_osi_operations['equipment'] = df_osi_operations['equipment'].str.strip().str.lower()
df_rigs['Rig-osiDEM'] = df_rigs['Rig-osiDEM'].str.strip().str.lower()

df_merged = df_osi_operations.merge(
    df_rigs,
    left_on='equipment',
    right_on='Rig-osiDEM',
    how='inner'  # или 'inner', если нужны только совпадающие строки
)


df_temp = df_merged.drop(columns=['equipment', 'Rig-osiDEM'])


df_temp['event'] = df_temp['event'].str.strip().str.lower()
df_operations_u['event'] = df_operations_u['event'].str.strip().str.lower()

df_operations_total = df_temp.merge(df_operations_u, on='event', how='left')

df_operations_total['Category'] = df_operations_total['Category'].fillna('Speed_losess')


# Убедимся, что endtime — datetime (хотя уже подтверждено)
df_operations_total['endtime'] = pd.to_datetime(df_operations_total['endtime'])
df_operations_total = df_operations_total.dropna(subset=['endtime'])

# Добавим год и месяц
df_operations_total['year'] = df_operations_total['endtime'].dt.year.astype('int16')
df_operations_total['month'] = df_operations_total['endtime'].dt.month.astype('int8')

pivot_df = df_operations_total.pivot_table(
    index=['DrillCompany', 'Rig-AcQuire', 'year', 'month', 'Tipe of circulation'],
    columns='Category',
    values='duration',
    aggfunc='sum',
    fill_value=0  # если нужно подставить 0 вместо NaN
).reset_index()

pivot_df = pivot_df.sort_values(
    by=['DrillCompany', 'Rig-AcQuire', 'year', 'month'],
    ascending=[True, True, True, True]
).reset_index(drop=True)




def hours_in_month(row):
    days = calendar.monthrange(row['year'], row['month'])[1]
    return days * 24

pivot_df['h_in_month'] = pivot_df.apply(hours_in_month, axis=1)

pivot_df.info()



# Преобразование ENDDATE в datetime формат
df_meters['ENDDATE'] = pd.to_datetime(df_meters['ENDDATE'], errors='coerce')

# Удаление строк с некорректными датами (NaT после преобразования)
df_meters.dropna(subset=['ENDDATE'], inplace=True)

df_meters



df_meters['year'] = df_meters['ENDDATE'].dt.year.astype('int16')
df_meters['month'] = df_meters['ENDDATE'].dt.month.astype('int8')


pivot2 = pd.pivot_table(
    df_meters,
    index=['year', 'month', 'DrillRig'],
    columns='HoleStatus',
    values='HOLEID',
    aggfunc='count',
    fill_value=0
)

# Добавим суммарную глубину
depth_sum = df_meters.groupby(['year', 'month', 'DrillRig'])['DEPTH'].sum()

# Объединяем всё
pivot2['DEPTH'] = depth_sum

# Сбросим индекс, чтобы было удобно смотреть
meters_pivot = pivot2.reset_index()

meters_pivot


pivot_df['Rig-AcQuire'] = pivot_df['Rig-AcQuire'].str.strip().str.upper()
meters_pivot['DrillRig'] = meters_pivot['DrillRig'].str.strip().str.upper()


total_merged_df = pd.merge(
    pivot_df,
    meters_pivot,
    left_on=['Rig-AcQuire', 'year', 'month'],
    right_on=['DrillRig', 'year', 'month'],
    how='left'  
)

total_merged_df.drop(columns='Rig-AcQuire', inplace=True)

total_merged_df



total_merged_df['Planned_production_time'] = total_merged_df['h_in_month'] - total_merged_df['Planned_downtime']

total_merged_df['Planned_factor'] = total_merged_df['Planned_production_time'] / total_merged_df['h_in_month']

# 1. Planned Production Time (PPT) с ограничением снизу
total_merged_df['Planned_production_time'] = (
    total_merged_df['h_in_month'] - total_merged_df['Planned_downtime']
).clip(lower=0)

# 2. Planned Factor 
total_merged_df['Planned_factor'] = (
    total_merged_df['Planned_production_time'] / total_merged_df['h_in_month']
).clip(lower=0)


total_merged_df.describe(include="all")



# Gross Operating Time (GOT)
total_merged_df['Gross_operating_time'] = total_merged_df['Planned_production_time'] - total_merged_df['Unplanned_downtime_losses']

# Availability (доступность)
total_merged_df['Availability'] = total_merged_df['Gross_operating_time'] / total_merged_df['h_in_month']

# Gross Operating Time (GOT) с ограничением снизу
total_merged_df['Gross_operating_time'] = (
    total_merged_df['Planned_production_time'] - total_merged_df['Unplanned_downtime_losses']
).clip(lower=0)

# Availability (доступность) с ограничением снизу
total_merged_df['Availability'] = (
    total_merged_df['Gross_operating_time'] / total_merged_df['h_in_month']
).clip(lower=0)


total_merged_df.info()



# Назначаем коэффициенты по типу циркуляции
circulation_coeffs = df_circ_data.set_index('Circ')['Standard avarage drilling, m/h'].to_dict()

# Присваиваем коэффициенты каждому ряду
total_merged_df['circulation_coeff'] = total_merged_df['Tipe of circulation'].map(circulation_coeffs)

# Вычисляем потенциальную глубину бурения на основе Gross Operating Time
total_merged_df['Potential_depth'] = (
    total_merged_df['Gross_operating_time'] * total_merged_df['circulation_coeff']
).round(1)



total_merged_df.head(3)



def calculate_net_operating_time(row):
    try:
        drilled = row['DEPTH']
        coeff = row['circulation_coeff']
        if pd.notnull(drilled) and pd.notnull(coeff) and coeff > 0:
            result = drilled / coeff
            return max(result, 0)  # отрицательные значения → 0
        else:
            return 0
    except:
        return 0

total_merged_df['Net_operating_time'] = total_merged_df.apply(calculate_net_operating_time, axis=1).round(1)



total_merged_df.head(2)


def calculate_performance(row):
    try:
        net_time = row['Net_operating_time']
        gross_time = row['Gross_operating_time']
        if pd.notnull(net_time) and pd.notnull(gross_time) and gross_time > 0:
            result = net_time / gross_time
            return max(min(result, 1), 0)
        else:
            return 0
    except:
        return 0

total_merged_df['Performance'] = total_merged_df.apply(calculate_performance, axis=1)


total_merged_df.head(3)


# Назначаем коэффициенты по типу циркуляции
well_drill_time_coef = df_circ_data.set_index('Circ')['time to well drill, h'].to_dict()

# Присваиваем коэффициенты каждому ряду
total_merged_df['Well drill coef'] = total_merged_df['Tipe of circulation'].map(well_drill_time_coef)

# Вычисляем quality Gross Operating Time
total_merged_df['Valuable Operating Time'] = (
    total_merged_df['Net_operating_time'] - total_merged_df['Well drill coef'] * total_merged_df['LIQUID']
).round(1).clip(lower=0)

def calculate_quality(row):
    try:
        net_time = row['Valuable Operating Time']
        gross_time = row['Net_operating_time']
        if pd.notnull(net_time) and pd.notnull(gross_time) and gross_time > 0:
            result = net_time / gross_time
            return max(result, 0)  # убрать отрицательные
        else:
            return 0
    except:
        return 0

total_merged_df['quality'] = total_merged_df.apply(calculate_quality, axis=1)


total_merged_df['OEE'] = (
    total_merged_df['Availability'] *
    total_merged_df['Performance'] *
    total_merged_df['quality']
)



total_merged_df.head(3)


total_merged_df['TRS'] = (
    total_merged_df['OEE'] * total_merged_df['Planned_factor']
)

total_merged_df.head(3)



total_merged_df.columns


total_merged_df.columns = [
    "Drill Company", "Year", "Month", "Type Of Circulation", "Planned Downtime",
    "Speed Losses", "Unplanned Downtime Losses", "H In Month", "Drill Rig",
    "ACCEPTED", "LIQUID", "NOT PROFITABLE", "Total Drilled", "Planned Production Time",
    "Planned Factor", "Gross Operating Time", "Availability", "Circulation Coeff",
    "Potential Depth", "Net Operating Time", "Performance", "Well drill coef", 
    "Valuable Operating Time", "Quality", "OEE", "TRS"
]


total_merged_df.describe(include='all')



total_merged_df['Speed Losses Calculated'] = (
    total_merged_df['Gross Operating Time'] - total_merged_df['Net Operating Time']
).clip(lower=0)  # Чтобы отрицательных не было


df_losses = pd.melt(
    total_merged_df,
    id_vars=[
        'Drill Company',
        'Drill Rig',
        'Year',
        'Month',
        'Type Of Circulation'
    ],
    value_vars=[
        'Planned Downtime',
        'Unplanned Downtime Losses',
        'Speed Losses Calculated',
        'Speed Losses'
    ],
    var_name='Loss Type',
    value_name='Loss Value'
)



df_productivity = pd.melt(
    total_merged_df,
    id_vars=[
        'Drill Company',
        'Drill Rig',
        'Year',
        'Month',
        'Type Of Circulation'
    ],
    value_vars=[
        'Planned Factor',
        'Availability',
        'Performance',
        'Quality',
        'OEE',
        'TRS'
    ],
    var_name='Productivity Type',
    value_name='Productivity Value'
)


df_productivity



df_operations_total['Category'].value_counts()


df_losses['Loss Type'].value_counts()


df_operations_total['Category'] = df_operations_total['Category'].replace({
    'Speed_losess': 'Speed Losses',
    'Planned_downtime': 'Planned Downtime',
    'Unplanned_downtime_losses': 'Unplanned Downtime Losses'
})

df_operations_total['Category'].value_counts()


df_operations_total.columns = [
    "Event",
    "Start Time",
    "End Time",
    "Duration",
    "Drill Company",
    "Rig",
    "Type Of Circulation",
    "Category",
    "Year",
    "Month"
]



df_operations_total['Event Category'] = df_operations_total['Event'].str.split('_').str[0].str.upper()
df_operations_total['Event'] = df_operations_total['Event'].str.split('_', n=1).str[1].str.capitalize()

df_events_duration = df_operations_total[
    [
        "Year",
        "Month",
        "Category",
        "Drill Company",
        "Rig",
        "Type Of Circulation",
        "Event Category",
        "Event",
        "Duration"
    ]
]
df_events_duration['Event'].value_counts()


df_events_duration.sort_values(
    by=['Year', 'Month', 'Rig', 'Category', 'Duration'],
    ascending=[False, True, True, True, False],
    inplace=True
)

df_events_duration.head(3)


df_meters.dropna(subset=['DrillCompany', 'DrillRig'], inplace=True)

df_meters.columns = [
    'Hole ID',
    'Drilling Company',
    'Drill Rig',
    'Purpose',
    'Status',
    'End_Date',
    'Depth_m',
    'Year',
    'Month'
]



import os

output_dir = 'k:/DOP/OED/METHOD&TOOLS/3 - PROJECTS/2 - ON GOING/2 - OE/2502 DIGITAL PROJECTS YEVGENIY/TRS/Python converted/'
os.makedirs(output_dir, exist_ok=True)

df_meters.to_csv(
    os.path.join(output_dir, 'df_meters.csv'),
    index=False,
    encoding='utf-8-sig'
)

df_events_duration.to_csv(
    os.path.join(output_dir, 'df_events_duration.csv'),
    index=False,
    encoding='utf-8-sig'
)




df_events_duration


df_losses.to_csv(
    os.path.join(output_dir, 'df_losses.csv'),
    index=False,
    encoding='utf-8-sig'
)


df_productivity.to_csv(
    os.path.join(output_dir, 'df_productivity.csv'),
    index=False,
    encoding='utf-8-sig'
)


dates = df_productivity[['Year', 'Month']].drop_duplicates().sort_values(['Year', 'Month']).reset_index(drop=True)


dates.to_csv(
    os.path.join(output_dir, 'df_dates.csv'),
    index=False,
    encoding='utf-8-sig'
)


