DRIVER = "ODBC Driver 17 for SQL Server"
SERVER = "BKZTKDSDB41,21433"

DATABASE_1 = "PIFD"
DATABASE_2 = "acqkzdem1"


QUERY_1 = """
SELECT 
    AFE.name COLLATE DATABASE_DEFAULT AS [equipment],
    AFET.name COLLATE DATABASE_DEFAULT AS [event],
    DATEADD(SECOND, (AVEvF.starttime / 10000000) - 63240134400 + 5 * 3600, '2005-01-01') AS starttime,
    IIF(AVEvF.endtime > 2638317188000000000, NULL, 
        DATEADD(SECOND, (AVEvF.endtime / 10000000) - 63240134400 + 5 * 3600, '2005-01-01')) AS endtime
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

QUERY_2 = """
SELECT
    HOLEID,
    DrillCompany,
    DrillRig,
    HolePurpose,
    HoleStatus,
    ENDDATE,
    CASE 
        WHEN TRY_CAST([InvoiceDepth] AS FLOAT) IS NOT NULL 
            AND TRY_CAST([InvoiceDepth] AS FLOAT) > 0 
            THEN CAST([InvoiceDepth] AS FLOAT)
        ELSE [DEPTH]
    END AS DEPTH
FROM [ACQKZDEM1].[dbo].[AV_AREVA_V_DRILLING_SPREADSHEET_DOP]
WHERE [Sourse_DB]='DEM' 
  AND NOT ([HoleStatus] = 'PROJECT' OR [HoleStatus]='CONSTRUCTION')
  AND ENDDATE >= '2020-01-01'
"""
