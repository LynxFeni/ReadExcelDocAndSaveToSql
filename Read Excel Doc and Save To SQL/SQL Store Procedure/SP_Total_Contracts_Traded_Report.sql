CREATE PROCEDURE [dbo].[SP_Total_Contracts_Traded_Report]
    @DateFrom DATE,
    @DateTo DATE
AS
BEGIN
    SET NOCOUNT ON;

    -- Temporary table to hold the aggregated data
    CREATE TABLE #DailyTotals (
        [FileDate] DATE,
        [TotalContractsTraded] INT
    );

    -- Insert total daily contracts traded into the temporary table
    INSERT INTO #DailyTotals ([FileDate], [TotalContractsTraded])
    SELECT 
        FileDate,
        SUM(ContractsTraded) AS TotalContractsTraded
    FROM 
        DailyMTM
    WHERE 
        FileDate BETWEEN @DateFrom AND @DateTo
    GROUP BY 
        FileDate;

    -- Select data and calculate the percentage of total contracts traded per day for each contract
    SELECT 
        D.FileDate AS [File Date],
        DMTM.[Contract],
        SUM(DMTM.ContractsTraded) AS [Contracts Traded],
        CAST((CAST(SUM(DMTM.ContractsTraded) AS DECIMAL(18,2)) / D.TotalContractsTraded) * 100 AS DECIMAL(18,2)) AS [% Of Total Contracts Traded] 
    FROM 
        DailyMTM DMTM
    JOIN 
        #DailyTotals D ON DMTM.FileDate = D.FileDate
    WHERE 
        DMTM.FileDate BETWEEN @DateFrom AND @DateTo
        AND DMTM.ContractsTraded > 0
    GROUP  BY 
        DMTM.FileDate, DMTM.[Contract],D.TotalContractsTraded,D.FileDate
	ORDER  BY 
        DMTM.FileDate, DMTM.[Contract];


    -- Drop the temporary table
    DROP TABLE #DailyTotals;
END;