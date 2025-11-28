SELECT DISTINCT  
      [tr_addr] as [Address],
      [tr_site] as [Site],
	  DATEPART(YEAR, [tr_effdate]) as yearnum,
      DATEPART(WEEK, [tr_effdate]) AS weeknum,
	  	    DATEPART(DAY, [tr_effdate]) AS daynum-- Adding week number column  

FROM [QADEE].[dbo].[tr_hist]  
WHERE [tr_type] = 'rct-po'  

UNION ALL  

SELECT   DISTINCT 
      [tr_addr] as [Address],  
      [tr_site] as [Site],
	  DATEPART(YEAR, [tr_effdate]) as yearnum,
      DATEPART(WEEK, [tr_effdate]) AS weeknum,
	  	    DATEPART(DAY, [tr_effdate]) AS daynum-- Adding week number column  

FROM [QADEE2798].[dbo].[tr_hist]  
WHERE [tr_type] = 'rct-po'  

ORDER BY weeknum;  -- Using weeknum for ordering instead of tr_part  