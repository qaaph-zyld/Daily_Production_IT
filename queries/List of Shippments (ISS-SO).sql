SELECT DISTINCT  
      [tr_type],  
      [tr_addr],  
      [tr_site],  
	  DATEPART(YEAR, [tr_effdate]) as yearnum,
      DATEPART(WEEK, [tr_effdate]) AS weeknum,
	    DATEPART(DAY, [tr_effdate]) AS daynum

FROM [QADEE].[dbo].[tr_hist]  
WHERE [tr_type] = 'iss-so'  

UNION ALL  

SELECT   DISTINCT
      [tr_type],  
      [tr_addr],    
      [tr_site],  
	  DATEPART(YEAR, [tr_effdate]) as yearnum,
      DATEPART(WEEK, [tr_effdate]) AS weeknum,
	    DATEPART(DAY, [tr_effdate]) AS daynum

FROM [QADEE2798].[dbo].[tr_hist]  
WHERE [tr_type] = 'iss-so'  

ORDER BY weeknum;  -- Using weeknum for ordering instead of tr_part  