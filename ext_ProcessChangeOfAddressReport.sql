IF OBJECT_ID('dbo.ext_ProcessChangeOfAddressReport') IS NOT NULL
BEGIN
	DROP PROCEDURE dbo.ext_ProcessChangeOfAddressReport
END

CREATE PROCEDURE [dbo].[ext_ProcessChangeOfAddressReport](@exportPath NVARCHAR(192))    
AS    
BEGIN    
    
SET NOCOUNT ON;    
/*=======================    
Created By Tyler T on 6/9/23 to import CoA report data from Clerk   
Modified on 11/3/23 - added section to remove all addresses without matching apt numbers  
========================*/    
BEGIN TRY    
    
DECLARE @tran_count_on_entry INT = @@TRANCOUNT,    
  @return_value INT = @@ERROR;    
  
/*=======================================================================  
If data successfully loaded into table from csv file -- > Begin  
=======================================================================*/  
    
IF EXISTS (SELECT * FROM dbo.ext_ChangeOfAddress)    
BEGIN;  
  
DECLARE @PropertyID INT,    
  @hx NVARCHAR(4),    
  @YearID SMALLINT = (SELECT YearID FROM xrYearColor WHERE IsCurrentFlag = 1 GROUP BY YearID);    
    
 IF OBJECT_ID('dbo.ChangeOfAddressRpt') IS NOT NULL    
 BEGIN    
  DROP TABLE ChangeOfAddressRpt    
 END    
    
 IF OBJECT_ID('tempdb..#tempOwnerComp') IS NOT NULL    
 BEGIN    
  DROP TABLE #tempOwnerComp    
 END    
    
 IF OBJECT_ID('tempdb..#tempAddressesFound') IS NOT NULL    
 BEGIN    
  DROP TABLE #tempAddressesFound    
 END    
    
 IF OBJECT_ID('tempdb..#tempNameFound') IS NOT NULL    
 BEGIN    
  DROP TABLE #tempNameFound    
 END    
    
 IF OBJECT_ID('tempdb..#RecordMatching') IS NOT NULL    
 BEGIN    
  DROP TABLE #RecordMatching    
 END    
  
/*=======================================================================    
 Begin creating CTEs gathering data for Owner record comparisons  
=======================================================================*/    
  
 ;WITH OwnerName AS    
 (    
 SELECT    
   OwnerID,     
   OwnerFormatted,    
   LTRIM(RTRIM(OwnerFirstName)) [OwnerFirstName],    
   OwnerMiddleName,    
   LTRIM(RTRIM(OwnerLastName)) [OwnerLastName],    
   IIF(NULLIF(Owner1DOB,'') IS NULL OR ISDATE(Owner1DOB) = 0,'1900-01-01',CAST(Owner1DOB AS DATETIME)) [OwnerDOB]    
 FROM    
	dbo.GetOwnersTable(@YearID,1)    
 )    
    
 ,PropertyOwner AS    
 (    
 SELECT    
   PropertyID,    
   OwnerID,     
   OwnerSequence,    
   IsPrivateOwnerFlag    
 FROM    
	dbo.GetPropertyOwnersTable(@YearID,1)    
 )    
    
 ,HomesteadProperty AS    
 (    
 SELECT    
   e.PropertyID,    
   e.YearID,    
   xr.Exemption    
 FROM     
	dbo.GetExemptionsTable(@YearID,1) e    
 JOIN    
	(SELECT xrExemptionID, Exemption FROM dbo.GetxrExemptionTable(1,@YearID,1)) xr ON xr.xrExemptionID = e.xrExemptionID    
 WHERE    
	1=1    
 AND    
	e.YearID = @YearID    
 AND    
	xr.Exemption = '01'    
 )    
    
 ,PropertyLocation AS    
 (    
  SELECT     
   PropertyID,    
   LTRIM(RTRIM(LocationStartNumber)) [LocationStartNumber],    
   StreetDirection,    
   CONCAT(LTRIM(RTRIM(StreetName)), ' ', LTRIM(RTRIM(StreetWay))) [StreetName],    
   ApartmentUnitNumber,    
   LTRIM(RTRIM(City)) [City],    
   'FL' [State],    
   Postal    
 FROM     
	dbo.GetFormattedLocationTable(1,@YearID,1)    
 )    
    
 SELECT    
   po.PropertyID,    
   IIF(a.Exemption IS NULL,'No','Yes') [Homestead],    
   po.OwnerSequence,    
   po.IsPrivateOwnerFlag,    
   o.OwnerFormatted,    
   o.OwnerFirstName,    
   o.OwnerMiddleName,    
   o.OwnerLastName,    
   CAST(o.OwnerDOB AS DATE) [OwnerDOB],    
   p.LocationStartNumber,    
   p.StreetDirection,    
   p.StreetName,    
   p.ApartmentUnitNumber,    
   p.City,    
   p.State,    
   p.Postal    
 INTO   
	#tempOwnerComp    
 FROM    
	PropertyOwner po    
 JOIN     
	OwnerName o ON o.OwnerID = po.OwnerID    
 LEFT JOIN    
	HomesteadProperty a ON a.PropertyID = po.PropertyID    
 JOIN    
	PropertyLocation p ON p.PropertyID = po.PropertyID    
 JOIN    
	(SELECT PropertyID FROM dbo.GetPropertiesTable(1,@YearID,1) WHERE InActiveFlag = 0 AND IsPersonalPropertyFlag = 0) pt ON pt.PropertyID = po.PropertyID    
 ORDER BY    
	po.PropertyID ASC,    
	po.OwnerSequence ASC    
    
/*=======================================================================*/  
 /* all CTE data into #tempOwnerComp to serve as search dataset */    
 --SELECT * FROM dbo.ext_ChangeOfAddress    
 --SELECT TOP 1000 * FROM #tempOwnerComp ORDER BY PropertyID ASC  
 /* Next a while loop is used to compare all records in table to #tempOwnercomp dataset */  
/*=======================================================================*/    
    
 CREATE TABLE #RecordMatching ([RM_ID] SMALLINT IDENTITY(1,1), PropertyID INT NULL, AddressMatch NVARCHAR(8) NULL, NameMatch NVARCHAR(8) NULL, DOBMatch NVARCHAR(8) NULL, ActiveHomestead NVARCHAR(8) NULL);    
    
 DECLARE @counter SMALLINT = 0;    
 DECLARE @ceiling SMALLINT = (SELECT COUNT(*) FROM dbo.ext_ChangeOfAddress);    
    
 WHILE @counter < @ceiling    
 BEGIN  
  
  SET @PropertyID = 0;    
    
  IF OBJECT_ID('tempdb..#tempAddressesFound') IS NOT NULL    
  BEGIN    
   DROP TABLE #tempAddressesFound    
  END    
    
  IF OBJECT_ID('tempdb..#tempNameFound') IS NOT NULL    
  BEGIN    
   DROP TABLE #tempNameFound    
  END    
    
  IF OBJECT_ID('tempdb..#tempComp') IS NOT NULL    
  BEGIN    
   DROP TABLE #tempComp    
  END    
    
  SELECT *  
  INTO   
	#tempComp    
  FROM     
	dbo.ext_ChangeOfAddress     
  ORDER BY     
	[ID] ASC    
  OFFSET @counter ROWS    
  FETCH NEXT 1 ROW ONLY    
  
 --SELECT * FROM #tempComp -- 1 record at a time    
   
/*=================================================  
Create dataset of matching addresses found in   
the #tempOwnerComp for the Coa record -- > #tempAddressesFound  
===================================================*/   
    
  SELECT *     
  INTO     
	#tempAddressesFound    
  FROM     
	#tempOwnerComp t    
  WHERE     
	1=1    
  AND     
	t.LocationStartNumber = (SELECT LTRIM(RTRIM(OldStreetNum)) [OldStreetNum] FROM #tempComp)    
  AND     
	t.StreetName LIKE '%' + (SELECT LTRIM(RTRIM(OldStreet)) FROM #tempComp) + '%'    
    
  --SELECT * FROM #tempAddressesFound    
  --SELECT * FROM #RecordMatching    
  
/*=================================================  
Added to delete all address records that do not have a matching apartment number  
===================================================*/  
  
  IF ((SELECT NULLIF(OldApt,'') FROM #tempComp) IS NOT NULL) AND ((SELECT COUNT(DISTINCT PropertyID) FROM #tempAddressesFound) > 1)  
  BEGIN  
  
  DECLARE @aptNum VARCHAR(32) = (SELECT REPLACE(REPLACE(OldApt,'-',''),'#','') FROM #tempComp);  
  
  DELETE FROM #tempAddressesFound  
  WHERE PATINDEX('%'+@aptNum+'%',REPLACE(REPLACE(ISNULL(ApartmentUnitNumber,''),'-',''),'#','')) <> 1  
  
  END  
   
/*===================================================  
#RecordMatching table is updated to reflect results of found addresses   
===================================================*/  
  
  IF EXISTS (SELECT * FROM #tempAddressesFound)  
   IF (SELECT DISTINCT TOP 1 RANK() OVER (PARTITION BY LocationStartNumber ORDER BY PropertyID ASC) [R] FROM #tempAddressesFound ORDER BY [R] DESC) = 1 -- if its one address/account (more than one owner)    
   BEGIN    
  
    SET @PropertyID = (SELECT PropertyID FROM #tempAddressesFound GROUP BY PropertyID);    
    SET @hx = (SELECT UPPER(Homestead) [hx] FROM #tempAddressesFound GROUP BY Homestead);    
    
     INSERT INTO #RecordMatching    
     (    
		PropertyID,    
		AddressMatch,    
		NameMatch,    
		DOBMatch,    
		ActiveHomestead    
     )    
     VALUES    
     (     
		@PropertyID,   -- PropertyID - int    
		N'YES', -- AddressMatch - nvarchar(8)    
		N'', -- NameMatch - nvarchar(8)    
		N'', -- DOBMatch - nvarchar(8)    
		@hx  -- ActiveHomestead - nvarchar(8)    
   )    
   END    
   ELSE           
  BEGIN -- if address found but more than account    
  
     INSERT INTO #RecordMatching    
   (    
    PropertyID,    
    AddressMatch,    
    NameMatch,    
    DOBMatch,    
    ActiveHomestead    
   )    
     VALUES    
   (   0,   -- PropertyID - int    
    N'YES', -- AddressMatch - nvarchar(8)    
    N'', -- NameMatch - nvarchar(8)    
    N'', -- DOBMatch - nvarchar(8)    
    N'NO'  -- ActiveHomestead - nvarchar(8)    
    )    
  
  END    
  ELSE   -- there were no accounts/addresses found    
  BEGIN    
  
    INSERT INTO #RecordMatching    
    (    
     PropertyID,    
     AddressMatch,    
     NameMatch,    
     DOBMatch,    
     ActiveHomestead    
    )    
    VALUES    
    (   0,   -- PropertyID - int    
     N'NO', -- AddressMatch - nvarchar(8)    
     N'', -- NameMatch - nvarchar(8)    
     N'', -- DOBMatch - nvarchar(8)    
     N'NO'  -- ActiveHomestead - nvarchar(8)    
     )   
      
   END    
   
/*=================================================  
Create dataset of matching names found in the #tempOwnerComp for the Coa record  
===================================================*/  
  
  SELECT *     
  INTO     
	#tempNameFound    
  FROM     
	#tempAddressesFound a    
  WHERE     
	a.OwnerLastName LIKE '%' + (SELECT LTRIM(RTRIM(LastName)) [LastName] FROM #tempComp) + '%'    
    
  --SELECT * FROM #tempNameFound  
  
/*===================================================  
#RecordMatching table is updated to reflect results of found last names   
===================================================*/  
     
  IF EXISTS (SELECT * FROM #tempNameFound)  -- if owner with matching last name found  
  BEGIN   
    
   SET @hx = (SELECT UPPER(Homestead) [hx] FROM #tempNameFound GROUP BY Homestead);    
   SET @PropertyID = (SELECT PropertyID FROM #tempNameFound GROUP BY PropertyID);    
    
   UPDATE #RecordMatching    
   SET NameMatch = 'YES',    
   ActiveHomestead = ISNULL(@hx,'NO'),    
   PropertyID = ISNULL(@PropertyID,0)    
   WHERE   
   [RM_ID] = (@counter + 1)    
    
  END    
  ELSE  -- if no matching owner last name record found  
  BEGIN    
  
   UPDATE #RecordMatching    
   SET NameMatch = 'NO',    
   PropertyID = ISNULL(@PropertyID,0)    
   WHERE   
   [RM_ID] = (@counter + 1)    
    
  END    
    
/*=================================================  
Compare DOB from Coa record to DOB of owners found with matching first name  
===================================================*/  
  
  DECLARE @DOBMatch NVARCHAR(12);    
    
  SET @DOBMatch = (    
  SELECT TOP 1    
   CASE   
   WHEN OwnerDOB = (SELECT DateOfBirth FROM #tempComp WHERE #tempComp.FirstName LIKE '%' + LTRIM(RTRIM(#tempNameFound.OwnerFirstName)) + '%') THEN 'Matching'    
   WHEN OwnerDOB > (SELECT DateOfBirth FROM #tempComp WHERE #tempComp.FirstName LIKE '%' + LTRIM(RTRIM(#tempNameFound.OwnerFirstName)) + '%') THEN 'Older'    
   WHEN OwnerDOB < (SELECT DateOfBirth FROM #tempComp WHERE #tempComp.FirstName LIKE '%' + LTRIM(RTRIM(#tempNameFound.OwnerFirstName)) + '%') THEN 'Younger'    
   ELSE 'NA' END [DOBComp]    
  FROM     
	#tempNameFound    
  WHERE     
	OwnerFirstName LIKE '%'+ (SELECT LTRIM(RTRIM(FirstName)) [FirstName] FROM #tempComp) + '%'    
   )    
    
/*===================================================  
#RecordMatching table is updated to reflect results DoB comparison  
===================================================*/  
    
  IF EXISTS (SELECT * FROM #tempNameFound)    
  BEGIN    
    
   UPDATE #RecordMatching    
   SET DOBMatch = (ISNULL(@DOBMatch,'NA'))    
   WHERE [RM_ID] = (@counter + 1)    
    
  END    
  ELSE    
  BEGIN  
  
  UPDATE #RecordMatching    
  SET DOBMatch = 'NA'    
  WHERE [RM_ID] = (@counter + 1)  
  
  END    
    
  SET @counter += 1    
    
 END;    
  
/*=================================================  
Loop completed --> next record in Coa table will be processed  
---------  
Portion below runs after while loop has ended,   
the #RecordMatching table is joined on the Coa table to produce statistical data stored in ext_CoaResults  
===================================================*/  
     
 --SELECT * FROM #RecordMatching    
    
 IF OBJECT_ID('tempdb..#temp_join') IS NOT NULL    
 BEGIN    
 DROP TABLE #temp_join    
 END    
    
 SELECT *     
 INTO     
	#temp_join    
 FROM    
	dbo.ext_ChangeOfAddress coa    
 JOIN     
	#RecordMatching rm ON rm.RM_ID = coa.ID    
    
 --SELECT * FROM #temp_join     
    
 BEGIN TRY;    
     
  INSERT INTO dbo.ext_CoaResults (CoaBatchImportID, NumberOfRecords, ReportMonth, AccountMatchPercentage, AddressMatchPercentage, LastNameMatchPercentage, DoBMatchPercentage, HomesteadPercentage)    
    
  SELECT     
   (SELECT ISNULL(MAX(CoaBatchImportID),0) + 1 FROM dbo.ext_CoaResults) [BatchImportID],    
   (SELECT COUNT(*) FROM #temp_join) [NumberofRecords],    
   (SELECT TOP 1 DATENAME(MONTH,ChangeRequestedOn) [RptMonth] FROM #temp_join ORDER BY [RptMonth] ASC) [Month],    
   ROUND(    
   (SELECT COUNT(PropertyID) FROM #temp_join WHERE PropertyID <> 0)     
   /     
   CAST((SELECT COUNT(PropertyID) FROM #temp_join) AS FLOAT) * 100,2) AS [AccountMatch%],    
   ROUND(    
   (SELECT COUNT(AddressMatch) FROM #temp_join WHERE AddressMatch = 'YES')     
   /     
   CAST((SELECT COUNT(AddressMatch) FROM #temp_join) AS FLOAT) * 100,2) AS [AddressMatch%],    
   ROUND(    
   (SELECT COUNT(NameMatch) FROM #temp_join WHERE NameMatch = 'YES')     
   /     
   CAST((SELECT COUNT(NameMatch) FROM #temp_join) AS FLOAT) * 100,2) AS [LastNameMatch%],    
   ROUND(    
   (SELECT COUNT(DOBMatch) FROM #temp_join WHERE DOBMatch = 'Matching')     
   /     
   CAST((SELECT COUNT(DOBMatch) FROM #temp_join) AS FLOAT) * 100,2) AS [DoBMatch%],    
   ROUND(    
   (SELECT COUNT(ActiveHomestead) FROM #temp_join WHERE ActiveHomestead = 'YES')     
   /     
   CAST((SELECT COUNT(ActiveHomestead) FROM #temp_join) AS FLOAT) * 100,2) AS [Homestead%]    
     
 END TRY    
 BEGIN CATCH    
    
 THROW;    
    
 END CATCH;    
    
/*=================================================  
The next section handles the data being pivoted and exported to a csv file and then converted to an Excel file  
 ===================================================*/  
    
 SET @Counter = 1;    
 DECLARE @RecordCount SMALLINT = (SELECT COUNT(*) FROM #temp_join);    
 DECLARE @join NVARCHAR(MAX) = '';    
 DECLARE @Field NVARCHAR(MAX) = '';    
 DECLARE @sql NVARCHAR(MAX);    
    
 WHILE @Counter < @RecordCount    
 BEGIN    
  DECLARE @Letter NVARCHAR(4) = (CHAR(@Counter / power(26,2) % 26 + 65) +      
    CHAR(@Counter / 26 % 26 + 65) +       
    CHAR(@Counter % 26 + 65)) ;     
  DECLARE @temp_field NVARCHAR(MAX) = '[Request]';    
    
  DECLARE @temp_join NVARCHAR(MAX) =    
  'JOIN (SELECT RequestorData, Request    
  FROM    
  (    
  SELECT     
   CAST(coa.ChangeRequestedOn AS NVARCHAR(250)) [ChangeRequestedOn],    
   CAST(coa.FormattedOwner AS NVARCHAR(250)) [FormattedRequestor],    
   CAST(coa.DateOfBirth AS NVARCHAR(250)) [DateOfBirth],    
   CAST(coa.FormattedOldAddress AS NVARCHAR(250)) [FormattedOldAddress],    
   CAST(coa.FormattedOldCityStZip AS NVARCHAR(250)) [FormattedOldCityStZip],    
   CAST(coa.FormattedNewAddress AS NVARCHAR(250)) [FormattedNewAddress],    
   CAST(coa.FormattedNewCityStZip AS NVARCHAR(250)) [FormattedNewCityStZip],    
   CAST(coa.PropertyID AS NVARCHAR(250)) [PropertyID],    
   CAST(coa.AddressMatch AS NVARCHAR(250)) [AddressMatch],    
   CAST(coa.NameMatch AS NVARCHAR(250)) [NameMatch],    
   CAST(coa.DOBMatch AS NVARCHAR(250)) [DOBMatch],    
   CAST(coa.ActiveHomestead AS NVARCHAR(250)) [ActiveHomestead]    
    
  FROM #temp_join coa    
  ORDER BY     
   ID ASC    
  OFFSET '+CAST(@Counter AS NVARCHAR(2))+' ROWS    
  FETCH NEXT 1 ROWS ONLY    
  ) x    
  UNPIVOT    
   (Request FOR RequestorData IN (ChangeRequestedOn, FormattedRequestor, DateOfBirth, FormattedOldAddress, FormattedOldCityStZip, FormattedNewAddress, FormattedNewCityStZip, PropertyID, AddressMatch, NameMatch, DOBMatch, ActiveHomestead)    
   ) AS unpvt) '+@Letter+' ON '+@Letter+'.RequestorData = unpvt.RequestorData    
   '    
    
  --PRINT @temp_join    
  SET @join = @join + @temp_join    
  --PRINT @join    
    
  SET @temp_field = CONCAT(',',@Letter,'.',@temp_field)    
  SET @Field = @Field + @temp_field    
  --PRINT @Field    
    
  SET @Counter = @Counter + 1    
 END    
    
 SET @sql =    
 '    
 SELECT unpvt.RequestorData, unpvt.Request '+@Field+'    
 FROM    
 (    
 SELECT     
  CAST(coa.ChangeRequestedOn AS NVARCHAR(250)) [ChangeRequestedOn],    
  CAST(coa.FormattedOwner AS NVARCHAR(250)) [FormattedRequestor],    
  CAST(coa.DateOfBirth AS NVARCHAR(250)) [DateOfBirth],    
  CAST(coa.FormattedOldAddress AS NVARCHAR(250)) [FormattedOldAddress],    
  CAST(coa.FormattedOldCityStZip AS NVARCHAR(250)) [FormattedOldCityStZip],    
  CAST(coa.FormattedNewAddress AS NVARCHAR(250)) [FormattedNewAddress],    
  CAST(coa.FormattedNewCityStZip AS NVARCHAR(250)) [FormattedNewCityStZip],    
  CAST(coa.PropertyID AS NVARCHAR(250)) [PropertyID],    
  CAST(coa.AddressMatch AS NVARCHAR(250)) [AddressMatch],    
  CAST(coa.NameMatch AS NVARCHAR(250)) [NameMatch],    
  CAST(coa.DOBMatch AS NVARCHAR(250)) [DOBMatch],    
  CAST(coa.ActiveHomestead AS NVARCHAR(250)) [ActiveHomestead]    
    
 FROM #temp_join coa    
 ORDER BY     
  ID ASC    
 OFFSET 0 ROWS    
 FETCH NEXT 1 ROWS ONLY    
 ) x    
 UNPIVOT    
  (Request FOR RequestorData IN (ChangeRequestedOn, FormattedRequestor, DateOfBirth, FormattedOldAddress, FormattedOldCityStZip, FormattedNewAddress, FormattedNewCityStZip, PropertyID, AddressMatch, NameMatch, DOBMatch, ActiveHomestead)    
  ) AS unpvt    
  ' + @join    
    
    
 SET @Field = REPLACE(@Field,']',' NVARCHAR(250)') + ' );';    
 SET @Field = REPLACE(@Field,'.[','_');    
 --PRINT @Field    
    
 DECLARE @tableName NVARCHAR(24) = 'ChangeOfAddressRpt';    
    
 DECLARE @Table NVARCHAR(MAX) = ' CREATE TABLE '+@tableName+' ( RequestorData NVARCHAR(250), Request NVARCHAR(250)';    
 SET @Table = @Table + @Field    
 --PRINT @Table    
 --PRINT @sql    
    
 EXEC (@Table);    
    
 INSERT INTO ChangeOfAddressRpt EXEC (@sql);    
    
 --SELECT * FROM ChangeOfAddressRpt    
    
 DECLARE @month NVARCHAR(24) = (SELECT TOP 1 MONTH(ChangeRequestedOn) FROM dbo.ext_ChangeOfAddress GROUP BY MONTH(ChangeRequestedOn) ORDER BY MONTH(ChangeRequestedOn) ASC);    
 DECLARE @year NVARCHAR(4) = (SELECT TOP 1 YEAR(ChangeRequestedOn) FROM dbo.ext_ChangeOfAddress GROUP BY YEAR(ChangeRequestedOn) ORDER BY YEAR(ChangeRequestedOn) ASC);    
 DECLARE @db NVARCHAR(64) = (SELECT DB_NAME());    
 DECLARE @csvName NVARCHAR(64) = CONCAT(@month, '-', @year, '_', N'Processed_CoA_Report.csv');    
 DECLARE @fullCsvPath NVARCHAR(256) = CONCAT(@exportPath,'\',@csvName);    
 DECLARE @fullXlsxPath NVARCHAR(256) = REPLACE(@fullCsvPath, 'csv', 'xlsx');    
    
 --PRINT @RecordCount;    
    
 DECLARE @ColumnCharacter TABLE (Number SMALLINT, Letter NVARCHAR(4));    
 INSERT INTO @ColumnCharacter    
 (    
  Number,    
  LETTER    
 )    
 VALUES    
 (1,'A'), (2,'B'), (3,'C'), (4,'D'), (5,'E'), (6,'F'), (7,'G'), (8,'H'), (9,'I'), (10,'J'), (11,'K'), (12,'L'),    
 (13,'M'), (14,'N'), (15,'O'), (16,'P'), (17,'Q'), (18,'R'), (19,'S'), (20,'T'), (21,'U'), (22,'V'), (23,'W'), (24,'X'),    
 (25,'Y'), (26,'Z'), (27,'AA'), (28,'AB'), (29,'AC'), (30,'AD'), (31,'AE'), (32,'AF'), (33,'AG'), (34,'AH'), (35,'AI'), (36,'AJ'),    
 (37,'AK'), (38,'AL'), (39,'AM'), (40,'AN'), (41,'AO'), (42,'AP'), (43,'AQ'), (44,'AR'), (45,'AS'), (46,'AT'), (47,'AU'), (48,'AV'),    
 (49,'AW'), (50,'AX'), (51,'AY'), (52,'AZ'), (53,'BA'), (54,'BB'), (55,'BC'), (56,'BD'), (57,'BE'), (58,'BF'), (59,'BG'), (60,'BH');    
    
    
 --SELECT * FROM @ColumnCharacter    
 SET @Letter = (SELECT Letter FROM @ColumnCharacter WHERE Number = (@RecordCount + 1));    
    
 SET @RecordCount = (SELECT COUNT(*) + 1 FROM dbo.ChangeOfAddressRpt);    
    
 EXEC dbo.ext_ExportDataToCsv @dbName = @db,          -- nvarchar(100)    
         @includeHeaders = 1, -- bit    
         @filePath = @exportPath,        -- nvarchar(512)    
         @tableName = N'ChangeOfAddressRpt',       -- nvarchar(100)    
         @reportName = @csvName,      -- nvarchar(100)    
         @delimiter = N'|'        -- nvarchar(4)    
    
    
 EXEC dbo.CSVtoXLSXwTable @fullCsvPath = @fullCsvPath,  -- varchar(512)    
        @fullXlsxPath = @fullXlsxPath, -- varchar(512)    
        @rowCount = @RecordCount,      -- int    
        @colCharacter = @Letter  -- varchar(4)    
    
    
END;    
    
END TRY    
BEGIN CATCH    
    
 DECLARE @param_values NVARCHAR(512) =    
   (    
    SELECT     
     ISNULL(@exportPath,'')  AS [exportPath],    
     ISNULL(@YearID,'') AS [YearID],    
     ISNULL(@tran_count_on_entry,'') AS [tran_count_on_entry]    
    FOR JSON PATH    
   );    
    
 EXEC @return_value = dbo.sp_handle_exception    
        @client_id = 1,     
        @source_procedure_id = @@procid,     
        @additional_info = @param_values;    
    
    
END CATCH    
    
RETURN @return_value;    
    
END