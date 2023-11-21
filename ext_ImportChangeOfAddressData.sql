IF OBJECT_ID('dbo.ext_ImportChangeOfAddressData') IS NOT NULL
BEGIN
	DROP PROCEDURE dbo.ext_ImportChangeOfAddressData
END

CREATE PROCEDURE [dbo].[ext_ImportChangeOfAddressData] (@filepath NVARCHAR(256))  
AS  
BEGIN  
/*=======================  
Created By Tyler T on 6/8/23 to import CoA report data from Clerk  
========================*/  
BEGIN TRY  
  
DECLARE @fileExists BIT = (SELECT dbo.doesFileExist (@filePath) as IsExists),  
  @tran_count_on_entry INT = @@TRANCOUNT,  
  @return_value INT = @@ERROR;  
  
IF @fileExists = 1  
 BEGIN  
  
 DECLARE @sql NVARCHAR(1000);  
  
 IF OBJECT_ID('dbo.ext_ChangeOfAddress') IS NOT NULL  
 BEGIN  
  DROP TABLE dbo.ext_ChangeOfAddress  
 END  
  
 CREATE TABLE dbo.ext_ChangeOfAddress (  
            FormattedOwner VARCHAR(256) NULL,  
            FirstName VARCHAR(64) NULL,  
            MiddleName VARCHAR(64) NULL,  
            LastName VARCHAR(64) NULL,  
            DateOfBirth DATE NULL,  
            FormattedOldAddress VARCHAR(512) NULL,  
            FormattedOldCityStZip VARCHAR(512) NULL,  
            OldStreetNum VARCHAR(32) NULL,  
            OldStreetDir VARCHAR(32) NULL,  
            OldStreet VARCHAR(64) NULL,  
            OldApt VARCHAR(32) NULL,  
            OldCity VARCHAR(32) NULL,  
            OldState VARCHAR(32) NULL,  
            OldZip VARCHAR(32) NULL,  
            FormattedNewAddress VARCHAR(512) NULL,  
            FormattedNewCityStZip VARCHAR(512) NULL,  
            ChangeRequestedOn DATETIME NULL);  
  
 SET @sql =   
 'BULK INSERT dbo.ext_ChangeOfAddress  
 FROM '''+@filepath+'''  
 WITH (  
   FIELDTERMINATOR = ''|'',  
   ROWTERMINATOR = ''\n''  
   );';  
  
 BEGIN TRY    
  EXEC (@sql);    
 END TRY    
  
 BEGIN CATCH    
  THROW;    
 END CATCH    
    
  
 ALTER TABLE dbo.ext_ChangeOfAddress  
 ADD [ID] SMALLINT IDENTITY(1,1);  
  
 --SELECT * FROM dbo.ext_ChangeOfAddress  
  
 UPDATE dbo.ext_ChangeOfAddress  
 SET LastName = REPLACE(LastName,'''','')  
 FROM dbo.ext_ChangeOfAddress  
  
 UPDATE ext_ChangeOfAddress  
 SET DateOfBirth = ''  
 WHERE DateOfBirth IS NULL  
  
 UPDATE ext_ChangeOfAddress  
 SET FormattedOldAddress = ''  
 WHERE FormattedOldAddress IS NULL  
  
 UPDATE ext_ChangeOfAddress  
 SET FormattedOldCityStZip = ''  
 WHERE FormattedOldCityStZip IS NULL  
  
 UPDATE ext_ChangeOfAddress  
 SET FormattedNewAddress = ''  
 WHERE FormattedNewAddress IS NULL  
  
 UPDATE ext_ChangeOfAddress  
 SET FormattedNewCityStZip = ''  
 WHERE FormattedNewCityStZip IS NULL  
  
 END  
ELSE  
 BEGIN  
  
 PRINT('File does not exist!')  
 RETURN;  
  
 END  
  
END TRY  
BEGIN CATCH  
  
  DECLARE @param_values NVARCHAR(512) =  
    (  
     SELECT   
      @filePath  AS [filePath],  
      @fileExists AS [fileExists],  
      @tran_count_on_entry AS [tran_count_on_entry]  
     FOR JSON PATH  
    );  
  
  EXEC @return_value = dbo.sp_handle_exception  
         @client_id = 1,   
         @source_procedure_id = @@procid,   
         @additional_info = @param_values;  
  
END CATCH  
  
RETURN @return_value;  
   
END