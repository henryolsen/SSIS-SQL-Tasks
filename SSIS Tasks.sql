/*  CREATED BY HENRY OLSEN, 8/21/2019 

THIS FILE CONTAINS SQL TASKS THAT CAN BE USED IN AN SSIS PACKAGE

SQL TASK 1: IS THE PACKAGE ALREADY RUNNING?
SQL TASK 2: START TIME OF PACKAGE
SQL TASK 3: END TIME OF PACKAGE
SQL TASK 4: OnError Event Handler (Update statement for tbDataLoaders if there is an error.)
SQL TASK 5: OnTaskFailed Event Handler (Send Email)

*/


--===================Creating the DataLoader table in SQL Server.==========

CREATE TABLE tbDataLoaders(
 DateLog VARCHAR(12), --Date the package was run in YYYYMMDD format
 PackageName VARCHAR(100), --Name of the SSIS Package being run
 StartDateTime DATETIME, --Start date and time when the package was run 
 EndDateTime DATETIME,  --End date and time when the package was run 
 ErrorMessage nvarchar(MAX), --If the package failed, an error message will be loaded into this column
 SuccessFlag BIT, --1 if the Package executed successfully
 RunTimeInSeconds decimal(13,5), --The amount of time the package ran in seconds
 MachineName nvarchar(45), --Name of the machine the package is running on
 UserName nvarchar(30) --Name of the user running the package
)

/* There are built in parameters in SSIS for the name of the package, error message, machine name, and user name.  Anything with a ? is where the built-in
parameters are placed.  This can be useful when reusing the following SQL Scripts in multiple packages.
*/ 

--===SQL TASK 1:  IS THE PACKAGE RUNNING ALREADY?

DECLARE @error nvarchar(300)
DECLARE @machine nvarchar(50)
DECLARE @username nvarchar(50)
DECLARE @packagename nvarchar(50)

set @machine = (SELECT TOP 1 MachineName FROM tbDataLoaders WHERE PackageName =? ORDER BY StartDateTime DESC)

set @username = (SELECT TOP 1 UserName FROM tbDataLoaders WHERE PackageName =? ORDER BY StartDateTime DESC)

set @packagename = (SELECT TOP 1 PackageName FROM tbDataLoaders WHERE PackageName =? ORDER BY StartDateTime DESC)

--The error message answers the following questions: What package is already running?  What machine is it running on?  What user is running the package?

set @error= @packagename + ' is currently being run by the following user: ' + @username + '.'
+ ' The machine that this package is being run on' + ' is ' + @machine + '.' 

/* If there is no end date time in the data loaders table for the package being run, we can assume that the package is already running and will throw an error.  
However, could there be other reasons why EndDateTime would be NULL?  One I can think of is if the package gets manually stopped in the middle of execution.
*/

IF EXISTS(SELECT * FROM tbDataLoaders WHERE EndDateTime IS NULL AND PackageName = ?) 
 THROW 51000, @error, 1;
 
--===SQL TASK 2:START TIME OF PACKAGE
 
 DECLARE @date NVARCHAR(12) = CONVERT(NVARCHAR(12),GETDATE(),112)
INSERT INTO tbDataLoaders (DateLog,PackageName, StartDateTime, MachineName, UserName)
VALUES (@date, ? , GETDATE(),?,?)

--===SQL TASK 3:END TIME OF PACKAGE

DECLARE @date VARCHAR(12) = CONVERT(VARCHAR(12),GETDATE(),112)
UPDATE tbDataLoaders
SET EndDateTime = GETDATE()
, SuccessFlag=1
, RunTimeInSeconds = CAST(CONCAT(DATEDIFF(ss, startdatetime, GETDATE()) 
,'.'
,DATEDIFF(ms, startdatetime, GETDATE())) AS decimal(13,5))
WHERE PackageName = ? AND EndDateTime IS NULL

--===SQL TASK 4:OnError Event Handler
--Updates the tbDataLoaders table with error information when an error is thrown off.

UPDATE tbDataLoaders
SET SuccessFlag=0
, EndDateTime = GETDATE()
, ErrorMessage = ?
,RunTimeInSeconds = CAST(CONCAT(DATEDIFF(ss, startdatetime, GETDATE()) 
,'.'
,DATEDIFF(ms, startdatetime, GETDATE())) AS decimal(13,5))
WHERE PackageName = ? AND EndDateTime IS NULL

--===SQL TASK 5:  OnTaskFailed Event Handler sending an automated email to multiple email addresses.  

--Cursor used to send email to multiple people.  The emails in the EMAIL_LIST table are those who have access to the package
DECLARE MDM_EMAIL_LIST_CURSOR CURSOR FOR
	SELECT PACKAGE_NAME, EMAIL_ADDRESS FROM HenryDB.dbo.EMAIL_LIST 
OPEN MDM_EMAIL_LIST_CURSOR

DECLARE @packageerror AS nvarchar(MAX)
DECLARE @packagename AS nvarchar(50)
DECLARE @errordesc AS nvarchar(MAX)
DECLARE @EmailAddress NVARCHAR(100)
DECLARE @ListOfRecipients NVARCHAR(MAX)


SET @packagename= ?
SET @errordesc= (SELECT TOP (1)
      ErrorMessage
   FROM [HenryDB].[dbo].[tbDataLoaders]
WHERE PackageName=?
   ORDER BY StartDateTime DESC)

--@packageerror is the error message that is being emailed
SET @packageerror= '<font size = 5><b>An error has occurred in the following package: </b></font>' + '&nbsp;' + 
+ '<font size="4" color="red">' + @packagename + '</font>' + '.' + '<br>' + '<br>' + 
'<font size = 5><b>The error is: </b></font>'
+ '&nbsp;' + '<font size="4" color="red">'+ @errordesc + '</font>'

--Cursor used to send email to multiple people.  

FETCH NEXT FROM MDM_EMAIL_LIST_CURSOR INTO @PackageName, @EmailAddress

WHILE @@FETCH_STATUS=0
	IF @PackageName=?
	BEGIN

		SET @ListOfRecipients= @EmailAddress + ';' 
		FETCH NEXT FROM MDM_EMAIL_LIST_CURSOR INTO @PackageName, @EmailAddress
		exec sp_send_dbmail    
       @recipients = @ListOfRecipients
     ,  @from_address =  'no-reply@churchofjesuschrist.org'     
     ,  @subject =  'Error'    
     ,  @body =  @packageerror
  ,@body_format = 'HTML'
  ,@query_result_separator= ' '
		
	END

	

CLOSE MDM_EMAIL_LIST_CURSOR
DEALLOCATE MDM_EMAIL_LIST_CURSOR




