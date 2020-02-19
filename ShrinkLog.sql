
USE [master]
GO

SET NOCOUNT ON

DECLARE @CreateJobs nvarchar(max)
DECLARE @BackupDirectory nvarchar(max)
DECLARE @OutputFileDirectory nvarchar(max)
DECLARE @LogToTable nvarchar(max)
DECLARE @Version numeric(18,10)
DECLARE @Error int
DECLARE @MaxVersion int

SET @CreateJobs          = 'Y'          -- Specify whether jobs should be created.
SET @BackupDirectory     = N'D:\MSSQL'	-- Specify the backup root directory.
SET @OutputFileDirectory = NULL         -- Specify the output file directory. If no directory is specified, then the SQL Server error log directory is used.
SET @LogToTable          = 'Y'          -- Log commands to a table.
SET @MaxVersion			 = 15

SET @Error = 0

SET @Version = CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)),CHARINDEX('.',CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max))) - 1) + '.' + REPLACE(RIGHT(CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)), LEN(CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max))) - CHARINDEX('.',CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)))),'.','') AS numeric(18,10))

IF IS_SRVROLEMEMBER('sysadmin') = 0
BEGIN
  RAISERROR('You need to be a member of the SysAdmin server role to install the solution.',16,1)
  SET @Error = @@ERROR
END

IF OBJECT_ID('tempdb..#Config') IS NOT NULL DROP TABLE #Config

CREATE TABLE #Config ([Name] nvarchar(max),
                      [Value] nvarchar(max))

IF @CreateJobs = 'Y' AND @OutputFileDirectory IS NULL AND SERVERPROPERTY('EngineEdition') <> 4 AND @Version < @MaxVersion
BEGIN
  IF @Version >= 11
  BEGIN
    SELECT @OutputFileDirectory = [path]
    FROM sys.dm_os_server_diagnostics_log_configurations
  END
  ELSE
  BEGIN
    SELECT @OutputFileDirectory = LEFT(CAST(SERVERPROPERTY('ErrorLogFileName') AS nvarchar(max)),LEN(CAST(SERVERPROPERTY('ErrorLogFileName') AS nvarchar(max))) - CHARINDEX('\',REVERSE(CAST(SERVERPROPERTY('ErrorLogFileName') AS nvarchar(max)))))
  END
END

IF @CreateJobs = 'Y' AND RIGHT(@OutputFileDirectory,1) = '\' AND SERVERPROPERTY('EngineEdition') <> 4
BEGIN
  SET @OutputFileDirectory = LEFT(@OutputFileDirectory, LEN(@OutputFileDirectory) - 1)
END

INSERT INTO #Config ([Name], [Value])
VALUES('CreateJobs', @CreateJobs)

INSERT INTO #Config ([Name], [Value])
VALUES('BackupDirectory', @BackupDirectory)

INSERT INTO #Config ([Name], [Value])
VALUES('OutputFileDirectory', @OutputFileDirectory)

INSERT INTO #Config ([Name], [Value])
VALUES('LogToTable', @LogToTable)

INSERT INTO #Config ([Name], [Value])
VALUES('DatabaseName', DB_NAME(DB_ID()))

INSERT INTO #Config ([Name], [Value])
VALUES('Error', CAST(@Error AS nvarchar))

IF OBJECT_ID('[dbo].[ShrinkLog]') IS NOT NULL DROP PROCEDURE [dbo].[ShrinkLog]

/****** Objet :  StoredProcedure [dbo].[sp_dba_shrinkLog] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[ShrinkLog] 

@Databases nvarchar(max),
@Execute nvarchar(max) = 'Y',
@LogToTable nvarchar(max)

AS 
BEGIN 


  ----------------------------------------------------------------------------------------------------
  --// Set options                                                                                //--
  ----------------------------------------------------------------------------------------------------

  SET NOCOUNT ON

  ----------------------------------------------------------------------------------------------------
  --// Declare variables                                                                          //--
  ----------------------------------------------------------------------------------------------------

  DECLARE @StartMessage nvarchar(max)
  DECLARE @EndMessage nvarchar(max)
  DECLARE @DatabaseMessage nvarchar(max)
  DECLARE @ErrorMessage nvarchar(max)

  DECLARE @CurrentID int
  DECLARE @CurrentDatabase nvarchar(max)
  DECLARE @CurrentDate datetime
  DECLARE @CurrentIsDatabaseAccessible bit
  DECLARE @CurrentMirroringRole nvarchar(max)
  
  DECLARE @CurrentCommand01 nvarchar(max)
  
  DECLARE @CurrentCommandOutput01 int
  
  DECLARE @tmpDatabases TABLE (ID int IDENTITY,
                               DatabaseName nvarchar(max),
                               DatabaseNameFS nvarchar(max),
                               DatabaseType nvarchar(max),
                               Selected bit,
                               Completed bit,
                               PRIMARY KEY(Selected, Completed, ID))

  DECLARE @SelectedDatabases TABLE (DatabaseName nvarchar(max),
                                    DatabaseType nvarchar(max),
                                    Selected bit)

 -- DECLARE @tmpDatabases TABLE (ID int IDENTITY PRIMARY KEY,
 --                              DatabaseName nvarchar(max),
 --                              Completed bit)

  DECLARE @Error int

  SET @Error = 0

  ----------------------------------------------------------------------------------------------------
  --// Log initial information                                                                    //--
  ----------------------------------------------------------------------------------------------------

  SET @StartMessage = 'DateTime: ' + CONVERT(nvarchar,GETDATE(),120) + CHAR(13) + CHAR(10)
  SET @StartMessage = @StartMessage + 'Server: ' + CAST(SERVERPROPERTY('ServerName') AS nvarchar) + CHAR(13) + CHAR(10)
  SET @StartMessage = @StartMessage + 'Version: ' + CAST(SERVERPROPERTY('ProductVersion') AS nvarchar) + CHAR(13) + CHAR(10)
  SET @StartMessage = @StartMessage + 'Edition: ' + CAST(SERVERPROPERTY('Edition') AS nvarchar) + CHAR(13) + CHAR(10)
  SET @StartMessage = @StartMessage + 'Procedure: ' + QUOTENAME(DB_NAME(DB_ID())) + '.' + (SELECT QUOTENAME(sys.schemas.name) FROM sys.schemas INNER JOIN sys.objects ON sys.schemas.[schema_id] = sys.objects.[schema_id] WHERE [object_id] = @@PROCID) + '.' + QUOTENAME(OBJECT_NAME(@@PROCID)) + CHAR(13) + CHAR(10)
  SET @StartMessage = @StartMessage + 'Parameters: @Databases = ' + ISNULL('''' + REPLACE(@Databases,'''','''''') + '''','NULL')
  SET @StartMessage = @StartMessage + ', @Execute = ' + ISNULL('''' + REPLACE(@Execute,'''','''''') + '''','NULL')
  SET @StartMessage = @StartMessage + CHAR(13) + CHAR(10)
  SET @StartMessage = REPLACE(@StartMessage,'%','%%')
  RAISERROR(@StartMessage,10,1) WITH NOWAIT

  ----------------------------------------------------------------------------------------------------
  --// Select databases                                                                           //--
  ----------------------------------------------------------------------------------------------------

  SET @Databases = REPLACE(@Databases, ', ', ',');

  WITH Databases1 (StartPosition, EndPosition, DatabaseItem) AS
  (
  SELECT 1 AS StartPosition,
         ISNULL(NULLIF(CHARINDEX(',', @Databases, 1), 0), LEN(@Databases) + 1) AS EndPosition,
         SUBSTRING(@Databases, 1, ISNULL(NULLIF(CHARINDEX(',', @Databases, 1), 0), LEN(@Databases) + 1) - 1) AS DatabaseItem
  WHERE @Databases IS NOT NULL
  UNION ALL
  SELECT CAST(EndPosition AS int) + 1 AS StartPosition,
         ISNULL(NULLIF(CHARINDEX(',', @Databases, EndPosition + 1), 0), LEN(@Databases) + 1) AS EndPosition,
         SUBSTRING(@Databases, EndPosition + 1, ISNULL(NULLIF(CHARINDEX(',', @Databases, EndPosition + 1), 0), LEN(@Databases) + 1) - EndPosition - 1) AS DatabaseItem
  FROM Databases1
  WHERE EndPosition < LEN(@Databases) + 1
  ),
  Databases2 (DatabaseItem, Selected) AS
  (
  SELECT CASE WHEN DatabaseItem LIKE '-%' THEN RIGHT(DatabaseItem,LEN(DatabaseItem) - 1) ELSE DatabaseItem END AS DatabaseItem,
         CASE WHEN DatabaseItem LIKE '-%' THEN 0 ELSE 1 END AS Selected
  FROM Databases1
  ),
  Databases3 (DatabaseItem, DatabaseType, Selected) AS
  (
  SELECT CASE WHEN DatabaseItem IN('ALL_DATABASES','SYSTEM_DATABASES','USER_DATABASES') THEN '%' ELSE DatabaseItem END AS DatabaseItem,
         CASE WHEN DatabaseItem = 'SYSTEM_DATABASES' THEN 'S' WHEN DatabaseItem = 'USER_DATABASES' THEN 'U' ELSE NULL END AS DatabaseType,
         Selected
  FROM Databases2
  ),
  Databases4 (DatabaseName, DatabaseType, Selected) AS
  (
  SELECT CASE WHEN LEFT(DatabaseItem,1) = '[' AND RIGHT(DatabaseItem,1) = ']' THEN PARSENAME(DatabaseItem,1) ELSE DatabaseItem END AS DatabaseItem,
         DatabaseType,
         Selected
  FROM Databases3
  )
  INSERT INTO @SelectedDatabases (DatabaseName, DatabaseType, Selected)
  SELECT DatabaseName,
         DatabaseType,
         Selected
  FROM Databases4
  OPTION (MAXRECURSION 0)

  INSERT INTO @tmpDatabases (DatabaseName, DatabaseNameFS, DatabaseType, Selected, Completed)
  SELECT [name] AS DatabaseName,
         REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE([name],'\',''),'/',''),':',''),'*',''),'?',''),'"',''),'<',''),'>',''),'|',''),' ','') AS DatabaseNameFS,
         CASE WHEN name IN('master','msdb','model') THEN 'S' ELSE 'U' END AS DatabaseType,
         0 AS Selected,
         0 AS Completed
  FROM sys.databases
  WHERE [name] <> 'tempdb'
  AND source_database_id IS NULL
  ORDER BY [name] ASC

  UPDATE tmpDatabases
  SET tmpDatabases.Selected = SelectedDatabases.Selected
  FROM @tmpDatabases tmpDatabases
  INNER JOIN @SelectedDatabases SelectedDatabases
  ON tmpDatabases.DatabaseName LIKE REPLACE(SelectedDatabases.DatabaseName,'_','[_]')
  AND (tmpDatabases.DatabaseType = SelectedDatabases.DatabaseType OR SelectedDatabases.DatabaseType IS NULL)
  WHERE SelectedDatabases.Selected = 1

  UPDATE tmpDatabases
  SET tmpDatabases.Selected = SelectedDatabases.Selected
  FROM @tmpDatabases tmpDatabases
  INNER JOIN @SelectedDatabases SelectedDatabases
  ON tmpDatabases.DatabaseName LIKE REPLACE(SelectedDatabases.DatabaseName,'_','[_]')
  AND (tmpDatabases.DatabaseType = SelectedDatabases.DatabaseType OR SelectedDatabases.DatabaseType IS NULL)
  WHERE SelectedDatabases.Selected = 0

  IF @Databases IS NULL OR NOT EXISTS(SELECT * FROM @SelectedDatabases) OR EXISTS(SELECT * FROM @SelectedDatabases WHERE DatabaseName IS NULL OR DatabaseName = '')
  BEGIN
    SET @ErrorMessage = 'The value for the parameter @Databases is not supported.' + CHAR(13) + CHAR(10) + ' '
    RAISERROR(@ErrorMessage,16,1) WITH NOWAIT
    SET @Error = @@ERROR
  END

  ----------------------------------------------------------------------------------------------------
  --// Check database names                                                                       //--
  ----------------------------------------------------------------------------------------------------

  SET @ErrorMessage = ''
  SELECT @ErrorMessage = @ErrorMessage + QUOTENAME(DatabaseName) + ', '
  FROM @tmpDatabases
  WHERE Selected = 1
  AND DatabaseNameFS = ''
  ORDER BY DatabaseName ASC
  IF @@ROWCOUNT > 0
  BEGIN
    SET @ErrorMessage = 'The names of the following databases are not supported: ' + LEFT(@ErrorMessage,LEN(@ErrorMessage)-1) + '.' + CHAR(13) + CHAR(10) + ' '
    RAISERROR(@ErrorMessage,16,1) WITH NOWAIT
    SET @Error = @@ERROR
  END

  SET @ErrorMessage = ''
  SELECT @ErrorMessage = @ErrorMessage + QUOTENAME(DatabaseName) + ', '
  FROM @tmpDatabases
  WHERE UPPER(DatabaseNameFS) IN(SELECT UPPER(DatabaseNameFS) FROM @tmpDatabases GROUP BY UPPER(DatabaseNameFS) HAVING COUNT(*) > 1)
  AND UPPER(DatabaseNameFS) IN(SELECT UPPER(DatabaseNameFS) FROM @tmpDatabases WHERE Selected = 1)
  AND DatabaseNameFS <> ''
  ORDER BY DatabaseName ASC
  OPTION (RECOMPILE)
  IF @@ROWCOUNT > 0
  BEGIN
    SET @ErrorMessage = 'The names of the following databases are not unique in the file system: ' + LEFT(@ErrorMessage,LEN(@ErrorMessage)-1) + '.' + CHAR(13) + CHAR(10) + ' '
    RAISERROR(@ErrorMessage,16,1) WITH NOWAIT
    SET @Error = @@ERROR
  END

    ----------------------------------------------------------------------------------------------------
  --// Check input parameters                                                                     //--
  ----------------------------------------------------------------------------------------------------

  IF @Execute NOT IN('Y','N') OR @Execute IS NULL
  BEGIN
    SET @ErrorMessage = 'The value for parameter @Execute is not supported.' + CHAR(13) + CHAR(10)
    RAISERROR(@ErrorMessage,16,1) WITH NOWAIT
    SET @Error = @@ERROR
  END

  ----------------------------------------------------------------------------------------------------
  --// Check error variable                                                                       //--
  ----------------------------------------------------------------------------------------------------

  IF @Error <> 0 GOTO Logging

  ----------------------------------------------------------------------------------------------------
  --// Execute backup commands                                                                    //--
  ----------------------------------------------------------------------------------------------------

  WHILE EXISTS (SELECT * FROM @tmpDatabases WHERE Completed = 0)
  BEGIN

    SELECT TOP 1 @CurrentID = ID,
                 @CurrentDatabase = DatabaseName
    FROM @tmpDatabases
    WHERE Completed = 0
    ORDER BY ID ASC

    IF EXISTS (SELECT * FROM sys.database_recovery_status WHERE database_id = DB_ID(@CurrentDatabase) AND database_guid IS NOT NULL)
    BEGIN
      SET @CurrentIsDatabaseAccessible = 1
    END
    ELSE
    BEGIN
      SET @CurrentIsDatabaseAccessible = 0
    END

    SELECT @CurrentMirroringRole = mirroring_role_desc
    FROM sys.database_mirroring
    WHERE database_id = DB_ID(@CurrentDatabase)

    -- Set database message
    SET @DatabaseMessage = 'DateTime: ' + CONVERT(nvarchar,GETDATE(),120) + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'Database: ' + QUOTENAME(@CurrentDatabase) + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'Status: ' + CAST(DATABASEPROPERTYEX(@CurrentDatabase,'Status') AS nvarchar) + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'Mirroring role: ' + ISNULL(@CurrentMirroringRole,'None') + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'Standby: ' + CASE WHEN DATABASEPROPERTYEX(@CurrentDatabase,'IsInStandBy') = 1 THEN 'Yes' ELSE 'No' END + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'Updateability: ' + CAST(DATABASEPROPERTYEX(@CurrentDatabase,'Updateability') AS nvarchar) + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'User access: ' + CAST(DATABASEPROPERTYEX(@CurrentDatabase,'UserAccess') AS nvarchar) + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'Is accessible: ' + CASE WHEN @CurrentIsDatabaseAccessible = 1 THEN 'Yes' ELSE 'No' END + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = @DatabaseMessage + 'Recovery model: ' + CAST(DATABASEPROPERTYEX(@CurrentDatabase,'Recovery') AS nvarchar) + CHAR(13) + CHAR(10)
    SET @DatabaseMessage = REPLACE(@DatabaseMessage,'%','%%')
    RAISERROR(@DatabaseMessage,10,1) WITH NOWAIT

    IF DATABASEPROPERTYEX(@CurrentDatabase,'Status') = 'ONLINE'
    AND NOT (DATABASEPROPERTYEX(@CurrentDatabase,'UserAccess') = 'SINGLE_USER' AND @CurrentIsDatabaseAccessible = 0)
    AND DATABASEPROPERTYEX(@CurrentDatabase,'IsInStandBy') = 0
    BEGIN

      -- Set variables
      SET @CurrentDate = GETDATE()

      declare @varlog sysname
      declare @shrinksize sysname
	  declare @count int
	  declare @iRow int
  
	  CREATE TABLE #tbl (
		RowID INT IDENTITY(1, 1) NOT NULL PRIMARY KEY,
		name varchar(max),
		used int
	  )
  
      exec('use ['+@CurrentDatabase+']; insert #tbl
		select
			name,
			case
				when cast((cntr_value/1024) as int) = 0 then 1
				else cast((cntr_value/1024) as int)
			end as used
		from
			['+@CurrentDatabase+'].dbo.sysfiles s,
			master.sys.dm_os_performance_counters m
		where
			s.groupid=0
			and m.counter_name=''Log File(s) Used Size (KB)''
			and m.instance_name='''+@CurrentDatabase+'''')
      
	  SET @count = @@ROWCOUNT
	  SET @iRow = 1
	  
	  WHILE @iRow <= @count 
	  BEGIN
		  select @varlog = name, @shrinksize = used from #tbl where RowID = @iRow
	      SELECT @CurrentCommand01 = 'use ['+ @CurrentDatabase+'];DBCC SHRINKFILE(''' + @varlog + ''','+ @shrinksize +')'
		  EXECUTE @CurrentCommandOutput01 = [dbo].[CommandExecute] @Command = @CurrentCommand01, @CommandType = 'ShrinkLog', @Mode = 1, @DatabaseName = @CurrentDatabase, @LogToTable = @LogToTable, @Execute = @Execute
          SET @Error = @@ERROR
          IF @Error <> 0 SET @CurrentCommandOutput01 = @Error
		  
	      SET @iRow = @iRow + 1
	  END
	
	  DROP TABLE #tbl
	
    END

    -- Update that the database is completed
    UPDATE @tmpDatabases
    SET Completed = 1
    WHERE ID = @CurrentID

    -- Clear variables
    SET @CurrentID = NULL
    SET @CurrentDatabase = NULL
    SET @CurrentDate = NULL
    SET @CurrentIsDatabaseAccessible = NULL
    SET @CurrentMirroringRole = NULL

    SET @CurrentCommand01 = NULL
    
    SET @CurrentCommandOutput01 = NULL
    
  END

  ----------------------------------------------------------------------------------------------------
  --// Log completing information                                                                 //--
  ----------------------------------------------------------------------------------------------------

  Logging:
  SET @EndMessage = 'DateTime: ' + CONVERT(nvarchar,GETDATE(),120)
  SET @EndMessage = REPLACE(@EndMessage,'%','%%')
  RAISERROR(@EndMessage,10,1) WITH NOWAIT

  ----------------------------------------------------------------------------------------------------

END
GO

IF (SELECT CAST([Value] AS int) FROM #Config WHERE Name = 'Error') = 0
AND (SELECT [Value] FROM #Config WHERE Name = 'CreateJobs') = 'Y'
AND SERVERPROPERTY('EngineEdition') <> 4
BEGIN
	DECLARE @BackupDirectory nvarchar(max)
	DECLARE @OutputFileDirectory nvarchar(max)
	DECLARE @LogToTable nvarchar(max)
	DECLARE @DatabaseName nvarchar(max)

	DECLARE @Version numeric(18,10)
	
	DECLARE @TokenServer nvarchar(max)
	DECLARE @TokenJobID nvarchar(max)
	DECLARE @TokenStepID nvarchar(max)
	DECLARE @TokenDate nvarchar(max)
	DECLARE @TokenTime nvarchar(max)
	DECLARE @TokenLogDirectory nvarchar(max)

	DECLARE @JobDescription nvarchar(max)
	DECLARE @JobCategory nvarchar(max)
	DECLARE @JobOwner nvarchar(max)
	
	DECLARE @JobName nvarchar(max)
	DECLARE @StepName01 nvarchar(max)
	DECLARE @StepName02 nvarchar(max)
	
	DECLARE @JobCommand01 nvarchar(max)
	DECLARE @JobCommand02 nvarchar(max)

	DECLARE @OutputFile01 nvarchar(max)
	DECLARE @OutputFile02 nvarchar(max)

	SET @Version = CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)),CHARINDEX('.',CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max))) - 1) + '.' + REPLACE(RIGHT(CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)), LEN(CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max))) - CHARINDEX('.',CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)))),'.','') AS numeric(18,10))

	IF @Version >= 9.002047
	BEGIN
		SET @TokenServer = '$' + '(ESCAPE_SQUOTE(SRVR))'
		SET @TokenJobID = '$' + '(ESCAPE_SQUOTE(JOBID))'
		SET @TokenStepID = '$' + '(ESCAPE_SQUOTE(STEPID))'
		SET @TokenDate = '$' + '(ESCAPE_SQUOTE(STRTDT))'
		SET @TokenTime = '$' + '(ESCAPE_SQUOTE(STRTTM))'
	END
	ELSE
	BEGIN
		SET @TokenServer = '$' + '(SRVR)'
		SET @TokenJobID = '$' + '(JOBID)'
		SET @TokenStepID = '$' + '(STEPID)'
		SET @TokenDate = '$' + '(STRTDT)'
		SET @TokenTime = '$' + '(STRTTM)'
	END

	IF @Version >= 12
	BEGIN
		SET @TokenLogDirectory = '$' + '(ESCAPE_SQUOTE(SQLLOGDIR))'
	END

	SELECT @BackupDirectory = Value
	FROM #Config
	WHERE [Name] = 'BackupDirectory'

	SELECT @OutputFileDirectory = Value
	FROM #Config
	WHERE [Name] = 'OutputFileDirectory'

	SELECT @LogToTable = Value
	FROM #Config
	WHERE [Name] = 'LogToTable'

	SELECT @DatabaseName = Value
	FROM #Config
	WHERE [Name] = 'DatabaseName'

	SET @JobDescription = 'Solution de Maintenance Ola Hallengren, modifiée par NourY Solutions'
	SET @JobCategory = 'Database Maintenance'
	SET @JobOwner = SUSER_SNAME(0x01)

	SET @JobName = 'DBA_SHRINK_USER_DATABASES_LOG'
	SET @StepName01 = 'DBA_BACKUP_USER_DATABASES_LOG'
	SET @JobCommand01 = 'sqlcmd -E -S ' + @TokenServer + ' -d ' + @DatabaseName + ' -Q "EXECUTE [dbo].[DatabaseBackup] @Databases = ''USER_DATABASES'', @Directory = ' + ISNULL('N''' + REPLACE(@BackupDirectory,'''','''''') + '''','NULL') + ', @BackupType = ''LOG'', @Verify = ''Y'', @CleanupTime = ''24'', @CheckSum = ''N''' + CASE WHEN @LogToTable = 'Y' THEN ', @LogToTable = ''Y''' ELSE '' END + '" -b'
	SET @OutputFile01 = COALESCE(@OutputFileDirectory,@TokenLogDirectory) + '\' + 'DatabaseBackupLog_' + @TokenJobID + '_' + @TokenStepID + '_' + @TokenDate + '_' + @TokenTime + '.txt'
	IF LEN(@OutputFile01) > 200 SET @OutputFile01 = COALESCE(@OutputFileDirectory,@TokenLogDirectory) + '\' + @TokenJobID + '_' + @TokenStepID + '_' + @TokenDate + '_' + @TokenTime + '.txt'
	IF LEN(@OutputFile01) > 200 SET @OutputFile01 = NULL

	SET @StepName02 = 'DBA_SHRINK_USER_DATABASES_LOG'
	SET @JobCommand02 = 'sqlcmd -E -S ' + @TokenServer + ' -d ' + @DatabaseName + ' -Q "EXECUTE [dbo].[ShrinkLog] @Databases = ''USER_DATABASES'', @LogToTable = ''Y''" -b'
	SET @OutputFile02 = COALESCE(@OutputFileDirectory,@TokenLogDirectory) + '\' + 'DatabaseShrinkLog_' + @TokenJobID + '_' + @TokenStepID + '_' + @TokenDate + '_' + @TokenTime + '.txt'
	IF LEN(@OutputFile02) > 200 SET @OutputFile02 = COALESCE(@OutputFileDirectory,@TokenLogDirectory) + '\' + @TokenJobID + '_' + @TokenStepID + '_' + @TokenDate + '_' + @TokenTime + '.txt'
	IF LEN(@OutputFile02) > 200 SET @OutputFile02 = NULL

	IF NOT EXISTS (SELECT * FROM msdb.dbo.sysjobs WHERE [name] = @JobName)
	BEGIN
		EXECUTE msdb.dbo.sp_add_job @job_name = @JobName, @description = @JobDescription, @category_name = @JobCategory, @owner_login_name = @JobOwner, @enabled = 0
		EXECUTE msdb.dbo.sp_add_jobstep @job_name = @JobName, @step_name = @StepName01, @subsystem = 'CMDEXEC', @command = @JobCommand01, @output_file_name = @OutputFile01, @on_success_action = 3
		EXECUTE msdb.dbo.sp_add_jobstep @job_name = @JobName, @step_name = @StepName02, @subsystem = 'CMDEXEC', @command = @JobCommand02, @output_file_name = @OutputFile02
		EXECUTE msdb.dbo.sp_add_jobserver @job_name = @JobName
		EXECUTE msdb.dbo.sp_add_schedule @schedule_name = N'SCHED_DBA_SHRINK_USER_DATABASES_LOG', @freq_type = 8, @freq_interval = 62, @freq_recurrence_factor = 1, @active_start_time = 190000
		EXECUTE msdb.dbo.sp_attach_schedule @job_name = @JobName, @schedule_name = N'SCHED_DBA_SHRINK_USER_DATABASES_LOG'
	END
END