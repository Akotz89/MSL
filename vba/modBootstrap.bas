Attribute VB_Name = "modBootstrap"
Option Compare Database
Option Explicit

Public Sub BuildMslTables()
    On Error GoTo ErrHandler

    CreateTableIfMissing "tblSites", _
        "CREATE TABLE tblSites (" & _
        "SiteID AUTOINCREMENT CONSTRAINT pk_tblSites PRIMARY KEY, " & _
        "SiteName TEXT(100) NOT NULL, " & _
        "ContractNumber TEXT(100), " & _
        "CDRLNumber TEXT(50), " & _
        "Customer TEXT(100), " & _
        "Active YESNO)"

    CreateTableIfMissing "tblPersonnel", _
        "CREATE TABLE tblPersonnel (" & _
        "PersonnelID AUTOINCREMENT CONSTRAINT pk_tblPersonnel PRIMARY KEY, " & _
        "SiteID LONG NOT NULL, " & _
        "FirstName TEXT(50) NOT NULL, " & _
        "LastName TEXT(50) NOT NULL, " & _
        "Role TEXT(50), " & _
        "ArrivalOnStation DATETIME, " & _
        "Active YESNO)"

    CreateTableIfMissing "tblReportPeriods", _
        "CREATE TABLE tblReportPeriods (" & _
        "PeriodID AUTOINCREMENT CONSTRAINT pk_tblReportPeriods PRIMARY KEY, " & _
        "SiteID LONG NOT NULL, " & _
        "ReportMonth DATETIME NOT NULL, " & _
        "ReportTitle TEXT(150), " & _
        "PreparedBy LONG, " & _
        "ReportStatus TEXT(25), " & _
        "CreatedDate DATETIME, " & _
        "LockedDate DATETIME)"

    CreateTableIfMissing "tblLogEntries", _
        "CREATE TABLE tblLogEntries (" & _
        "LogID AUTOINCREMENT CONSTRAINT pk_tblLogEntries PRIMARY KEY, " & _
        "PeriodID LONG NOT NULL, " & _
        "SiteID LONG NOT NULL, " & _
        "LogDate DATETIME NOT NULL, " & _
        "Category TEXT(50) NOT NULL, " & _
        "SystemArea TEXT(100), " & _
        "Title TEXT(255) NOT NULL, " & _
        "Description LONGTEXT, " & _
        "Resolution LONGTEXT, " & _
        "LoggedBy LONG, " & _
        "EventCount LONG, " & _
        "ResponseTimeMinutes LONG, " & _
        "ReturnToServiceMinutes LONG, " & _
        "FollowUpRequired YESNO, " & _
        "CreatedDate DATETIME, " & _
        "ModifiedDate DATETIME)"

    CreateTableIfMissing "tblLogEntrySteps", _
        "CREATE TABLE tblLogEntrySteps (" & _
        "StepID AUTOINCREMENT CONSTRAINT pk_tblLogEntrySteps PRIMARY KEY, " & _
        "LogID LONG NOT NULL, " & _
        "StepNumber LONG NOT NULL, " & _
        "StepText LONGTEXT NOT NULL, " & _
        "StepType TEXT(50))"

    CreateTableIfMissing "tblTrainingStats", _
        "CREATE TABLE tblTrainingStats (" & _
        "TrainingID AUTOINCREMENT CONSTRAINT pk_tblTrainingStats PRIMARY KEY, " & _
        "PeriodID LONG NOT NULL, " & _
        "LogID LONG, " & _
        "TrainingTopic TEXT(150), " & _
        "MembersTrained LONG, " & _
        "HoursTrained DOUBLE)"

    CreateTableIfMissing "tblSystemStatus", _
        "CREATE TABLE tblSystemStatus (" & _
        "StatusID AUTOINCREMENT CONSTRAINT pk_tblSystemStatus PRIMARY KEY, " & _
        "PeriodID LONG NOT NULL, " & _
        "MissionsSupported LONG, " & _
        "MissionHoursSupported DOUBLE, " & _
        "TotalSystemDownTimeMinutes LONG, " & _
        "SystemUpTimePercent DOUBLE, " & _
        "Notes LONGTEXT)"

    CreateTableIfMissing "tblPIOSupport", _
        "CREATE TABLE tblPIOSupport (" & _
        "PIOID AUTOINCREMENT CONSTRAINT pk_tblPIOSupport PRIMARY KEY, " & _
        "PeriodID LONG NOT NULL, " & _
        "Direction TEXT(25) NOT NULL, " & _
        "HardwareCount LONG, " & _
        "SoftwareCount LONG, " & _
        "EngineeringCount LONG, " & _
        "Notes LONGTEXT)"

    CreateTableIfMissing "tblAuditTrail", _
        "CREATE TABLE tblAuditTrail (" & _
        "AuditID AUTOINCREMENT CONSTRAINT pk_tblAuditTrail PRIMARY KEY, " & _
        "TableName TEXT(100) NOT NULL, " & _
        "RecordID LONG, " & _
        "ActionName TEXT(50) NOT NULL, " & _
        "ActionBy TEXT(100), " & _
        "ActionDate DATETIME, " & _
        "Details LONGTEXT)"

    CreateIndexes
    InsertStarterData

    MsgBox "MSL tables were created or verified successfully.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BuildMslTables failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub CreateTableIfMissing(ByVal tableName As String, ByVal createSql As String)
    If Not TableExists(tableName) Then
        CurrentDb.Execute createSql, dbFailOnError
    End If
End Sub

Private Function TableExists(ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef

    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf

    TableExists = False
End Function

Private Sub CreateIndexes()
    On Error Resume Next

    CurrentDb.Execute "CREATE INDEX ix_tblPersonnel_SiteID ON tblPersonnel (SiteID)", dbFailOnError
    CurrentDb.Execute "CREATE INDEX ix_tblReportPeriods_SiteMonth ON tblReportPeriods (SiteID, ReportMonth)", dbFailOnError
    CurrentDb.Execute "CREATE INDEX ix_tblLogEntries_PeriodID ON tblLogEntries (PeriodID)", dbFailOnError
    CurrentDb.Execute "CREATE INDEX ix_tblLogEntries_Category ON tblLogEntries (Category)", dbFailOnError
    CurrentDb.Execute "CREATE INDEX ix_tblLogEntrySteps_LogID ON tblLogEntrySteps (LogID)", dbFailOnError
    CurrentDb.Execute "CREATE INDEX ix_tblTrainingStats_PeriodID ON tblTrainingStats (PeriodID)", dbFailOnError
    CurrentDb.Execute "CREATE INDEX ix_tblSystemStatus_PeriodID ON tblSystemStatus (PeriodID)", dbFailOnError
    CurrentDb.Execute "CREATE INDEX ix_tblPIOSupport_PeriodID ON tblPIOSupport (PeriodID)", dbFailOnError
End Sub

Private Sub InsertStarterData()
    On Error Resume Next

    If DCount("*", "tblSites") = 0 Then
        CurrentDb.Execute "INSERT INTO tblSites (SiteName, ContractNumber, CDRLNumber, Customer, Active) " & _
                          "VALUES ('Texas', 'FA4890-23-F-0026', 'A001', 'TENAX Technologies', True)", dbFailOnError
    End If
End Sub
