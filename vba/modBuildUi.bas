Attribute VB_Name = "modBuildUi"
Option Compare Database
Option Explicit

' Builds the first working Access UI so you do not have to hand-place every field.
' Run from the Immediate Window:
'   BuildMslApp

Public Sub BuildMslApp()
    On Error GoTo ErrHandler

    BuildMslTables
    InsertMslStarterRecords
    BuildMslQueries
    BuildFrmLogEntry
    BuildFrmLogEntrySteps
    BuildFrmSystemStatus
    BuildFrmPIOSupport
    BuildFrmDashboard

    MsgBox "MSL app shell built. Open frmDashboard to start.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BuildMslApp failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub InsertMslStarterRecords()
    On Error GoTo ErrHandler

    Dim siteId As Long
    Dim periodId As Long
    Dim aaronId As Long

    If DCount("*", "tblSites", "SiteName='Texas'") = 0 Then
        CurrentDb.Execute "INSERT INTO tblSites (SiteName, ContractNumber, CDRLNumber, Customer, Active) VALUES ('Texas', 'FA4890-23-F-0026', 'A001', 'TENAX Technologies', True)", dbFailOnError
    End If

    siteId = DLookup("SiteID", "tblSites", "SiteName='Texas'")

    AddPersonnelIfMissing siteId, "Aaron", "Kotz", "Site Lead", #7/1/2023#
    AddPersonnelIfMissing siteId, "Dale", "Gordon", "CSO", #7/1/2023#
    AddPersonnelIfMissing siteId, "Josh", "Williams", "CSO", #7/1/2023#
    AddPersonnelIfMissing siteId, "David", "Roberts", "CTS", #7/1/2023#
    AddPersonnelIfMissing siteId, "Collin", "Holladay", "CTS", #8/1/2024#

    aaronId = DLookup("PersonnelID", "tblPersonnel", "FirstName='Aaron' AND LastName='Kotz'")

    If DCount("*", "tblReportPeriods", "SiteID=" & siteId & " AND ReportMonth=#11/1/2025#") = 0 Then
        CurrentDb.Execute "INSERT INTO tblReportPeriods (SiteID, ReportMonth, ReportTitle, PreparedBy, ReportStatus, CreatedDate) VALUES (" & siteId & ", #11/1/2025#, 'Texas - November 2025', " & aaronId & ", 'Draft', Now())", dbFailOnError
    End If

    periodId = DLookup("PeriodID", "tblReportPeriods", "SiteID=" & siteId & " AND ReportMonth=#11/1/2025#")

    If DCount("*", "tblSystemStatus", "PeriodID=" & periodId) = 0 Then
        CurrentDb.Execute "INSERT INTO tblSystemStatus (PeriodID, MissionsSupported, MissionHoursSupported, TotalSystemDownTimeMinutes, SystemUpTimePercent, Notes) VALUES (" & periodId & ", 26, 483.6, 0, 100, '')", dbFailOnError
    End If

    If DCount("*", "tblPIOSupport", "PeriodID=" & periodId & " AND Direction='Requested'") = 0 Then
        CurrentDb.Execute "INSERT INTO tblPIOSupport (PeriodID, Direction, HardwareCount, SoftwareCount, EngineeringCount, Notes) VALUES (" & periodId & ", 'Requested', 0, 0, 0, 'No support was requested during this period.')", dbFailOnError
    End If

    If DCount("*", "tblPIOSupport", "PeriodID=" & periodId & " AND Direction='Received'") = 0 Then
        CurrentDb.Execute "INSERT INTO tblPIOSupport (PeriodID, Direction, HardwareCount, SoftwareCount, EngineeringCount, Notes) VALUES (" & periodId & ", 'Received', 0, 0, 0, 'No support was received during this period.')", dbFailOnError
    End If

    Exit Sub

ErrHandler:
    MsgBox "InsertMslStarterRecords failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub AddPersonnelIfMissing(ByVal siteId As Long, ByVal firstName As String, ByVal lastName As String, ByVal roleName As String, ByVal arrivalDate As Date)
    Dim whereClause As String

    whereClause = "SiteID=" & siteId & " AND FirstName='" & SqlSafe(firstName) & "' AND LastName='" & SqlSafe(lastName) & "'"

    If DCount("*", "tblPersonnel", whereClause) = 0 Then
        CurrentDb.Execute "INSERT INTO tblPersonnel (SiteID, FirstName, LastName, Role, ArrivalOnStation, Active) VALUES (" & _
                          siteId & ", '" & SqlSafe(firstName) & "', '" & SqlSafe(lastName) & "', '" & SqlSafe(roleName) & "', #" & Format(arrivalDate, "m/d/yyyy") & "#, True)", dbFailOnError
    End If
End Sub

Public Sub BuildMslQueries()
    SaveQuery "qryReportPeriodHeader", _
        "SELECT rp.PeriodID, rp.SiteID, s.SiteName, s.ContractNumber, s.CDRLNumber, s.Customer, rp.ReportMonth, rp.ReportTitle, rp.PreparedBy, rp.ReportStatus " & _
        "FROM tblSites AS s INNER JOIN tblReportPeriods AS rp ON s.SiteID = rp.SiteID;"

    SaveQuery "qryPersonnelDisplay", _
        "SELECT PersonnelID, LastName & ', ' & FirstName & ' - ' & Role AS DisplayName FROM tblPersonnel WHERE Active=True ORDER BY LastName, FirstName;"

    SaveQuery "qryPeriodDisplay", _
        "SELECT PeriodID, ReportTitle FROM tblReportPeriods ORDER BY ReportMonth DESC;"

    SaveQuery "qrySiteDisplay", _
        "SELECT SiteID, SiteName FROM tblSites WHERE Active=True ORDER BY SiteName;"

    SaveQuery "qryLogEntriesDisplay", _
        "SELECT le.LogID, rp.ReportTitle, s.SiteName, le.LogDate, le.Category, le.SystemArea, le.Title, le.Description, le.Resolution, p.LastName & ', ' & p.FirstName AS LoggedByName, le.EventCount, le.ResponseTimeMinutes, le.ReturnToServiceMinutes, le.FollowUpRequired " & _
        "FROM ((tblLogEntries AS le LEFT JOIN tblReportPeriods AS rp ON le.PeriodID = rp.PeriodID) LEFT JOIN tblSites AS s ON le.SiteID = s.SiteID) LEFT JOIN tblPersonnel AS p ON le.LoggedBy = p.PersonnelID " & _
        "ORDER BY le.LogDate DESC, le.LogID DESC;"
End Sub

Private Sub SaveQuery(ByVal queryName As String, ByVal sqlText As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete queryName
    On Error GoTo 0
    CurrentDb.CreateQueryDef queryName, sqlText
End Sub

Public Sub BuildFrmLogEntry()
    Dim frm As Form
    Dim formName As String
    Dim topPos As Long

    formName = "frmLogEntry"
    DeleteFormIfExists formName

    Set frm = CreateForm
    frm.Caption = "MSL Log Entry"
    frm.RecordSource = "tblLogEntries"
    frm.Width = 11000

    AddTitle frm, "MSL Log Entry", 300, 200

    topPos = 800
    AddCombo frm, "cboPeriodID", "Period", "PeriodID", "SELECT PeriodID, ReportTitle FROM tblReportPeriods ORDER BY ReportMonth DESC;", 300, topPos, 2400, 4500, 2, "0in;3in"
    topPos = topPos + 420
    AddCombo frm, "cboSiteID", "Site", "SiteID", "SELECT SiteID, SiteName FROM tblSites WHERE Active=True ORDER BY SiteName;", 300, topPos, 2400, 3000, 2, "0in;2in"
    topPos = topPos + 420
    AddText frm, "txtLogDate", "Log Date", "LogDate", 300, topPos, 2400, 2200
    topPos = topPos + 420
    AddCombo frm, "cboCategory", "Category", "Category", "Administration;Troubleshooting;Training;PIO Requested;PIO Received;Site Lead Note", 300, topPos, 2400, 3000, 1, "2.5in", "Value List"
    topPos = topPos + 420
    AddText frm, "txtSystemArea", "System Area", "SystemArea", 300, topPos, 2400, 4500
    topPos = topPos + 420
    AddText frm, "txtTitle", "Title", "Title", 300, topPos, 2400, 6500
    topPos = topPos + 420
    AddLongText frm, "txtDescription", "Description", "Description", 300, topPos, 2400, 6500, 1000
    topPos = topPos + 1120
    AddLongText frm, "txtResolution", "Resolution / Result", "Resolution", 300, topPos, 2400, 6500, 800
    topPos = topPos + 920
    AddCombo frm, "cboLoggedBy", "Logged By", "LoggedBy", "SELECT PersonnelID, LastName & ', ' & FirstName & ' - ' & Role AS DisplayName FROM tblPersonnel WHERE Active=True ORDER BY LastName, FirstName;", 300, topPos, 2400, 4500, 2, "0in;3in"
    topPos = topPos + 420
    AddText frm, "txtEventCount", "Event Count", "EventCount", 300, topPos, 2400, 1500
    topPos = topPos + 420
    AddText frm, "txtResponseTimeMinutes", "Response Time Minutes", "ResponseTimeMinutes", 300, topPos, 2400, 1500
    topPos = topPos + 420
    AddText frm, "txtReturnToServiceMinutes", "Return To Service Minutes", "ReturnToServiceMinutes", 300, topPos, 2400, 1500
    topPos = topPos + 420
    AddCheck frm, "chkFollowUpRequired", "Follow Up Required", "FollowUpRequired", 300, topPos, 2400

    AddHiddenBoundText frm, "txtCreatedDate", "CreatedDate"
    AddHiddenBoundText frm, "txtModifiedDate", "ModifiedDate"

    AddButton frm, "btnSave", "Save", 300, topPos + 700, "=SaveCurrentRecord()"
    AddButton frm, "btnNew", "New Entry", 1600, topPos + 700, "=GoNewRecord()"
    AddButton frm, "btnClose", "Close", 3100, topPos + 700, "=CloseCurrentForm()"

    DoCmd.Save acForm, frm.Name
    DoCmd.Rename formName, acForm, frm.Name
    DoCmd.Close acForm, formName, acSaveYes
End Sub

Public Sub BuildFrmLogEntrySteps()
    Dim frm As Form
    Dim formName As String
    Dim topPos As Long

    formName = "frmLogEntrySteps"
    DeleteFormIfExists formName

    Set frm = CreateForm
    frm.Caption = "Log Entry Steps"
    frm.RecordSource = "tblLogEntrySteps"
    frm.DefaultView = 1
    frm.Width = 9000

    topPos = 400
    AddText frm, "txtLogID", "Log ID", "LogID", 300, topPos, 1600, 1200
    topPos = topPos + 420
    AddText frm, "txtStepNumber", "Step Number", "StepNumber", 300, topPos, 1600, 1200
    topPos = topPos + 420
    AddLongText frm, "txtStepText", "Step Text", "StepText", 300, topPos, 1600, 6000, 800
    topPos = topPos + 920
    AddCombo frm, "cboStepType", "Step Type", "StepType", "Action;Result;Note", 300, topPos, 1600, 2200, 1, "2in", "Value List"

    DoCmd.Save acForm, frm.Name
    DoCmd.Rename formName, acForm, frm.Name
    DoCmd.Close acForm, formName, acSaveYes
End Sub

Public Sub BuildFrmSystemStatus()
    Dim frm As Form
    Dim formName As String
    Dim topPos As Long

    formName = "frmSystemStatus"
    DeleteFormIfExists formName

    Set frm = CreateForm
    frm.Caption = "Monthly System Status"
    frm.RecordSource = "tblSystemStatus"
    frm.Width = 9000

    AddTitle frm, "Monthly System Status", 300, 200
    topPos = 800
    AddCombo frm, "cboPeriodID", "Period", "PeriodID", "SELECT PeriodID, ReportTitle FROM tblReportPeriods ORDER BY ReportMonth DESC;", 300, topPos, 2600, 4500, 2, "0in;3in"
    topPos = topPos + 420
    AddText frm, "txtMissionsSupported", "Missions Supported", "MissionsSupported", 300, topPos, 2600, 1800
    topPos = topPos + 420
    AddText frm, "txtMissionHoursSupported", "Mission Hours Supported", "MissionHoursSupported", 300, topPos, 2600, 1800
    topPos = topPos + 420
    AddText frm, "txtTotalSystemDownTimeMinutes", "Down Time Minutes", "TotalSystemDownTimeMinutes", 300, topPos, 2600, 1800
    topPos = topPos + 420
    AddText frm, "txtSystemUpTimePercent", "System Up Time Percent", "SystemUpTimePercent", 300, topPos, 2600, 1800
    topPos = topPos + 420
    AddLongText frm, "txtNotes", "Notes", "Notes", 300, topPos, 2600, 5200, 1000

    AddButton frm, "btnSave", "Save", 300, topPos + 1250, "=SaveCurrentRecord()"
    AddButton frm, "btnClose", "Close", 1600, topPos + 1250, "=CloseCurrentForm()"

    DoCmd.Save acForm, frm.Name
    DoCmd.Rename formName, acForm, frm.Name
    DoCmd.Close acForm, formName, acSaveYes
End Sub

Public Sub BuildFrmPIOSupport()
    Dim frm As Form
    Dim formName As String
    Dim topPos As Long

    formName = "frmPIOSupport"
    DeleteFormIfExists formName

    Set frm = CreateForm
    frm.Caption = "Program Office Support"
    frm.RecordSource = "tblPIOSupport"
    frm.Width = 9000

    AddTitle frm, "Program Office Support", 300, 200
    topPos = 800
    AddCombo frm, "cboPeriodID", "Period", "PeriodID", "SELECT PeriodID, ReportTitle FROM tblReportPeriods ORDER BY ReportMonth DESC;", 300, topPos, 2400, 4500, 2, "0in;3in"
    topPos = topPos + 420
    AddCombo frm, "cboDirection", "Direction", "Direction", "Requested;Received", 300, topPos, 2400, 2200, 1, "2in", "Value List"
    topPos = topPos + 420
    AddText frm, "txtHardwareCount", "Hardware Count", "HardwareCount", 300, topPos, 2400, 1500
    topPos = topPos + 420
    AddText frm, "txtSoftwareCount", "Software Count", "SoftwareCount", 300, topPos, 2400, 1500
    topPos = topPos + 420
    AddText frm, "txtEngineeringCount", "Engineering Count", "EngineeringCount", 300, topPos, 2400, 1500
    topPos = topPos + 420
    AddLongText frm, "txtNotes", "Notes", "Notes", 300, topPos, 2400, 5200, 1000

    AddButton frm, "btnSave", "Save", 300, topPos + 1250, "=SaveCurrentRecord()"
    AddButton frm, "btnClose", "Close", 1600, topPos + 1250, "=CloseCurrentForm()"

    DoCmd.Save acForm, frm.Name
    DoCmd.Rename formName, acForm, frm.Name
    DoCmd.Close acForm, formName, acSaveYes
End Sub

Public Sub BuildFrmDashboard()
    Dim frm As Form
    Dim formName As String
    Dim topPos As Long

    formName = "frmDashboard"
    DeleteFormIfExists formName

    Set frm = CreateForm
    frm.Caption = "MSL Dashboard"
    frm.Width = 9000

    AddTitle frm, "MSL Logbook Dashboard", 300, 250

    topPos = 1000
    AddButton frm, "btnLogEntry", "New / Edit Log Entry", 600, topPos, "=OpenMslForm('frmLogEntry')", 2600
    topPos = topPos + 520
    AddButton frm, "btnSystemStatus", "System Status", 600, topPos, "=OpenMslForm('frmSystemStatus')", 2600
    topPos = topPos + 520
    AddButton frm, "btnPIOSupport", "Program Office Support", 600, topPos, "=OpenMslForm('frmPIOSupport')", 2600
    topPos = topPos + 520
    AddButton frm, "btnViewLogs", "View Log Table", 600, topPos, "=OpenMslQuery('qryLogEntriesDisplay')", 2600
    topPos = topPos + 520
    AddButton frm, "btnClose", "Close", 600, topPos, "=CloseCurrentForm()", 2600

    DoCmd.Save acForm, frm.Name
    DoCmd.Rename formName, acForm, frm.Name
    DoCmd.Close acForm, formName, acSaveYes
End Sub

Private Function AddTitle(ByVal frm As Form, ByVal captionText As String, ByVal leftPos As Long, ByVal topPos As Long) As Control
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , leftPos, topPos, 7000, 360)
    ctl.Name = "lblTitle"
    ctl.Caption = captionText
    ctl.FontSize = 18
    ctl.FontBold = True
    Set AddTitle = ctl
End Function

Private Function AddLabel(ByVal frm As Form, ByVal labelText As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthVal As Long) As Control
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , leftPos, topPos, widthVal, 260)
    ctl.Caption = labelText
    ctl.TextAlign = 3
    Set AddLabel = ctl
End Function

Private Function AddText(ByVal frm As Form, ByVal controlName As String, ByVal labelText As String, ByVal controlSource As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal labelWidth As Long, ByVal controlWidth As Long) As Control
    Dim ctl As Control
    AddLabel frm, labelText, leftPos, topPos + 30, labelWidth
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , controlSource, leftPos + labelWidth + 200, topPos, controlWidth, 300)
    ctl.Name = controlName
    Set AddText = ctl
End Function

Private Function AddLongText(ByVal frm As Form, ByVal controlName As String, ByVal labelText As String, ByVal controlSource As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal labelWidth As Long, ByVal controlWidth As Long, ByVal controlHeight As Long) As Control
    Dim ctl As Control
    AddLabel frm, labelText, leftPos, topPos + 30, labelWidth
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , controlSource, leftPos + labelWidth + 200, topPos, controlWidth, controlHeight)
    ctl.Name = controlName
    ctl.EnterKeyBehavior = True
    ctl.ScrollBars = 2
    Set AddLongText = ctl
End Function

Private Function AddCombo(ByVal frm As Form, ByVal controlName As String, ByVal labelText As String, ByVal controlSource As String, ByVal rowSource As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal labelWidth As Long, ByVal controlWidth As Long, ByVal columnCount As Integer, ByVal columnWidths As String, Optional ByVal rowSourceType As String = "Table/Query") As Control
    Dim ctl As Control
    AddLabel frm, labelText, leftPos, topPos + 30, labelWidth
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , controlSource, leftPos + labelWidth + 200, topPos, controlWidth, 300)
    ctl.Name = controlName
    ctl.RowSourceType = rowSourceType
    ctl.RowSource = rowSource
    ctl.BoundColumn = 1
    ctl.ColumnCount = columnCount
    ctl.ColumnWidths = columnWidths
    ctl.LimitToList = True
    Set AddCombo = ctl
End Function

Private Function AddCheck(ByVal frm As Form, ByVal controlName As String, ByVal labelText As String, ByVal controlSource As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal labelWidth As Long) As Control
    Dim ctl As Control
    AddLabel frm, labelText, leftPos, topPos + 30, labelWidth
    Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , controlSource, leftPos + labelWidth + 200, topPos, 300, 300)
    ctl.Name = controlName
    Set AddCheck = ctl
End Function

Private Function AddButton(ByVal frm As Form, ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal onClickExpression As String, Optional ByVal controlWidth As Long = 1100) As Control
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , leftPos, topPos, controlWidth, 360)
    ctl.Name = controlName
    ctl.Caption = captionText
    ctl.OnClick = onClickExpression
    Set AddButton = ctl
End Function

Private Sub AddHiddenBoundText(ByVal frm As Form, ByVal controlName As String, ByVal controlSource As String)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , controlSource, 0, 0, 100, 100)
    ctl.Name = controlName
    ctl.Visible = False
End Sub

Private Sub DeleteFormIfExists(ByVal formName As String)
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveNo
    DoCmd.DeleteObject acForm, formName
    On Error GoTo 0
End Sub

Private Function SqlSafe(ByVal value As String) As String
    SqlSafe = Replace(value, "'", "''")
End Function

Public Function SaveCurrentRecord() As Boolean
    On Error GoTo ErrHandler

    If Screen.ActiveForm.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
    End If

    SaveCurrentRecord = True
    MsgBox "Saved.", vbInformation
    Exit Function

ErrHandler:
    MsgBox "Save failed: " & Err.Description, vbExclamation
    SaveCurrentRecord = False
End Function

Public Function GoNewRecord() As Boolean
    On Error GoTo ErrHandler
    DoCmd.GoToRecord , , acNewRec
    GoNewRecord = True
    Exit Function

ErrHandler:
    MsgBox "Unable to go to a new record: " & Err.Description, vbExclamation
    GoNewRecord = False
End Function

Public Function CloseCurrentForm() As Boolean
    On Error Resume Next
    DoCmd.Close acForm, Screen.ActiveForm.Name
    CloseCurrentForm = True
End Function

Public Function OpenMslForm(ByVal formName As String) As Boolean
    On Error GoTo ErrHandler
    DoCmd.OpenForm formName
    OpenMslForm = True
    Exit Function

ErrHandler:
    MsgBox "Unable to open " & formName & ": " & Err.Description, vbExclamation
    OpenMslForm = False
End Function

Public Function OpenMslQuery(ByVal queryName As String) As Boolean
    On Error GoTo ErrHandler
    DoCmd.OpenQuery queryName
    OpenMslQuery = True
    Exit Function

ErrHandler:
    MsgBox "Unable to open " & queryName & ": " & Err.Description, vbExclamation
    OpenMslQuery = False
End Function
