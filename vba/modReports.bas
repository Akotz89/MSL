Attribute VB_Name = "modReports"
Option Compare Database
Option Explicit

Public Function CalculateSystemUptime(ByVal missionHours As Double, ByVal downTimeMinutes As Long) As Double
    On Error GoTo ErrHandler

    Dim totalMinutes As Double

    totalMinutes = missionHours * 60

    If totalMinutes <= 0 Then
        CalculateSystemUptime = 100
    Else
        CalculateSystemUptime = Round(((totalMinutes - downTimeMinutes) / totalMinutes) * 100, 2)
    End If

    Exit Function

ErrHandler:
    CalculateSystemUptime = 0
End Function

Public Function BuildMonthlyReportFileName(ByVal siteName As String, ByVal reportMonth As Date) As String
    Dim safeSiteName As String

    safeSiteName = Replace(siteName, " ", "_")
    safeSiteName = Replace(safeSiteName, "/", "-")
    safeSiteName = Replace(safeSiteName, "\", "-")
    safeSiteName = Replace(safeSiteName, ":", "-")

    BuildMonthlyReportFileName = safeSiteName & "_" & Format(reportMonth, "yyyy_mm") & "_Monthly_Report.pdf"
End Function

Public Sub PreviewMonthlyReport(ByVal periodId As Long)
    On Error GoTo ErrHandler

    DoCmd.OpenReport "rptMonthlySiteLog", acViewPreview, , "PeriodID = " & periodId
    Exit Sub

ErrHandler:
    MsgBox "Unable to preview the monthly report: " & Err.Description, vbExclamation
End Sub

Public Sub ExportMonthlyReport(ByVal periodId As Long, ByVal outputFolder As String)
    On Error GoTo ErrHandler

    Dim siteName As String
    Dim reportMonth As Date
    Dim filePath As String

    siteName = Nz(DLookup("SiteName", "qryReportPeriodHeader", "PeriodID = " & periodId), "Site")
    reportMonth = Nz(DLookup("ReportMonth", "qryReportPeriodHeader", "PeriodID = " & periodId), Date)

    If Right(outputFolder, 1) <> "\" Then
        outputFolder = outputFolder & "\"
    End If

    filePath = outputFolder & BuildMonthlyReportFileName(siteName, reportMonth)

    DoCmd.OpenReport "rptMonthlySiteLog", acViewPreview, , "PeriodID = " & periodId
    DoCmd.OutputTo acOutputReport, "rptMonthlySiteLog", acFormatPDF, filePath, False
    DoCmd.Close acReport, "rptMonthlySiteLog"

    MsgBox "Monthly report exported:" & vbCrLf & filePath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Unable to export the monthly report: " & Err.Description, vbExclamation
End Sub
