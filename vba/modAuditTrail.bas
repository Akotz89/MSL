Attribute VB_Name = "modAuditTrail"
Option Compare Database
Option Explicit

Public Function GetCurrentUserName() As String
    On Error Resume Next
    GetCurrentUserName = Environ$("USERNAME")
    If Len(Trim(GetCurrentUserName)) = 0 Then
        GetCurrentUserName = CurrentUser()
    End If
End Function

Public Sub WriteAudit(ByVal tableName As String, ByVal recordId As Long, ByVal actionName As String, Optional ByVal details As String = "")
    On Error GoTo ErrHandler

    Dim sqlText As String

    sqlText = "INSERT INTO tblAuditTrail (TableName, RecordID, ActionName, ActionBy, ActionDate, Details) " & _
              "VALUES (" & _
              "'" & Replace(tableName, "'", "''") & "', " & _
              recordId & ", " & _
              "'" & Replace(actionName, "'", "''") & "', " & _
              "'" & Replace(GetCurrentUserName(), "'", "''") & "', " & _
              "Now(), " & _
              "'" & Replace(details, "'", "''") & "')"

    CurrentDb.Execute sqlText, dbFailOnError
    Exit Sub

ErrHandler:
    ' Do not block user work because audit logging failed.
End Sub
