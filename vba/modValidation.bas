Attribute VB_Name = "modValidation"
Option Compare Database
Option Explicit

Public Function IsBlank(ByVal value As Variant) As Boolean
    IsBlank = (Len(Trim(Nz(value, ""))) = 0)
End Function

Public Function ValidateLogEntryForm(ByVal frm As Form) As Boolean
    On Error GoTo ErrHandler

    ValidateLogEntryForm = False

    If IsBlank(frm!LogDate) Then
        MsgBox "Log Date is required.", vbExclamation
        frm!LogDate.SetFocus
        Exit Function
    End If

    If IsBlank(frm!Category) Then
        MsgBox "Category is required.", vbExclamation
        frm!Category.SetFocus
        Exit Function
    End If

    If IsBlank(frm!Title) Then
        MsgBox "Title is required.", vbExclamation
        frm!Title.SetFocus
        Exit Function
    End If

    If Nz(frm!EventCount, 0) <= 0 Then
        frm!EventCount = 1
    End If

    Select Case Nz(frm!Category, "")
        Case "Troubleshooting"
            If IsBlank(frm!Resolution) Then
                MsgBox "Troubleshooting entries should include a resolution or current status.", vbExclamation
                frm!Resolution.SetFocus
                Exit Function
            End If
        Case "Training"
            ' Training detail validation belongs on the training subform/table.
        Case "Site Lead Note"
            If IsBlank(frm!Description) Then
                MsgBox "Site lead notes require a description.", vbExclamation
                frm!Description.SetFocus
                Exit Function
            End If
    End Select

    ValidateLogEntryForm = True
    Exit Function

ErrHandler:
    MsgBox "Unable to validate this log entry: " & Err.Description, vbExclamation
End Function
