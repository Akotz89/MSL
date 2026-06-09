Attribute VB_Name = "Form_frmLogEntry"
Option Compare Database
Option Explicit

Private Sub Form_Load()
    ApplyCategoryLayout
End Sub

Private Sub Form_Current()
    ApplyCategoryLayout
End Sub

Private Sub Category_AfterUpdate()
    ApplyCategoryLayout
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    If IsNull(Me!LogDate) Then Me!LogDate = Date
    If IsNull(Me!EventCount) Then Me!EventCount = 1
    If IsNull(Me!CreatedDate) Then Me!CreatedDate = Now()
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Not ValidateLogEntryForm(Me) Then
        Cancel = True
        Exit Sub
    End If

    Me!ModifiedDate = Now()
End Sub

Private Sub ApplyCategoryLayout()
    On Error Resume Next

    Dim selectedCategory As String
    selectedCategory = Nz(Me!Category, "")

    Me!Resolution.Visible = (selectedCategory = "Troubleshooting")
    Me!ResponseTimeMinutes.Visible = (selectedCategory = "Troubleshooting" Or selectedCategory = "Administration")
    Me!ReturnToServiceMinutes.Visible = (selectedCategory = "Troubleshooting" Or selectedCategory = "Administration")
    Me!FollowUpRequired.Visible = (selectedCategory = "Troubleshooting" Or selectedCategory = "Site Lead Note")
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrHandler

    If Me.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
    End If

    MsgBox "Log entry saved.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Unable to save log entry: " & Err.Description, vbExclamation
End Sub

Private Sub btnNew_Click()
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub
