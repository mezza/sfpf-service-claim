VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthForm 
   Caption         =   "SFP Service Claim"
   ClientHeight    =   2000
   ClientLeft      =   2000
   ClientTop       =   -5920
   ClientWidth     =   7000
   OleObjectBlob   =   "MonthForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MonthForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    MonthForm.Hide
    Unload MonthForm
End Sub

Private Sub ProceedButton_Click()
    Range("README!F5").Value = MonthCombo.Value
    If MonthCombo.Value = "" Then
        MsgBox "You must select a month!  Please try again"
        Exit Sub
    Else
        MonthForm.Hide
        Unload MonthForm
        Call Parameters.URL_Get_Query
    End If
End Sub

Private Sub UserForm_Activate()
On Error GoTo errortrap
    'Setup array for months that can be selected
    Dim myarray(0, 4) As String
    myarray(0, 0) = Worksheets("README").Range("E27").Value
    myarray(0, 1) = Worksheets("README").Range("E28").Value
    myarray(0, 2) = Worksheets("README").Range("E29").Value
    myarray(0, 3) = Worksheets("README").Range("E30").Value
    myarray(0, 4) = Worksheets("README").Range("E31").Value
    MonthCombo.Column() = myarray
    'Set default month selection to previous month
    MonthCombo.Value = Range("README!F7").Value
    Exit Sub

errortrap:

    MsgBox "Error is: " & Err

End Sub

Private Sub UserForm_Terminate()
    MonthForm.Hide
    Unload MonthForm
End Sub
