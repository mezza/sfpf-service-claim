VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
'SoftCalc 28-06-13 left functional for Expenses sheet

  Dim f As Boolean
  Target.Interior.ColorIndex = 0 'No Fill
  Select Case Target.Column
    ' If user selects a Task then we need to set TORTASKID and reset GRANT
    Case S_TASK_INDEX
      f = SetTorTaskId(Target.Worksheet.Name, Target.row)
    ' If user enters a Date ensure it's valid
    Case S_DATE_INDEX
      f = IsValidDate(Target)
    ' If user selects a Grant then we need to set GRANTID
    Case S_GRANT_CODE_INDEX
      f = SetGrantCodeID(Target.Worksheet.Name, Target.row)
    ' If user enters a Report then ensure it's meaningful
    Case S_REPORT_INDEX
      f = IsReportOk(Target)
    ' Try to calculate the hours worked if needed
    Case S_START_TIME_INDEX
      f = CalculateHours(Target)
    Case S_END_TIME_INDEX
      f = CalculateHours(Target)
    Case S_HOURS_INDEX
      f = CheckHours(Target)
  End Select

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    Dim PS As Positions
    
    Application.ScreenUpdating = False

    If ActiveCell.row = 1 Then Exit Sub
    
    If ActiveCell.Column = S_TOR_INDEX Or ActiveCell.Column = S_PROJECT_INDEX Or ActiveCell.Column = S_TASK_INDEX Or ActiveCell.Column = S_GRANT_CODE_INDEX Then
        'Unload ServicesABCForm
        PS = PositionForm(ServicesABCForm, ActiveCell, 0, 0, cstFhpFormLeftCellRight, cstFvpFormTopCellTop)
        ServicesABCForm.Top = PS.FrmTop
        ServicesABCForm.Left = PS.FrmLeft
        ServicesABCForm.Show Modal
    End If
    
    Application.ScreenUpdating = True

End Sub

    
