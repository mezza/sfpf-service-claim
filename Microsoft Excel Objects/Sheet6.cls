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
    
'SoftCalc 28/06/13 switched off by to allow form entry
GoTo Skip

    Dim flag As Boolean
    Dim PS As Positions
    
    ' Fetch choices for TOR column
    If Target.Column = S_TOR_INDEX And Target.row <> 1 Then
      flag = GetTORs(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Project column
    If Target.Column = S_PROJECT_INDEX And Target.row <> 1 Then
      flag = GetProjects(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Task column
    If Target.Column = S_TASK_INDEX And Target.row <> 1 Then
      flag = GetTask(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Grant code column
    If Target.Column = S_GRANT_CODE_INDEX And Target.row <> 1 Then
      flag = GetGrantCode(Target.Worksheet.Name, Target.row)
    End If

Skip:

Application.ScreenUpdating = False

    'If ActiveCell.Column > 3 Then
    '    ServicesABCForm.Hide
    '    End
    'End If
    
    If ActiveCell.row = 1 Then Exit Sub
    
    If ActiveCell.Column = 1 Or ActiveCell.Column = 2 Or ActiveCell.Column = 3 Then
        'Unload ServicesABCForm
        PS = PositionForm(ServicesABCForm, ActiveCell, 0, 0, cstFhpFormLeftCellRight, cstFvpFormTopCellTop)
        ServicesABCForm.Top = PS.FrmTop
        ServicesABCForm.Left = PS.FrmLeft
        ServicesABCForm.Show Modal
    End If


End Sub

    
