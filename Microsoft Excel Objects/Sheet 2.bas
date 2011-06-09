Private Sub Worksheet_Change(ByVal Target As Excel.Range)
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
    Dim flag As Boolean
    
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
    
End Sub



