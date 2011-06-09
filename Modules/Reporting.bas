' Sub RemoveValidation()
'     Worksheets("Services").Range(S_PROJECT & 2 & ":" & S_PROJECT & 10).Select
'     With Selection.Validation
'         .Delete
'     End With
' End Sub

Sub PrepareReport()
  Dim f As Long
  f = ValidateSheets
  If f > 0 Then
    MsgBox "There are " & f & " errors. Please check the Red cells."
  Else
    CompileReport
  End If 
End Sub

Private Sub CompileReport()
  Application.ScreenUpdating = FALSE
  Sheets.Add.Name = "Reports2"
  Sheets("Reports2").Move after:=Worksheets("Report")
  
  'Set header row and formats
  With Sheets("Reports2")
    Range("A1").Value = "TORs"
    Range("B1").Value = "Task"
    Range("C1").Value = "Collated submissions"
    Range("D1").Value = "Report"
    Range("E1").Value = "TORTASKID"
    Range("A1:E1").Select
    Selection.Font.Bold = True
    Columns("A:D").Select
    Selection.ColumnWidth = 35.0
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    Columns("E").Select
    Selection.ColumnWidth = 7
  End With
  
  Worksheets("Parameters").Activate
  Dim torItem As Range
  Dim report_row As Long
  Dim msg As String
  Dim rep As String
  Dim res_count As Double
  Dim search_criteria As Range
  
  On Error GoTo myError

  report_row = 2
  For Each torItem In Worksheets("Parameters").Range(P_TORTASKIDs & 2, Range(P_TORTASKIDs & 2).End(xlDown))
    msg = ""
    rep = ""
    ' Are there matches in Services? Set msg
    Worksheets("Services").Activate
    With Worksheets("Services").Range(S_TORTASKID & 2, Range(S_TORTASKID & 2).End(xlDown))
      Set search_result = .Find(torItem.Value, lookin:=xlValues)
      If Not search_result Is Nothing Then
        firstAddress = search_result.Address
        Do
          msg = msg & Worksheets("Services").Range(S_REPORT & search_result.row).Value & Chr(10)
          Set search_result = .FindNext(search_result)
        Loop While Not search_result Is Nothing And search_result.Address <> firstAddress
      End If
    End With
    
    ' Are there matches in Expenses? Set msg
    Worksheets("Expenses").Activate
    With Worksheets("Expenses").Range(E_TORTASKID & 2, Range(E_TORTASKID & 2).End(xlDown))
      Set search_result = .Find(torItem.Value, lookin:=xlValues)
      If Not search_result Is Nothing Then
        firstAddress = search_result.Address
        Do
          msg = msg & Worksheets("Expenses").Range(E_DESCRIPTION & search_result.row).Value & Chr(10)
          Set search_result = .FindNext(search_result)
        Loop While Not search_result Is Nothing And search_result.Address <> firstAddress
      End If
    End With
    
    ' Are there matches in Reports? Set rep
    Worksheets("Report").Activate
    With Worksheets("Report").Range(R_TORTASKID & 2, Range(R_TORTASKID & 2).End(xlDown))
      Set search_result = .Find(torItem.Value, lookin:=xlValues)
      If Not search_result Is Nothing Then
        firstAddress = search_result.Address
        Do
          rep = rep & Worksheets("Report").Range(R_EDITED_REPORT & search_result.row).Value & Chr(10)
          Set search_result = .FindNext(search_result)
        Loop While Not search_result Is Nothing And search_result.Address <> firstAddress
      End If
    End With

    
    ' Update the Report lines
    If msg <> "" Then
        Worksheets("Reports2").Range(R_TOR & report_row) = Worksheets("Parameters").Range(P_TORs2 & torItem.row)
        Worksheets("Reports2").Range(R_TASK & report_row) = Worksheets("Parameters").Range(P_TORs2_TASKS & torItem.row)
        Worksheets("Reports2").Range(R_COLLATED_SUBMISSIONS & report_row) = msg
        If rep <> "" Then
          Worksheets("Reports2").Range(R_EDITED_REPORT & report_row) = rep
        End If
        Worksheets("Reports2").Range(R_TORTASKID & report_row) = torItem.Value
        report_row = report_row + 1
    End If
  Next torItem
  f = DeleteSheet("Report")
  Worksheets(Worksheets("Reports2").Index).Name = "Report"
  Worksheets("Report").Select
  Worksheets("Report").Range("E2", Range("E2").End(xlDown)).Interior.ColorIndex = 5
  MsgBox "Your report is ready for entry. Please enter an edited report in all cells in Column D and upload to Planet."
myError:
  If Err.Number <> 0 Then
    f = DeleteSheet("Reports2")
    MsgBox "There was a problem generating the Report. Please contact support."
  End If
  Application.ScreenUpdating = True
End Sub
