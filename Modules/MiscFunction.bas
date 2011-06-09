'Looks up Currencies from named range
Public Function GetCurrency(ByVal sheet As String, ByVal row As Long) As Boolean
  
  Application.EnableEvents = False

  Worksheets(sheet).Range(E_CURRENCY & row).Select
  With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=Currencies" 
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
  End With    
   
myError:
  If Err.Number <> 0 Then
    MsgBox "Error when setting Currencies"
  End If
  Application.EnableEvents = True
End Function

'Looks up Expense categories from named range
Public Function GetExpensesCategory(ByVal sheet As String, ByVal row As Long) As Boolean
  
  Application.EnableEvents = False

  Worksheets(sheet).Range(E_CATEGORY & row).Select
  With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ExpenseCategories" 
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
  End With    
   
myError:
  If Err.Number <> 0 Then
    MsgBox "Error when setting Expense categories"
  End If
  Application.EnableEvents = True
End Function

'Looks up TOR options from named range
Public Function GetTORs(ByVal sheet As String, ByVal row As Long) As Boolean
  
  Application.EnableEvents = False

  ' This currently assumes that both Services and Expenses sheet have the same structure TOR/Project/Task
  Worksheets(sheet).Range(S_TOR & row).Select
  With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=TORs" 
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
  End With    
   
myError:
  If Err.Number <> 0 Then
    MsgBox "Error setting TOR items"
  End If
  Application.EnableEvents = True
End Function

'Looks up Project options from named range
Public Function GetProjects(ByVal sheet As String, ByVal row As Long) As Boolean
  
  Application.EnableEvents = False
  
  ' This currently assumes that both Services and Expenses sheet have the same structure TOR/Project/Task
  Worksheets(sheet).Range(S_PROJECT & row).Select
  With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=Projects" 
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
  End With    
   
myError:
  If Err.Number <> 0 Then
    MsgBox "Error setting Project options"
  End If
  Application.EnableEvents = True
End Function

'Uses Match and CountIf to set the available Tasks
Public Function GetTask(ByVal sheet As String, ByVal row As Long) As Boolean
  Application.EnableEvents = False
    
  On Error GoTo myError
    
  Dim raw_value As String
  Dim lookup_value As String
  Dim search_range As Range
  Dim match_location_offset As Double
  Dim matches As Double
  Dim result_tor_range As String
    
  If Trim(sheet) = "Services" Then
    Col_TORs = S_TOR
    Col_Project = S_PROJECT
    Col_Task = S_TASK
  End If
  If Trim(sheet) = "Expenses" Then
    Col_TORs = E_TOR
    Col_Project = E_PROJECT
    Col_Task = E_TASK
  End If
    
  ' Need to check whether TOR or Project has been chosen
  If Trim(CStr(Worksheets(sheet).Range(Col_TORs & row).Value)) <> "" and _
    Trim(CStr(Worksheets(sheet).Range(Col_Project & row).Value)) <> "" Then
    MsgBox "Please select EITHER a TOR item or a Project. Not both"
    Application.EnableEvents = True
    Exit Function
  End If
    
  ' Check whether TOR or Project has been chosen and set lookup value and search range
  If Trim(CStr(Worksheets(sheet).Range(Col_TORs & row).Value)) <> "" Then
    raw_value = CStr(Worksheets(sheet).Range(Col_TORs & row).Value)
    Res_Task = P_TORs2_TASKS
    Set search_range = Range("TORTasks").Columns(1)
  End If
    
  If Trim(CStr(Worksheets(sheet).Range(Col_Project & row).Value)) <> "" Then
    raw_value = CStr(Worksheets(sheet).Range(Col_Project & row).Value)
    Res_Task = P_PROJECTS2_TASKS
    Set search_range = Range("ProjectTasks").Columns(1)
  End If
    
  ' No need to do anything if neither TOR or Project has been set
  If raw_value = "" Then
    Application.EnableEvents = True
    Exit Function
  End If 
    
  lookup_value = Left(raw_value, 48) & "*"

  match_location_offset = Application.WorksheetFunction.Match(lookup_value, search_range, 0)
  matches = Application.WorksheetFunction.CountIf(search_range, lookup_value)
  result_tor_range = "=Parameters!$" & Res_Task & "$" & (match_location_offset + 1) & ":$" & Res_Task & "$" & (match_location_offset + matches)
    
  Worksheets(sheet).Range(Col_Task & row).Select
    With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=result_tor_range
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
  End With    

    
myError:
  If Err.Number <> 0 Then
    MsgBox "No Tasks found for selected TOR item or Project."
  End If
  Application.EnableEvents = True
End Function

'Uses Match and CountIf to set the available Grant codes
Public Function GetGrantCode(ByVal sheet As String, ByVal row As Long) As Boolean
  Application.EnableEvents = False
    
  On Error GoTo myError
    
  Dim lookup_value As Double
  Dim search_range As Range
  Dim match_location_offset As Double
  Dim matches As Double
  Dim result_tor_range As String
    
  If Trim(sheet) = "Services" Then
    Col_TORTASKID = S_TORTASKID
    Col_GrantCode = S_GRANT_CODE
  End If
  If Trim(sheet) = "Expenses" Then
    Col_TORTASKID = E_TORTASKID
    Col_GrantCode = E_GRANTCODE
  End If
    
  lookup_value = Worksheets(sheet).Range(Col_TORTASKID & row).Value
    
  ' No need to do anything if neither TOR or Project has been set
  If NOT lookup_value > 0 Then
    Application.EnableEvents = True
    Exit Function
  End If 
    
  Set search_range = Range("NodeIDGrants").Columns(1)

  match_location_offset = Application.WorksheetFunction.Match(lookup_value, search_range, 0)
  matches = Application.WorksheetFunction.CountIf(search_range, lookup_value)
  result_tor_range = "=Parameters!$" & P_ID_GRANTS_2 & "$" & (match_location_offset + 1) & ":$" & P_ID_GRANTS_2 & "$" & (match_location_offset + matches)
    
  Worksheets(sheet).Range(Col_GrantCode & row).Select
    With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=result_tor_range
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
  End With    

    
myError:
  If Err.Number <> 0 Then
    MsgBox "No Grants found for selected Task."
  End If
  Application.EnableEvents = True
End Function

' Function to look up project_node_id for selected Task and set TORTASKID on Services and Expenses sheets
Public Function SetTorTaskId(ByVal sheet As String, ByVal row As Long) As Boolean
  Application.EnableEvents = False
            
  Dim Col_Task As String
  Dim Col_TORTASKID As String
  Dim lookup_value As String
  Dim res as Variant
        
  On Error GoTo myError
    
  If Trim(sheet) = "Services" Then
    Col_Task = S_TASK
    Col_TORTASKID = S_TORTASKID
    Col_Grant = S_GRANT_CODE
    Col_GrantID = S_GRANTCODEID
  End If
  If Trim(sheet) = "Expenses" Then
    Col_Task = E_TASK
    Col_TORTASKID = E_TORTASKID
    Col_Grant = E_GRANTCODE
    Col_GrantID = E_GRANTCODEID
  End If
    
  ' If nothing has been entered in the Task column then exit
  If Trim(CStr(Worksheets(sheet).Range(Col_Task & row).Value)) = "" Then
    Worksheets(sheet).Range(Col_TORTASKID & row).Value = ""
    Worksheets(sheet).Range(Col_Grant & row).Value = ""
    Worksheets(sheet).Range(Col_GrantID & row).Value = ""
    Application.EnableEvents = True
    Exit Function
  End If
  
  ' Otherwise let's lookup the selected Task in P_TASKS_IDs_1 and P_TASKS_IDs_2
  lookup_value = LEFT(Worksheets(sheet).Range(Col_Task & row).Value,48) & "*"
  res = Application.WorksheetFunction.VLookup(lookup_value, Range("TaskNodeIDs"), 2, False)
  Worksheets(sheet).Range(Col_TORTASKID & row).Value = res
  Worksheets(sheet).Range(Col_TORTASKID & row).Interior.ColorIndex = 4
  
  myError:
    If Err.Number <> 0 Then
      MsgBox "Error setting TORTASKID."
    End If
    
  Worksheets(sheet).Range(Col_Grant & row).Value = ""
  Worksheets(sheet).Range(Col_GrantID & row).Value = ""
  Application.EnableEvents = True
End Function

'Function to look up proposal_id for selected Grant and set GRANTCODEID on Services and Expenses sheets
Public Function SetGrantCodeID(ByVal sheet As String, ByVal row As Long) As Boolean
  Application.EnableEvents = False
            
  Dim Col_Task As String
  Dim Col_TORTASKID As String
  Dim lookup_value As String
  Dim res as Variant
        
  On Error GoTo myError
    
  If Trim(sheet) = "Services" Then
    Col_GrantCode = S_GRANT_CODE
    Col_GrantCodeID = S_GRANTCODEID
  End If
  If Trim(sheet) = "Expenses" Then
    Col_GrantCode = E_GRANTCODE
    Col_GrantCodeID = E_GRANTCODEID
  End If
    
  ' If nothing has been entered in the Grant code column then exit
  If NOT Worksheets(sheet).Range(Col_GrantCode & row).Value > 0 Then
    Worksheets(sheet).Range(Col_GrantCodeID & row).Value = ""
    Application.EnableEvents = True
    Exit Function
  End If
  
  ' Otherwise let's lookup the selected GrantCode in P_GRANT_IDs_1 and P_GRANT_IDs_2
  lookup_value = Worksheets(sheet).Range(Col_GrantCode & row).Value
  res = Application.WorksheetFunction.VLookup(lookup_value, Range("GrantIDs"), 2, False)
  Worksheets(sheet).Range(Col_GrantCodeID & row).Value = res
  Worksheets(sheet).Range(Col_GrantCodeID & row).Interior.ColorIndex = 4
  
myError:
  If Err.Number <> 0 Then
    MsgBox "Error setting GRANTCODEID."
  End If
    
  Application.EnableEvents = True
End Function

'Function to validate all sheets and highlight cells with issues
Public Function ValidateSheets() As Long
  Application.EnableEvents = False

  Dim lastRow As Long
  Dim invalidCells as Long
  
  lastRow = Worksheets("Services").Cells.SpecialCells(xlCellTypeLastCell).row
  invalidCells = 0
  
  'Task, Date, Hours worked, Grant code, Report, TORTASKID, GRANTCODEID for Services
  invalidCells = invalidCells + ValidateColumn("Services", S_TASK, lastRow)
  invalidCells = invalidCells + ValidateColumn("Services", S_DATE, lastRow)
  invalidCells = invalidCells + ValidateColumn("Services", S_HOURS, lastRow)
  invalidCells = invalidCells + ValidateColumn("Services", S_GRANT_CODE, lastRow)
  invalidCells = invalidCells + ValidateColumn("Services", S_REPORT, lastRow)
  invalidCells = invalidCells + ValidateColumn("Services", S_TORTASKID, lastRow)
  invalidCells = invalidCells + ValidateColumn("Services", S_GRANTCODEID, lastRow)
  
  lastRow = Worksheets("Expenses").Cells.SpecialCells(xlCellTypeLastCell).row

  'Task, Date, US amount, Description, Expenses Category, Receipt page ID, Grant code, TORTASKID, GRANTCODEID for Expenses
  invalidCells = invalidCells + ValidateColumn("Expenses", E_TASK, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_DATE, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_US_AMOUNT, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_DESCRIPTION, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_CATEGORY, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_RECEIPT_PAGE, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_GRANTCODE, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_TORTASKID, lastRow)
  invalidCells = invalidCells + ValidateColumn("Expenses", E_GRANTCODEID, lastRow)
    
  'Check for any existing cells that are coloured Red due to validation failures
  invalidCells = invalidCells + CheckForReds  

  Application.EnableEvents = True
  ValidateSheets = invalidCells
End Function

'Function to check a given column for null entries and style them red as well as return their count
Public Function ValidateColumn(sheetToCheck As String, columnToCheck As String, lastRow As Long) As Long
  Application.EnableEvents = False
  Dim invalidRange as Range
  On Error GoTo myError
  Set invalidRange = Worksheets(sheetToCheck).Range(columnToCheck & 2, columnToCheck & lastRow).SpecialCells(xlCellTypeBlanks)
  invalidRange.Interior.ColorIndex = 3
  ValidateColumn = invalidRange.Count
myError:
  Application.EnableEvents = True
End Function

' Delete named sheet
Public Function DeleteSheet(ByVal sheet as String) As Boolean
    Application.DisplayAlerts = False
    Sheets(sheet).Delete
    Application.DisplayAlerts = True
End Function

Public Function IsValidDate(ByVal Target As Excel.Range) As Boolean
  Application.EnableEvents = False
  If Target.Value <> "" Then
    If IsDate(Target.Value) Then
      If DateDiff("m", CDate(Target.Value), Now()) > 3 Then
        Target.Interior.ColorIndex = 3 'Red
        Target.Select
        MsgBox "Dates older than three months are not allowed."
      Else
        Target.Interior.ColorIndex = 0 'No Fill
      End If
    Else
      Target.Interior.ColorIndex = 3 'Red
      Target.Select
      MsgBox "Date is invalid."
    End If
  End If
  Application.EnableEvents = True
End Function

Public Function IsReportOk(ByVal Target As Excel.Range) As Boolean
  If Target.Value <> "" And Len(Target.Value) < 5 Then
    Target.Interior.ColorIndex = 3 'Red
    Target.Select
    MsgBox "Report must be a longer than 5 characters."
  Else
    Target.Interior.ColorIndex = 0 'No Fill
  End If
End Function

Public Function CalculateHours(ByVal Target As Excel.Range) As Boolean
  Application.EnableEvents = False
  On Error GoTo myError
  
  'If no start time and end time supplied then exit
  If IsEmpty(Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row).Value) and _
    IsEmpty(Worksheets(Target.Worksheet.Name).Range(S_END_TIME & Target.row).Value) Then
      Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row, S_HOURS & Target.row).Interior.ColorIndex = 0
      Worksheets(Target.Worksheet.Name).Range(S_HOURS & Target.row).Formula = ""
      Application.EnableEvents = True
      Exit Function
  End If
    
  'If both are specified then set the formula and exit
  If Not(IsEmpty(Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row).Value)) and _
    Not(IsEmpty(Worksheets(Target.Worksheet.Name).Range(S_END_TIME & Target.row).Value)) Then
      Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row, S_HOURS & Target.row).Interior.ColorIndex = 0
      Worksheets(Target.Worksheet.Name).Range(S_HOURS & Target.row).Formula = "=(RC[-1]-RC[-2])*24"
      If Worksheets(Target.Worksheet.Name).Range(S_HOURS & Target.row).Value <= 0 Then
        Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row, S_HOURS & Target.row).Interior.ColorIndex = 3 'Red
        MsgBox "Please check the times, as you've worked zero or less hours!"
      End If
      Application.EnableEvents = True
      Exit Function
  End If
  
myError:
  If Err.Number <> 0 Then
    MsgBox "Error when calculating hours worked"
  End If
   
  Application.EnableEvents = True
End Function

Public Function CheckHours(ByVal Target As Excel.Range) As Boolean
  Application.EnableEvents = False
  On Error GoTo myError
  
  If IsEmpty(Target.Value) Then
    Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row, S_HOURS & Target.row).Interior.ColorIndex = 0
    Application.EnableEvents = True
    Exit Function
  Else
    If Target.Value <= 0 Then
      Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row, S_HOURS & Target.row).Interior.ColorIndex = 3
      MsgBox "Please check the times, as you've worked zero or less hours!"
      Target.Interior.ColorIndex = 3 'Red
    End If
    If Target.Value > 0 Then
      Worksheets(Target.Worksheet.Name).Range(S_START_TIME & Target.row, S_HOURS & Target.row).Interior.ColorIndex = 0
    End If
  End If
  
myError:
   If Err.Number <> 0 Then
        MsgBox "Error when checking hours worked"
   End If
   
   Application.EnableEvents = True
End Function

Public Function CheckUSAmount(ByVal Target As Excel.Range) As Boolean
    If Target.Value <> "" Then
        If IsNumeric(Target.Value) Then
          Target.Interior.ColorIndex = 0 'No Fill
        Else
            Target.Interior.ColorIndex = 3 'Red
            Target.Select
            MsgBox "US amount must be a valid number."
        End If
    End If
End Function

Public Function CheckDescription(ByVal Target As Excel.Range) As Boolean
    If Target.Value <> "" And Len(Target.Value) < 5 Then
        Target.Interior.ColorIndex = 3 'Red
        Target.Select
        MsgBox "Description must be greater than 5 characters."
    Else
        Target.Interior.ColorIndex = 0 'No Fill
    End If
End Function

Public Function CheckReceiptPageID(ByVal Target As Excel.Range) As Boolean
        If Target.Value <> "" Then
            If IsNumeric(Target.Value) Then
                    If CDbl(Target.Value) >= 0 And CDbl(Target.Value) <= 100 Then
                        Target.Interior.ColorIndex = 0 'No Fill
                    Else
                        Target.Interior.ColorIndex = 3 'Red
                        Target.Select
                        MsgBox "Receipt page ID must be between 0 and 100."
                    End If
            Else
                Target.Interior.ColorIndex = 3 'Red
                Target.Select
                MsgBox "Receipt page ID must be a valid number."
            End If
        End If
End Function

Public Function CheckForReds() As Long
    Dim Cll As Range
    Dim res As Long
    
    res = 0
    
    For Each Cll In Worksheets("Services").Range(S_TOR & 2, S_GRANTCODEID & Worksheets("Services").Cells.SpecialCells(xlCellTypeLastCell).row)
        If Cll.Interior.ColorIndex = 3 Then
            res = res + 1
        End If
    Next Cll
    
    For Each Cll In Worksheets("Expenses").Range(E_TOR & 2, E_GRANTCODEID & Worksheets("Expenses").Cells.SpecialCells(xlCellTypeLastCell).row)
        If Cll.Interior.ColorIndex = 3 Then
            res = res + 1
        End If
    Next Cll
    
    CheckForReds = res
End Function

