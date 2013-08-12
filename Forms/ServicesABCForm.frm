VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ServicesABCForm 
   Caption         =   "Services ABC Entry"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   -880
   ClientWidth     =   16920
   OleObjectBlob   =   "ServicesABCForm.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "ServicesABCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ABCcombo_Change()

End Sub

Private Sub CancelButton_Click()
    ServicesABCForm.Hide
    Unload ServicesABCForm
End Sub

Private Sub ServicesABCForm_Terminate()
    ServicesABCForm.Hide
    Unload ServicesABCForm
End Sub

Private Sub ClearButton_Click()
    ActiveCell.Value = ""
    ServicesABCForm.Hide
End Sub

Private Sub ProceedButton_Click()
    ActiveCell.Value = ABCcombo.Value
    ServicesABCForm.Hide
    Unload ServicesABCForm
End Sub

Private Sub UserForm_Activate()
On Error GoTo trap
    Dim columnName As String

    If ActiveCell.Worksheet.Name = "Services" Then
        If ActiveCell.Column = S_TOR_INDEX Then
            columnName = "TOR"
        ElseIf ActiveCell.Column = S_PROJECT_INDEX Then
            columnName = "Project"
        ElseIf ActiveCell.Column = S_TASK_INDEX Then
            columnName = "Task"
        ElseIf ActiveCell.Column = S_GRANT_CODE_INDEX Then
            columnName = "Grant"
        End If
    ElseIf ActiveCell.Worksheet.Name = "Expenses" Then
         If ActiveCell.Column = E_TOR_INDEX Then
            columnName = "TOR"
        ElseIf ActiveCell.Column = E_PROJECT_INDEX Then
            columnName = "Project"
        ElseIf ActiveCell.Column = E_TASK_INDEX Then
            columnName = "Task"
        ElseIf ActiveCell.Column = E_GRANTCODE_INDEX Then
            columnName = "Grant"
        ElseIf ActiveCell.Column = E_CURRENCY_INDEX Then
            columnName = "Currency"
        ElseIf ActiveCell.Column = E_CATEGORY_INDEX Then
            columnName = "Category"
        End If
    End If
    TitleText = columnName & " selection"

    ABCcombo.Clear
    
    Dim arrayABC() As String
    Dim i As Integer
    i = 0
    Dim taskList As Variant
    Dim taskSize As Integer
    Dim taskRange As Range
    Dim grantList As Variant
    Dim grantSize As Integer
    Dim grantRange As Range
    
    ' Populate TOR DDL
    If columnName = "TOR" Then
    
        ReDim arrayABC(0, Worksheets("Parameters").Range("TORs").Rows.Count - 1)
        For Each c In Worksheets("Parameters").Range("TORs")
            arrayABC(0, i) = c.Value
            i = i + 1
            
        Next c
        
    ' Populate Project DDL
    ElseIf columnName = "Project" Then
    
        ReDim arrayABC(0, Worksheets("Parameters").Range("Projects").Rows.Count - 1)
        For Each c In Worksheets("Parameters").Range("Projects")
            arrayABC(0, i) = c.Value
            i = i + 1
        Next c
    
    ' Populate Tasks DDL
    ElseIf columnName = "Task" Then
        
        taskList = getTaskForForm(ActiveCell.Worksheet.Name, ActiveCell.row)
        If IsNull(taskList) Then
            ServicesABCForm.Hide
            Exit Sub
        End If
        
        Set taskRange = Worksheets("Parameters").Range(taskList)
        
        'UBound doesn't work well when taskList is a single cell
        'taskSize = UBound(taskList)
        taskSize = taskRange.Rows.Count
        ReDim arrayABC(0, taskSize - 1)
        For Each c In taskRange
            arrayABC(0, i) = c
            i = i + 1
        Next c
        
    ' Populate Grants DDL
    ElseIf columnName = "Grant" Then
    
        grantList = getGrantForForm(ActiveCell.Worksheet.Name, ActiveCell.row)
        If IsNull(grantList) Then
            ServicesABCForm.Hide
            Exit Sub
        End If
        
        Set grantRange = Worksheets("Parameters").Range(grantList)
            
        grantSize = grantRange.Rows.Count
        Debug.Print grantSize
        ReDim arrayABC(0, grantSize - 1)
        For Each c In grantRange
            arrayABC(0, i) = c
            i = i + 1
        Next c
        
    ' Populate Currency DDL
    ElseIf columnName = "Currency" Then
    
        ReDim arrayABC(0, Worksheets("Parameters").Range("Currencies").Rows.Count - 1)
        For Each c In Worksheets("Parameters").Range("Currencies")
            arrayABC(0, i) = c.Value
            i = i + 1
            
        Next c
    
    ' Populate Category DDL
    ElseIf columnName = "Category" Then
    
        ReDim arrayABC(0, Worksheets("Parameters").Range("ExpenseCategories").Rows.Count - 1)
        For Each c In Worksheets("Parameters").Range("ExpenseCategories")
            arrayABC(0, i) = c.Value
            i = i + 1
            
        Next c
    
    End If
    
    ABCcombo.Column() = arrayABC
    ' Set column widths as wider than combobox size to allow scrolling
    ABCcombo.ColumnWidths = 1500
    
Exit Sub
trap:
MsgBox "Error is: " & Err

End Sub

' Looks up the Task for the selected Project or TOR
Public Function getTaskForForm(ByVal sheet As String, ByVal row As Long) As Variant
    Application.EnableEvents = False
    'Debug.Print "Sheet: " & sheet
    'Debug.Print "Row: " & row
    On Error GoTo myError
    
    Dim raw_value As String
    Dim lookup_value As String
    Dim search_range As Range
    Dim match_location_offset As Double
    Dim matches As Double
    Dim result_tor_range_string As String
    
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
    If Trim(CStr(Worksheets(sheet).Range(Col_TORs & row).Value)) <> "" And _
      Trim(CStr(Worksheets(sheet).Range(Col_Project & row).Value)) <> "" Then
        'Debug.Print "BOTH TOR and Project selected"
        MsgBox "Please select EITHER a TOR item or a Project. Not both"
        GoTo exitGracefully
    End If
    
    ' Check whether TOR or Project has been chosen and set lookup value and search range
    If Trim(CStr(Worksheets(sheet).Range(Col_TORs & row).Value)) <> "" Then
        raw_value = CStr(Worksheets(sheet).Range(Col_TORs & row).Value)
        Res_Task = P_TORs2_TASKS
        Set search_range = Worksheets("Parameters").Range("TORTasks").Columns(1)
    End If
    
    If Trim(CStr(Worksheets(sheet).Range(Col_Project & row).Value)) <> "" Then
        raw_value = CStr(Worksheets(sheet).Range(Col_Project & row).Value)
        Res_Task = P_PROJECTS2_TASKS
        Set search_range = Worksheets("Parameters").Range("ProjectTasks").Columns(1)
    End If
    
    ' No need to do anything if neither TOR or Project has been set
    If raw_value = "" Then
        'Debug.Print "Exiting function prematurely"
        MsgBox "Please select a Project or TOR item"
        GoTo exitGracefully
    End If
    
    lookup_value = Left(raw_value, 48) & "*"
    
    'Debug.Print "Search range: " & search_range.Address
    'Debug.Print "Lookup: " & lookup_value
    
    match_location_offset = Application.WorksheetFunction.Match(lookup_value, search_range, 0)
    matches = Application.WorksheetFunction.CountIf(search_range, lookup_value)
    
    result_tor_range_string = "$" & Res_Task & "$" & (match_location_offset + 1) & ":$" & Res_Task & "$" & (match_location_offset + matches)
    getTaskForForm = result_tor_range_string
    Application.EnableEvents = True
    Exit Function
    
exitGracefully:
    getTaskForForm = Null
    Application.EnableEvents = True
    Exit Function
        
myError:
    
    MsgBox "No Tasks found for selected TOR item or Project"
    getTaskForForm = Null
    Application.EnableEvents = True
    
End Function

Public Function getGrantForForm(ByVal sheet As String, ByVal row As Long) As Variant
    Application.EnableEvents = False
    
    On Error GoTo myError
    
    Dim lookup_value As Double
    Dim search_range As Range
    Dim match_location_offset As Double
    Dim matches As Double
    Dim result_tor_range_string As String
    
    If Trim(sheet) = "Services" Then
        Col_TORTASKID = S_TORTASKID
        Col_GrantCode = S_GRANT_CODE
    End If
    If Trim(sheet) = "Expenses" Then
        Col_TORTASKID = E_TORTASKID
        Col_GrantCode = E_GRANTCODE
    End If
    
    lookup_value = Worksheets(sheet).Range(Col_TORTASKID & row).Value
    Debug.Print "MEZZA " & lookup_value
    
    ' No need to do anything if neither TOR or Project has been set
    If Not lookup_value > 0 Then
        GoTo myError
    End If
    
    Set search_range = Range("NodeIDGrants").Columns(1)

    match_location_offset = Application.WorksheetFunction.Match(lookup_value, search_range, 0)
    matches = Application.WorksheetFunction.CountIf(search_range, lookup_value)
    result_tor_range_string = "$" & P_ID_GRANTS_2 & "$" & (match_location_offset + 1) & ":$" & P_ID_GRANTS_2 & "$" & (match_location_offset + matches)
    getGrantForForm = result_tor_range_string
    Application.EnableEvents = True
    Exit Function
    
myError:
    MsgBox "No Grants found for selected Task."
    getGrantForForm = Null
    Application.EnableEvents = True
End Function


