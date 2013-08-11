VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ServicesABCForm 
   Caption         =   "Services ABC Entry"
   ClientHeight    =   1740
   ClientLeft      =   2000
   ClientTop       =   -4400
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

    If ABCcombo.Value <> "" Then ActiveCell.Value = ABCcombo.Value
    
    If ActiveCell.Column = 2 Then
    
        'feed for column C
        Range("README!AQ50").Value = ABCcombo.Value
        
    End If

End Sub

Private Sub CancelButton_Click()
    ServicesABCForm.Hide
    Unload ServicesABCForm
End Sub

Private Sub ServicesABCForm_Terminate()
    ServicesABCForm.Hide
    Unload ServicesABCForm
End Sub

Private Sub SelectButton_Click()

    ServicesABCForm.Hide
    If ActiveCell.Column = 1 Or ActiveCell.Column = 2 Then
        ActiveCell.Offset(0, 1).Select
    ElseIf ActiveCell.Column = 3 Then
    
        'MODIFY TO (0,-2) TO RETURN TO COL A
        ActiveCell.Offset(0, 1).Select
        ServicesABCForm.Hide
        End
    End If

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

    'If ActiveCell.Column > 3 Then
    '    ServicesABCForm.Hide
    '    End
    'End If
    'Dim PS As Positions

    'PS = PositionForm(ServicesABCForm, ActiveCell, 0, 0, cstFhpFormLeftCellRight, cstFvpFormTopCellTop)
    'ServicesABCForm.Top = PS.FrmTop
    'ServicesABCForm.Left = PS.FrmLeft

    If ActiveCell.Column = 1 Then
        TitleText = Range("Services!A1").Value & " selection"
        
    ElseIf ActiveCell.Column = 2 Then
        TitleText = Range("Services!B1").Value & " selection"
        
    ElseIf ActiveCell.Column = 3 Then
        TitleText = Range("Services!C1").Value & " selection"
    
    End If

    ABCcombo.Clear
    
    Dim arrayABC() As String
    Dim i As Integer
    i = 0
    Dim taskList As Variant
    Dim tasksize As Integer
    
    
    ' Populate TOR DDL
    If ActiveCell.Column = 1 Then
    
        ReDim arrayABC(0, Worksheets("Parameters").Range("TORs").Rows.Count - 1)
        For Each c In Worksheets("Parameters").Range("TORs")
            arrayABC(0, i) = c.Value
            i = i + 1
            
        Next c
        
    ' Populate Project DDL
    ElseIf ActiveCell.Column = 2 Then
    
        ReDim arrayABC(0, Worksheets("Parameters").Range("Projects").Rows.Count - 1)
        For Each c In Worksheets("Parameters").Range("Projects")
            arrayABC(0, i) = c.Value
            i = i + 1
        Next c
    
    ' Populate Tasks DDL
    ElseIf ActiveCell.Column = 3 Then
        
        taskList = getTaskForForm(ActiveCell.Worksheet.Name, ActiveCell.row)
        If IsNull(taskList) Then
            ServicesABCForm.Hide
            Exit Sub
        End If
        
        tasksize = UBound(taskList)
        
        ReDim arrayABC(0, tasksize - 1)
        For Each c In taskList
            arrayABC(0, i) = c
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
    Dim result_tor_range As Range
    
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
    
    result_tor_range_string = Res_Task & "$" & (match_location_offset + 1) & ":$" & Res_Task & "$" & (match_location_offset + matches)
    
    Set result_tor_range = Worksheets("Parameters").Range(result_tor_range_string)
    getTaskForForm = result_tor_range.Value
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
