Attribute VB_Name = "Parameters"
Sub Update_Parameters_Dialog()
    MonthForm.Show
End Sub

Sub URL_Get_Query()

  On Error GoTo myError

  If IsNull(Range("README!F5")) Then
    MsgBox "There was a problem getting data from Planet. Please contact support."
    Exit Sub
  End If
    
  Dim str As String
  Dim url As String
  
  url = ThisWorkbook.CustomDocumentProperties("PLANET_URL") & ThisWorkbook.CustomDocumentProperties("ACCESS_KEY") & "/" & Range("README!F6").Value
  
  Application.Cursor = xlWait
  
  
  Dim f As Boolean

  Sheets.Add.Name = "MySheet"
  Sheets("MySheet").Move after:=Worksheets(Worksheets.Count)
  
  With Worksheets("MySheet").QueryTables.Add(Connection:="URL;" & url, Destination:=Range("A1"))
    .BackgroundQuery = True
    .TablesOnlyFromHTML = True
    .RefreshStyle = xlOverwriteCells
    .Refresh BackgroundQuery:=False
    .SavePassword = False
    .SaveData = True
  End With
  
  ' Check that the new sheet has more than one row. If not, then find the error.
  Dim lastRow As Long
  lastRow = Worksheets("MySheet").Cells.SpecialCells(xlCellTypeLastCell).row
        
  If lastRow = 1 Then
    MsgBox Worksheets("MySheet").Range("A1").Value
    f = DeleteSheet("MySheet")
    Worksheets("README").Select
    Application.Cursor = xlDefault
    Exit Sub
  End If
        
myError:
  If Err.Number <> 0 Then
    f = DeleteSheet("MySheet")
    Worksheets("README").Select
    MsgBox "There was a problem getting data from Planet. Please contact support."
    Application.Cursor = xlDefault
    Exit Sub
  End If
                
  'Delete Parameters sheet and rename new sheet
  DeleteSheet ("Parameters")
  Worksheets(Worksheets.Count).Name = "Parameters"

  'Once the Parameters have been updated, then update all named ranges and hide the Parameters sheet
  DefineServiceClaimRanges
  Worksheets("Parameters").Visible = False
        
  Worksheets("README").Select
  MsgBox "Your parameters have been updated successfully. Please save the file and continue completing it."
  Application.Cursor = xlDefault

End Sub

'Updates the named ranges required
Private Sub DefineServiceClaimRanges()
  With Worksheets("Parameters")
    'Delete ranges
    DeleteNamedRange ("TORs")
    DeleteNamedRange ("TORTasks")
    DeleteNamedRange ("Projects")
    DeleteNamedRange ("ProjectTasks")
    DeleteNamedRange ("TaskNodeIDs")
    DeleteNamedRange ("NodeIDGrants")
    DeleteNamedRange ("GrantIDs")
    DeleteNamedRange ("Currencies")
    DeleteNamedRange ("ExpenseCategories")
    'Create ranges
    ThisWorkbook.Names.Add "TORs", Range(P_TORs & 2, Range(P_TORs & 1).End(xlDown))
    ThisWorkbook.Names.Add "TORTasks", Range(P_TORs2 & 2, Range(P_TORs2_TASKS & 1).End(xlDown))
    ThisWorkbook.Names.Add "Projects", Range(P_PROJECTS & 2, Range(P_PROJECTS & 1).End(xlDown))
    ThisWorkbook.Names.Add "ProjectTasks", Range(P_PROJECTS2 & 2, Range(P_PROJECTS2_TASKS & 1).End(xlDown))
    ThisWorkbook.Names.Add "TaskNodeIDs", Range(P_TASKS_IDs_1 & 2, Range(P_TASKS_IDs_2 & 1).End(xlDown))
    ThisWorkbook.Names.Add "NodeIDGrants", Range(P_ID_GRANTS_1 & 2, Range(P_ID_GRANTS_2 & 1).End(xlDown))
    ThisWorkbook.Names.Add "GrantIDs", Range(P_GRANT_IDs_1 & 2, Range(P_GRANT_IDs_2 & 1).End(xlDown))
    ThisWorkbook.Names.Add "Currencies", Range(P_CURRENCIES & 2, Range(P_CURRENCIES & 1).End(xlDown))
    ThisWorkbook.Names.Add "ExpenseCategories", Range(P_EXPENSECATEGORIES & 2, Range(P_EXPENSECATEGORIES & 1).End(xlDown))
  End With
End Sub

' Deletes the specified named range if it exists
Function DeleteNamedRange(TheName As String) As Boolean
  If NameExists(TheName) = True Then
    ThisWorkbook.Names.Item(TheName).Delete
  End If
End Function

' Checks whether the named range exists
Function NameExists(TheName As String) As Boolean
    On Error Resume Next
    NameExists = Len(Names(TheName).Name) <> 0
End Function
