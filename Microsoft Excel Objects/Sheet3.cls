VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
  Dim f As Boolean
  Target.Interior.ColorIndex = 0 'No Fill
  
  Select Case Target.Column
    ' If user selects a Task then we need to set TORTASKID
    Case E_TASK_INDEX
      f = SetTorTaskId(Target.Worksheet.Name, Target.row)
    ' If user enters a Date ensure it's valid
    Case E_DATE_INDEX
      f = IsValidDate(Target)
    ' If user enters a US Amount ensure it's valid
    Case E_US_AMOUNT_INDEX
      f = CheckUSAmount(Target)
    ' If user enters an expense Description make sure it's meaningful
    Case E_DESCRIPTION_INDEX
      f = CheckDescription(Target)
    ' Ensure Receipt page numbers are valid
    Case E_RECEIPT_PAGE_INDEX
      f = CheckReceiptPageID(Target)
    ' If user selects a Grant then we need to set GRANTID
    Case E_GRANTCODE_INDEX
      f = SetGrantCodeID(Target.Worksheet.Name, Target.row)
  End Select
  
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    Dim flag As Boolean
    
    ' Fetch choices for TOR column
    If Target.Column = E_TOR_INDEX And Target.row <> 1 Then
      flag = GetTORs(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Currencies
    If Target.Column = E_CURRENCY_INDEX And Target.row <> 1 Then
      flag = GetCurrency(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Project column
    If Target.Column = E_PROJECT_INDEX And Target.row <> 1 Then
      flag = GetProjects(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Task column
    If Target.Column = E_TASK_INDEX And Target.row <> 1 Then
      flag = GetTask(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Grant code column
    If Target.Column = E_GRANTCODE_INDEX And Target.row <> 1 Then
      flag = GetGrantCode(Target.Worksheet.Name, Target.row)
    End If
    ' Fetch choices for Expense categories
    If Target.Column = E_CATEGORY_INDEX And Target.row <> 1 Then
      flag = GetExpensesCategory(Target.Worksheet.Name, Target.row)
    End If
    
End Sub







