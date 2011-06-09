Private Sub Workbook_Open()
  'Hide Parameters sheet if open
  Sheets("Parameters").Visible = False
  'Add custom colours
  'Light red
  ActiveWorkbook.Colors(5) = RGB(218,150,148)
  'Light blue
  ActiveWorkbook.Colors(4) = RGB(184,204,228)
End Sub

