VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
  'Hide Parameters sheet if open
  Sheets("Parameters").Visible = False
  'Add custom colours
  'Light red
  ActiveWorkbook.Colors(5) = RGB(218, 150, 148)
  'Light blue
  ActiveWorkbook.Colors(4) = RGB(184, 204, 228)
  
  'Checked the named ranges we need exist otherwise alert user
  If NameExists("TORs") = False Then
    MsgBox "You need to update your Parameters before you can start using this sheet."
  End If

End Sub


