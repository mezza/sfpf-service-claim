Attribute VB_Name = "modFormPositioner"
Option Explicit
Option Compare Text

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module Name: modFormPositioner
' Date: 22-Sept-2002
' Author: Chip Pearson, www.cpearson.com, chip@cpearson.com
' Copyright: (c) Copyright 2002, Charles H Pearson.
'
' Description:  Calculates to position to display
' a userform relative to a cell.
'
' Usage:
'   Declare a variable of type Positions:
'       Dim PS As Positions
'   Call the PositionForm function, passing it the following
'   parameters:
'       WhatForm        The userform object
'
'       AnchorRange     The cell relative to which the form
'                       should be displayed.
'
'       NudgeRight      Optional: Number of points to nudge the
'                       for to the right.  This is useful with
'                       bordered range.  Typically, this should
'                       be 0, but may be positive or negative.
'
'       NudgeDown       Optional: Number of points to nudge the
'                       for downward.  This is useful with
'                       bordered range.  Typically, this should
'                       be 0, but may be positive or negative.
'
'       HorizOrientation:   Optional: One of the following values:
'            cstFhpNull             = Left of screen
'            cstFhpAppCenter        = Center of Excel screen
'            cstFhpAuto             = Automatic (recommended and default)
'
'            cstFhpFormLeftCellLeft     = left edge of form at left edge of cell
'            cstFhpFormLeftCellRight    = left edge of form at right edge of cell
'            cstFhpFormLeftCellCenter   = left edge of form at center of cell
'
'            cstFhpFormRightCellLeft    = right edge of form at left edge of cell
'            cstFhpFormRightCellRight   = right edge of form at right edge of cell
'            cstFhpFormRightCellCenter  = right edge of form at center of cell
'
'            cstFhpFormCenterCellLeft   = center of form at left edge of cell
'            cstFhpFormCenterCellRight  = center of form at right edge of cell
'            cstFhpFormCenterCellCenter = center of form at center of cell
'
'       VertOrientation     Optional: One of the following values:
'
'            cstFvpNull                 = Top of screen
'            cstFvpAppCenter            = Center of Excel screen
'            cstFvpAuto                 = Automatic (recommended and default)
'
'            cstFvpFormTopCellTop       = top edge of form at top edge of cell
'            cstFvpFormTopCellBottom    = top edge of form at bottom edge of cell
'            cstFvpFormTopCellCenter    = top edge of form at center of cell
'
'            cstFvpFormBottomCellTop    = bottom edge of form at top of edge of cell
'            cstFvpFormBottomCellBottom = bottom edge of form at bottom edge of cell
'            cstFvpFormBottomCellCenter = bottom edge of form at center of cell
'
'            cstFvpFormCenterCellTop    = center of form at top of cell
'            cstFvpFormCenterCellBottom = center of form at bottom of cell
'            cstFvpFormCenterCellCenter = center of form at center of cell
'
'   For example:
'       PS = PositionForm (UserForm1,Range("C12"),0,0,cstFvpAuto,cstFhpAuto)
'
'   Then, position the form using the values from PS:
'        UserForm1.Top = PS.FrmTop
'        UserForm1.Left = PS.FrmLeft
'   Finally, show the form:
'        UserForm1.Show vbModal
'
'   In summary, the code would look like
'
'       Dim PS As Positions
'       PS = PositionForm (UserForm1,ActiveCell,0,0,cstFvpAuto,cstFhpAuto)
'       UserForm1.Top = PS.FrmTop
'       UserForm1.Left = PS.FrmLeft
'       UserForm1.Show vbModal
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Type: Positions
'
' We store everything in a structure so that we can easily
' pass things around from on procedure to another.  Otherwise,
' we'd quickly run out of stack space passing to the
' optimazation procedures.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type Positions
    
    FrmTop As Single        ' Userform
    FrmLeft As Single
    FrmHeight As Single
    FrmWidth As Single
    
    RngTop As Single        ' Passed in cell
    RngLeft As Single
    RngWidth  As Single
    RngHeight As Single
    
    
    AppTop As Single        'Application
    AppLeft As Single
    AppWidth  As Single
    AppHeight As Single
    
    WinTop As Single        ' Window
    WinLeft As Single
    WinWidth  As Single
    WinHeight As Single
    
    Cell1Top As Single      ' 1st cell in visible range
    Cell1Left As Single
    Cell1Width As Single
    Cell1Height As Single

    LastCellTop As Single   ' last visible cell in window
    LastCellLeft As Single
    LastCellWidth As Single
    LastCellHeight As Single

    BaseLeft As Single      ' the are the screen based coordinates for the upper left corner
    BaseTop As Single       ' of cell.

    VComp As Single         ' compensations for displayed object (toolbars, headers, etc)
    HComp As Single

    NudgeDown As Single     ' allow the user to nudge the positioning by a few pixels.
    NudgeRight As Single
        
#If VBA6 Then
    OrientationH As cstFormHorizontalPosition
    OrientationV As cstFormVerticalPosition
#Else
    OrientationH As Long
    OrientationV As Long
#End If
        
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' End TYPE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#If VBA6 Then
    Public Enum cstFormHorizontalPosition
        cstFhpNull = -2             ' X = 0, left of screen
        cstFhpAppCenter = -1
        cstFhpAuto = 0
        
        cstFhpFormLeftCellLeft
        cstFhpFormLeftCellRight
        cstFhpFormLeftCellCenter
        
        cstFhpFormRightCellLeft
        cstFhpFormRightCellRight
        cstFhpFormRightCellCenter
        
        cstFhpFormCenterCellLeft
        cstFhpFormCenterCellRight
        cstFhpFormCenterCellCenter
    End Enum
        
    Public Enum cstFormVerticalPosition
        cstFvpNull = -2             ' Y = 0, top of screen
        cstFvpAppCenter = -1
        cstFvpAuto = 0
    
        cstFvpFormTopCellTop
        cstFvpFormTopCellBottom
        cstFvpFormTopCellCenter
        
        cstFvpFormBottomCellTop
        cstFvpFormBottomCellBottom
        cstFvpFormBottomCellCenter
    
        cstFvpFormCenterCellTop
        cstFvpFormCenterCellBottom
        cstFvpFormCenterCellCenter
    End Enum
    
#Else
    
    Public Const cstFhpNull As Long = -2                ' X = 0, left of screen
    Public Const cstFhpAppCenter  As Long = -1
    Public Const cstFhpAuto  As Long = 0
        
    Public Const cstFhpFormLeftCellLeft  As Long = 1
    Public Const cstFhpFormLeftCellRight  As Long = 2
    Public Const cstFhpFormLeftCellCenter  As Long = 3
        
    Public Const cstFhpFormRightCellLeft  As Long = 4
    Public Const cstFhpFormRightCellRight  As Long = 5
    Public Const cstFhpFormRightCellCenter  As Long = 6
        
    Public Const cstFhpFormCenterCellLeft  As Long = 7
    Public Const cstFhpFormCenterCellRight  As Long = 8
    Public Const cstFhpFormCenterCellCenter  As Long = 9
    
    Public Const cstFvpNull  As Long = -2                ' Y = 0, top of screen
    Public Const cstFvpAppCenter  As Long = -1
    Public Const cstFvpAuto  As Long = 0
    
    Public Const cstFvpFormTopCellTop As Long = 1
    Public Const cstFvpFormTopCellBottom  As Long = 2
    Public Const cstFvpFormTopCellCenter  As Long = 3
        
    Public Const cstFvpFormBottomCellTop  As Long = 4
    Public Const cstFvpFormBottomCellBottom  As Long = 5
    Public Const cstFvpFormBottomCellCenter  As Long = 6
    
    Public Const cstFvpFormCenterCellTop  As Long = 7
    Public Const cstFvpFormCenterCellBottom  As Long = 8
    Public Const cstFvpFormCenterCellCenter  As Long = 9
   
#End If

Public Const cColHeaderHeight As Single = 9
Public Const cRowHeaderWidth As Single = 20
Public Const cDefaultWindowFrameHeight As Single = 26
Public Const cDefaultWindowFrameWidth As Single = 6
Public Const cDefaultCmdBarHeight = 26
Private Const cLeftBump = 5
Private Const cRightBump = 0
Private Const cUpBump = 0
Private Const cDownBump = 0

#If VBA6 Then
Function PositionForm(WhatForm As Object, AnchorRange As Range, _
    Optional NudgeRight As Single = 0, Optional NudgeDown As Single = 0, _
    Optional ByVal HorizOrientation As cstFormHorizontalPosition = cstFhpAuto, _
    Optional ByVal VertOrientation As cstFormVerticalPosition = cstFvpAuto) As Positions

#Else
Function PositionForm(WhatForm As Object, AnchorRange As Range, _
    Optional NudgeRight As Single = 0, Optional NudgeDown As Single = 0, _
    Optional ByVal HorizOrientation As Long = cstFhpAuto, _
    Optional ByVal VertOrientation As Long = cstFvpAuto) As Positions

#End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PositionForm
'
' The positions the form on the screen according to the specified
' parameters. It returns a Position structure.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim CmdBar As Office.CommandBar
Dim DefaultCmdBarHeight As Single

Dim VCmdArr(0 To 100) As Single  ' hold our command bar widths -- assume fewer that 20 rows
Dim HCmdArr(0 To 100) As Single  ' of command bars.

Dim HasVisibleWindow As Boolean
Dim Win As Window
Dim PS As Positions

Dim Ndx As Long

Dim ColHeaderHeight As Single: ColHeaderHeight = cColHeaderHeight
Dim RowHeaderWidth As Single: RowHeaderWidth = cRowHeaderWidth
Dim DefaultWindowFrameHeight As Single: DefaultWindowFrameHeight = cDefaultWindowFrameHeight
Dim DefaultWindowFrameWidth As Single: DefaultWindowFrameWidth = cDefaultWindowFrameWidth

PS.OrientationH = HorizOrientation
PS.OrientationV = VertOrientation
PS.NudgeRight = NudgeRight
PS.NudgeDown = NudgeDown
'
' If Excel is minimized, set to 0,0 and get out.  The caller should NOT be
' displaying a form when XL is minimized.
'
If Application.WindowState = xlMinimized Then
    WhatForm.Top = 0
    WhatForm.Left = 0
    PS.FrmTop = 0
    PS.FrmWidth = 0
    PS.OrientationH = cstFhpNull
    PS.OrientationV = cstFvpNull
    Exit Function
End If
'
' If the AnchorRange is not within the visible range of the activewindow,
' then force the form to be displayed as AppCenter.
'
If Application.Intersect(AnchorRange, ActiveWindow.VisibleRange) Is Nothing Then
    HorizOrientation = cstFhpAppCenter
    VertOrientation = cstFvpAppCenter
End If
'
' If there are no windows visible, force AppCenter.
'
For Each Win In Application.Windows
    If Win.Visible = True Then
        HasVisibleWindow = True
        Exit For
    End If
Next Win

If HasVisibleWindow = False Then
    HorizOrientation = cstFhpAppCenter
    VertOrientation = cstFvpAppCenter
End If
'
' get our object coordinates.
'
With Application
    PS.AppTop = .Top
    PS.AppLeft = .Left
    PS.AppWidth = .Width
    PS.AppHeight = .Height
End With

With Application.ActiveWindow
    PS.WinTop = .Top
    PS.WinLeft = .Left
    PS.WinWidth = .Width
    PS.WinHeight = .Height
    With .VisibleRange.Cells(1, 1)
        PS.Cell1Top = .Top
        PS.Cell1Left = .Left
        PS.Cell1Height = .Height
        PS.Cell1Width = .Width
    End With
    With .VisibleRange
        PS.LastCellTop = .Cells(.Cells.Count).Top
        PS.LastCellLeft = .Cells(.Cells.Count).Left
        PS.LastCellWidth = .Cells(.Cells.Count).Width
        PS.LastCellHeight = .Cells(.Cells.Count).Height
    End With
End With

With AnchorRange
    PS.RngTop = .Top
    PS.RngLeft = .Left
    PS.RngWidth = .Width
    PS.RngHeight = .Height
End With

PS.FrmHeight = WhatForm.Height
PS.FrmWidth = WhatForm.Width
'
' we'll assume that the application's caption bar and the formula
' bar are the same height as the menu bar.  If we can't figure that out, use 26 as a default.
'
If Application.CommandBars.ActiveMenuBar.Visible = True Then
    DefaultCmdBarHeight = Application.CommandBars.ActiveMenuBar.Height
Else
    DefaultCmdBarHeight = cDefaultCmdBarHeight
End If
'
' We have to have a compenstating factor for command bars. Load an array
' with the heights of visible command bars. The index into the array is
' the RowIndex of the command bar, so we won't "double dip" if two or more
' command bars occupy the same row.
'
For Each CmdBar In Application.CommandBars
    With CmdBar
        If (.Visible = True) And (.Position = msoBarTop) Or (.Position = msoBarMenuBar) Then
            If .RowIndex > 0 Then
                VCmdArr(.RowIndex) = .Height
            End If
        End If
        If (.Visible = True) And (.Position = msoBarLeft) Then
            If .RowIndex > 0 Then
                HCmdArr(.RowIndex) = .Width
            End If
        End If
    End With
Next CmdBar
'
' Now, add up the values in the array so that we can
' get the compensation neeed for toolbars on the
' top and left side of the screen.
'
For Ndx = LBound(VCmdArr) To UBound(VCmdArr)
    PS.VComp = PS.VComp + VCmdArr(Ndx)
Next Ndx

For Ndx = LBound(HCmdArr) To UBound(HCmdArr)
    PS.HComp = PS.HComp + HCmdArr(Ndx)
Next Ndx

'''''''''''''''''''''''''''''''''''''''''''''''''''
' VERTICAL COMPENSATION
'''''''''''''''''''''''''''''''''''''''''''''''''''
If Application.DisplayFullScreen = True Then
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' FULL SCREEN VERTICAL COMPENSATION - OK
    '''''''''''''''''''''''''''''''''''''''''''''''
    PS.VComp = DefaultCmdBarHeight
    '
    ' compensate for the rown and column headers
    '
    If ActiveWindow.DisplayHeadings = True Then
        PS.VComp = PS.VComp + ColHeaderHeight
    Else
        PS.VComp = PS.VComp - (0.666667 * ColHeaderHeight)
    End If
    
    ' no formula bar compensation is required since the
    ' formula bar is not displayed in full-screen mode.

Else
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' NORMAL SCREEN VERTICAL COMPENSATION
    '''''''''''''''''''''''''''''''''''''''''''''''
    '
    ' compensate for the rown and column headers
    '
    If ActiveWindow.DisplayHeadings = True Then
        PS.VComp = PS.VComp + ColHeaderHeight
    Else
        PS.VComp = PS.VComp - (0.666667 * ColHeaderHeight)
    End If
    '
    ' compenstate for formula bar
    '
    If Application.DisplayFormulaBar = True Then
        PS.VComp = PS.VComp + DefaultCmdBarHeight
    Else
        PS.VComp = PS.VComp + (ColHeaderHeight * 1.5)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''
' HORIZONTAL COMPENSATION
'''''''''''''''''''''''''''''''''''''''''''''''''''
If Application.DisplayFullScreen = True Then
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' FULL SCREEN HORIZONTAL COMPENSATION
    '''''''''''''''''''''''''''''''''''''''''''''''
    PS.HComp = 0
    '''''''''''''''''''''''''''''''''''''''''''''''
'Else
    ' do nothing -- HComp is already correct.
End If
'
' compensate for the row and column headers
'
If ActiveWindow.DisplayHeadings = True Then
    PS.HComp = PS.HComp + RowHeaderWidth
Else
    PS.HComp = PS.HComp
End If


'''''''''''''''''''''''''''''''''''''''''''''''
' Now, adjust for the window
'''''''''''''''''''''''''''''''''''''''''''''''
Select Case Application.ActiveWindow.WindowState

    Case xlMaximized
        '
        ' in the case of a maximized window, the action Window.Top
        ' and Window.Left properties will be negative.  Here,
        ' we want 0. as the basis for the window.
        '
        PS.WinTop = 0
        PS.WinLeft = 0
    
    Case xlMinimized
        '
        ' In a minimized window, display in the center of
        ' applicaiton. Force the form to the center of the
        ' application.
        '
        HorizOrientation = cstFhpAppCenter
        VertOrientation = cstFvpAppCenter
        
    Case xlNormal
        PS.WinTop = Abs(ActiveWindow.Top) + DefaultWindowFrameHeight
        PS.WinLeft = Abs(ActiveWindow.Left) + DefaultWindowFrameWidth
    
    Case Else
        ' shouldn't happen
End Select

'''''''''''''''''''''''''''''''''''''''''''''''
' Calculate our BaseLeft and BaseRight values.
' We'll use these as the base relative to which
' the form will actually be positioned.
'
' BaseLeft = LEFT edge of cell
' BaseTop= TOP edge of cell
'
'''''''''''''''''''''''''''''''''''''''''''''''
PS.BaseLeft = PS.AppLeft + PS.WinLeft + PS.HComp + (PS.RngLeft - PS.Cell1Left) + PS.NudgeRight
PS.BaseTop = PS.AppTop + PS.WinTop + PS.VComp + (PS.RngTop - PS.Cell1Top) + PS.NudgeDown

Select Case HorizOrientation

    Case cstFhpNull
        PS.FrmLeft = 0
        
    Case cstFhpAuto
        OptimizeH PS
    
    Case cstFhpFormLeftCellLeft
        PS.FrmLeft = PS.BaseLeft + cLeftBump

    Case cstFhpFormLeftCellRight
        PS.FrmLeft = PS.BaseLeft + PS.RngWidth

    Case cstFhpFormLeftCellCenter
        PS.FrmLeft = PS.BaseLeft + (PS.RngWidth / 2)

    Case cstFhpFormRightCellLeft
        PS.FrmLeft = PS.BaseLeft - PS.FrmWidth
    
    Case cstFhpFormRightCellRight
        PS.FrmLeft = PS.BaseLeft + PS.RngWidth
        
    Case cstFhpFormRightCellCenter
        PS.FrmLeft = PS.BaseLeft + (PS.RngWidth / 2) - PS.FrmWidth

    Case cstFhpFormCenterCellLeft
        PS.FrmLeft = PS.BaseLeft - (PS.FrmWidth / 2)

    Case cstFhpFormCenterCellRight
        PS.FrmLeft = PS.BaseLeft + PS.RngWidth - (PS.FrmWidth / 2)

    Case cstFhpFormCenterCellCenter
        PS.FrmLeft = PS.BaseLeft + (PS.RngWidth / 2) - (PS.FrmWidth / 2)

    Case cstFhpAppCenter    ' same as Case Else
        PS.FrmLeft = PS.AppLeft + (PS.AppWidth / 2) - (PS.FrmWidth / 2)
    
    Case Else               ' same as Case cstFhpAppCenter
        PS.FrmLeft = PS.AppLeft + (PS.AppWidth / 2) - (PS.FrmWidth / 2)
    
End Select


Select Case VertOrientation

    Case cstFvpNull
        PS.FrmTop = 0
        
    Case cstFvpAuto
        OptimizeV PS
    
    Case cstFvpFormTopCellTop
        PS.FrmTop = PS.BaseTop
    
    Case cstFvpFormTopCellBottom
        PS.FrmTop = PS.BaseTop + PS.RngHeight

    Case cstFvpFormTopCellCenter
        PS.FrmTop = PS.BaseTop + (PS.RngHeight / 2)

    Case cstFvpFormBottomCellTop
        PS.FrmTop = PS.BaseTop - PS.FrmHeight

    Case cstFvpFormBottomCellBottom
        PS.FrmTop = PS.BaseTop + PS.RngHeight - PS.FrmHeight

    Case cstFvpFormBottomCellCenter
        PS.FrmTop = PS.BaseTop - PS.FrmHeight + (PS.RngHeight / 2)

    Case cstFvpFormCenterCellTop
        PS.FrmTop = PS.BaseTop - (PS.FrmHeight / 2)

    Case cstFvpFormCenterCellBottom
        PS.FrmTop = PS.BaseTop + PS.RngHeight - (PS.FrmHeight / 2)
    
    Case cstFvpFormCenterCellCenter
        PS.FrmTop = PS.BaseTop + (PS.RngHeight / 2) - (PS.FrmHeight / 2)
        
    Case cstFvpAppCenter    ' same as case else
        PS.FrmTop = PS.AppTop + (PS.AppHeight / 2) - (PS.FrmHeight / 2)
    
    Case Else               ' same as cstFvpAppCenter
        PS.FrmTop = PS.AppTop + (PS.AppHeight / 2) - (PS.FrmHeight / 2)
        
End Select

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Finally, after all that, Move the form to the proper Left and Top
' coordinates.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
WhatForm.Move PS.FrmLeft, PS.FrmTop
PositionForm = PS

End Function

Private Sub OptimizeH(P As Positions)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure optimizes the horizontal position
' of the form.  It MUST define SOME (even arbirary)
' horizontal position.  First, we try to fit the
' form to the right of the cell. If this is unsuccessful,
' we try to fit the form on the left side of the cell.
' If this is unsuccessful, we try to fit the form centered
' to the cell.  If this proves unsuccessful, we
' are stuck with centering the form within the
' application.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WinRight As Single
Dim WinLeft As Single

WinLeft = P.Cell1Left
WinRight = P.LastCellLeft + P.LastCellWidth

' The default horizontal position of the form is aligned on the
' right size of the range.

If P.RngLeft + P.RngWidth + P.FrmWidth < WinRight Then
    P.FrmLeft = P.BaseLeft + P.RngWidth + cLeftBump
    P.OrientationH = cstFhpFormLeftCellRight
    Exit Sub
End If

' If we can't fit it on the right, try the left
'
If P.RngLeft - P.FrmWidth > WinLeft Then
    P.FrmLeft = P.BaseLeft - P.FrmWidth
    P.OrientationH = cstFhpFormRightCellLeft
    Exit Sub
End If

' If we can't fit it on the left, try the center
'
If (P.RngLeft + (P.RngWidth / 2) + (P.FrmWidth / 2) <= WinRight) And _
    (P.RngLeft + (P.RngWidth / 2) - (P.FrmWidth / 2) >= WinLeft) Then
        P.FrmLeft = P.BaseLeft + (P.RngWidth / 2) - (P.FrmWidth / 2)
        P.OrientationH = cstFhpFormCenterCellCenter
        Exit Sub
End If

' If it won't fit on the in the center, we have to go with AppCenter.
'
P.FrmLeft = P.AppLeft + (P.AppWidth / 2) - (P.FrmWidth / 2)
P.OrientationH = cstFhpAppCenter

End Sub

Private Sub OptimizeV(P As Positions)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure optimizes the horizontal position
' of the form.  It MUST define SOME (even arbirary)
' horizontal position.
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WinTop As Single
Dim WinBottom As Single

WinBottom = P.LastCellTop + P.LastCellHeight
WinTop = P.Cell1Top

' The default position is top aligned. See if we have room
' below.
'
If P.RngTop + P.FrmHeight <= WinBottom Then
    P.FrmTop = P.BaseTop
    P.OrientationV = cstFvpFormTopCellTop
    Exit Sub
End If

' If there is no room below, See if we have room above.
'
If P.RngTop - P.FrmHeight >= WinTop Then
    P.FrmTop = P.BaseTop - P.FrmHeight
    P.OrientationV = cstFvpFormTopCellTop
    Exit Sub
End If

' If there is no room above, try the center
'
If (P.RngTop + (P.RngHeight / 2) - (P.FrmHeight / 2) >= WinTop) And _
    (P.RngTop + (P.RngHeight / 2) + (P.FrmHeight / 2) <= WinBottom) Then
    P.FrmTop = P.BaseTop + P.RngTop + (P.RngHeight / 2)
    P.OrientationV = cstFvpFormCenterCellCenter
    Exit Sub
End If

' If we can't put it anywhere else, we have to go with AppCenter
'
P.FrmTop = P.AppTop + (P.AppHeight / 2) - (P.FrmHeight / 2)
P.OrientationV = cstFvpAppCenter

End Sub
