Attribute VB_Name = "modDirList"
'Code submitted by Tom Pydeski
'Routines to set the tab stops for a list box and
'to optionally resize the list and the form for the contents
Option Explicit
'listbox function module put together by Tom Pydeski
Declare Function GetDialogBaseUnits Lib "user32" () As Long
Global FormWindowHwnd As Long
Global ListBoxHwnd As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXVSCROLL = 2 ' Return the width of a vertical scrollbar.
Private Const SM_CYHSCROLL = 3 'Return the height of a horizontal scrollbar
Private Const SM_CXEDGE = 45 'Return the width of a 3D window border.
Dim lVerticalScrollbarWidth As Long
Dim ScrollbarWidth As Long
Dim Win3DWidth As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEFAULT_GUI_FONT = 17
Private Const GDI_ERROR = &HFFFF
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Type SIZE
    cx As Long
    cy As Long
End Type
Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31
Dim LongestListItem As String
'
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam$) As Long
Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
Public TabTot As Long 'Listbox API constant
Public Const LB_ADDSTRING = &H180
Public Const LB_INSERTSTRING = &H181
Public Const LB_DELETESTRING = &H182
Public Const LB_SELITEMRANGEEX = &H183
Public Const LB_RESETCONTENT = &H184
Public Const LB_SetSEL = &H185
Public Const LB_SetCURSEL = &H186
Public Const LB_GetSEL = &H187
Public Const LB_GetCURSEL = &H188
Public Const LB_GetText = &H189
Public Const LB_GetTextLen = &H18A
Public Const LB_GetCount = &H18B
Public Const LB_SELECTSTRING = &H18C
Public Const LB_DIR = &H18D
Public Const LB_GetTOPINDEX = &H18E
Public Const LB_FINDSTRING = &H18F
Public Const LB_GetSelCOUNT = &H190
Public Const LB_GetSelItems = &H191
Public Const LB_SetTabSTOPS = &H192
Public Const LB_GetHORIZONTALEXTENT = &H193
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SetCOLUMNWIDTH = &H195
Public Const LB_ADDFILE = &H196
Public Const LB_SETTOPINDEX = &H197
Public Const LB_GetITEMRECT = &H198
Public Const LB_GetITEMDATA = &H199
Public Const LB_SetITEMDATA = &H19A
Public Const LB_SELITEMRANGE = &H19B
Public Const LB_SetANCHORINDEX = &H19C
Public Const LB_GetANCHORINDEX = &H19D
Public Const LB_SetCARETINDEX = &H19E
Public Const LB_GetCARETINDEX = &H19F
Public Const LB_SetITEMHEIGHT = &H1A0
Public Const LB_GetITEMHEIGHT = &H1A1
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SetLOCALE = &H1A5
Public Const LB_GetLOCALE = &H1A6
Public Const LB_SetCOUNT = &H1A7
Public Const LB_MSGMAX = &H1A8
Public Const LB_ITEMFROMPOINT = &H1A9
'
Global NumTabs As Byte
Global MaxTabs As Long
Dim arylist() As String
Global TabArray() As Long
Global LLen As Long
Global LongLen() As Long
Global LongLine()
Global LstT()
'
'A different way to find the width and height of each list item
'was written by John Calvert and is at
'http://msdn.microsoft.com/msdnmag/issues/1200/combo/combo.asp
Dim hDC As Long
Dim lFont As Long
Dim lFontOld As Long
Dim lResult As Long
Dim uSize As SIZE
Dim SpaceWidth As Long
Dim i As Integer
Dim j As Integer
Dim tabnum As Integer
Global DlgWidthUnits As Integer
Dim ShowScroll As Byte
Dim ListWidth As Long
'these are used elsewhere
Global eTitle$
Global eMess$
Global mError As Long
Global Inits As Byte
Global ListText As String
Global LText As String
'
Dim intGreatestLen As Integer
Dim lngGreatestWidth As Long
Dim lineWidth As Long
Dim lCount As Integer
Dim litemHeight As Long
Dim NewHeight As Long
Dim ParTop As Integer
Dim TotWidth As Integer
Dim MaxWidth As Integer

Public Sub AddScroll(list As ListBox)
'adds horizontal scrollbar to listbox if necessary
Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
'Find Longest Text in Listbox
For i = 0 To list.ListCount - 1
    If Len(list.list(i)) > Len(list.list(intGreatestLen)) Then
        intGreatestLen = i
    End If
Next i
'Get Twips
lngGreatestWidth = list.Parent.TextWidth(list.list(intGreatestLen) + Space(1))
'Space(1) is used to prevent the last Character from being cut off
'Convert to Pixels
lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
'Use api to add scrollbar
SendMessage list.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
End Sub

Sub SetTabs(list As ListBox)
On Error GoTo Oops
'usage: SetTabs lstFind
Dim intGreatestLen As Integer
Dim lngGreatestWidth As Long
'Find Longest Text in Listbox
Dim ThisTab As Long
If list.ListCount = 0 Then Exit Sub
Screen.MousePointer = 11
ReDim LstT(1)
ReDim LongLine(1)
Dim LText As String
MaxTabs = 0
'Get the Dialog Width Units
DlgWidthUnits = (GetDialogBaseUnits() Mod 65536) / 2
'
For i = 0 To list.ListCount - 1
    'split the list data by tabs
    arylist = Split(list.list(i), vbTab)
    'count the number of tabs in each line
    NumTabs = UBound(arylist)
    If NumTabs = 0 Then GoTo notabs
    If NumTabs > MaxTabs Then
        'store the highest number of tabs
        MaxTabs = NumTabs
        ReDim LstT(MaxTabs)
        ReDim LongLine(MaxTabs)
        ReDim LongLen(MaxTabs)
    End If
    For j = 0 To NumTabs
        LstT(j) = Trim(arylist(j))
        LText = LstT(j) & " "
        'Determine the width of the string for each tabbed 'column'
        'LLen = Len(LText)
        'multipy the length of the item by the average space
        'DlgWidthUnits is the average character width
        'so all we have to do is multiply the # of characters by this #
        'it seems to be the most reliable method
        'i did not think it would work with various fonts, but it appears that it does
        LLen = (Len(LText) * DlgWidthUnits) + 5
        If LLen > LongLen(j) Then
            LongLen(j) = LLen
            LongLine(j) = LText
        End If
        'Debug.Print LText; "="; LLen,
    Next j
    'Debug.Print
notabs:
Next i
'
ReDim TabArray(0 To MaxTabs) As Long
TabArray(0) = 0
TabTot = 0
'first tabstop should be zero
'Debug.Print "---------------------------------------"
For tabnum = 1 To MaxTabs
    ThisTab = LongLen(tabnum - 1)
    TabArray(tabnum) = TabTot + ThisTab
    TabTot = TabTot + ThisTab
    'Debug.Print tabnum, TabArray(tabnum), lngGreatestWidth, LongLine(tabnum - 1)
Next tabnum
'Debug.Print "Total = "; TabTot
'
'clear any existing tabs
Call SendMessage(list.hwnd, LB_SetTabSTOPS, 0&, ByVal 0&)
'set list tabstops
Call SendMessage(list.hwnd, LB_SetTabSTOPS, MaxTabs + 1, TabArray(0))
list.Refresh
Screen.MousePointer = 0
GoTo Exit_SetTabs
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine SetTabs "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in SetTabs"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_SetTabs:
End Sub

Sub ResizeList(list As ListBox, Optional SizeForm As Boolean)
'pardon the mess with this sub, but there are many different ways
'to try to determine the proper width and they don't all work the same way
'therefore, we needed to try different things and did not want to delete the others...
'
'sizeform is is boolean which tells the form to resize or not
On Error GoTo Oops
Dim i As Integer
'
'set the parent form's font to the listbox's font
'we need to do this so that the textwidth function works on the proper font
list.Parent.Font.Name = list.Font.Name
list.Parent.Font.Bold = list.Font.Bold
list.Parent.Font.SIZE = list.Font.SIZE
lCount = list.ListCount
If lCount = 0 Then Exit Sub
Inits = 0
Screen.MousePointer = 11
'
ListBoxHwnd& = list.hwnd
'Get a handle to the device context for the control
hDC = GetDC(ListBoxHwnd&)
lFont = SendMessage(ListBoxHwnd&, WM_GETFONT, 0, ByVal 0)
'Select the font in to the device context, and retain prior font
lFontOld = SelectObject(hDC, lFont)
'If (lFontOld = 0) Or (lFontOld = GDI_ERROR) Then GoTo nodc
'
'find height of each list item
'below is replaced by the new dc method
'litemheight = List.Parent.TextHeight("Test")
lResult = GetTextExtentPoint32(hDC, list.list(0), Len(list.list(0)), uSize)
If (lResult <> 0) Then
    'Return the string length
    litemHeight = uSize.cy * Screen.TwipsPerPixelY
Else
    litemHeight = list.Parent.TextHeight("Test")
End If
'if listbox style is checked then we must add more for each item
NewHeight = ((lCount + 1) * (litemHeight + (list.Style * 22)))
'initialize scrollbar width
lVerticalScrollbarWidth = 0
seth:
If list.Parent.WindowState = vbNormal And SizeForm = True Then
    ParTop = list.Parent.Top
    If list.Top + NewHeight + ParTop + 700 > Screen.Height Then
        If ParTop <> 0 Then
            list.Parent.Top = 0
            GoTo seth
        End If
        NewHeight = (Screen.Height - (list.Parent.Top + list.Top)) - 800
        ShowScroll = 1
    End If
Else
    If list.Top + NewHeight + 700 > Screen.Height Then
        NewHeight = (list.Parent.Height - list.Top) - 900
        ShowScroll = 1
    End If
End If
list.Height = NewHeight
'
'Get the Dialog Width Units
'DlgWidthUnits = (GetDialogBaseUnits() Mod 65536) / 2
'Find Longest Text in Listbox
Dim newWid As Integer
lngGreatestWidth = 0
For i = 0 To list.ListCount - 1
    ListText = list.list(i) & Space(2)
    'above might mess up with some fonts
    lineWidth = list.Parent.TextWidth(ListText)
    'below is another way of doing it, but i could not get it to work
    'newWid = Len(ListText) * DlgWidthUnits * Screen.TwipsPerPixelX
    'this only works for sytem font used in dialog boxes
    'i think this messses up depending on the # of tabs...
    'lineWidth = 2 * newWid
    If lineWidth > lngGreatestWidth Then
        intGreatestLen = i
        lngGreatestWidth = lineWidth
        LongestListItem = ListText
        'Debug.Print tStr$; "="; linewidth
    End If
Next i
'we should do below if each list item has different # of tabs
'what if we try replacing the tabs with spaces?
'ltext = Replace(List.List(intGreatestLen), vbTab, "") + String$(MaxTabs - 1, vbTab)
'i don't know if it really is a good approach
LText = LongestListItem
GoTo nodc
'
'Determine the width of the string
'this method messes up with tabs
lResult = GetTextExtentPoint32(hDC, LText, Len(LText), uSize)
If (lResult <> 0) Then
    'Return the string length
    lngGreatestWidth = uSize.cx * Screen.TwipsPerPixelX
Else
    lngGreatestWidth = list.Parent.TextWidth(LText)
End If
'below is to be replaced by the new dc method, but dc method does not work...
lngGreatestWidth = list.Parent.TextWidth(LText)
'Space(1) is used to prevent the last Character from being cut off
'
nodc:
TotWidth = lngGreatestWidth
'Debug.Print totwidth, ltext
checkpos:
'this is from still another example for the enhanced list box
'make sure we are still on the screen
MaxWidth = list.Parent.Left + TotWidth + 220
If MaxWidth > Screen.Width Then
    If list.Parent.Left > 0 Then
        list.Parent.Left = Screen.Width - MaxWidth
        If list.Parent.Left < 0 Then list.Parent.Left = 0
        GoTo checkpos
    End If
    TotWidth = (Screen.Width - list.Parent.Left) - 250
End If
'Account for the window border, 1 pixel either side
'probably more because it is 3d
Win3DWidth = GetSystemMetrics(SM_CXEDGE) * Screen.TwipsPerPixelX
TotWidth = TotWidth + (2 * (Win3DWidth))
ScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
'Account for the scrollbar width plus the 3d border * 2 + a fudge factor
lVerticalScrollbarWidth = ShowScroll * (ScrollbarWidth + (2 * Win3DWidth)) ' * Screen.TwipsPerPixelX
'
'allow for scrollbar
list.Width = TotWidth + lVerticalScrollbarWidth + 200
'if we cant see the list then add the horizontal scroll
If MaxWidth > list.Width Then
    SendMessage list.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth / Screen.TwipsPerPixelX, 0
    list.Height = list.Height + 400 'account for the scrollbar
End If
If list.Parent.WindowState = vbNormal And SizeForm = True Then
    list.Parent.Height = list.Height + list.Top + 750  'add menu area +status bar
    list.Parent.Width = list.Width + list.Left + 100
End If
list.Refresh
'Reset the device context font and delete the temporary font. Ignore any errors.
SelectObject hDC, lFontOld
DeleteObject lFont
'Release the device context handle. Ignore any errors.
ReleaseDC ListBoxHwnd&, hDC
Inits = 1
Screen.MousePointer = 0
GoTo Exit_ResizeList
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine ResizeList "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in ResizeList"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_ResizeList:
End Sub

Sub SetListWidth(list As ListBox)
'below is from OddityX submission to psc
'it seems to be wider than necessary
'fix the width of the list
'Determine the width of the string
'Get a handle to the device context for the control
'i think this was actually from john calvert's msdn article
Dim hDC As Long
Dim lFont As Long
Dim lFontOld As Long
Dim lResult As Long
Dim uSize As SIZE
Dim ListBoxHwnd As Long
ListBoxHwnd& = list.hwnd
hDC = GetDC(ListBoxHwnd&)
lFont = SendMessage(ListBoxHwnd&, WM_GETFONT, 0, ByVal 0)
'Select the font in to the device context, and retain prior font
lFontOld = SelectObject(hDC, lFont)
If (lFontOld = 0) Or (lFontOld = GDI_ERROR) Then GoTo nodc
'Get the Dialog Width Units
DlgWidthUnits = (GetDialogBaseUnits() Mod 65536) / 2
'Find Longest Text in Listbox
ListWidth = 0
ShowScroll = 1
For i = 0 To list.ListCount - 1
    ListText = list.list(i)
    'above might mess up with some fonts
    'lineWidth = DlgWidthUnits * Len(ListText & " ") * Screen.TwipsPerPixelX
    'above too narrow
    'Determine the width of the string
    lResult = GetTextExtentPoint32(hDC, ListText & " ", Len(ListText & " "), uSize)
    If (lResult <> 0) Then
        'Return the string length
        lineWidth = uSize.cx * Screen.TwipsPerPixelX
    Else
        lineWidth = DlgWidthUnits * Len(ListText & " ") * Screen.TwipsPerPixelX
    End If
    If lineWidth > ListWidth Then
        intGreatestLen = i
        ListWidth = lineWidth
        LongestListItem = ListText
        'Debug.Print tStr$; "="; linewidth
    End If
Next i
GoTo nodc23
'
'Determine the width of the string
LongestListItem = LongestListItem & " "
lResult = GetTextExtentPoint32(hDC, LongestListItem, Len(LongestListItem), uSize)
If (lResult <> 0) Then
    'Return the string length
    ListWidth = uSize.cx * Screen.TwipsPerPixelX
Else
    GoTo nodc
End If
nodc23:
'Reset the device context font and delete the temporary font. Ignore any errors.
SelectObject hDC, lFontOld
DeleteObject lFont
'Release the device context handle. Ignore any errors.
ReleaseDC ListBoxHwnd&, hDC
'Account for the window border, 1 pixel either side
'probably more because it is 3d
'ListWidth = ListWidth + 2
Win3DWidth = (GetSystemMetrics(SM_CXEDGE)) * Screen.TwipsPerPixelX
ListWidth = ListWidth + (2 * (Win3DWidth)) + 10
'Account for the scrollbar width plus a fudge factor
lVerticalScrollbarWidth = ShowScroll * ((GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) + (2 * (Win3DWidth)) + 4)
'set the width of the list to a little more than the longest entry
ListWidth = ListWidth + lVerticalScrollbarWidth
list.Width = ListWidth
'below was used to move the list and resize the main window, but we won't do that here
'wWidth = ListWidth
'adjust the main window
'MoveWindow FormWindowHwnd&, wLeft, wTop, wWidth, wHeight, True
'reposition the list
'MoveWindow ListBoxHwnd&, 1, 0, ListWidth, ListHeight, True
'Debug.Print " final="; ListWidth
nodc:
End Sub


