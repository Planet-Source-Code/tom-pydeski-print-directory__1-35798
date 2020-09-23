VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDirPrint 
   Caption         =   "Tom Pydeski's Directory Printer"
   ClientHeight    =   5115
   ClientLeft      =   4230
   ClientTop       =   3180
   ClientWidth     =   11130
   Icon            =   "DirPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFileOnly 
      Caption         =   "FileName Only"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Displays UnSorted File Name Only without size or date info..."
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox lstUnsort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   15
      ToolTipText     =   "Listing of the files in the selected directory."
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8400
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox DirPath 
      Height          =   405
      Left            =   2040
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      ToolTipText     =   "Type a new path and press enter."
      Top             =   0
      Width           =   6195
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1245
      Left            =   600
      TabIndex        =   13
      ToolTipText     =   "Print Preview"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2196
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"DirPrint.frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   3600
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Listing of the files in the selected directory."
      Top             =   840
      Width           =   7455
   End
   Begin VB.Frame SortFrame 
      Height          =   470
      Left            =   3600
      TabIndex        =   7
      Top             =   360
      Width           =   6500
      Begin VB.OptionButton optSort 
         Caption         =   "FileType"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1650
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Click to sort by File's extension..."
         Top             =   150
         Width           =   1600
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Modified"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3250
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Click to sort by Date last modified..."
         Top             =   150
         Width           =   1600
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   50
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click to sort by FileName..."
         Top             =   150
         Width           =   1600
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4850
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to sort by File's size..."
         Top             =   150
         Width           =   1600
      End
   End
   Begin VB.Frame Files 
      Height          =   4710
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3570
      Begin VB.ComboBox Filter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "DirPrint.frx":0FB8
         Left            =   50
         List            =   "DirPrint.frx":0FBF
         TabIndex        =   6
         Text            =   "Filter"
         ToolTipText     =   "Select a filter for the files to display..."
         Top             =   4320
         Width           =   3375
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   50
         TabIndex        =   3
         ToolTipText     =   "Select a drive..."
         Top             =   120
         Width           =   3420
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3690
         Left            =   50
         TabIndex        =   2
         ToolTipText     =   "Select the directory to display..."
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   5715
         Left            =   3360
         MousePointer    =   9  'Size W E
         TabIndex        =   14
         ToolTipText     =   "Click and drag to resize th file list."
         Top             =   120
         Width           =   75
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "DirList v1.23"
      FileName        =   "*.txt"
      Filter          =   "*.txt"
   End
   Begin VB.Label Label1 
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   435
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mAttr 
         Caption         =   "Get File Attributes"
      End
      Begin VB.Menu mSave 
         Caption         =   "&Save Directory"
         Shortcut        =   ^S
      End
      Begin VB.Menu mbar 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "&Print Directory"
         Shortcut        =   ^P
      End
      Begin VB.Menu mPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mprintsel 
         Caption         =   "Printer Setup"
      End
      Begin VB.Menu msep 
         Caption         =   "-"
      End
      Begin VB.Menu mExp 
         Caption         =   "Launch &Explorer"
      End
      Begin VB.Menu mfind 
         Caption         =   "Do Windows &Find "
      End
      Begin VB.Menu msep1 
         Caption         =   "-"
      End
      Begin VB.Menu mRecent 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mbar2 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mHelpFile 
         Caption         =   "&Help"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmDirPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code submitted by Tom Pydeski
'My friend Carlos was after me to write something to print and save directories
'that he could then use to make labels for cd's full of mp3's.
'i looked at what was on psc and found one but modified it greatly.
'sorry, i don't know who the original author was, but this is not even close to what he had.
'instead of just listing the files, the list contains the size and date of the file
'this was accomplished using tabs to separate the data within each list item.
'the tabstops then had to be set for the longest entry in each column
'i then added the option buttons to sort by filename; size; extension; or date.
'Additionally, the 10 last directories are stored as well as any additional filters
'for the file list.  An option for displaying the file names only will use a different
'unsorted listbox (you can't change the sort property at run time) to contain only the
'file names in the order they appear in the directory (un-sorted)
'If a right click menuitem is added to explorer, the program can be launched to print the
'directory.  try putting a directory name in the command line option (goto project; properties;
'and the make tab and put something like C:\WINDOWS\CONFIG\ in the command line arguments
'The original had drag and drop, so i left that in.
'
Option Explicit
'directory print/listbox put together by Tom Pydeski
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
Private Const SW_SHOWNORMAL = 1
Private Const EM_SetTabSTOPS = &HCB
'
Dim RightClick
Dim AppDir
Dim AppName
Dim Ffilter As String
Dim FileExt$()
Dim FileExtIn$
Dim FileBase$
Dim FName$
Dim fi$
Dim CreateT$
Dim AccessT$
Dim WriteT$
Dim FileSize$
Dim SearchPath$
Dim FileNo
Dim FileMax
Const ATTR_DIRECTORY = 16
Private Const MAX_PATH = 260
'
Dim FileText$
Dim FileN$
Dim Flen As Long
Dim Ignores As Byte
Dim LastSort As Integer
Dim RevSort As Byte
Dim listd$()
Dim TotalSize As Long
Dim Directory As String
Private mbResizing As Boolean 'flag to indicate whether mouse left button is pressed down
Dim OriTop As Integer
Dim OriLeft As Integer
Dim OriHeight As Integer
Dim OriWidth As Integer
Dim F As Byte
Dim i As Integer
Dim j As Integer
Dim PathIn$
Dim RecMax As Byte
Dim Commandin$
Dim FileOnly$
Dim lpAppName As String
Dim lpFileName As String
Dim lonStatus  As Long
Dim RecKey$
Dim nStringLen As Integer
Dim RecStr$
Dim Def$
Dim rKey As Integer
Dim PrinterPicked As Byte
Dim fcName As String
Dim Elaps$
Dim Sizes$

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
'put in some default filters
Filter.Clear
Filter.AddItem "*.*"
Filter.AddItem "*.ico"
Filter.AddItem "*.bmp;*.jpg;*.jpeg;*.gif;*.ico;*.sys"
Filter.AddItem "*.mp3;*.wav;*.ram"
Filter.AddItem "*.exe;*.com;*.sys"
Filter.AddItem "*.ini"
Filter.AddItem "*.txt;*.doc"
Filter.AddItem "*.eml"
Filter.Text = "*.*"
'
On Error GoTo nofile
'read the filters we have stored already
F = FreeFile
ChDir App.Path
Open "Filters.ini" For Input As #F
Dim filterin$
Do
    Input #F, filterin$
    'don't add the filter if it already exists
    For j = 0 To Filter.ListCount - 1
        If UCase(filterin$) = UCase(Filter.list(j)) Then GoTo skip
    Next j
    Filter.AddItem filterin$
skip:
Loop Until EOF(F)
Close #F
'
nofile:
GetRecent
'it is possible to launch this from a right click so we need to
'get the directory it is launching
If Command$ <> "" Then
    Commandin$ = Command$
    Dim comlen As Byte, quote As Byte, dirt2 As Byte
    comlen = Len(Commandin$)
    quote = InStr(Commandin$, Chr$(34))
    If quote > 0 Then
        FName$ = Mid$(Commandin$, 2, comlen - 2)
    Else
        FName$ = Commandin$
    End If
    Caption = "Lauching " + FName$
    PathIn$ = ""
    dirt2 = 1
checkp:
    PathIn$ = StripPath(FName$)
    If PathIn$ <> "" Then
        ChDir PathIn$
    End If
    'End If
    Dir1.Path = PathIn$
    File1.Path = PathIn$
End If
endload:
Screen.MousePointer = 11
FillList
'Ignores = 1
'optSort(1).Value = True
LastSort = -1
Ignores = 0
Label2.Top = Files.Top
Label2.Height = Files.Height
Label2.MousePointer = vbSizeWE
Inits = 1
FormWinRegPos Me
ResizeFiles
Refresh
If Command$ <> "" Then
    PrinterPicked = 1 'print to default printer
    mprint_Click
    Refresh
    Screen.MousePointer = 0
    Unload Me
End If
Screen.MousePointer = 0
End Sub

Private Function FullPath(sText As String) As String
'Check to see if a \ is needed at the end of the path or not
Dim sMyFormat As String
sMyFormat = DirPath.Text
If Right(sMyFormat, 1) <> "\" Then sMyFormat = sMyFormat & "\"
FullPath = sMyFormat & sText
End Function

Private Sub filter_Change()
On Error Resume Next
If Inits = 0 Then Exit Sub
If Ignores = 1 Then Exit Sub
File1.Pattern = Filter.Text
FillList
End Sub

Private Sub filter_Click()
File1.Pattern = Filter.Text
FillList
End Sub

Private Sub Filter_KeyDown(KeyCode As Integer, Shift As Integer)
Ignores = 1
'if a filter is typed in and enter is hit, get the list
If KeyCode = 13 Then
    Ignores = 0
    File1.Pattern = Filter.Text
    FillList
End If
End Sub

Private Sub filter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    For i = 0 To Filter.ListCount - 1
        If UCase(Filter.Text) = UCase(Filter.list(i)) Then Exit Sub
        Filter.AddItem Filter.Text, 0
    Next i
    SaveFilters
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Text1.Visible = False
End Sub

Private Sub Form_Resize()
'adjust the sizes to match the available real estate
If Me.WindowState = vbMinimized Then Exit Sub
ResizeFiles
Me.Files.Height = Me.ScaleHeight - Me.Files.Top - 100
Filter.Top = Me.Files.Height - Filter.Height - 100
Dir1.Height = (Filter.Top - 100) - Dir1.Top
DirPath.Width = (Me.Width - DirPath.Left) - 100
Label2.Top = Files.Top
Label2.Height = Files.Height
Label2.Width = 200
Label2.Left = Files.Width - Label2.Width
Label2.MousePointer = vbSizeWE
If Inits = 0 Then Exit Sub
lstFiles.Height = (Me.ScaleHeight - lstFiles.Top)
With lstUnsort
    .Left = lstFiles.Left
    .Top = SortFrame.Top + 100
    .Height = lstFiles.Height + SortFrame.Height
    .Width = lstFiles.Width
End With
'
Text1.Move 0, 0, Me.Width - 100, Me.Height - 700
End Sub

Sub ResizeFiles()
'resize the frame and its contents
'Files.Width = (Me.ScaleWidth - Label2.Left) - 100
lstFiles.Left = Files.Width + 50
lstFiles.Width = (Me.ScaleWidth - lstFiles.Left) - 25
SortFrame.Left = lstFiles.Left
File1.Width = Files.Width - 100
Dir1.Width = File1.Width - 50
Drive1.Width = Dir1.Width
Filter.Width = Dir1.Width
'DoEvents
Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
mexit_Click
End Sub

Private Sub chkFileOnly_Click()
If chkFileOnly.Value = Checked Then
    With lstUnsort
        .Clear
        .Visible = False
        .Left = lstFiles.Left
        .Top = SortFrame.Top + 100
        .Height = lstFiles.Height + SortFrame.Height
    End With
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
    ListSubDirs DirPath
    'SetListWidth lstUnsort
    lstUnsort.Visible = True
    SetTabs lstUnsort
    ResizeList lstUnsort, True
Else
    lstUnsort.Visible = False
    FillList
    If LastSort > 0 Then optSort_Click (LastSort)
End If
'SetListWidth lstUnsort
lstFiles.Visible = Not lstUnsort.Visible
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then mbResizing = True
'Debug.Print File1.Left - (Dir1.Left + Dir1.width)
OriLeft = Files.Left
OriWidth = Files.Width
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mbResizing = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'resize controls while the left mousebutton is pressed down
If mbResizing Then
    Dim nx As Single
    nx = Files.Width + X
    'Caption = X & " " & nx
    If nx > Me.ScaleWidth - 500 Then Exit Sub
    Files.Width = nx
    Label2.Left = Files.Width - Label2.Width
    ResizeFiles
End If
End Sub

Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'get the name of the file the mouse is over and display it as a tip
Dim P As Long
Dim XPosition As Long, YPosition As Long
XPosition = CLng(X / Screen.TwipsPerPixelX)
YPosition = CLng(Y / Screen.TwipsPerPixelY)
P = SendMessage(lstFiles.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((YPosition * 65536) + XPosition))
If P < lstFiles.ListCount Then
    lstFiles.ToolTipText = Replace(lstFiles.list(P), vbTab, " : ") & " x=" & X / Screen.TwipsPerPixelX & "Y=" & Y / Screen.TwipsPerPixelY
End If
If P > 0 And P < lstFiles.ListCount Then
    lstFiles.ListIndex = P
End If
lstFiles.SetFocus
End Sub

Private Sub lstfiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'get the name of the file the mouse is over and display it as a tip
'also right click produces the popup menu
On Error GoTo Oops
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
Dim P$
Dim SelIt As Integer
Dim NewIndex As Long
lXPoint = CLng(X / Screen.TwipsPerPixelX)
lYPoint = CLng(Y / Screen.TwipsPerPixelY)
RightClick = Button
If Button = 2 Then
    SelIt = (Y \ Me.TextHeight("Test"))
    NewIndex = SelIt + lstFiles.TopIndex
    If NewIndex < lstFiles.ListCount Then
        lstFiles.ListIndex = NewIndex
    End If
    lIndex = SendMessage(lstFiles.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
    FName$ = Left(lstFiles.list(lstFiles.ListIndex), InStr(1, lstFiles.list(lstFiles.ListIndex), vbTab) - 1)
    P$ = File1.Path
    DirPath.Text = P$
    If Right$(P$, 1) <> "\" Then P$ = P$ & "\"
    If FName$ <> "" Then
        FName$ = P$ + FName$
    End If
    Caption = FName$
    PopupMenu mfile
End If
GoTo Exit_lstfiles_MouseDown
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine lstfiles_MouseDown "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in lstfiles_MouseDown"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_lstfiles_MouseDown:
End Sub

Private Sub lstUnsort_Click()
lstUnsort.Visible = False
End Sub

Private Sub mAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mAttr_Click()
Dim em$
'show the file attributes
Directory = StripPath(FName$)
Dim File As String
File = FileOnly
If File = "" Then
    MsgBox "No File Selected", vbOKOnly + vbInformation, "File Explorer"
    Exit Sub
End If
WriteT$ = Format(FileDateTime(FName$), "mm/dd/yyyy hh:nnam/pm")
FileSize$ = AddSpace(Format(FileLen(FName$), "#,# Bytes"), 10)
em$ = "File Name: " & File & vbCrLf
em$ = em$ & "File Directory: " & Directory & vbCrLf
em$ = em$ & "File Extension: " & Right(File, 4) & vbCrLf
em$ = em$ & "File Size: " & FileSize$ & vbCrLf
em$ = em$ & "Last Modified: " & WriteT$
MsgBox em$, vbOKOnly + vbInformation
End Sub

Sub mexit_Click()
'save the settings
FormWinRegPos Me, True
SaveFilters
Unload Me
End
End Sub

Private Sub mfind_Click()
DoFind
End Sub

Private Sub mexp_Click()
Dim sCom As String
'launch explorer
sCom = "c:\windows\explorer.exe " + Dir1.Path
Shell sCom, vbNormalFocus
End Sub

Private Sub mHelpFile_Click()
MsgBox "Sorry, the only help available is common sense and tooltips." & vbCrLf & "Move the mouse over a control for its tip."
End Sub

Private Sub msave_Click()
Dim cFileName As String
mPrintPreview_Click
CommonDialog1.InitDir = DirPath
FName$ = Replace(DirPath, "\", "~")
FName$ = Replace(FName$, "c:", "", , , vbTextCompare)
CommonDialog1.FileName = FName$
CommonDialog1.Action = 2
cFileName = CommonDialog1.FileName
'Format the save to filename
If InStr(1, cFileName, ".") = 0 Then cFileName = cFileName & ".txt"
If cFileName <> "" And cFileName <> "*.txt" Then Text1.SaveFile cFileName, rtfText
Text1.Visible = False
End Sub

Private Sub Dir1_Change()
Dim NextFile As Byte
Dim newB As Byte
Dim newB2 As Byte
Dim MenuNum As Integer
'
If Inits = 0 Then Exit Sub
' Change File List Box to display new subdirectory
Screen.MousePointer = 11
Me.chkFileOnly.Value = 0
lstUnsort.Visible = False
'
File1.Path = Dir1.Path
File1.Refresh
DoEvents
DirPath.Text = File1.Path
ChDir App.Path
'save the new directory and all the recents
Open "DirPrint.ini" For Output As #1
Print #1, Dir1.Path
Close
If File1.ListCount > 0 Then File1.ListIndex = 0
'
RecMax = mRecent.UBound
lpFileName = App.Path & "\" & "DirPrint.ini"
lpAppName = "Directories"
lonStatus = WritePrivateProfileString(lpAppName, "Recent0", Dir1.Path, lpFileName)
If mRecent(0).Caption = Dir1.Path Then
    Exit Sub
End If
'check if the path is a derivative of the previous path
'dont save c:\windows if the pathin is c:\windows\system
'since we have to back out through the parent again
'first check if the new path is a subfolder of the last path
newB = InStr(1, Dir1.Path, mRecent(0).Caption, vbTextCompare)
'also check if the last path is a subfolder of the new path
newB2 = InStr(1, mRecent(0).Caption, Dir1.Path, vbTextCompare)
If newB > 0 Then
    'mRecent(0).Caption = Dir1.Path
End If
If newB2 > 0 Then
    GoTo dexit
End If
If RecMax < 9 And Dir1.Path <> "" Then
    RecMax = RecMax + 1
    Load mRecent(RecMax)
    mRecent(RecMax).Visible = True
    'Debug.Print "New = "; mRecent(RecMax).Caption
End If
' Copy RecentFile1 to RecentFile2, and so on.
If RecMax > 1 Then
    For MenuNum = RecMax To 1 Step -1
        mRecent(MenuNum).Caption = mRecent(MenuNum - 1).Caption
        mRecent(MenuNum).Visible = True
    Next MenuNum
End If
mRecent(0).Caption = Dir1.Path
'Debug.Print "--------->"; RecMax
'For i = 0 To RecMax
'    Debug.Print mRecent(i).Caption
'Next i
'Debug.Print
'
NextFile = 1
'save the recent menus
For i = 1 To RecMax '3 To 1 Step -1
    If mRecent(i).Caption <> "" And UCase(mRecent(i).Caption) <> UCase(Dir1.Path) Then
        RecKey$ = "Recent" & NextFile
        RecStr$ = mRecent(i).Caption
        lonStatus = WritePrivateProfileString(lpAppName, RecKey$, RecStr$, lpFileName)
        NextFile = NextFile + 1
        If NextFile = 10 Then Exit For
    End If
Next
'
'remove the dupes from the recent file list
For i = RecMax To 1 Step -1
    For j = i - 1 To 1 Step -1
        If UCase(mRecent(i).Caption) = UCase(mRecent(j).Caption) Then
            'we are already on the list
            mRecent(i).Caption = ""
            mRecent(i).Visible = False
            Exit For
        End If
    Next j
Next i
dexit:
FillList
If LastSort > 0 Then optSort_Click (LastSort)
Screen.MousePointer = 0
End Sub

Private Sub Dir1_Click()
Dim ButtonIn As Byte
If ButtonIn = 2 Then
    'right click code
    ButtonIn = 0
    Exit Sub
End If
File1.Path = Dir1.list(Dir1.ListIndex)
File1.Refresh
End Sub


Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'i don't remember why i did this...
If KeyCode = vbKeyV And Shift = 2 Then
    PathIn$ = Clipboard.GetText
    If Dir(PathIn$, vbDirectory) <> "" Then
        Dir1.Path = PathIn$
    End If
End If
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dir1.Path = Dir1.list(Dir1.ListIndex)
    DirPath.Text = File1.Path
    Dir1.Refresh
End If
If KeyAscii = 8 Then
    Dir1.Path = Dir1.list(Dir1.ListIndex - 1)
    DirPath.Text = File1.Path
    Dir1.Refresh
End If
Dir1.SetFocus
End Sub

Private Sub mprintsel_Click()
Screen.MousePointer = 11
'select the printer
CommonDialog1.CancelError = True
On Error GoTo Oops
CommonDialog1.ShowPrinter
PrinterPicked = 1
GoTo mexit
Oops:
If Err.Number = 32755 Then GoTo mexit
'RETRY=4,ABORT=3,IGNORE=5
eMess$ = "Error # " + Str$(Err) + " - " + Error$
mError = MsgBox(eMess$, 2, "mprintsel_Click")
If mError = 4 Then Resume
If mError = 5 Then Resume Next
mexit:
Screen.MousePointer = 11
End Sub

Sub mprint_Click()
mPrintPreview_Click
If PrinterPicked = 0 Then
    mprintsel_Click
End If
Printer.Font.Name = Text1.Font.Name
Printer.Font.Bold = Text1.Font.Bold
Printer.Font.SIZE = Text1.Font.SIZE
Text1.SelPrint Printer.hDC
Text1.Visible = False
mPrintPreview.Checked = False
End Sub

Sub mPrintPreview_Click()
Dim tString As String
Dim list As ListBox
'the only issue here is that the tab stops don't line up like the list box
'i tried to send the settabs message, but it did not work
mPrintPreview.Checked = Not mPrintPreview.Checked
If mPrintPreview.Checked = True Then
    Text1.Text = ""
    'LongLen(0) = LongLen(0) + 10
    If chkFileOnly = 0 Then
        tString = "Directory of " & Dir1.Path & vbCrLf & vbCrLf
        SetTabs lstFiles
        If lstFiles.SelCount > 0 Then
            For i = 0 To lstFiles.ListCount - 1
                If lstFiles.Selected(i) = True Then
                    listd$() = Split(lstFiles.list(i), vbTab)
                    'now lets pad all the file names to the same length
                    listd$(0) = listd$(0) & String(LongLen(0) - Len(listd$(0)), " ")
                    listd$(1) = AddSpace(listd$(1), LongLen(1) - Len(listd$(1)))
                    listd$(2) = AddSpace(listd$(2), LongLen(2) - Len(listd$(2)))
                    tString = tString & Join(listd$(), vbTab) & vbCrLf
                End If
            Next i
        Else
            For i = 0 To lstFiles.ListCount - 1
                listd$() = Split(lstFiles.list(i), vbTab)
                'now lets pad all the file names to the same length
                listd$(0) = listd$(0) & String(LongLen(0) - Len(listd$(0)), " ")
                listd$(1) = AddSpace(listd$(1), LongLen(1) - Len(listd$(1)))
                listd$(2) = AddSpace(listd$(2), LongLen(2) - Len(listd$(2)))
                tString = tString & Join(listd$(), vbTab) & vbCrLf
            Next i
        End If
        tString = tString & vbCrLf
        tString = tString & lstFiles.ListCount & " Total Files "
        tString = tString & AddSpace(Format(TotalSize, "#,#"), 12) & " bytes" & vbCrLf & vbCrLf
        tString = tString & "Report Generated on " & Date & vbCrLf
    Else
        If lstUnsort.SelCount > 0 Then
            For i = 0 To lstUnsort.ListCount - 1
                If lstUnsort.Selected(i) = True Then
                    tString = tString & lstUnsort.list(i) & vbCrLf
                End If
            Next i
        Else
            For i = 0 To lstUnsort.ListCount - 1
                tString = tString & lstUnsort.list(i) & vbCrLf
            Next i
        End If
    End If
    Text1.Move 0, 0, Me.Width - 100, Me.Height - 700
    'Text1.Font.Name = lstFiles.Font.Name
    'Text1.Font.Bold = lstFiles.Font.Bold
    'Text1.Font.SIZE = lstFiles.Font.SIZE
    Text1.Text = tString
    Text1.Visible = True
    Text1.ZOrder 0
    Refresh
    DoEvents
Else
    Text1.Visible = False
End If
End Sub

Private Sub mRecent_Click(index As Integer)
On Error Resume Next
Dir1.Path = mRecent(index).Caption
End Sub

Private Sub optSort_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'setup the sort parameters
If index = LastSort Then
    RevSort = 1 - RevSort
    optSort_Click (index)
Else
    RevSort = 0
End If
End Sub

Sub optSort_Click(index As Integer)
Dim lCount As Integer
Dim lItem As Integer
Dim extLoc As Integer
Dim Ext$
'this routine created by Tom Pydeski
LastSort = index
If Ignores = 1 Then Exit Sub
'list item contains name; date; and size
'we will also sort by extension
Ignores = 1
Dim SortKey$(4)
Dim lists$()
Screen.MousePointer = 11
lCount = lstFiles.ListCount - 1
If lCount < 0 Then Exit Sub
ReDim lists$(lCount)
lstFiles.Enabled = False
lstFiles.Visible = False
For lItem = lCount To 0 Step -1
    'get the list info for each item
    LText = lstFiles.list(lItem)
    'split the info by the tabs
    listd$() = Split(LText, vbTab)
    extLoc = InStr(listd$(0), ".")
    Ext$ = Mid$(listd$(0), extLoc + 1)
    'get the split data
    fcName = Trim(listd$(0))
    WriteT$ = Trim(listd$(1))
    FileSize$ = Trim(listd$(2))
    'create an elapsed key to help sort the dates
    Elaps$ = AddZero(DateDiff("s", WriteT$, Now), 10)
    If FileSize$ <> "" Then Sizes$ = AddZero(Str$(Val(Int(FileSize$))), 10)
    'setup the various sort keys
    SortKey$(0) = fcName
    SortKey$(1) = Elaps$
    SortKey$(2) = Sizes$
    SortKey$(3) = Ext$
    'add the sortkey to the front of each list item
    'since the sort property is true, the list will sort based on the data
    'added as the sort key (which is delimited by a tab)
    fcName = SortKey$(index) & vbTab & fcName & vbTab & WriteT$ & vbTab & FileSize$
    'now remove the item
    lstFiles.RemoveItem (lItem)
    'add the new info to the array
    lists$(lItem) = fcName
Next lItem
'now that the array is built, let's put the data back into the list box
For lItem = 0 To lCount
    lstFiles.AddItem lists$(lItem)
Next lItem
If RevSort = 1 Then
    'reverse the order
    'this is done by adding the list data in opposite order
    For lItem = lCount To 0 Step -1
        lists$(lItem) = lstFiles.list(lItem)
        lstFiles.RemoveItem (lItem)
    Next lItem
    For lItem = 0 To lCount
        lstFiles.AddItem lists$(lItem), 0
    Next lItem
End If
'now remove the sorting key
'this leaves the list data in the current locations, but
'removes the extra text we used to sort.
For rKey = 0 To lstFiles.ListCount - 1
    listd$() = Split(lstFiles.list(rKey), vbTab)
    If UBound(listd) = 3 Then
        lstFiles.list(rKey) = listd$(1) & vbTab & listd$(2) & vbTab & listd$(3)
    End If
Next rKey
lstFiles.Enabled = True
lstFiles.Visible = True
Ignores = 0
Screen.MousePointer = 0
End Sub

Private Sub DirPath_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dir1.Path = DirPath.Text
End If
End Sub

Private Sub DirPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'this is part of the original example
Dim txt As String
Dim FName As Variant
Dim MyPos As Integer
MyPos = 0
'Go through the list of file names in the list
'of selected files. This will cause just the last
'selected one to be displayed in the path text box.
For Each FName In Data.Files
    txt = txt & FName & vbCrLf
Next FName
If InStr(1, txt, vbCrLf) Then
    DirPath.Text = Left(txt, Len(txt) - 2)
Else
    DirPath.Text = txt
End If
Dir1.Path = DirPath.Text
Effect = vbDropEffectNone
FillList
End Sub

Sub FillList()
On Error GoTo Oops
Dim fn As Integer
Dim fNm As String
Screen.MousePointer = 11
lstFiles.Enabled = False
lstFiles.Visible = False
lstFiles.Clear
ChDir File1.Path
TotalSize = 0
'build the list
For fn = 0 To File1.ListCount - 1
    fcName = File1.list(fn)
    fNm = File1.Path & "\" & (fcName)
    If Right(File1.Path, 1) <> "\" Then
        fNm = File1.Path & "\" & (fcName)
    Else
        fNm = File1.Path & (fcName)
    End If
    TotalSize = TotalSize + FileLen(fNm)
    WriteT$ = Format(FileDateTime(fNm), "mm/dd/yyyy hh:nnam/pm")
    FileSize$ = AddSpace(Format(FileLen(fNm), "#,#"), 10)
    Elaps$ = AddZero(DateDiff("s", WriteT$, Now), 10)
    fcName = Elaps$ & vbTab & fcName & vbTab & WriteT$ & vbTab & FileSize$
    lstFiles.AddItem fcName
Next fn
'
remsortkey:
'now remove the sorting key
Ignores = 1
For rKey = 0 To lstFiles.ListCount - 1
    listd$() = Split(lstFiles.list(rKey), vbTab)
    If UBound(listd) = 3 Then
        lstFiles.list(rKey) = listd$(1) & vbTab & listd$(2) & vbTab & listd$(3)
    End If
Next rKey
nosort:
SetTabs lstFiles
AddScroll lstFiles
'below does not really do anything since we later
'resize the list to fit in the available form space
ResizeList lstFiles
If Me.WindowState = vbNormal Then
    Me.Width = lstFiles.Width + lstFiles.Left + 200
    If Me.Width < Me.SortFrame.Left + Me.SortFrame.Width + 200 Then
        Me.Width = Me.SortFrame.Left + Me.SortFrame.Width + 200
    End If
    If Me.Left + Me.Width > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
        If Me.Left < 0 Then Me.Left = 0
        If Me.Width > Screen.Width Then
            Me.Width = Screen.Width
        End If
    End If
End If
'resize the buttons
SortFrame.Width = lstFiles.Width
'set the size of the sort buttons
'i re-arranged the position of the buttons
'file name and type share the first column
optSort(0).Width = (Me.TextWidth(LongLine(0) & "  ") / 2) + 100
optSort(3).Width = optSort(0).Width
optSort(1).Width = Me.TextWidth(LongLine(1) & "  ") + 100
For i = 0 To 3
    If optSort(i).Width < 1000 Then
        optSort(i).Width = 1000
    End If
Next i
optSort(3).Left = optSort(0).Left + optSort(0).Width
optSort(1).Left = optSort(3).Left + optSort(3).Width
optSort(2).Left = optSort(1).Left + optSort(1).Width
'give same buffer to the last button
optSort(2).Width = (Me.SortFrame.Width - optSort(2).Left) - optSort(0).Left
'
lstFiles.Enabled = True
lstFiles.Visible = True
Ignores = 0
GoTo Exit_FillList
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine FillList "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in FillList"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_FillList:
Screen.MousePointer = 0
End Sub

Sub DoFind()
'do the windows find
DoEvents
ShellExecute 0, "find", "C:\", vbNullString, vbNullString, SW_SHOW
DoEvents
End Sub

Sub SaveFilters()
'save the data we have entered as filters
Dim Savedfils$
Savedfils$ = ""
F = FreeFile
ChDir App.Path
Open "Filters.ini" For Output As #F
For i = 0 To Filter.ListCount - 1
    'check if we already saved it
    If InStr(1, Savedfils$, Filter.list(i), vbTextCompare) = 0 Then
        Print #F, Filter.list(i)
        Savedfils$ = Savedfils$ & Filter.list(i) & ","
    End If
Next i
Close #F
End Sub

Public Sub FormWinRegPos(pMyForm As Form, Optional pbSave As Boolean)
'this is from another psc submission
'sorry, but i don't know who originally submitted it.
On Error GoTo Oops
'This Procedure will Either Retrieve or Save Form Posn values
'Best used on Form Load and Unload or QueryUnLoad
With pMyForm
    If pbSave Then
        'If Saving then do this...
        'If Form was minimized or Maximized then Closed Need to set Back to Normal
        'Or previous non Max or Min State then Save Posn Parameters
        If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
            .WindowState = vbNormal
        End If
        'Save AppName...FrmName...KeyName...Value
        SaveSetting App.EXEName, .Name, "Top", .Top
        SaveSetting App.EXEName, .Name, "Left", .Left
        SaveSetting App.EXEName, .Name, "Height", .Height
        SaveSetting App.EXEName, .Name, "Width", .Width
        SaveSetting App.EXEName, .Name, "fWidth", Files.Width
    Else
        'If Not Saveing Must Be Getting ..
        'Need to ref AppName...FrmName...KeyName
        '(If nothing Stored Use The Exisiting Form value)
        .Top = GetSetting(App.EXEName, .Name, "Top", .Top)
        .Left = GetSetting(App.EXEName, .Name, "Left", .Left)
        .Height = GetSetting(App.EXEName, Name, "Height", .Height)
        .Width = GetSetting(App.EXEName, .Name, "Width", .Width)
        Files.Width = GetSetting(App.EXEName, .Name, "fWidth", Files.Width)
    End If
End With
GoTo Exit_FormWinRegPos
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine FormWinRegPos "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in FormWinRegPos"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_FormWinRegPos:
End Sub

Sub GetRecent()
'retrieve the recent file list
On Error GoTo nofile2
F = FreeFile
Open "DirPrint.ini" For Input As #F
Input #F, PathIn$
Close #F
Drive1.Drive = Left(PathIn$, 3)
Dir1.Path = PathIn$
File1.Path = PathIn$
DirPath.Text = Dir1.Path
If File1.ListCount > 0 Then File1.ListIndex = 0
RecMax = mRecent.UBound
mRecent(0).Visible = True
lpAppName = "Directories"
lpFileName = App.Path & "\" & "DirPrint.ini"
For i = 0 To 9
    RecKey$ = "Recent" & i
    nStringLen = 255
    RecStr$ = Space(255) 'String(nStringLen, Chr$(0)) ' Buffer String
    lonStatus = GetPrivateProfileString(lpAppName, RecKey$, "", RecStr$, Len(RecStr$), lpFileName)
    RecStr$ = Trim(RecStr$)
    RecStr$ = Left(RecStr$, Len(RecStr$) - 1)
    If RecStr$ = "" Then Exit For
    If RecStr$ = "-" Then Exit For
    If i > RecMax Then
        Load mRecent(i)
    End If
    mRecent(i).Visible = True
    mRecent(i).Caption = RecStr$
Next
Def$ = mRecent(0).Caption
If Def$ <> "" Then Dir1.Path = Def$
nofile2:
End Sub

Function StripPath(FilePathIn As String)
'below starts from the right side and is quicker
Dim ppos As Integer
ppos = InStrRev(FilePathIn, "\")
StripPath = Left$(FilePathIn, ppos)
FileOnly$ = Mid$(FilePathIn, ppos + 1)
pathend:
End Function

Function AddZero(StrIn$, intDigits)
StrIn$ = Trim(StrIn$)
If Len(StrIn$) >= intDigits Then
    AddZero = StrIn$
    Exit Function
End If
AddZero = String(intDigits - Len(StrIn$), "0") & StrIn$
End Function

Function AddSpace(StrIn$, intDigits)
StrIn$ = Trim(StrIn$)
If Len(StrIn$) >= intDigits Then
    AddSpace = StrIn$
    Exit Function
End If
AddSpace = String(intDigits - Len(StrIn$), " ") & StrIn$
End Function

Private Sub ListSubDirs(Path)
'this is an ollllld routine.
'this gets the files in the order they appear.
'this was put in for cd's with lists of mp3's
'if we sort them in any way, they will affect the playlist
'
Dim COUNT, D(), i, DIRNAME  ' Declare variables.
Screen.MousePointer = 11
DIRNAME = Dir(Path, ATTR_DIRECTORY) ' Get first directory name.
'Iterate through PATH, caching all subdirectories in D()
Do While DIRNAME <> ""
    If DIRNAME <> "." And DIRNAME <> ".." Then
        If DIRNAME <> "C:\Pagefile.sys" Then
            If GetAttr(Path & DIRNAME) = ATTR_DIRECTORY Then
                '= ATTR_DIRECTORY Then
                If (COUNT Mod 10) = 0 Then
                    ReDim Preserve D(COUNT + 10)    ' Resize the array.
                End If
                COUNT = COUNT + 1   ' Increment counter.
                D(COUNT) = DIRNAME
            Else
                FileMax = FileMax + 1
                Caption = Str$(FileMax) + " Files"
                FName$ = Replace(DIRNAME, ",", vbTab)
                FName$ = Replace(FName$, "-", vbTab)
                lstUnsort.AddItem FName$
                'EXECUTED ONLY IF NOT DIRECTORY
            End If
        End If
    End If
    DIRNAME = Dir   ' Get another directory name.
    DoEvents
Loop
' Now recursively iterate through each cached subdirectory.
'un comment this if you want to get all of the files, not just those
'in the parent directory
'For i = 1 To COUNT
'    'list1.AddItem PATH & D(I)   ' Put name in list box.
'    ListSubDirs Path & D(i) & "\"
'Next i
Screen.MousePointer = 0
End Sub

