VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrettySort 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pretty Sort"
   ClientHeight    =   7455
   ClientLeft      =   2835
   ClientTop       =   1680
   ClientWidth     =   9645
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   960
      Tag             =   "Rd"
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Index           =   0
      Left            =   180
      MaxLength       =   65500
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   900
      Width           =   9315
   End
   Begin VB.Frame fraTitle 
      BackColor       =   &H80000005&
      Height          =   780
      Left            =   195
      TabIndex        =   10
      Top             =   30
      Width           =   9225
      Begin VB.Image imgState 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Left            =   8700
         Picture         =   "PrettySort.frx":0000
         Top             =   315
         Width           =   195
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pretty File Names - Natural Numeric Sort"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   30
         TabIndex        =   11
         Top             =   225
         Width           =   9000
      End
   End
   Begin VB.Frame fraBorder 
      BackColor       =   &H80000005&
      Height          =   960
      Left            =   195
      TabIndex        =   9
      Top             =   6390
      Width           =   9225
      Begin VB.OptionButton optExt 
         BackColor       =   &H80000005&
         Caption         =   "Group By Extension"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   250
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.OptionButton optExt 
         BackColor       =   &H80000005&
         Caption         =   "Group By Folder"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   545
         Width           =   1845
      End
      Begin VB.CheckBox chkDesc 
         BackColor       =   &H80000005&
         Caption         =   "Descending Order"
         Height          =   255
         Left            =   2070
         TabIndex        =   8
         Top             =   540
         Width           =   1785
      End
      Begin VB.CheckBox chkCaps 
         BackColor       =   &H80000005&
         Caption         =   "List Capitals First"
         Height          =   255
         Left            =   2070
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Open file to sort"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3975
         TabIndex        =   0
         Top             =   300
         Width           =   1665
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sort file names"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         TabIndex        =   1
         Top             =   300
         Width           =   1665
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quit program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7545
         TabIndex        =   2
         Top             =   300
         Width           =   1425
      End
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Index           =   1
      Left            =   180
      MaxLength       =   65500
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   3690
      Width           =   9315
   End
   Begin VB.Image imgMax 
      Height          =   165
      Left            =   9750
      Picture         =   "PrettySort.frx":010E
      Top             =   1890
      Width           =   165
   End
   Begin VB.Image imgNorm 
      Height          =   165
      Left            =   9720
      Picture         =   "PrettySort.frx":021C
      Top             =   2490
      Width           =   165
   End
End
Attribute VB_Name = "frmPrettySort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private file_name As String

Private Sub Form_Load()
    ' Set the default Common Dialog properties
    With dlgCommonDialog
        .CancelError = True
        .InitDir = App.Path
        .Filter = "Text files (*.txt)|*.txt"
        .DefaultExt = "txt"
        ' If the user does not include an extension then '.txt' will
        ' be appended to the filename.
        .FilterIndex = 1 ' text files
        ' If you set more than one file type to the Filter property
        ' you can set the FilterIndex property (indexed from 1) to
        ' specify the default filter in the dialogs 'Files of type'
        ' combo box and determine the initial file types displayed
        ' in the file select listbox.
    End With
End Sub

Private Sub cmdOpen_Click()
    With dlgCommonDialog
        .DialogTitle = "Open File To Sort..."
        .FileName = vbNullString
        .Flags = 4096 + 4 ' File must exist, no read-only checkbox
        On Error GoTo dlgCancelHandler
        .ShowOpen
        file_name = .FileName
    End With
    ' Clear the sorted display
    txtDisplay(1).Text = vbNullString
    ' Call the OpenFile function
    txtDisplay(0).Text = OpenFile(file_name)
    ' Display the filename in the titlebar
    Me.Caption = " Pretty Sort - " & file_name
dlgCancelHandler:
End Sub

Private Sub cmdSort_Click()
    Dim sA() As String
    Screen.MousePointer = vbHourglass
    If txtDisplay(0).Text <> vbNullString Then
        txtDisplay(1).Text = vbNullString
        sA = Split(txtDisplay(0).Text, vbCrLf)
        txtDisplay(1).Text = PrettySort(sA)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Function OpenFile(sFileSpec As String) As String
    ' Handle errors if they occur
    On Error GoTo GetFileError
    Dim iFile As Integer
    iFile = FreeFile
    ' Open in binary mode, let others read but not write
    Open sFileSpec For Binary Access Read Lock Write As #iFile
    ' Allocate the length first
    OpenFile = Space$(LOF(iFile))
    ' Get the file in one chunk
    Get #iFile, , OpenFile
GetFileError:
    Close #iFile ' Close the file
End Function

' Find the True option from a control array of OptionButtons
Private Function GetOption(opts As Object) As Long
    ' Assume no option set True
    'GetOption = -1
    On Error GoTo GetOptionFail
    Dim opt As OptionButton
    For Each opt In opts
        If opt.Value Then
            GetOption = opt.Index
            Exit Function
        End If
    Next
GetOptionFail:
End Function

Private Function PrettySort(aFileNames() As String) As String
    Dim lb As Long, ub As Long
    Dim Grouping As Long
    Dim CapsFirst As Boolean
    CapsFirst = CBool(chkCaps.Value = 1)
    Grouping = GetOption(optExt)
    SortOrder = (chkDesc.Value * -2) + 1 '0 >> 1 : 1 >> -1
    lb = LBound(aFileNames)
    ub = UBound(aFileNames)
    strPrettyFileNames aFileNames, lb, ub, CapsFirst, Grouping
    'strPrettyNumSort aFileNames, lb, ub, CapsFirst
    'strPrettySort aFileNames, lb, ub, CapsFirst
    PrettySort = Join(aFileNames, vbCrLf)
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    ElseIf Me.WindowState = vbMaximized Then
        Set imgState.Picture = imgNorm.Picture
    Else
        Set imgState.Picture = imgMax.Picture
    End If
    fraTitle.Left = (Me.Width - fraTitle.Width) \ 2
    txtDisplay(0).Width = Me.Width - 400
    txtDisplay(1).Width = Me.Width - 400
    fraBorder.Left = (Me.Width - fraBorder.Width) \ 2
    fraBorder.Top = Me.Height - 1485
    txtDisplay(1).Height = ((fraBorder.Top - txtDisplay(1).Top) - 30)
End Sub

Private Sub imgState_Click()
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

Private Sub optExt_Click(Index As Integer)
    If txtDisplay(1).Text <> vbNullString Then
        cmdSort_Click
    End If
End Sub

Private Sub chkCaps_Click()
    If txtDisplay(1).Text <> vbNullString Then
        cmdSort_Click
    End If
End Sub

Private Sub chkDesc_Click()
    If txtDisplay(1).Text <> vbNullString Then
        cmdSort_Click
    End If
End Sub
