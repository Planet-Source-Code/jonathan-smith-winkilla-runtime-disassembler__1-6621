VERSION 5.00
Begin VB.Form frmWinKilla 
   Caption         =   "WîÑkÎLLå"
   ClientHeight    =   6765
   ClientLeft      =   3390
   ClientTop       =   1200
   ClientWidth     =   4815
   Icon            =   "frmWinKilla.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   4815
   Begin VB.Frame Frame3 
      Caption         =   "Display Text"
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   4455
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Text"
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtOut 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Coordinates:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Print Text on Window:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Window Title"
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   4455
      Begin VB.TextBox txtNewText 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdChangeText 
         Caption         =   "Change Text"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Set Window Text:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.ListBox lstWins 
      Height          =   1035
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Winshot"
      Height          =   1335
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
      Begin VB.OptionButton optDestroy 
         Caption         =   "Destroy Processes"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optHide 
         Caption         =   "Hide Window"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optLoad 
         Caption         =   "Load Window"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optKill 
         Caption         =   "Kill Window"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Make It So"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtHwnd 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Window List"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Programmed by ULTiMaTuM"
      BeginProperty Font 
         Name            =   "Acidic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Top             =   6360
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "hWnd of window to kill (handle):"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2265
   End
End
Attribute VB_Name = "frmWinKilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WîÑkÎLLå v1.0.7
'programmed by ULTiMaTuM (a.k.a. Jon)
'copyright © 1999, all rights reserved
'
'This source code and project is distributed
'for educational purposes only. If you plan on
'using this code in an application, first
'contact me for any questions you might have.
'I don't care if you don't include me in the
'credits, just tell me thanks or something.

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal lSw As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
'NT-compatible function
'Use this ONLY if you are running Windows NT
Private Declare Function NTPostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_QUIT = &H12

Private Sub cmdChangeColor_Click()
    'I took this subroutine out because it
    'wasn't doing what I wanted it to do
    Dim hWndVal As Long
    Dim hDCVal As Long
    Trim txtHwnd.Text
    If txtHwnd.Text = "" Then
        MsgBox "Tu es tonto! You forgot to enter an hWnd value!", vbCritical, "OH MY GOD!!!"
        Exit Sub
    End If
    hWndVal = txtHwnd.Text
    hDCVal = GetDC(hWndVal)
    x = SetBkColor(hDCVal, picColor.BackColor)
        
End Sub

Private Sub cmdChangeText_Click()
    Dim hWndVal As Long
    Trim txtHwnd.Text
    If txtHwnd.Text = "" Then
        MsgBox "Tu es tonto! You forgot to enter an hWnd value!", vbCritical, "OH MY GOD!!!"
        Exit Sub
    End If
    hWndVal = txtHwnd.Text
    x = SetWindowText(hWndVal, txtNewText.Text)
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
    
End Sub

Private Sub cmdOk_Click()
    Dim hWndVal As Long
    
    Const SW_SHOW = 5
    Const SW_HIDE = 0
    
    Trim txtHwnd.Text
    If txtHwnd.Text = "" Then
        MsgBox "Tu es tonto! You forgot to enter an hWnd value!", vbCritical, "OH MY GOD!!!"
        Exit Sub
    End If
    hWndVal = Val(txtHwnd.Text)
    If optKill.Value = True Then
        x = DestroyWindow(hWndVal)
    ElseIf optHide.Value = True Then
        x = ShowWindow(hWndVal, SW_HIDE)
    ElseIf optLoad.Value = True Then
        x = ShowWindow(hWndVal, SW_SHOW)
    Else
        'If you are running Windows NT, use the following:
        'x = NTPostMessage(hWndVal, WM_QUIT, 0, 0&)
        x = PostMessage(hWndVal, WM_QUIT, 0, 0&)
    End If
    
End Sub

Private Sub cmdPrint_Click()
    Dim hWndVal As Long
    Dim hDCVal As Long
    Dim xVal As Long
    Dim yVal As Long
    Dim size As Long
    
    Trim txtHwnd.Text
    If txtHwnd.Text = "" Then
        MsgBox "Tu es tonto! You forgot to enter an hWnd value!", vbCritical, "OH MY GOD!!!"
        Exit Sub
    End If
    hWndVal = Val(txtHwnd.Text)
    xVal = Val(txtX.Text)
    yVal = Val(txtY.Text)
    size = Len(txtOut.Text)
    hDCVal = GetDC(hWndVal)
    x = TextOut(hDCVal, xVal, yVal, txtOut.Text, size)
    
End Sub

Private Sub cmdRefresh_Click()
    x = FillTaskListBox(lstWins)
    
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
    'picColor.BackColor = cdlColor.Color
    
End Sub

Private Sub lstWins_Click()
    txtHwnd.Text = lstWins.ItemData(lstWins.ListIndex)
    
End Sub

'Some more obsolete functions from version 1.0.0
''''''''''''''''''''''''''''''''''''''''''''''''

'Private Sub TaskTimer_Timer()
'    Windows = FillTaskListBox(lstWins)
'
'End Sub
'
'Private Sub picColor_Click()
'    On Error Resume Next
'    cdlColor.DialogTitle = "Color Selection"
'    cdlColor.ShowColor
'    If Err = 0 Then
'        picColor.BackColor = cdlColor.Color
'    Else
'        Exit Sub
'    End If
'
'End Sub
'
'Private Sub txtBlue_Change()
'    R = Val(txtRed.Text)
'    G = Val(txtGreen.Text)
'    B = Val(txtBlue.Text)
'    picColor.BackColor = RGB(R, G, B)
'End Sub
'
'Private Sub txtGreen_Change()
'    R = Val(txtRed.Text)
'    G = Val(txtGreen.Text)
'    B = Val(txtBlue.Text)
'    picColor.BackColor = RGB(R, G, B)
'End Sub
'
'Private Sub txtRed_Change()
'    R = Val(txtRed.Text)
'    G = Val(txtGreen.Text)
'    B = Val(txtBlue.Text)
'    picColor.BackColor = RGB(R, G, B)
'
'End Sub
