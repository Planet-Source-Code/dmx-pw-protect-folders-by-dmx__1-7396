VERSION 5.00
Begin VB.Form frmPassWord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password"
   ClientHeight    =   1080
   ClientLeft      =   5970
   ClientTop       =   2295
   ClientWidth     =   2910
   ControlBox      =   0   'False
   Icon            =   "PassWord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program will check to see if a window is open and if the password has been entered
'for the window to stay open.  There are a few problems with this program though
'If the folder is renamed then it will not werk!  I don't know how to protect a folder, but if
'you do then IM me at xMEminemMx or e-mail me the code at DMXXMD44@fcmail.com
'That is about the only problem there is, I think.  I have tested this code to make sure
'everything werks, so there shouldn't be any bugs, but if there is IM me and let me know
'about them, and I will fix them.  Jus press Control+H and Replace -Michael's- with the
'name of your folder that you want to protect and everything should werk fine!
'Also if you like this code please vote for it on PlanetSourceCode, Later DmX

'API Calls
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'Constants
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const WM_CLOSE = &H10
'Variables
Dim A As Long, B As Long, C As Long, G As Long
'Function that tests to see if window is open
Function Test() As Long
A = FindWindow("CabinetWClass", "Michael's")
If A <> 0 Then
Test = 1
Else
Test = 0
End If
End Function
'Tests to see if password is correct or not
Private Sub Command1_Click()
If Text1.Text = "Blue84" Then
    G = 1
    A = FindWindow("CabinetWClass", "Michael's")
    ShowWin A
    Me.Hide
    Text1.Text = ""
Else
    MsgBox "Wrong!!", vbokony, "Incorrect Password"
End If
End Sub
'Closes window that is opened and hides form
Private Sub Command2_Click()
A = FindWindow("CabinetWClass", "Michael's")
B = SendMessageByNum(A, WM_CLOSE, 0, 0)
Me.Hide
End Sub
'Hides program from being visible and tasks running
Private Sub Form_Load()
App.TaskVisible = False
Me.Hide
End Sub
'Check to see if window is open and if password has been entered
Private Sub Timer1_Timer()
C = Test
If C = 1 Then
    If G = 1 Then
        Do
        DoEvents
        C = Test
        If Test = 0 Then G = 0
        Loop Until Test = 0
    Else
        A = FindWindow("CabinetWClass", "Michael's")
        HideWin A
        Me.Show
    End If
End If
End Sub
'Hide Window
Sub HideWin(Window As Long)
    ShowWindow Window, SW_HIDE
End Sub
'Show Window
Sub ShowWin(Window As Long)
    ShowWindow Window, SW_SHOW
End Sub
