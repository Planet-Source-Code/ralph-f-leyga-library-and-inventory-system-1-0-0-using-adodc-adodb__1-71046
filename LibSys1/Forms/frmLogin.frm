VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1275
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   753.312
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Libsys1.xpcmdbutton cmdok 
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      caption         =   "Log In"
      font            =   "frmLogin.frx":1762
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1050
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   165
      Width           =   2565
   End
   Begin Libsys1.xpcmdbutton cmdcancel 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      caption         =   "Cancel"
      font            =   "frmLogin.frx":178A
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    End
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtPassword = "password" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
