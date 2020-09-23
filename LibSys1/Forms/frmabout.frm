VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmabout.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "www.rleyga.phpnet.us"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Text or Call: 09057805663 or E-mail: ralphleyga@yahoo.cpm"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "For Comments or Suggestion:"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Ralph F. Leyga"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2008 by RLeyga Software, Inc."
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmabout.frx":55D1
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

