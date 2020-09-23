VERSION 5.00
Begin VB.Form frmfine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fine"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   -840
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Save"
      font            =   "frmfine.frx":1762
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Fine:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmfine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set rs = New ADODB.Recordset

rs.Open "Select * from tfine", db, 3, 3

Text1.Text = rs!fine

Text2.Text = rs!fine

Set rs = Nothing

End Sub

Private Sub xpcmdbutton1_Click()

Set rs = New ADODB.Recordset

rs.Open "Select * from tfine where fine ='" & Text2.Text & "'", db, 3, 3

With rs
        
        .Fields("fine") = Text1.Text
        
        .Update
        
End With

Set rs = Nothing

MsgBox "Fine is save!", vbInformation

Unload Me

End Sub
