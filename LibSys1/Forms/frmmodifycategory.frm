VERSION 5.00
Begin VB.Form frmmodifycategory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modify Category"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   Icon            =   "frmmodifycategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Save"
      font            =   "frmmodifycategory.frx":1762
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Cancel"
      font            =   "frmmodifycategory.frx":178A
   End
   Begin VB.Label Label1 
      Caption         =   "Category:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmmodifycategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set rs = New ADODB.Recordset

rs.Open "Select * from tcategory where category='" & frmcategory.Text1.Text & "'", db, 3, 3

Text1.Text = rs!category

Text2.Text = rs!Description

Set rs = Nothing




End Sub

Private Sub xpcmdbutton1_Click()

If Text1.Text <> "" And Text2.Text <> "" Then

Set rs = New ADODB.Recordset

rs.Open "Select * from tcategory where category='" & frmcategory.Text1.Text & "'", db, 3, 3

With rs

        .Fields("category") = Text1.Text
        .Fields("description") = Text2.Text
        .Update
        
End With

Set rs = Nothing

MsgBox "Fine is save!", vbInformation

frmcategory.Timer1.Enabled = True

Unload Me

Else

MsgBox "All fields are required!", vbInformation

End If

End Sub

Private Sub xpcmdbutton2_Click()

Unload Me

End Sub
