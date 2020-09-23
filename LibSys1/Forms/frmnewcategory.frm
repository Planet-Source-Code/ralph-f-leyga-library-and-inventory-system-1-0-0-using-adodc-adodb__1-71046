VERSION 5.00
Begin VB.Form frmnewcategory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Category"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4965
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
   ScaleHeight     =   1575
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Save"
      font            =   "frmnewcategory.frx":0000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Cancel"
      font            =   "frmnewcategory.frx":0028
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Category:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmnewcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub xpcmdbutton1_Click()

If Text1.Text <> "" And Text2.Text <> "" Then

Set rs = New ADODB.Recordset

rs.Open "Select * from tcategory", db, 3, 3

With rs
        .AddNew
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
