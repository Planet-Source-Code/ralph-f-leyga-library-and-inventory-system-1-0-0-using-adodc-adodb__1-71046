VERSION 5.00
Begin VB.Form frmnewborrower 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Borrower"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   5085
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
   ScaleHeight     =   2955
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tlastname 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox tfirstname 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox tmiddlename 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox tage 
      Height          =   285
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox tgender 
      Height          =   315
      ItemData        =   "frmnewborrower.frx":0000
      Left            =   1560
      List            =   "frmnewborrower.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox tyearandsection 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "&Save"
      font            =   "frmnewborrower.frx":0014
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   2520
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "&Cancel"
      font            =   "frmnewborrower.frx":003C
   End
   Begin VB.Label Label1 
      Caption         =   "Lastname:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Firstname:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Middlename:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Age:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Gender:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Year and Section:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frmnewborrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub xpcmdbutton1_Click()

If tlastname.Text <> "" And tfirstname.Text <> "" And tmiddlename.Text <> "" And tage.Text <> "" And tgender.Text <> "" And tyearandsection.Text <> "" Then
Set rs = New ADODB.Recordset
rs.Open "Select * from tborrowers", db, 3, 3
With rs
        .AddNew
        .Fields("lastname") = tlastname.Text
        .Fields("firstname") = tfirstname.Text
        .Fields("middlename") = tmiddlename.Text
        .Fields("age") = tage.Text
        .Fields("gender") = tgender.Text
        .Fields("yearandsec") = tyearandsection.Text
        .Update
End With
Set rs = Nothing
MsgBox "Borrower is Added!", vbInformation
frmBorrower.Timer1.Enabled = True
Unload Me
Else
MsgBox "All fields are required!", vbExclamation
End If
End Sub

Private Sub xpcmdbutton2_Click()
Unload Me
End Sub
