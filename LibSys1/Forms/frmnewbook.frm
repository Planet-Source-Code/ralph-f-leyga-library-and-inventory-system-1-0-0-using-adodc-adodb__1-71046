VERSION 5.00
Begin VB.Form frmnewbook 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Book"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmnewbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tPrice 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox tyearpublish 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox tcategory 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox tauthor 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox tTitle 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   240
      Width           =   4095
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2760
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "&Save"
      font            =   "frmnewbook.frx":1762
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2760
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "&Cancel"
      font            =   "frmnewbook.frx":178A
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   5400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label5 
      Caption         =   "Price:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Year Publish:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Author:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Title:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmnewbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cbo
End Sub

Private Sub xpcmdbutton1_Click()

If ttitle.Text <> "" And tauthor.Text <> "" And tcategory.Text <> "" And tyearpublish.Text <> "" And tPrice.Text <> "" Then
Set rs = New ADODB.Recordset
rs.Open "Select * from tbooks", db, 3, 3
With rs
        .AddNew
        .Fields("title") = ttitle.Text
        .Fields("author") = tauthor.Text
        .Fields("category") = tcategory.Text
        .Fields("yearpublish") = tyearpublish.Text
        .Fields("price") = tPrice.Text
        .Update
End With
Set rs = Nothing
MsgBox "Book is Added!", vbInformation
frmbook.Timer1.Enabled = True
Unload Me
Else
MsgBox "All fields are required!", vbExclamation
End If

End Sub

Private Sub xpcmdbutton2_Click()
Unload Me
End Sub
Public Sub cbo()

Set rs = New ADODB.Recordset

tcategory.Clear

rs.Open "Select * from tcategory order by category asc", db, 3, 3

    Do Until rs.EOF
    
        tcategory.AddItem rs!category
        
        rs.MoveNext
        
    Loop
    
    Set rs = Nothing
End Sub
