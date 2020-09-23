VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnewreturn 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Return Books"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4830
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
   ScaleHeight     =   4590
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tfine 
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   4080
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Save"
      font            =   "frmnewreturn.frx":0000
   End
   Begin VB.TextBox tday 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker tdatereturn 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   33023
      Format          =   51052545
      CurrentDate     =   39502
   End
   Begin VB.TextBox tdatedue 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox tdateissue 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox ttitle 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox tfullname 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   4080
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Cancel"
      font            =   "frmnewreturn.frx":0028
   End
   Begin VB.Label Label8 
      Caption         =   "Fine:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4560
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   240
      X2              =   4560
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label7 
      Caption         =   "Days of Before Due:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Date Return:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Date Due:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Date Issue:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Book Title:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Borrower Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Select Out Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmnewreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()

Set rs = New ADODB.Recordset

rs.Open "Select * from qborrowbook where outbooknumber=" & Combo1.Text & "", db, 3, 3

tfullname.Text = rs!fullname

ttitle.Text = rs!Title

tdateissue.Text = rs!dateissue

tdatedue.Text = rs!datedue

Set rs = Nothing

End Sub

Private Sub Form_Load()

cbo

Set rs = New ADODB.Recordset

rs.Open "Select * from tfine", db, 3, 3

Text1.Text = rs!fine

Set rs = Nothing

End Sub

Private Sub tday_Change()

tfine.Text = Val(tday.Text) * Val(Text1.Text)

End Sub

Private Sub xpcmdbutton1_Click()
If Combo1.Text <> "" And tfine.Text <> "" Then
Set rs = New ADODB.Recordset
rs.Open "Select * from treturn", db, 3, 3
With rs
        .AddNew
        .Fields("fullname") = tfullname.Text
        .Fields("title") = ttitle.Text
        .Fields("dateissue") = tdateissue.Text
        .Fields("datedue") = tdatedue.Text
        .Fields("datereturn") = tdatereturn.Value
        .Fields("fine") = tfine.Text
        .Update
End With
Set rs = Nothing
Set rs = New ADODB.Recordset
rs.Open "Delete * from tborrowbook where outbooknumber=" & Combo1.Text & "", db, 3, 3
Set rs = Nothing
frmreturnbook.Timer1.Enabled = True
MsgBox "The book is Successfully Return", vbInformation
Unload Me
Else
MsgBox "All fields are required!", vbInformation
End If

End Sub

Private Sub xpcmdbutton2_Click()

Unload Me

End Sub

Public Sub cbo()

Set rs = New ADODB.Recordset

Combo1.Clear

rs.Open "Select * from tborrowbook order by outbooknumber asc", db, 3, 3

    Do Until rs.EOF
    
        Combo1.AddItem rs!outbooknumber
        
        rs.MoveNext
        
    Loop
    
    Set rs = Nothing
    
End Sub
