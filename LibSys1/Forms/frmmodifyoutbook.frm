VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmodifyoutbook 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modify Out Books"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmmodifyoutbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Save"
      font            =   "frmmodifyoutbook.frx":1762
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   33023
      CalendarTrailingForeColor=   4210752
      Format          =   50855937
      CurrentDate     =   39502
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   33023
      CalendarTrailingForeColor=   4210752
      Format          =   50855937
      CurrentDate     =   39502
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2280
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Cancel"
      font            =   "frmmodifyoutbook.frx":178A
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "Select Book Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label u 
      Caption         =   "Select Borrower's ID :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Date Issue:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Date Due:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   240
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmmodifyoutbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cbo
Set rs = New ADODB.Recordset

rs.Open "Select * from tborrowbook where outbooknumber=" & frmborrowbook.Text1.Text & "", db, 3, 3

Combo1.Text = rs!booknumber

Combo2.Text = rs!borrowerid

DTPicker1.Value = rs!dateissue

DTPicker2.Value = rs!datedue

Set rs = Nothing

End Sub

Private Sub xpcmdbutton1_Click()
On Error GoTo err
If Combo1.Text <> "" And Combo2.Text <> "" Then

Set rs = New ADODB.Recordset

rs.Open "Select * from tborrowbook where outbooknumber=" & frmborrowbook.Text1.Text & "", db, 3, 3

With rs

        .Fields("Booknumber") = Combo1.Text
        
        .Fields("borrowerid") = Combo2.Text
        
        .Fields("dateissue") = DTPicker1.Value
        
        .Fields("datedue") = DTPicker2.Value
        
        .Update
        
End With

Set rs = Nothing

MsgBox "Data is save!", vbInformation

frmborrowbook.Timer1.Enabled = True

Unload Me

Else

MsgBox "All Fields are required.", vbExclamation

End If

Exit Sub

err:

MsgBox "Book is not Available!", vbInformation


End Sub

Private Sub xpcmdbutton2_Click()

Unload Me

End Sub

Public Sub cbo()

Set rs = New ADODB.Recordset

Combo1.Clear

rs.Open "Select * from tbooks order by booknumber asc", db, 3, 3

    Do Until rs.EOF
    
        Combo1.AddItem rs!booknumber
        
        rs.MoveNext
        
    Loop
    
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset

Combo2.Clear

rs.Open "Select * from tborrowers order by borrowerid asc", db, 3, 3

    Do Until rs.EOF
    
        Combo2.AddItem rs!borrowerid
        
        rs.MoveNext
        
    Loop
    
    Set rs = Nothing
End Sub
