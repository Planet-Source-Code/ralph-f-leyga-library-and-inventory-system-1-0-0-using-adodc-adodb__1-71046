VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmborrowbook 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Book Borrowed"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmborrowbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8760
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmborrowbook.frx":1762
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmborrowbook.frx":79FC
      Left            =   4080
      List            =   "frmborrowbook.frx":7A0C
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   12
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7680
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Out Books"
      TabPicture(0)   =   "frmborrowbook.frx":7A35
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "xpcmdbutton5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "xpcmdbutton4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "xpcmdbutton3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "xpcmdbutton2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ListView1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "xpcmdbutton1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin Libsys1.xpcmdbutton xpcmdbutton1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   5160
         Width           =   1455
         _extentx        =   2566
         _extenty        =   661
         caption         =   "New Out Book"
         font            =   "frmborrowbook.frx":7A51
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Out Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Borrower Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Book borrowed"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Issue"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Due"
            Object.Width           =   2540
         EndProperty
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton2 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   5160
         Width           =   1455
         _extentx        =   2566
         _extenty        =   661
         caption         =   "Modify Out Book"
         font            =   "frmborrowbook.frx":7A79
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton3 
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   5160
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         caption         =   "Remove Out book"
         font            =   "frmborrowbook.frx":7AA1
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton4 
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   5160
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "Print"
         font            =   "frmborrowbook.frx":7AC9
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton5 
         Height          =   375
         Left            =   6240
         TabIndex        =   9
         Top             =   5160
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "Cancel"
         font            =   "frmborrowbook.frx":7AF1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   8640
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   240
         X2              =   8640
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Text:"
         Height          =   255
         Left            =   5160
         TabIndex        =   4
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Search by:"
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Total Out Books:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   4680
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmborrowbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.Text <> "" Then
Text3.Enabled = True
End If
End Sub

Private Sub Command1_Click()

dbase

ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from qborrowbook where " & Combo1.Text & " like '" & Text3.Text & "%' order by title"
        
        .Open criteria, db, 3, 3
        
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !outbooknumber, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !fullname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !Title
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !dateissue
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !datedue
            
            .MoveNext
            
            Loop
            
        .Close
        
    End With
    
  Set rs = Nothing

End Sub

Private Sub Form_Load()

outbooks

Set rs = New ADODB.Recordset

rs.Open "Select count(outbooknumber) as totalcount from tborrowbook", db, 3, 3

Text2.Text = rs!totalcount

Set rs = Nothing


End Sub

Private Sub ListView1_Click()

Text1.Text = ListView1.SelectedItem.Text

End Sub

Private Sub Text3_Change()
Command1_Click
End Sub

Private Sub Timer1_Timer()

Form_Load

Text1.Text = ""

Timer1.Enabled = False


End Sub

Private Sub xpcmdbutton1_Click()

frmnewoutbook.Show vbModal

End Sub

Private Sub xpcmdbutton2_Click()

If Text1.Text <> "" Then

frmmodifyoutbook.Show vbModal

Else

MsgBox "Select a data!", vbInfomation

End If

End Sub

Private Sub xpcmdbutton3_Click()

On Error GoTo error

If Text1.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text1.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tborrowbook where outbooknumber=" & Text1.Text & "", db, 3, 3

MsgBox "Data is remove.", vbInformation

Set rs = Nothing

Timer1.Enabled = True

Text1.Text = ""

End If

Else

MsgBox "No Information Selected!", vbExclamation

End If

Exit Sub

error:

        MsgBox "No Active Record!", vbExclamation


End Sub

Private Sub xpcmdbutton4_Click()

Set rs = New ADODB.Recordset

rs.Open "Select * from qborrowbook", db, 3, 3

Set DataReport3.DataSource = rs

DataReport3.Show vbModal

End Sub

Private Sub xpcmdbutton5_Click()

Unload Me

End Sub

Public Sub outbooks()

On Error Resume Next

ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from qborrowbook order by fullname asc"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !outbooknumber, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !fullname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !Title
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !dateissue
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !datedue
            
            .MoveNext
            
            Loop
                                    
        .Close
        
    End With
    
Set rs = Nothing

End Sub

