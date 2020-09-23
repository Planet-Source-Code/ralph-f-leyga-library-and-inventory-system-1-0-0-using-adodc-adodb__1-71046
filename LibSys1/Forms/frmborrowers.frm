VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBorrower 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Borrower General Information"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmborrowers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   -240
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
            Picture         =   "frmborrowers.frx":1762
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Borrowers Information"
      TabPicture(0)   =   "frmborrowers.frx":7384
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "xpcmdbutton6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "xpcmdbutton1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "xpcmdbutton5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "xpcmdbutton4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "xpcmdbutton2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListView1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   5760
         TabIndex        =   14
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmborrowers.frx":73A0
         Left            =   3480
         List            =   "frmborrowers.frx":73B0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   5640
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   12
         Top             =   5640
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6720
         TabIndex        =   10
         Top             =   6120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   270
         TabIndex        =   1
         Top             =   570
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8705
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
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Borrower's Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Age"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Gender"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Year And Section"
            Object.Width           =   2540
         EndProperty
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton2 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   6120
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Remove"
         font            =   "frmborrowers.frx":73D4
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton4 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   6120
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Print"
         font            =   "frmborrowers.frx":73FC
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton5 
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   6120
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Cancel"
         font            =   "frmborrowers.frx":7424
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton1 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   6120
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Modify"
         font            =   "frmborrowers.frx":744C
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton6 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   6120
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&New"
         font            =   "frmborrowers.frx":7474
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Text:"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Search by:"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Total Info:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   5640
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   270
         X2              =   8550
         Y1              =   5970
         Y2              =   5970
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   270
         X2              =   8550
         Y1              =   5970
         Y2              =   5970
      End
   End
End
Attribute VB_Name = "frmBorrower"
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
'On Error Resume Next
ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from qborrower where " & Combo1.Text & " like '" & Text3.Text & "%' order by bname"
        
        .Open criteria, db, 3, 3
        
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !borrowerid, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !bname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !age
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !gender
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !yearandsec
            
            .MoveNext
            
            Loop
            
        .Close
        
    End With
    
  Set rs = Nothing

End Sub

Private Sub Form_Load()

borrowerslist

Set rs = New ADODB.Recordset

rs.Open "Select count(borrowerid) as totalcount from tborrowers", db, 3, 3

Text2.Text = rs!totalcount

Set rs = Nothing

End Sub
Public Sub borrowerslist()

On Error Resume Next

ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from qborrower order by bname asc"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !borrowerid, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !bname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !age
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !gender
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !yearandsec
            
            
            .MoveNext
            
            Loop
            
                
                        
        .Close
        
    End With
    
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

If Text1.Text <> "" Then

frmmodifyborrower.Show vbModal

Else

MsgBox "Select a information", vbInformation

End If

End Sub

Private Sub xpcmdbutton2_Click()

On Error GoTo error

If Text1.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text1.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tborrowers where borrowerid=" & Text1.Text & "", db, 3, 3

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
rs.Open "Select * from qborrower", db, 3, 3
Set DataReport1.DataSource = rs
DataReport1.Show vbModal
End Sub

Private Sub xpcmdbutton5_Click()
Unload Me
End Sub

Private Sub xpcmdbutton6_Click()
frmnewborrower.Show vbModal
End Sub
