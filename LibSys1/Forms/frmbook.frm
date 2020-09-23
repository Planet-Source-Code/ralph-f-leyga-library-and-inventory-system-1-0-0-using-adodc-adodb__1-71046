VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmbook 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Book General Information"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   0
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
            Picture         =   "frmbook.frx":1762
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Book General Information"
      TabPicture(0)   =   "frmbook.frx":79FC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
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
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6600
         TabIndex        =   6
         Top             =   5760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   4
         Top             =   5280
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmbook.frx":7A18
         Left            =   3480
         List            =   "frmbook.frx":7A2B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   5640
         TabIndex        =   2
         Top             =   5760
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8281
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Book Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Author"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Year Publish"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton2 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   5760
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Remove"
         font            =   "frmbook.frx":7A5C
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton4 
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   5760
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Print"
         font            =   "frmbook.frx":7A84
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton5 
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   5760
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Cancel"
         font            =   "frmbook.frx":7AAC
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton1 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   5760
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Modify"
         font            =   "frmbook.frx":7AD4
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton6 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   5760
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&New"
         font            =   "frmbook.frx":7AFC
      End
      Begin VB.Label Label1 
         Caption         =   "Total Info:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Search by:"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Text:"
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   5280
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   8880
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   240
         X2              =   8880
         Y1              =   5640
         Y2              =   5640
      End
   End
End
Attribute VB_Name = "frmbook"
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
    
        criteria = "Select * from tbooks where " & Combo1.Text & " like '" & Text3.Text & "%' order by title"
        
        .Open criteria, db, 3, 3
        
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !booknumber, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !Title
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !author
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !category
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !Yearpublish
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = !Price
            
            .MoveNext
            
            Loop
            
        .Close
        
    End With
    
  Set rs = Nothing
  
End Sub

Private Sub Form_Load()

booklist

Set rs = New ADODB.Recordset

rs.Open "Select count(booknumber) as totalcount from tbooks", db, 3, 3

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

If Text1.Text <> "" Then

frmmodifybook.Show vbModal

Else

MsgBox "Select a book!", vbInformation

End If

End Sub

Private Sub xpcmdbutton2_Click()

On Error GoTo error

If Text1.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text1.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tbooks where booknumber=" & Text1.Text & "", db, 3, 3

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
rs.Open "Select * from tbooks", db, 3, 3
Set DataReport2.DataSource = rs
DataReport2.Show vbModal

End Sub

Private Sub xpcmdbutton5_Click()

Unload Me

End Sub

Public Sub booklist()

On Error Resume Next

ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from tbooks order by title asc"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !booknumber, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !Title
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !author
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !category
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !Yearpublish
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = !Price
            
            .MoveNext
            
            Loop
            
                
                        
        .Close
        
    End With
    
Set rs = Nothing
End Sub

Private Sub xpcmdbutton6_Click()

frmnewbook.Show vbModal

End Sub
