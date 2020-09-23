VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmreturnbook 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Return Book"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8655
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
   ScaleHeight     =   5565
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   0
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
            Picture         =   "frmreturnbook.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Return Book"
      TabPicture(0)   =   "frmreturnbook.frx":5C22
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "xpcmdbutton5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "xpcmdbutton4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "xpcmdbutton1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin Libsys1.xpcmdbutton xpcmdbutton1 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   4440
         Width           =   735
         _extentx        =   1296
         _extenty        =   661
         caption         =   "New"
         font            =   "frmreturnbook.frx":5C3E
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   3960
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5953
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Return Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Book Title"
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
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date Return"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fine"
            Object.Width           =   2540
         EndProperty
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton4 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   4440
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "Print"
         font            =   "frmreturnbook.frx":5C66
      End
      Begin Libsys1.xpcmdbutton xpcmdbutton5 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   4440
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "Cancel"
         font            =   "frmreturnbook.frx":5C8E
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   7920
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   240
         X2              =   7920
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label1 
         Caption         =   "Total Book Return This Day:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   3960
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmreturnbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

returnlist
Set rs = New ADODB.Recordset

rs.Open "Select count(ReturnNumber) as totalcount from treturn", db, 3, 3

Text1.Text = rs!totalcount

Set rs = Nothing
End Sub

Private Sub Timer1_Timer()
Form_Load
Timer1.Enabled = False
End Sub

Private Sub xpcmdbutton1_Click()

frmnewreturn.Show vbModal

End Sub

Private Sub xpcmdbutton4_Click()

Set rs = New ADODB.Recordset

rs.Open "Select * from treturn", db, 3, 3

Set DataReport4.DataSource = rs

DataReport4.Show vbModal


End Sub

Private Sub xpcmdbutton5_Click()

Unload Me

End Sub

Public Sub returnlist()

'On Error Resume Next

ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from treturn order by fullname asc"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !returnnumber, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !fullname
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !Title
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = !dateissue
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = !datedue
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = !datereturn
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = !fine
            
            .MoveNext
            
            Loop
            
                
                        
        .Close
        
    End With
    
Set rs = Nothing

End Sub
