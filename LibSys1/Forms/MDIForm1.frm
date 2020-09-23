VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   "Library System Version 1.0"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7095
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":1762
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1800
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":189A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E5C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2A054
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C6CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D0E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7095
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   33023
         CalendarTrailingForeColor=   12640511
         Format          =   20381696
         CurrentDate     =   39502
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6030
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "MDIForm1.frx":345E2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Time Log-in:"
            TextSave        =   "Time Log-in:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "4:40 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MDIForm1.frx":3497C
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "6/27/2001"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File Maintenance"
      Begin VB.Menu borrowersinformation 
         Caption         =   "&Borrowers Information"
         Shortcut        =   {F1}
      End
      Begin VB.Menu bookinformation 
         Caption         =   "&Books Information"
         Shortcut        =   {F2}
      End
      Begin VB.Menu borrowedbooks 
         Caption         =   "&Borrowed Books"
         Shortcut        =   {F3}
      End
      Begin VB.Menu returnbook 
         Caption         =   "&Returned Book"
         Shortcut        =   {F4}
      End
      Begin VB.Menu bar01 
         Caption         =   "-"
      End
      Begin VB.Menu closeme 
         Caption         =   "&Close Application"
      End
   End
   Begin VB.Menu settings 
      Caption         =   "&Settings"
      Begin VB.Menu category 
         Caption         =   "&Category"
         Shortcut        =   {F5}
      End
      Begin VB.Menu fine 
         Caption         =   "&Fine"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu about 
      Caption         =   "&About"
      Begin VB.Menu bar02 
         Caption         =   "-"
      End
      Begin VB.Menu dev 
         Caption         =   "Developer"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bookinformation_Click()
frmbook.Show vbModal
End Sub

Private Sub borrowedbooks_Click()
frmborrowbook.Show vbModal
End Sub

Private Sub borrowersinformation_Click()
frmBorrower.Show vbModal
End Sub

Private Sub category_Click()
frmcategory.Show vbModal
End Sub

Private Sub closeme_Click()
End
End Sub

Private Sub dev_Click()
frmabout.Show vbModal
End Sub

Private Sub fine_Click()
frmfine.Show vbModal
End Sub

Private Sub MDIForm_Load()
dbase
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub returnbook_Click()
frmreturnbook.Show vbModal
End Sub

Private Sub Timer1_Timer()
frmLogin.Show vbModal
Timer1.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1:
        borrowersinformation_Click
Case 2:
        bookinformation_Click
Case 3:
        borrowedbooks_Click
Case 4:
        returnbook_Click
Case 5:
       category_Click
Case 6:
        fine_Click

End Select

End Sub
