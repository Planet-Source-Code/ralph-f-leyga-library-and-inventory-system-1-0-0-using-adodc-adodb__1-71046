VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcategory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Category"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmcategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   2040
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   3600
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
            Picture         =   "frmcategory.frx":1762
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   735
      _extentx        =   1296
      _extenty        =   661
      caption         =   "New"
      font            =   "frmcategory.frx":79FC
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   8820
      EndProperty
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3360
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Modify"
      font            =   "frmcategory.frx":7A24
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton3 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Remove"
      font            =   "frmcategory.frx":7A4C
   End
   Begin Libsys1.xpcmdbutton xpcmdbutton4 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3360
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "Cancel"
      font            =   "frmcategory.frx":7A74
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7200
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
categorylist
End Sub

Private Sub ListView1_Click()

Text1.Text = ListView1.SelectedItem.Text

End Sub

Private Sub Timer1_Timer()

Form_Load

Text1.Text = ""

Timer1.Enabled = False

End Sub

Private Sub xpcmdbutton1_Click()

frmnewcategory.Show vbModal

End Sub

Private Sub xpcmdbutton2_Click()

If Text1.Text <> "" Then

frmmodifycategory.Show vbModal

Else

MsgBox "Select a category!", vbInformation

End If

End Sub

Private Sub xpcmdbutton3_Click()

On Error GoTo error

If Text1.Text <> "" Then

Dim repp As String

repp = MsgBox("Do you want to remove " & Text1.Text & " ?", vbYesNo, "Confirm Delete")

If repp = vbYes Then

Set rs = New ADODB.Recordset

rs.Open "Delete * from tcategory where category=" & Text1.Text & "", db, 3, 3

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

Unload Me

End Sub

Public Sub categorylist()

On Error Resume Next

ListView1.ListItems.Clear

Dim criteria As String

Set rs = New ADODB.Recordset

    With rs
    
        criteria = "Select * from tcategory order by category asc"
        
            .Open criteria, db, 3, 3
                
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !category, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !Description
            
            
            .MoveNext
            
            Loop
            
                
                        
        .Close
        
    End With
    
Set rs = Nothing
End Sub
