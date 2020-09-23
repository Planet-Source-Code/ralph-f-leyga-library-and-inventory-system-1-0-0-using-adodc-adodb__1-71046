Attribute VB_Name = "Module1"
'==========Developed by: Ralph F. Leyga============
'visit: www.rleyga.phpnet.us
'e-mail: ralphleyga@yahoo.com
'text or call: 09057805663


Public db As New ADODB.Connection

Public rs As New ADODB.Recordset

Public bol As Boolean

Public Sub dbase()

Set db = New ADODB.Connection

  db.CursorLocation = adUseClient
  
  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= Dbase.mdb ;Persist Security Info=False;Jet OLEDB"
  
End Sub

