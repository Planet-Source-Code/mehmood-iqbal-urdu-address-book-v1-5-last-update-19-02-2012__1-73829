Attribute VB_Name = "DB_Con"

'Database Connection variables
Public Conn As Connection
Public RecSource As Recordset

Option Explicit


Sub Connect()

'Connection to Database
Set Conn = New ADODB.Connection
With Conn
  .ConnectionString = App.Path & "\Database.UDB"  'Database file
  .Provider = "Microsoft.Jet.OLEDB.4.0"
  .Open
End With

End Sub

Sub Set_RS()

'Database Perameter Set & Table Selection
Set RecSource = New ADODB.Recordset
  With RecSource
   .ActiveConnection = Conn
   .LockType = adLockOptimistic
   .CursorType = adOpenKeyset
   .Open "Urdu_Address_Book"                 ' Table Name
  End With

End Sub

Sub Set_DB_Fields(FRM As Form)

 'Data Sources (Fields) for Objects
 
 Set FRM.TBX(0).DataSource = RecSource
 FRM.TBX(0).DataField = "FName"      ' First Name
 
 Set FRM.TBX(1).DataSource = RecSource
 FRM.TBX(1).DataField = "LName"      ' Last Name
 
 Set FRM.TBX(2).DataSource = RecSource
 FRM.TBX(2).DataField = "Nick"       ' Nick Name
 
 Set FRM.TBX(3).DataSource = RecSource
 FRM.TBX(3).DataField = "Fa-Name"    ' Father's Name
 
 Set FRM.TBX(4).DataSource = RecSource
 FRM.TBX(4).DataField = "City"       ' City Name
 
 Set FRM.TBX(5).DataSource = RecSource
 FRM.TBX(5).DataField = "Provence"   ' Provence Name
 
 Set FRM.TBX(6).DataSource = RecSource
 FRM.TBX(6).DataField = "Country"    ' Country Name
 
  Set FRM.TBX(7).DataSource = RecSource
 FRM.TBX(7).DataField = "Education"  ' Education
 
 Set FRM.TBX(8).DataSource = RecSource
 FRM.TBX(8).DataField = "Ocupation"  ' Ocupation
 
 Set FRM.Text1.DataSource = RecSource
 FRM.Text1.DataField = "HPhone"        ' Home's Phone Number
 
 Set FRM.Text2.DataSource = RecSource
 FRM.Text2.DataField = "CPhone"        ' Cell Phone Number
 
 Set FRM.Text3.DataSource = RecSource
 FRM.Text3.DataField = "Email"         ' Email Address

End Sub

Sub Reset_DB_Con(FRM As Form)

     'Reset Database Connection as Initial, When Called
     Conn.Close
     DB_Con.Connect
     DB_Con.Set_RS
     DB_Con.Set_DB_Fields FRM
     Initialize.Invisible_TBXS FRM
     Initialize.Visible_LBLS FRM

End Sub
