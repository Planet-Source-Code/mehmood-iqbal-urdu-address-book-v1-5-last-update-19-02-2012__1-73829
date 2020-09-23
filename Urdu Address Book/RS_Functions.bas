Attribute VB_Name = "RS_Functions"
'RS related variables
Public Abs_Pos As Integer
Public IsLast_Rec As Boolean
Public Rec_Count As Integer

Option Explicit


Sub Add(FRM As Form)

  'One or more Text fields are empty
  If FRM.TBX(0).Text = "" Or FRM.TBX(1).Text = "" Or _
  FRM.TBX(2).Text = "" Or FRM.TBX(3).Text = "" Or _
  FRM.TBX(4).Text = "" Or FRM.TBX(5).Text = "" Or _
  FRM.TBX(6).Text = "" Or FRM.TBX(7).Text = "" Or _
  FRM.TBX(8).Text = "" Or FRM.Text1.Text = "" Or _
  FRM.Text2.Text = "" Or FRM.Text3.Text = "" Then
  
  'Shoe up Error Message
  MSG_BOX.Show 1
  Exit Sub
  
  Else
  
  'All is well, then Add a new record via SQL Statment
  Conn.Execute "Insert Into Urdu_Address_Book Values(" _
  & Rec_Count & ",'" & FRM.TBX(0).Text & "','" _
  & FRM.TBX(1).Text & "','" & FRM.TBX(2).Text & "','" _
  & FRM.TBX(3).Text & "','" & FRM.TBX(4).Text & "','" _
  & FRM.TBX(5).Text & "','" & FRM.TBX(6).Text & "','" _
  & FRM.TBX(7).Text & "','" & FRM.TBX(8).Text & "'," _
  & FRM.Text1.Text & "," & FRM.Text2.Text & ",'" & FRM.Text3.Text & "')"
  
  'After Sucessfull Saving, Do
  FRM.CMD7.Enabled = False
  FRM.CMD1.Enabled = True
  FRM.CMD2.Enabled = True
  FRM.CMD3.Enabled = True
  FRM.CMD4.Enabled = True
  FRM.CMD5.Enabled = True
  FRM.CMD6.Enabled = True
  FRM.CMD9.Enabled = True
  Rec_Count = 0
  
  'Show up Sucess Message
  MSG_BOX.Show 5
  
  End If

End Sub

Sub Update(FRM As Form)

  'One or more Text fields are empty
  If FRM.TBX(0).Text = "" Or FRM.TBX(1).Text = "" Or _
  FRM.TBX(2).Text = "" Or FRM.TBX(3).Text = "" Or _
  FRM.TBX(4).Text = "" Or FRM.TBX(5).Text = "" Or _
  FRM.TBX(6).Text = "" Or FRM.TBX(7).Text = "" Or _
  FRM.TBX(8).Text = "" Or FRM.Text1.Text = "" Or _
  FRM.Text2.Text = "" Or FRM.Text3.Text = "" Then
  MSG_BOX.Show 1
  Else
  
  'All is well, then Update record via SQL Statment
  Conn.Execute "Update Urdu_Address_Book Set Fname='" _
  & FRM.TBX(0).Text & "', LName='" & FRM.TBX(1).Text & _
  "', Nick='" & FRM.TBX(2).Text & "', [Fa-Name]='" _
  & FRM.TBX(3).Text & "', City='" & FRM.TBX(4).Text & "', Provence='" _
  & FRM.TBX(5).Text & "', Country='" & FRM.TBX(6).Text & "', Education='" _
  & FRM.TBX(7).Text & "', Ocupation='" & FRM.TBX(8).Text & "', HPhone=" _
  & FRM.Text1.Text & ", CPhone=" & FRM.Text2.Text & ", Email='" & FRM.Text3.Text _
  & "' Where SN=" & Rec_Count & ";"
  
  'Record Updated Sucessfully Then, Do
  FRM.CMD7.Enabled = False
  FRM.CMD1.Enabled = True
  FRM.CMD2.Enabled = True
  FRM.CMD3.Enabled = True
  FRM.CMD4.Enabled = True
  FRM.CMD5.Enabled = True
  FRM.CMD6.Enabled = True
  FRM.CMD9.Enabled = True
  Rec_Count = 0
  
  'Show up Sucess Message
  MSG_BOX.Show 6
  
  End If

End Sub

Public Sub RecDel(Rec_Number As Integer)

'Set the Absoulute Position
Abs_Pos = Rec_Number

'Delete a Record, with reference got
Conn.Execute "Delete From Urdu_Address_Book Where SN=" & Rec_Number

'Show Sucess message
MSG_BOX.Show 4

End Sub

Sub ReNew_Serial(Del_Case As Integer)

'Reset Serial Number field, to numberize records
Select Case Del_Case

'Deleted record was First one, so start from first
Case 0

      'Requery RecordSet
      RecSource.Requery

      'Setup variables
      Dim A As Integer
      Dim B As String
      Dim C As Integer

      'Setup variable values
      A = RecSource.RecordCount
      B = RecSource.Fields(1)
      C = 1

      'Setup Serial from 1 to Last, usind Field(1) data reference
      Do Until A = 0
    
          Conn.Execute "Update Urdu_Address_Book Set SN=" & C & " Where FName='" & B & "';"
          A = A - 1
          C = C + 1
          RecSource.MoveNext
    
          'Check if End Of File (EOF) True
          If RecSource.EOF = True Then
           
          'Move to first & stop Serial
          RecSource.MoveFirst
          Exit Sub
    
          End If
    
          B = RecSource.Fields(1)
    
      Loop

'Deleted record was from Mid, So start Serial from that
Case 1

      'Requery RecordSet
      RecSource.Requery
      
      'Setup variables
      Dim j As Integer
      Dim K As String
      Dim L As Integer
    
      'Move to Last deleted record position
      RecSource.Move (Abs_Pos - 1), 1
      
      'Setup variable's values
      j = RecSource.RecordCount
      K = RecSource.Fields(1)
      L = Abs_Pos
      
      'Setup Serial next from Deleted record
      Do Until j = 0
    
          Conn.Execute "Update Urdu_Address_Book Set SN=" & L & " Where FName='" & K & "';"
          j = j - 1
          L = L + 1
          RecSource.MoveNext
    
          'Check if End of File (EOF) True
          If RecSource.EOF = True Then
    
          'Move to first & stop Serial
          RecSource.MoveFirst
          Exit Sub
    
          End If
    
          K = RecSource.Fields(1)
    
      Loop


End Select

End Sub

Sub Process_Deletation()

'Start processing Delete
Select Case Rec_Count

Case 0

      'Delete Record
      RS_Functions.RecDel RecSource.Fields(0)
      
      'But check if that was Last Number record
      Check_IsLast
      
      'Select a Case if that record was Last or Another
      Select Case IsLast_Rec
      
      Case True
            
            'That was Last Number record
            Exit Sub
      
      Case False
      
            'Deleted record was not last Number, Was before last
            If Abs_Pos = 1 Then
             
                      'If that was First One
                      RS_Functions.ReNew_Serial 0
                      
            Else
             
                      'If that was not First or Last, that was from Mid
                      RS_Functions.ReNew_Serial 1
            
            End If

      End Select

End Select


End Sub

Sub Check_IsLast()

'Check if Deleted record was Last Number

If Abs_Pos = RecSource.RecordCount Then

        'Was Last Number record
        IsLast_Rec = True
    
Else

        'Was not Last One
        IsLast_Rec = False

End If


End Sub
