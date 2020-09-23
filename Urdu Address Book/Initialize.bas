Attribute VB_Name = "Initialize"

Sub Btn_Captions(FRM As Form)

'Set up Main Form Captions in Urdu
'Frm.CMD1.Caption = ChrW$(&H628) & ChrW$(&H639) & ChrW$(&H62F) & ChrW(&H20) & ChrW$(&H648) & ChrW$(&H627) & ChrW$(&H644) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) & ChrW$(&H688)
'Frm.CMD2.Caption = ChrW$(&H67E) & ChrW$(&H6C1) & ChrW$(&H644) & ChrW$(&H6D2) & ChrW(&H20) & ChrW$(&H648) & ChrW$(&H627) & ChrW$(&H644) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) & ChrW$(&H688)
'Frm.CMD3.Caption = ChrW$(&H633) & ChrW$(&H628) & ChrW(&H20) & ChrW$(&H633) & ChrW$(&H6D2) & ChrW(&H20) & ChrW$(&H622) & ChrW$(&H62E) & ChrW$(&H631) & ChrW$(&H6CC)
'Frm.CMD4.Caption = ChrW$(&H633) & ChrW$(&H628) & ChrW(&H20) & ChrW$(&H633) & ChrW$(&H6D2) & ChrW(&H20) & ChrW$(&H67E) & ChrW$(&H6C1) & ChrW$(&H644) & ChrW$(&H627)
FRM.CMD5.Caption = ChrW$(&H646) & ChrW$(&H6CC) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) & ChrW$(&H688) & ChrW$(&H6D4) & ChrW$(&H6D4) & ChrW$(&H6D4)
FRM.CMD6.Caption = ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) & ChrW$(&H688) & ChrW(&H20) & ChrW$(&H62E) & ChrW$(&H62A) & ChrW$(&H645) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA)
FRM.CMD7.Caption = ChrW$(&H645) & ChrW$(&H62D) & ChrW$(&H641) & ChrW$(&H648) & ChrW$(&H638) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA)
FRM.CMD8.Caption = ChrW$(&H628) & ChrW$(&H646) & ChrW$(&H62F) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA)
FRM.CMD9.Caption = ChrW$(&H62A) & ChrW$(&H628) & ChrW$(&H62F) & ChrW$(&H6CC) & ChrW$(&H644) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA)

End Sub

Sub Lbl_Captions(FRM As Form)

'Set up Main Form captions in Urdu
FRM.Label(0).Caption = ChrW$(&H3A) & ChrW$(&H645) & ChrW$(&H631) & ChrW$(&H6A9) & ChrW$(&H632) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H646) & ChrW$(&H627) & ChrW$(&H645)
FRM.Label(1).Caption = ChrW$(&H3A) & ChrW$(&H622) & ChrW$(&H62E) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H646) & ChrW$(&H627) & ChrW$(&H645)
FRM.Label(2).Caption = ChrW$(&H3A) & ChrW$(&H639) & ChrW$(&H64F) & ChrW$(&H631) & ChrW$(&H641)
FRM.Label(3).Caption = ChrW$(&H3A) & ChrW$(&H648) & ChrW$(&H627) & ChrW$(&H644) & ChrW$(&H62F) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H646) & ChrW$(&H627) & ChrW$(&H645)
FRM.Label(4).Caption = ChrW$(&H3A) & ChrW$(&H634) & ChrW$(&H6C1) & ChrW$(&H631)
FRM.Label(5).Caption = ChrW$(&H3A) & ChrW$(&H635) & ChrW$(&H648) & ChrW$(&H628) & ChrW$(&H6C1)
FRM.Label(6).Caption = ChrW$(&H3A) & ChrW$(&H645) & ChrW$(&H64F) & ChrW$(&H644) & ChrW$(&H6A9)
FRM.Label(7).Caption = ChrW$(&H3A) & ChrW$(&H62A) & ChrW$(&H639) & ChrW$(&H644) & ChrW$(&H6CC) & ChrW$(&H645)
FRM.Label(8).Caption = ChrW$(&H3A) & ChrW$(&H67E) & ChrW$(&H6CC) & ChrW$(&H634) & ChrW$(&H6C1)
FRM.Label(9).Caption = ChrW$(&H3A) & ChrW$(&H6AF) & ChrW$(&H6BE) & ChrW$(&H631) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H641) & ChrW$(&H648) & ChrW$(&H646) & ChrW(&H20) & ChrW$(&H646) & ChrW$(&H645) & ChrW$(&H628) & ChrW$(&H631) & ChrW$(&H2190)
FRM.Label(10).Caption = ChrW$(&H3A) & ChrW$(&H633) & ChrW$(&H64E) & ChrW$(&H6CC) & ChrW$(&H644) & ChrW(&H20) & ChrW$(&H641) & ChrW$(&H648) & ChrW$(&H646) & ChrW(&H20) & ChrW$(&H646) & ChrW$(&H645) & ChrW$(&H628) & ChrW$(&H631)
FRM.Label(11).Caption = ChrW$(&H3A) & ChrW$(&H627) & ChrW$(&H650) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H6CC) & ChrW$(&H644) & ChrW(&H20) & ChrW$(&H627) & ChrW$(&H6CC) & ChrW$(&H688) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H633)

End Sub

Sub CMD_Captions(FRM As Form)

'Set up Message Form's Command Captions
FRM.CMD(0).Caption = ChrW$(&H21) & ChrW(&H20) & ChrW$(&H6C1) & ChrW$(&H627) & ChrW$(&H6BA)
FRM.CMD(1).Caption = ChrW$(&H21) & ChrW(&H20) & ChrW$(&H646) & ChrW$(&H6C1) & ChrW$(&H6CC) & ChrW$(&H6BA)
FRM.CMD(2).Caption = ChrW$(&H21) & ChrW(&H20) & ChrW$(&H679) & ChrW$(&H6BE) & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW(&H20) & ChrW$(&H6C1) & ChrW$(&H6D2)

End Sub

Sub Invisible_TBXS(FRM As Form)

'Make invisible all Textboxes
For i = 0 To 8

       FRM.TBX(i).Visible = False
       
Next i


FRM.Text1.Visible = False
FRM.Text2.Visible = False
FRM.Text3.Visible = False

End Sub

Sub Visible_TBXS(FRM As Form)

'Make invisible all Textboxes
For i = 0 To 8

       FRM.TBX(i).Visible = True
       
Next i


FRM.Text1.Visible = True
FRM.Text2.Visible = True
FRM.Text3.Visible = True


End Sub


Sub Invisible_LBLS(FRM As Form)

'Make invisible Labels, Containing Data
For i = 0 To 11

      FRM.LBL(i).Visible = False
    
      'Also Rectangular shapes behind Labels
      FRM.RECT_SHP(i).Visible = False
      
Next i

End Sub

Sub Visible_LBLS(FRM As Form)

'Make visible Labels, Containing Data
For i = 0 To 11

      FRM.LBL(i).Visible = True
      
      'Also Rectangular shapes behind Labels
      FRM.RECT_SHP(i).Visible = True
      
Next i

End Sub

