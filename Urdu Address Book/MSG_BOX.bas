Attribute VB_Name = "MSG_BOX"

'Runtime variables
Public Ref As Integer
Public Mode As Integer
Dim i As Integer

Option Explicit


Sub Show(Number As Integer)

'Select a case in which  Message orderd to diaplay

Select Case Number

Case 1  'Incomplete Data
  
     'Set Reference & Mode
     Ref = 1
     Mode = 1
     
     'Make all Image boxes Invisible
     For i = 0 To 5
     
          MSG_BX.IMG(i).Visible = False
          
     Next i
    
     'Setup other items
     MSG_BX.IMG(0).Visible = True
     MSG_BX.CMD(0).Visible = False
     MSG_BX.CMD(1).Visible = False
     MSG_BX.CMD(2).Visible = True
     
     'Set an Urdu Message
     MSG_BX.LBL.Caption = ChrW$(&H645) & ChrW$(&H6A9) & ChrW$(&H645) _
                          & ChrW$(&H644) & ChrW(&H20) & ChrW$(&H688) & ChrW$(&H6CC) & _
                          ChrW$(&H679) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H62F) & _
                          ChrW$(&H627) & ChrW$(&H62E) & ChrW$(&H644) & ChrW(&H20) & _
                          ChrW$(&H646) & ChrW$(&H6C1) & ChrW$(&H6CC) & ChrW$(&H6BA) _
                          & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H6CC) & ChrW$(&H627) _
                          & ChrW(&H20) & ChrW$(&H6AF) & ChrW$(&H6CC) & ChrW$(&H627) _
                          & ChrW(&H20) & ChrW$(&H21) & ChrW(&H20) & ChrW$(&H628) & _
                          ChrW$(&H631) & ChrW$(&H627) & ChrW$(&H6C1) & ChrW$(&H650) _
                          & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H6C1) & ChrW$(&H631) _
                          & ChrW$(&H628) & ChrW$(&H627) & ChrW$(&H646) & ChrW$(&H6CC) _
                          & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H6A9) & ChrW$(&H645) & _
                          ChrW$(&H644) & ChrW(&H20) & ChrW$(&H688) & ChrW$(&H6CC) & _
                          ChrW$(&H679) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H62F) & _
                          ChrW$(&H627) & ChrW$(&H62E) & ChrW$(&H644) & ChrW(&H20) & _
                          ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA) & ChrW$(&H6D4)

     MSG_BX.Caption = App.Title
     
     'Show up Message
     MSG_BX.Show vbModal

Case 2    'No record into the Database
    
     'Set Reference only
     Ref = 2
     
     'Make all Image boxes Invisible
     For i = 0 To 5
     
         MSG_BX.IMG(i).Visible = False
          
     Next i
    
     'Setup other items
     MSG_BX.IMG(1).Visible = True
     MSG_BX.CMD(0).Visible = False
     MSG_BX.CMD(1).Visible = False
     MSG_BX.CMD(2).Visible = True

     'Setup an Urdu Message
     MSG_BX.LBL.Caption = ChrW$(&H21) & ChrW(&H20) & ChrW$(&H688) & ChrW$(&H6CC) & ChrW$(&H679) & ChrW$(&H627) _
                          & ChrW(&H20) & ChrW$(&H628) & ChrW$(&H6CC) & ChrW$(&H633) _
                          & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H6CC) & ChrW$(&H6BA) _
                          & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H648) & ChrW$(&H626) _
                          & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H631) & ChrW$(&H6CC) _
                          & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) & ChrW$(&H688) _
                          & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H648) & ChrW$(&H62C) _
                          & ChrW$(&H648) & ChrW$(&H62F) & ChrW(&H20) & ChrW$(&H646) _
                          & ChrW$(&H6C1) & ChrW$(&H6CC) & ChrW$(&H6BA)


     MSG_BX.Caption = App.Title
     
     'Show up Message
     MSG_BX.Show vbModal

Case 3

     'Set Reference, Mode & Rec_Count (To recognize a condition only)
     Ref = 3
     Mode = 2
     Rec_Count = 0
     
     'Make Invisible all Image Boxes
     For i = 0 To 5
     
         MSG_BX.IMG(i).Visible = False
          
     Next i
    
     'Setup other items
     MSG_BX.IMG(2).Visible = True
     MSG_BX.CMD(0).Visible = True
     MSG_BX.CMD(1).Visible = True
     MSG_BX.CMD(2).Visible = False
     
     'Setup an Urdu Message
     MSG_BX.LBL.Caption = ChrW$(&H6A9) & ChrW$(&H6CC) & ChrW$(&H627) & ChrW(&H20) _
                          & ChrW$(&H622) & ChrW$(&H67E) & ChrW(&H20) & ChrW$(&H648) & ChrW$(&H627) _
                          & ChrW$(&H642) & ChrW$(&H639) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H6CC) _
                          & ChrW$(&H6C1) & ChrW(&H20) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6A9) _
                          & ChrW$(&H627) & ChrW$(&H631) & ChrW$(&H688) & ChrW(&H20) & ChrW$(&H688) _
                          & ChrW$(&H6CC) & ChrW$(&H679) & ChrW$(&H627) & ChrW$(&H628) & ChrW$(&H6CC) _
                          & ChrW$(&H633) & ChrW(&H20) & ChrW$(&H633) & ChrW$(&H6D2) & ChrW(&H20) _
                          & ChrW$(&H62E) & ChrW$(&H62A) & ChrW$(&H645) & ChrW(&H20) & ChrW$(&H6A9) _
                          & ChrW$(&H631) & ChrW$(&H646) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H686) _
                          & ChrW$(&H627) & ChrW$(&H6C1) & ChrW$(&H62A) & ChrW$(&H6D2) & ChrW(&H20) _
                          & ChrW$(&H6C1) & ChrW$(&H6CC) & ChrW$(&H6BA) & ChrW$(&H61F)
     
     
     
     MSG_BX.Caption = App.Title
     
     'Show up Message
     MSG_BX.Show vbModal


Case 4

     'Set Reference & Mode
     Ref = 4
     Mode = 2
     
     'Make Invisible all Image Boxes
     For i = 0 To 5
     
         MSG_BX.IMG(i).Visible = False
          
     Next i
    
     'Setup other items
     MSG_BX.IMG(3).Visible = True
     MSG_BX.CMD(0).Visible = False
     MSG_BX.CMD(1).Visible = False
     MSG_BX.CMD(2).Visible = True
     
     'Setup an Urdu Message
     MSG_BX.LBL.Caption = ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) _
                          & ChrW$(&H688) & ChrW(&H20) & ChrW$(&H62E) & ChrW$(&H62A) & ChrW$(&H645) _
                          & ChrW(&H20) & ChrW$(&H6C1) & ChrW$(&H648) & ChrW(&H20) & ChrW$(&H686) _
                          & ChrW$(&H64F) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H6C1) _
                          & ChrW$(&H6D2) & ChrW(&H20) & ChrW$(&H6D4)
     
     
     MSG_BX.Caption = App.Title
     
     'Show up Message
     MSG_BX.Show vbModal
     

Case 5

     'Set Reference & Mode
     Ref = 5
     Mode = 3
     
     'Make invisible all Image Boxes
     For i = 0 To 5
     
         MSG_BX.IMG(i).Visible = False
          
     Next i
    
     'Setup other items
     MSG_BX.IMG(4).Visible = True
     MSG_BX.CMD(0).Visible = False
     MSG_BX.CMD(1).Visible = False
     MSG_BX.CMD(2).Visible = True
     
     'Setup an Urdu Message
     MSG_BX.LBL.Caption = ChrW$(&H646) & ChrW$(&H6CC) & ChrW$(&H627) & ChrW(&H20) & ChrW$(&H631) _
                          & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) & ChrW$(&H688) _
                          & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H645) & ChrW$(&H6CC) _
                          & ChrW$(&H627) & ChrW$(&H628) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H633) _
                          & ChrW$(&H6D2) & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H62D) & ChrW$(&H641) _
                          & ChrW$(&H648) & ChrW$(&H638) & ChrW(&H20) & ChrW$(&H6C1) & ChrW$(&H648) _
                          & ChrW(&H20) & ChrW$(&H686) & ChrW$(&H64F) & ChrW$(&H6A9) & ChrW$(&H627) _
                          & ChrW(&H20) & ChrW$(&H6C1) & ChrW$(&H6D2) & ChrW$(&H6D4)
     
     
     MSG_BX.Caption = App.Title
     
     'Show up Message
     MSG_BX.Show vbModal
     
Case 6

     'Set Reference & Mode
     Ref = 6
     Mode = 2
     
     'Make Invisible all Image Boxes
     For i = 0 To 5
     
         MSG_BX.IMG(i).Visible = False
          
     Next i
    
     'Setup other items
     MSG_BX.IMG(5).Visible = True
     MSG_BX.CMD(0).Visible = False
     MSG_BX.CMD(1).Visible = False
     MSG_BX.CMD(2).Visible = True
     
     'Setup an Urdu Message
     MSG_BX.LBL.Caption = ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6A9) & ChrW$(&H627) & ChrW$(&H631) _
                          & ChrW$(&H688) & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H6CC) & ChrW$(&H6BA) _
                          & ChrW(&H20) & ChrW$(&H62A) & ChrW$(&H628) & ChrW$(&H62F) & ChrW$(&H6CC) _
                          & ChrW$(&H644) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H627) _
                          & ChrW$(&H645) & ChrW$(&H6CC) & ChrW$(&H627) & ChrW$(&H628) & ChrW$(&H6CC) _
                          & ChrW(&H20) & ChrW$(&H633) & ChrW$(&H6D2) & ChrW(&H20) & ChrW$(&H645) _
                          & ChrW$(&H62D) & ChrW$(&H641) & ChrW$(&H648) & ChrW$(&H638) & ChrW(&H20) _
                          & ChrW$(&H6C1) & ChrW$(&H648) & ChrW(&H20) & ChrW$(&H686) & ChrW$(&H6A9) _
                          & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H6C1) & ChrW$(&H6D2) & ChrW$(&H6D4)
     
     
     MSG_BX.Caption = App.Title
     
     'Show up Message
     MSG_BX.Show vbModal

End Select

End Sub
