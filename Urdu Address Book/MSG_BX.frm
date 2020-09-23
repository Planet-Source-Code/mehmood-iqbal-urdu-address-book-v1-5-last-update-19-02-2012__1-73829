VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form MSG_BX 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1695
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5550
   Icon            =   "MSG_BX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Urdu_Address_Book.jcbutton CMD 
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jameel Noori Nastaleeq"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Yes"
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD 
      Height          =   495
      Index           =   1
      Left            =   1088
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jameel Noori Nastaleeq"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "No"
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD 
      Height          =   495
      Index           =   2
      Left            =   2025
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jameel Noori Nastaleeq"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "OK"
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      ColorScheme     =   2
   End
   Begin VB.Image IMG 
      Height          =   480
      Index           =   5
      Left            =   4800
      Picture         =   "MSG_BX.frx":0A02
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IMG 
      Height          =   480
      Index           =   4
      Left            =   4800
      Picture         =   "MSG_BX.frx":0F78
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IMG 
      Height          =   480
      Index           =   3
      Left            =   4800
      Picture         =   "MSG_BX.frx":15DA
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IMG 
      Height          =   480
      Index           =   2
      Left            =   4800
      Picture         =   "MSG_BX.frx":1C84
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IMG 
      Height          =   480
      Index           =   1
      Left            =   4800
      Picture         =   "MSG_BX.frx":2323
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IMG 
      Height          =   480
      Index           =   0
      Left            =   4800
      Picture         =   "MSG_BX.frx":2930
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSForms.Label LBL 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Size            =   "8916;873"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   285
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "MSG_BX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMD_Click(Index As Integer)

'Select a case when Command Button clicked
Select Case Index

'YES Button clicked
Case 0
      
      'Set a Reference Number
      Rec_Count = 0
      
      'Hide Message Box form
      MSG_BX.Hide
      
      'Process Deletation
      RS_Functions.Process_Deletation
      'Exit Sub

'NO Button clicked
Case 1
 
      MSG_BX.Hide

'OK Button clicked
Case 2

      MSG_BX.Hide

End Select

End Sub

Private Sub Form_Activate()

'To play a sound
'Check for the Reference, in which form activated

'Windows Media Player control placed on Main Form So,
'That will be used as containng form
Select Case Ref

Case 1

      Sound.Critical Main_Form

Case 2

      Sound.Critical Main_Form

Case 3

      Sound.Exclamation Main_Form

Case 4

      Sound.Information Main_Form

Case 5

      Sound.Information Main_Form

Case 6

      Sound.Information Main_Form

End Select

End Sub
