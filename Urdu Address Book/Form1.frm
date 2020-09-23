VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Main_Form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Urdu Address Book"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Urdu_Address_Book.jcbutton CMD7 
      Height          =   495
      Left            =   4140
      TabIndex        =   12
      Top             =   4275
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   873
      ButtonStyle     =   3
      Enabled         =   0   'False
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
      Caption         =   "Save Record or Save Changes"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD3 
      Height          =   495
      Left            =   1058
      TabIndex        =   11
      Top             =   3765
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "9"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD4 
      Height          =   495
      Left            =   2543
      TabIndex        =   10
      Top             =   3765
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   ":"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD1 
      Height          =   495
      Left            =   1058
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "3"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD2 
      Height          =   495
      Left            =   2550
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "4"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD5 
      Height          =   495
      Left            =   4140
      TabIndex        =   7
      Top             =   3240
      Width           =   2940
      _ExtentX        =   5186
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
      Caption         =   "New Rec"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD9 
      Height          =   495
      Left            =   4140
      TabIndex        =   6
      Top             =   3780
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
      Caption         =   "Edit Rec"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD6 
      Height          =   495
      Left            =   5625
      TabIndex        =   5
      Top             =   3780
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
      Caption         =   "Delete Rec"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Urdu_Address_Book.jcbutton CMD8 
      Height          =   495
      Left            =   1058
      TabIndex        =   4
      Top             =   4275
      Width           =   2940
      _ExtentX        =   5186
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
      Caption         =   "Exit"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   3
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   2295
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4500
      MaxLength       =   15
      TabIndex        =   0
      Top             =   2070
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4500
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2565
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   11
      Left            =   180
      TabIndex        =   45
      Top             =   2295
      Width           =   2385
      Size            =   "4207;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   10
      Left            =   4500
      TabIndex        =   44
      Top             =   2565
      Width           =   2340
      Size            =   "4128;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   9
      Left            =   4500
      TabIndex        =   43
      Top             =   2070
      Width           =   2340
      Size            =   "4128;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   8
      Left            =   405
      TabIndex        =   42
      Top             =   765
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   7
      Left            =   2925
      TabIndex        =   41
      Top             =   765
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   6
      Left            =   405
      TabIndex        =   40
      Top             =   1350
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   6
      Left            =   360
      Top             =   1380
      Width           =   1200
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   5
      Left            =   2970
      TabIndex        =   39
      Top             =   1350
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   4
      Left            =   5715
      TabIndex        =   38
      Top             =   1350
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   3
      Left            =   5670
      TabIndex        =   37
      Top             =   765
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   2
      Left            =   405
      TabIndex        =   36
      Top             =   135
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   1
      Left            =   2970
      TabIndex        =   35
      Top             =   135
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   1
      Left            =   2925
      Top             =   180
      Width           =   1200
   End
   Begin MSForms.Label LBL 
      Height          =   375
      Index           =   0
      Left            =   5625
      TabIndex        =   34
      Top             =   150
      Width           =   1215
      Size            =   "2143;661"
      BorderStyle     =   1
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   225
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   0
      Left            =   5580
      Top             =   180
      Width           =   1200
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP1 
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   11
      Left            =   2640
      TabIndex        =   24
      Top             =   2280
      Width           =   1095
      Size            =   "1931;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   10
      Left            =   6975
      TabIndex        =   23
      Top             =   2565
      Width           =   975
      Size            =   "1720;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   9
      Left            =   6960
      TabIndex        =   22
      Top             =   2040
      Width           =   975
      Size            =   "1720;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   8
      Left            =   1920
      TabIndex        =   21
      Top             =   840
      Width           =   495
      Size            =   "873;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   7
      Left            =   4440
      TabIndex        =   20
      Top             =   840
      Width           =   735
      Size            =   "1296;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   19
      Top             =   1440
      Width           =   495
      Size            =   "873;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   18
      Top             =   1440
      Width           =   735
      Size            =   "1296;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   17
      Top             =   1440
      Width           =   735
      Size            =   "1296;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   3
      Left            =   7200
      TabIndex        =   16
      Top             =   840
      Width           =   735
      Size            =   "1296;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   15
      Top             =   240
      Width           =   495
      Size            =   "873;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   14
      Top             =   240
      Width           =   735
      Size            =   "1296;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label 
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   13
      Top             =   240
      Width           =   735
      Size            =   "1296;661"
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   2
      Left            =   360
      Top             =   165
      Width           =   1200
   End
   Begin MSForms.TextBox TBX 
      Height          =   540
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   90
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;952"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   3
      Left            =   5625
      Top             =   795
      Width           =   1200
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   3
      Left            =   5400
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   4
      Left            =   5670
      Top             =   1380
      Width           =   1200
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   5
      Left            =   2925
      Top             =   1380
      Width           =   1200
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   7
      Left            =   2880
      Top             =   810
      Width           =   1200
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   8
      Left            =   360
      Top             =   810
      Width           =   1200
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   4
      Left            =   5400
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   5
      Left            =   2640
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   7
      Left            =   2640
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox TBX 
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      BorderStyle     =   1
      Size            =   "2990;873"
      BorderColor     =   4194304
      SpecialEffect   =   0
      FontName        =   "Jameel Noori Nastaleeq"
      FontHeight      =   240
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   9
      Left            =   4455
      Top             =   2115
      Width           =   2370
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   10
      Left            =   4455
      Top             =   2610
      Width           =   2325
   End
   Begin VB.Shape RECT_SHP 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   405
      Index           =   11
      Left            =   135
      Top             =   2325
      Width           =   2415
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' By Author:                                                    ''
'' I'm glade to upload this code on PSC, Because Urdu Database   ''
'' mathods are not commonly available on the online world.       ''
'' I also searched more for that but no source code was available''
'' for anyone. All Open source programmers was asking to make or ''
'' get a mathod or for a sample database.                        ''                               ''
'' So, for that reason, i've personally tried to make a database ''
'' that recognize Urdu Script like language characters in  a     ''
'' MS Access Database using VB6. And after a long search on the  ''
'' online wrold, i've got more different ideas, And when those   ''
'' ideas combined togather, a sucessfull Urdu Database Managment ''
'' Syatem (DBMS) appeared. This Project is only a Sample of that.''
'' You can use & modify it in you projects for your database     ''
'' managment systems.                                            ''
'' I hope, you'll like this effort. I'll wait for your Comments  ''
'' & votes on Planet-Source-Code.Com & on my email.              ''
''                                                               ''
'' Changes in Last Update (Feb 2012):                            ''
''                                                               ''
'' 1: Source code Maximum Optimized                              ''
'' 2: Frontend outlook changed with a new & professional look    ''
'' 3: DB Functions enhanced                                      ''
'' 4: Common Urdu Message Box introduced                         ''
'' 5: Common Moduling System used                                ''
'' 6: Urdu Phonetic Keyboard Layout Module Updated               ''
'' 7: Urdu Captions via HEX Values, (Picture methode removed)    ''
''                                                               ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Your feedback is very valueable & usefull for me, So,  '
'Please don't forget to give feedback & suggestions you '
'have. You can also contact me on my email if you have  '
'any problem in use of this project. I'll feel glade to '
'guide you for better work, as i can.                   '
'                                                       '
'Thank You.                                             '
'Regards,                                               '
'              Muhammd Mehmood Iqbal                    '
'               ME_IQ_TM@Yahoo.Com                      '
'                 +92-313-5324352                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Veriables, That Will Be Used in the Whole Project

Public DB_Mode As Integer
Public IsLast_Rec As Boolean


Private Sub CMD1_Click()

'Goto Next Record
RecSource.MoveNext

'If no record existing Next then, stay on last
If RecSource.EOF Then

RecSource.MoveLast

End If

End Sub

Private Sub CMD2_Click()

'Goto Previous Record
RecSource.MovePrevious

'If no record existing previous then, stay on first
If RecSource.BOF Then

RecSource.MoveFirst

End If

End Sub

Private Sub CMD3_Click()

'Goto Last Record
RecSource.MoveLast

End Sub

Private Sub CMD4_Click()

'Goto First Record
RecSource.MoveFirst

End Sub

Private Sub CMD5_Click()

'Set Database Mode to "Add"
DB_Mode = 0

'Make invisible all data containing Labels
Initialize.Invisible_LBLS Me

'Make visible all Textboxes
Initialize.Visible_TBXS Me


'Clear All TBX(es to Enter New Data
TBX(0).Text = ""
TBX(1).Text = ""
TBX(2).Text = ""
TBX(3).Text = ""
TBX(4).Text = ""
TBX(5).Text = ""
TBX(6).Text = ""
TBX(7).Text = ""
TBX(8).Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

'Disabling CMD Buttons to Avoide From Runtime Errors
CMD1.Enabled = False
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = True
CMD9.Enabled = False

'Change Caption to "Save Record"
CMD7.Caption = ChrW$(&H645) & ChrW$(&H62D) & ChrW$(&H641) & ChrW$(&H648) & ChrW$(&H638) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA)
TBX(0).SetFocus

'New Record Count (SN)
Rec_Count = RecSource.RecordCount + 1


End Sub
 
Private Sub CMD6_Click()

'Check if database file have no record
If RecSource.BOF = True Or RecSource.EOF = True Then

     'If no record found then show Error message
     MSG_BOX.Show 2
     Exit Sub

End If

'Confirming to Delete a Record
MSG_BOX.Show 3

End Sub

Private Sub CMD7_Click()

'Saving a New Record
If DB_Mode = 0 Then

      RS_Functions.Add Me
      
  
'Updating an Existing Record
ElseIf DB_Mode = 1 Then

      'Call Update Function
      RS_Functions.Update Me

      'Set Command7 Caption to "Save Record"
      If CMD7.Enabled = False Then

         CMD7.Caption = ChrW$(&H645) & ChrW$(&H62D) & ChrW$(&H641) & ChrW$(&H648) & ChrW$(&H638) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA)

      End If

End If


End Sub

Private Sub CMD8_Click()

'Unload Message Form
Unload MSG_BX

'Unload Main Form
Unload Me

End Sub

Private Sub CMD9_Click()

'Set Databse Mode to "Update"
DB_Mode = 1

Initialize.Invisible_LBLS Me
Initialize.Visible_TBXS Me

'Update Button Clicked Then, Do
CMD1.Enabled = False
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = True
CMD9.Enabled = False

'Change Command7 caption to "Save Changes"
CMD7.Caption = ChrW$(&H62A) & ChrW$(&H628) & ChrW$(&H62F) & ChrW$(&H6CC) & ChrW$(&H644) & ChrW$(&H6CC) & ChrW(&H20) & ChrW$(&H645) & ChrW$(&H62D) & ChrW$(&H641) & ChrW$(&H648) & ChrW$(&H638) & ChrW(&H20) & ChrW$(&H6A9) & ChrW$(&H631) & ChrW$(&H6CC) & ChrW$(&H6BA)

'Get record position from "SN" field
Rec_Count = RecSource.Fields(0)

End Sub


Private Sub Form_Activate()

'Check for the Modes, in which conditions Main Form activated

'Incomplete Data
If MSG_BOX.Mode = 1 Then

     'Make Visible all Textboxes
     Initialize.Visible_TBXS Me
     TBX(0).SetFocus

'Record Updated, Deleted etc
ElseIf MSG_BOX.Mode = 2 Then

     'Make Invisible all Textboxes
     Initialize.Invisible_TBXS Me
     
     'Make visible all Data containing Labels
     Initialize.Visible_LBLS Me
     
'New record Added
ElseIf MSG_BOX.Mode = 3 Then

     'Reset Database (Conn, RecSource)
     DB_Con.Reset_DB_Con Me

Else

    'At First Run, Initialize all Captions
    Initialize.Btn_Captions Me
    Initialize.Lbl_Captions Me
    Initialize.CMD_Captions MSG_BX

End If


End Sub

Private Sub Form_Load()

'Start to Application Initializing

'Make Database connection
DB_Con.Connect

'Set RecordSet
DB_Con.Set_RS

'Set Data Fields
DB_Con.Set_DB_Fields Me

'If Connection Sucess Then Show Sucess Message
If Conn.State Then

      'You Can Enter Here a Msgbox That Inform About Sucessfull Connection

End If


End Sub


Private Sub Form_Unload(Cancel As Integer)

'Unload Message Form, before Closing Main Form
Unload MSG_BX

End Sub

Private Sub TBX_Change(Index As Integer)

'Set Label Caption equaleant to Textbox.Text
LBL(Index).Caption = TBX(Index).Text

End Sub

Private Sub TBX_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)

'Call to Urdu Phonetic keyboard Layout Module to write Urdu
Urdu_Phonetic_Keyboard_Layout.KeyDown TBX(Index), KeyCode

End Sub

Private Sub TBX_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)

'Call to Urdu Phonetic keyboard Layout Module to write Urdu
Urdu_Phonetic_Keyboard_Layout.KeyPress TBX(Index), KeyAscii

End Sub

Private Sub Text1_Change()

'Set changes in Textbox Text to Label Captions
LBL(9).Caption = Text1.Text

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

'Allow Numeric keys only
If KeyAscii < 48 Or KeyAscii > 57 Then

   'Also Allow '+' key
   If KeyAscii = 43 Then
   
      Exit Sub
   
   Else

      'If other key pressed, make it nothing
      KeyAscii = 0
   
   End If

End If

End Sub


Private Sub Text2_Change()

''Set changes in Textbox Text to Label Captions
LBL(10).Caption = Text2.Text

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

'Allow only Numeric Keys
If KeyAscii < 48 Or KeyAscii > 57 Then

   'Also allow '+' key
   If KeyAscii = 43 Then
   
      Exit Sub
   
   Else
   
      'If other key pressed, make it nothing
      KeyAscii = 0
   
   End If

End If

End Sub

Private Sub Text3_Change()

'Set changes in Textbox Text to Label Captions
LBL(11).Caption = Text3.Text

End Sub
