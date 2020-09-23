VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H008A4922&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MFG EX TOOLTIPS 1.0"
   ClientHeight    =   8925
   ClientLeft      =   270
   ClientTop       =   585
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   60
      Text            =   "Try Me"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   59
      Text            =   "Try Me"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   58
      Text            =   "Try Me"
      Top             =   4440
      Width           =   615
   End
   Begin VB.ComboBox CboFont 
      Height          =   315
      ItemData        =   "FormDemo.frx":0000
      Left            =   2520
      List            =   "FormDemo.frx":0002
      TabIndex        =   56
      Text            =   "cboFont"
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   7200
      TabIndex        =   51
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   225
      Left            =   7080
      TabIndex        =   50
      Text            =   "E-mail Me"
      Top             =   720
      Width           =   1695
   End
   Begin VB.OptionButton OptionBack 
      BackColor       =   &H008A4922&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   8760
      TabIndex        =   47
      Top             =   4680
      Width           =   255
   End
   Begin VB.PictureBox PictureBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   9120
      Picture         =   "FormDemo.frx":0004
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   46
      Top             =   4560
      Width           =   330
   End
   Begin VB.OptionButton OptionBack 
      BackColor       =   &H008A4922&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   45
      Top             =   3960
      Width           =   255
   End
   Begin VB.PictureBox PictureBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   9120
      Picture         =   "FormDemo.frx":3BBE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   44
      Top             =   3840
      Width           =   330
   End
   Begin VB.OptionButton OptionBack 
      BackColor       =   &H008A4922&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   43
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox PictureBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   9120
      Picture         =   "FormDemo.frx":6A32
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   42
      Top             =   3240
      Width           =   330
   End
   Begin VB.OptionButton OptionBack 
      BackColor       =   &H008A4922&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   41
      Top             =   2520
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.PictureBox PictureBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   9120
      Picture         =   "FormDemo.frx":6D96
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   40
      Top             =   2400
      Width           =   330
   End
   Begin VB.ComboBox CboImageSize 
      Height          =   315
      ItemData        =   "FormDemo.frx":E646
      Left            =   2520
      List            =   "FormDemo.frx":E659
      TabIndex        =   38
      Text            =   "cboSize"
      Top             =   6720
      Width           =   1815
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2520
      Picture         =   "FormDemo.frx":E68F
      ScaleHeight     =   720
      ScaleWidth      =   1785
      TabIndex        =   36
      Top             =   2040
      Width           =   1815
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Try Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   960
         TabIndex        =   37
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2520
      Picture         =   "FormDemo.frx":F559
      ScaleHeight     =   720
      ScaleWidth      =   1785
      TabIndex        =   34
      Top             =   4200
      Width           =   1815
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Try Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         TabIndex        =   35
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2520
      Picture         =   "FormDemo.frx":10423
      ScaleHeight     =   720
      ScaleWidth      =   1785
      TabIndex        =   32
      Top             =   3120
      Width           =   1815
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Try Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   960
         TabIndex        =   33
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2520
      Picture         =   "FormDemo.frx":112ED
      ScaleHeight     =   720
      ScaleWidth      =   1785
      TabIndex        =   30
      Top             =   5280
      Width           =   1815
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Try Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   960
         TabIndex        =   31
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2520
      Picture         =   "FormDemo.frx":121B7
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   28
      Top             =   960
      Width           =   1815
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Try Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   960
         TabIndex        =   29
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      ItemData        =   "FormDemo.frx":13081
      Left            =   2520
      List            =   "FormDemo.frx":13091
      TabIndex        =   27
      Text            =   "cboStyle"
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   6000
      TabIndex        =   25
      Text            =   "Try Me"
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5160
      TabIndex        =   23
      Text            =   "Try Me"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   22
      Text            =   "Try Me"
      Top             =   6240
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageListC 
      Left            =   9000
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":130CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":13667
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":13C01
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":1419B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":15D95
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":1798F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":17F29
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":19B23
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormDemo.frx":1B71D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Textwelcome 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "Try Me"
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Text            =   "Try Me"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Text            =   "Try Me"
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Text            =   "Try Me"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Text            =   "Try Me"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   1
      Text            =   "Try Me"
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8805
      Left            =   0
      Picture         =   "FormDemo.frx":1F797
      ScaleHeight     =   8775
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox TGEColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   55
         Top             =   4800
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "End"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   4800
         Width           =   735
      End
      Begin VB.PictureBox TGSColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   53
         Top             =   4440
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   4440
         Width           =   735
      End
      Begin VB.CheckBox CheckShadow 
         Caption         =   "TTShadow"
         Height          =   255
         Left            =   480
         TabIndex        =   48
         Top             =   8400
         Width           =   1215
      End
      Begin VB.CheckBox CheckBalloon 
         Caption         =   "TTBalloon"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   8040
         Width           =   1215
      End
      Begin VB.PictureBox ShapeColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   240
         ScaleHeight     =   105
         ScaleWidth      =   1545
         TabIndex        =   24
         Top             =   6720
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Back"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Text "
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.PictureBox TTextColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   13
         Top             =   3720
         Width           =   855
      End
      Begin VB.PictureBox TBackcolor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   12
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox ColorPicker 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C8&
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   60
         Picture         =   "FormDemo.frx":216A8
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   123
         TabIndex        =   11
         Top             =   5160
         Width           =   1845
      End
      Begin MSComctlLib.Slider SliderT 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   7440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Max             =   100
         SelStart        =   100
         TickStyle       =   3
         Value           =   100
      End
      Begin MSComctlLib.Slider SliderTT 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Max             =   200
         SelStart        =   25
         TickStyle       =   3
         Value           =   25
      End
      Begin MSComctlLib.Slider SliderTTT 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Max             =   2000
         SelStart        =   500
         TickStyle       =   3
         Value           =   500
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H008A4922&
         Caption         =   "Delay ToolTip Kill Time"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Top             =   2520
         Width           =   1980
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H008A4922&
         Caption         =   "Delay ToolTip Pop Time"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   1680
         Width           =   1980
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H008A4922&
         Caption         =   "ExTooltip Transparency"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   17
         Top             =   7080
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H008A4922&
         Caption         =   "ExTooltip Colors"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   3240
         Width           =   1980
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Font Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   57
      Top             =   7920
      Width           =   1530
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Fill Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   49
      Top             =   7200
      Width           =   1380
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Image Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   39
      Top             =   6510
      Width           =   1620
   End
   Begin VB.Label LabelN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mario Flores G"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   7200
      TabIndex        =   10
      Top             =   7680
      Width           =   1410
   End
   Begin VB.Label LabelN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sistec_de_juarez@hotmail.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   6360
      TabIndex        =   9
      Top             =   7920
      Width           =   3045
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mario Flores Ex ToolTip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   8640
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//-- Mario Flores Gonzalez
'//-- ExToolTip Class 1.0
'//-- E:mail : sistec_de_juarez@hotmail.com


'//-- Updates at http://www.geocities.com/sistec_de_juarez/ExTooltip/


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Const DefaultFont = "Arial"  '//-- Default Font For Demo

Private C        As ExToolTip        '//-- ExTooltip Class
Private MDown    As Boolean          '//-- Flag Used in ColorPicker for MouseDown Capture
Private MMove    As Boolean          '//-- Flag Used in ColorPicker for MouseMove Capture


Private Sub Form_Initialize()
InitCommonControls                   '//-- Registers and initializes the common control window classes.
End Sub

Private Sub Form_Terminate()
Set C = Nothing                      '//-- Release all the system and memory resources associated with Class.
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MMove Then MMove = False
End Sub

Private Sub Form_Load()
Dim enumFonts    As Long

For enumFonts = 0 To VB.Screen.FontCount - 1                                    '//--Enumerate All System Fonts
    CboFont.AddItem Screen.Fonts(enumFonts)                                     '//--Add Each Font To ComboBox
    If Screen.Fonts(enumFonts) = DefaultFont Then CboFont.ListIndex = enumFonts '//--Demo Default Font as 1st Item in ComboList.
Next enumFonts
 
Set C = New ExToolTip             '//-- Creates a new instance of theclass.

cboStyle.ListIndex = 0            '//-- Combo 1st Item = Solid
CboImageSize.ListIndex = 3        '//-- Combo 1st Item = TTIcon48
OptionBack_Click 0                '//-- Assign Demo Image as ExTooltip BackGround Picture

End Sub

Private Sub CboFont_Change()
C.Font.Name = CboFont.List(CboFont.ListIndex)
End Sub

Private Sub CheckBalloon_Click()
C.ToolTipStyle = IIf(CheckBalloon.Value = 1, 1, 0)
End Sub

Private Sub cboStyle_Click()
C.BackStyle = cboStyle.ListIndex + 1
End Sub


Private Sub SliderTT_Click()
C.DelayTime = SliderTT.Value
End Sub

Private Sub SliderTTT_Click()
C.KillTime = SliderTTT.Value
End Sub

Private Sub CheckShadow_Click()
C.Shadow = IIf(CheckShadow.Value = 1, True, False)
End Sub

Private Sub OptionBack_Click(Index As Integer)
Set C.Picture = PictureBack(Index)
End Sub

Private Sub ApplyDemoValues() '//--Default Values used in Demo
               
        '//-- Remember each time a new Tooltip is created if the
        '     user doesn't specify any parameters values like this ones,
        '     the default values are going to be the Extooltip default values.(See ExTooltip Class_Initialize).
        
        C.DelayTime = SliderTT.Value
        C.KillTime = SliderTTT.Value
        C.BackColor = TBackcolor.BackColor
        C.TextColor = TTextColor.BackColor
        C.GradientColorStart = TGSColor.BackColor
        C.GradientColorEnd = TGEColor.BackColor
        C.BackStyle = cboStyle.ListIndex + 1
        C.Font.Name = CboFont.List(CboFont.ListIndex)
        C.Shadow = IIf(CheckShadow.Value = 1, True, False)
        C.ToolTipStyle = IIf(CheckBalloon.Value = 1, 1, 0)

End Sub


'==============================================================================
'A Picker Color Function that enables the user to extract a color from a Bitmap
'==============================================================================

Private Sub ColorPicker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim R      As Integer
Dim G      As Integer
Dim B      As Integer
Dim PixCol As Long

'Assign Color
PixCol = GetPixel(ColorPicker.hdc, X, Y)

'Convert to RGB
R = PixCol Mod 256
B = Int(PixCol / 65536)
G = (PixCol - (B * 65536) - R) / 256

'Sanity Checks
If R < 0 Then R = 0
If G < 0 Then G = 0
If B < 0 Then B = 0

'Visual Color Table
ShapeColor.BackColor = RGB(R, G, B)

If Option1.Value = True Then
    C.TextColor = ShapeColor.BackColor
    TTextColor.BackColor = ShapeColor.BackColor
ElseIf Option2.Value = True Then
    C.BackColor = ShapeColor.BackColor
    TBackcolor.BackColor = ShapeColor.BackColor
ElseIf Option3.Value = True Then
    C.GradientColorStart = ShapeColor.BackColor
    TGSColor.BackColor = ShapeColor.BackColor
ElseIf Option4.Value = True Then
    C.GradientColorEnd = ShapeColor.BackColor
    TGEColor.BackColor = ShapeColor.BackColor
End If

MDown = True

End Sub

Private Sub ColorPicker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MDown Then ColorPicker_MouseDown Button, Shift, X, Y
End Sub

Private Sub ColorPicker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MDown = False
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = TTIcon72
          C.BackColor = &H8000&
          C.TextColor = &HFF00&
          C.Font.Name = "Tahoma"
          C.Font.Size = 8
          C.ShowToolTip Command1.hwnd, "Exit Demo", _
          "Remember to Visit Web Page for Updates." _
          , ImageListC.ListImages(9).Picture, SliderT.Value
    End If
        
End Sub

Private Sub CboImageSize_DropDown()
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = 16
          C.ShowToolTip CboImageSize.hwnd, "Warning!", _
          "Using diferent Image size can change Image Aspect." _
          , ImageListC.ListImages(2).Picture, SliderT.Value
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = Choose(CboImageSize.ListIndex + 1, 16, 24, 32, 48, 72)
          C.ShowToolTip Picture2.hwnd, "Microsoft Access", _
          "Microsoft Access is a powerful program to create and manage your databases." & _
          vbCrLf & "It has many built in features to assist you in manage and viewing your information." _
          , Picture2.Picture, SliderT.Value
    End If
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = Choose(CboImageSize.ListIndex + 1, 16, 24, 32, 48, 72)
          C.ShowToolTip Picture3.hwnd, "Microsoft Excel", _
          "Microsoft Excel allows you to create professional spreadsheets and charts." & _
          vbCrLf & "It performs numerous functions and formulas to assist you in your projects." _
          , Picture3.Picture, SliderT.Value
    End If
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = Choose(CboImageSize.ListIndex + 1, 16, 24, 32, 48, 72)
          C.ShowToolTip Picture4.hwnd, "Microsoft PowerPoint", _
          "Microsoft PowerPoint is a powerful tool to create professional looking presentations and slide shows." & _
          vbCrLf & "PowerPoint allows you to build presentations from scratch or by using the easy to use wizard." _
          , Picture4.Picture, SliderT.Value
    End If
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = Choose(CboImageSize.ListIndex + 1, 16, 24, 32, 48, 72)
          C.ShowToolTip Picture5.hwnd, "Microsoft Word", _
          "Microsoft Word is a powerful tool to create professional looking documents." _
          , Picture5.Picture, SliderT.Value
    End If
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = Choose(CboImageSize.ListIndex + 1, 16, 24, 32, 48, 72)
          C.ShowToolTip Picture6.hwnd, "Microsoft Publisher", _
          "Microsoft Publisher helps you easily create," & vbCrLf & _
          "customize and publish materials such as newsletters," & vbCrLf & _
          "brochures, flyers, catalogs, and Web sites." _
          , Picture6.Picture, SliderT.Value
    End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.ShowToolTip Text1.hwnd, "Caps Lock is On ", _
          "Having Caps Lock on may cause you to enter your password" & _
          vbCrLf & "incorrectly." & _
          vbCrLf & _
          vbCrLf & "You should press Caps Lock to turn it off before entering your" & _
          vbCrLf & "password.", TTI_WARNING, SliderT.Value
      End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.ShowToolTip Text2.hwnd, "Did you forget your password?", _
          "You can click the ? button to see your password hint." & _
          vbCrLf & _
          vbCrLf & "Please type your password again." & _
          vbCrLf & "Be sure yo use the correct uppercase and lowercase letters." _
          , TTI_ERROR, SliderT.Value
      End If
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.ShowToolTip Text3.hwnd, "Password Hint:", _
          "Visual Basic Best ToolTip" _
          , TTI_INFO, SliderT.Value
      End If
End Sub

Private Sub Text4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.IconSize = 72
          C.ShowToolTip Text4.hwnd, "Planet Source Code", "Mario Flores ExToolTip " & _
          vbCrLf & "Open Source" _
          , ImageListC.ListImages(4).Picture, SliderT.Value
       End If
End Sub

Private Sub Text5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.BackColor = vbWhite
          C.TextColor = vbBlack
          C.IconSize = 72
          C.ShowToolTip Text5.hwnd, "ExToolTip", "sistec_de_juarez@hotmail.com", ImageListC.ListImages(5).Picture, SliderT.Value
    End If
End Sub

Private Sub Text6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.BackColor = &H40C0&
          C.TextColor = &H80FF&
          C.IconSize = 16
          C.ShowToolTip Text6.hwnd, "WinZip", "Visit www.winzip.com", ImageListC.ListImages(6).Picture, 70
    End If
End Sub

Private Sub Text7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.BackColor = &HC8&
          C.TextColor = vbWhite
          C.IconSize = 72
          C.ShowToolTip Text7.hwnd, "Drink Coke", "Visit www.cocacola.com", ImageListC.ListImages(7).Picture, SliderT.Value
    End If
End Sub

Private Sub Text8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.Font.Name = "Verdana"
          C.Font.Size = 7
          C.IconSize = 72
          C.BackColor = &HC00000
          C.TextColor = &HFFFF00
          C.ShowToolTip Text8.hwnd, "E-MAIL ME", _
          "sistec_de_juarez@hotmail.com", ImageListC.ListImages(8).Picture, SliderT.Value
      End If
End Sub

Private Sub Text9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.Font.Size = 10
          C.Font.Italic = True
          C.ShowToolTip Text9.hwnd, "ExToolTips 1.0", _
          "   · Custom Icon ToolTips" & _
          vbCrLf & "   · Custom Ballon Styles" & _
          vbCrLf & "   · Transparency in ToolTips" & _
          vbCrLf & "   · Support For MultiLine Text" & _
          vbCrLf & "   · Designed for Windows NT,98,ME,2000,2003,XP ", 0, SliderT.Value
      End If
End Sub

Private Sub Text10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
           C.ShowToolTip Text10.hwnd, "Do you Like Custom Icons?", _
          "Display any Custom Icon on Screen.", _
          ImageListC.ListImages(1).Picture, SliderT.Value
      End If
End Sub

Private Sub Text11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.ShowToolTip Text11.hwnd, "Do you Like Custom Icons?", _
          "Display any Custom Icon on Screen.", _
          ImageListC.ListImages(3).Picture, SliderT.Value
      End If
End Sub

Private Sub Text12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
           C.ShowToolTip Text12.hwnd, "Do you Like Custom Icons?", _
          "Display any Custom Icon on Screen.", _
          ImageListC.ListImages(2).Picture, SliderT.Value
      End If
End Sub

Private Sub Textwelcome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Not MMove Or C.Alive = False Then
          MMove = True
          ApplyDemoValues
          C.ShowToolTip Textwelcome.hwnd, "ExToolTips 1.0", _
          "   · Custom Icon ToolTips" & _
          vbCrLf & "   · Custom Ballon Styles" & _
          vbCrLf & "   · Transparency in ToolTips" & _
          vbCrLf & "   · Support For MultiLine Text" & _
          vbCrLf & "   · You can use any Font in the System" & _
          vbCrLf & "   · Designed for Windows NT,98,ME,2000,2003,XP ", 0, SliderT.Value
      End If
End Sub
