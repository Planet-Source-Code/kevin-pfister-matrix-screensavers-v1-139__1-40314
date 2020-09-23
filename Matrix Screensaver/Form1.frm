VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matrix Settings"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSSettings 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Falling Code"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdFrame"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "Frame4"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "More Falling Code"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Call Tracing"
      TabPicture(3)   =   "Form1.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Label6"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame9 
         Caption         =   "Falling Code Colour"
         Height          =   855
         Left            =   -74880
         TabIndex        =   38
         Top             =   2520
         Width           =   5055
         Begin VB.OptionButton optCol 
            Caption         =   "Blue"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   41
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optCol 
            Caption         =   "Red"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optCol 
            Caption         =   "Green"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   39
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame Rate Limiter"
         Height          =   735
         Left            =   -74880
         TabIndex        =   36
         Top             =   1680
         Width           =   5055
         Begin MSComctlLib.Slider SldFrameRate 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   100
            SelStart        =   80
            TickStyle       =   3
            Value           =   80
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "BackGround Image"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   5055
         Begin VB.CheckBox ChkBackImage 
            Caption         =   "Use Background Image"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.TextBox TxtImagePath 
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Text            =   "C:\Agent.jpg"
            Top             =   600
            Width           =   3735
         End
         Begin VB.CommandButton CmdBrowse 
            Caption         =   "Browse..."
            Height          =   375
            Left            =   3960
            TabIndex        =   33
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Number"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   28
         Top             =   840
         Width           =   5055
         Begin VB.CheckBox ChkRandom 
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TxtPhoneNumber 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   120
            MaxLength       =   11
            TabIndex        =   29
            Text            =   "00000000000"
            Top             =   600
            Width           =   4815
         End
      End
      Begin VB.CommandButton CmdFrame 
         Caption         =   "Frame Rate"
         Height          =   375
         Left            =   -72000
         TabIndex        =   27
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "FontSize"
         Height          =   735
         Left            =   -72000
         TabIndex        =   24
         Top             =   2040
         Width           =   2175
         Begin VB.TextBox txtsize 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Text            =   "12"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Which of the 3 screensavers would you like to use?"
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   5055
         Begin VB.OptionButton OptScreen 
            Caption         =   "Falling Code"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton OptScreen 
            Caption         =   "Call Tracing"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   22
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton OptScreen 
            Caption         =   "Knock, Knock"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   21
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Drop Options"
         Height          =   2955
         Left            =   -74880
         TabIndex        =   8
         Top             =   420
         Width           =   2775
         Begin MSComctlLib.Slider SldMaxDropLength 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   510
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            _Version        =   393216
            LargeChange     =   1
            Min             =   10
            Max             =   100
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
         End
         Begin MSComctlLib.Slider SldWait 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   1215
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   500
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
         End
         Begin MSComctlLib.Slider SldDroppingCols 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   100
            SelStart        =   20
            TickStyle       =   3
            Value           =   20
         End
         Begin MSComctlLib.Slider SldFading 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   2580
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            LargeChange     =   1
            Min             =   1
            SelStart        =   4
            TickStyle       =   3
            Value           =   4
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Drop Length"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Wait Before Clearing"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Number of Dropping Columns"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   2070
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fading Speed"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Options"
         Height          =   615
         Left            =   -72000
         TabIndex        =   6
         Top             =   420
         Width           =   2175
         Begin VB.CheckBox ChkFromTop 
            Caption         =   "From Top"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   1  'Checked
            Width           =   1860
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Colour Options"
         Height          =   855
         Left            =   -72000
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
         Begin VB.CheckBox ChkFade 
            Caption         =   "Fading(Much Slower!!)"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1980
         End
         Begin VB.CheckBox ChkMultCols 
            Caption         =   "Multiple Colours(Slower)"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Value           =   1  'Checked
            Width           =   1980
         End
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Needs to be run at 1024 by 768 to work normally"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   "3 Screensavers made to emulate scenes from the Matrix film"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label8 
         Caption         =   "MATRIX Screensavers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label11 
         Caption         =   "Email Address:"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Yet_Another_Idiot@Hotmail.com"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "Form1.frx":04B2
         Top             =   480
         Width           =   240
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   480
      Width           =   1275
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WhichScreensaver As Integer
Dim FallCol As Integer

Private Sub ChkFade_Click()
    'Only enables the fading speed if the Fading option has been checked
    SldFading.Enabled = ChkFade.Value
    Label4.Enabled = ChkFade.Value
End Sub

Sub SaveSets()
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "MaxDrop", SldMaxDropLength.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "BeforeClean", SldWait.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "DropsRunning", SldDroppingCols.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "FadeSpeed", SldFading.Value)
    
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "FromTop", ChkFromTop.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "Random", ChkRandom.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "StrNumber", TxtPhoneNumber.Text)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "Which", WhichScreensaver)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "Colour", FallCol)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "Size", txtsize.Text)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "UseImage", ChkBackImage.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "BckImage", TxtImagePath.Text)
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "Frame Rate", SldFrameRate.Value)
    
    Call SaveSetting("Kevin Pfister's Matrix", "Colour", "Fade", ChkFade.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Colour", "MColours", ChkMultCols.Value)
End Sub

Private Sub ChkMultCols_Click()
    'Only enable the Fading Option and Fading Speed, if Multiple IntColours is checked
    ChkFade.Enabled = ChkMultCols.Value
    If ChkFade.Value = 1 Then
        Label4.Enabled = ChkMultCols.Value
        SldFading.Enabled = ChkMultCols.Value
    End If
    If ChkBackImage.Enabled = True And ChkMultCols.Enabled = False Then
        ChkBackImage.Enabled = False
    End If
End Sub

Private Sub ChkRandom_Click()
    If ChkRandom.Value = 1 Then
        TxtPhoneNumber.Enabled = False
    Else
        TxtPhoneNumber.Enabled = True
    End If
End Sub

Private Sub ChkBackImage_GotFocus()
    TxtImagePath.Enabled = ChkBackImage.Enabled
End Sub

Private Sub CmdBrowse_Click()
    CD1.ShowOpen
    TxtImagePath.Text = CD1.FileName
End Sub

Private Sub CmdCancel_Click()
    End 'Exit without saving
End Sub

Private Sub CmdFrame_Click()
    MsgBox "Matrix FrameRate" & vbCrLf & Str(GetSetting("Kevin Pfister's Matrix", "Speed", "FrameRate", 0)) & "FPS"
End Sub

Private Sub CmdOk_Click()
    Call SaveSets
    End 'Save settings and then exit
End Sub


Private Sub Form_Load()
    FrmConfig.Caption = "Matrix Settings ~ V" & App.Major & "." & App.Minor & "." & App.Revision
    'retieve settings
    SldMaxDropLength.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "MaxDrop", 100)
    SldWait.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "BeforeClean", 100)
    SldDroppingCols.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "DropsRunning", 20)
    SldFading.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "FadeSpeed", 2)
    
    ChkFromTop.Value = GetSetting("Kevin Pfister's Matrix", "Options", "FromTop", 1)
    ChkRandom.Value = GetSetting("Kevin Pfister's Matrix", "Options", "Random", 1)
    TxtPhoneNumber.Text = GetSetting("Kevin Pfister's Matrix", "Options", "StrNumber", "0000000000")
    txtsize.Text = GetSetting("Kevin Pfister's Matrix", "Options", "Size", "12")
    WhichScreensaver = Val(GetSetting("Kevin Pfister's Matrix", "Options", "Which", 0))
    OptScreen(WhichScreensaver).Value = True
    FallCol = GetSetting("Kevin Pfister's Matrix", "Options", "Colour", 1)
    optCol(FallCol).Value = True
    ChkBackImage.Value = Val(GetSetting("Kevin Pfister's Matrix", "Options", "UseImage", 1))
    TxtImagePath.Text = GetSetting("Kevin Pfister's Matrix", "Options", "BckImage", "C:\Agent.jpg")
    SldFrameRate.Value = GetSetting("Kevin Pfister's Matrix", "Options", "Frame Rate", 100)
    
    ChkFade.Value = GetSetting("Kevin Pfister's Matrix", "Colour", "Fade", 0)
    SldFading.Enabled = ChkFade.Value
    Label4.Enabled = ChkFade.Value
    ChkMultCols.Value = GetSetting("Kevin Pfister's Matrix", "Colour", "MColours", 1)
End Sub


Private Sub optCol_Click(Index As Integer)
    FallCol = Index
End Sub

Private Sub OptScreen_Click(Index As Integer)
    WhichScreensaver = Index
End Sub

