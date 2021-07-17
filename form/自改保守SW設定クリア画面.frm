VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHoshuSwClear 
   BorderStyle     =   0  'なし
   Caption         =   "バージョン管理"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   14.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer tmrMail 
      Left            =   2280
      Top             =   8040
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "クリア実行"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      TabIndex        =   3
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "改札機"
      ForeColor       =   &H8000000D&
      Height          =   6855
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   11535
      Begin VB.CommandButton cmdHHisentaku 
         Caption         =   " 表示コーナ   全号機 非選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   7800
         TabIndex        =   8
         Top             =   1680
         Width           =   2000
      End
      Begin VB.CommandButton cmdHSentaku 
         Caption         =   " 表示コーナ   全号機  選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5640
         TabIndex        =   7
         Top             =   1680
         Width           =   2000
      End
      Begin VB.CommandButton cmdZHisentaku 
         Caption         =   "  全コーナ    全号機 非選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3480
         TabIndex        =   6
         Top             =   1680
         Width           =   2000
      End
      Begin VB.CommandButton cmdZSentaku 
         Caption         =   "  全コーナ    全号機 選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   2000
      End
      Begin TabDlg.SSTab tabCorner 
         Height          =   2535
         Left            =   1320
         TabIndex        =   9
         Top             =   3000
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4471
         _Version        =   393216
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   794
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "自改保守SW設定クリア画面.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ChkGoki(15)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "ChkGoki(14)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "ChkGoki(13)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "ChkGoki(12)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "ChkGoki(11)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "ChkGoki(10)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "ChkGoki(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "ChkGoki(8)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "ChkGoki(7)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "ChkGoki(6)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ChkGoki(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "ChkGoki(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "ChkGoki(3)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "ChkGoki(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "ChkGoki(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "ChkGoki(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "  "
         TabPicture(1)   =   "自改保守SW設定クリア画面.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ChkGoki(31)"
         Tab(1).Control(1)=   "ChkGoki(30)"
         Tab(1).Control(2)=   "ChkGoki(29)"
         Tab(1).Control(3)=   "ChkGoki(28)"
         Tab(1).Control(4)=   "ChkGoki(27)"
         Tab(1).Control(5)=   "ChkGoki(26)"
         Tab(1).Control(6)=   "ChkGoki(25)"
         Tab(1).Control(7)=   "ChkGoki(24)"
         Tab(1).Control(8)=   "ChkGoki(23)"
         Tab(1).Control(9)=   "ChkGoki(22)"
         Tab(1).Control(10)=   "ChkGoki(21)"
         Tab(1).Control(11)=   "ChkGoki(20)"
         Tab(1).Control(12)=   "ChkGoki(19)"
         Tab(1).Control(13)=   "ChkGoki(18)"
         Tab(1).Control(14)=   "ChkGoki(17)"
         Tab(1).Control(15)=   "ChkGoki(16)"
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "  "
         TabPicture(2)   =   "自改保守SW設定クリア画面.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ChkGoki(47)"
         Tab(2).Control(1)=   "ChkGoki(46)"
         Tab(2).Control(2)=   "ChkGoki(45)"
         Tab(2).Control(3)=   "ChkGoki(44)"
         Tab(2).Control(4)=   "ChkGoki(43)"
         Tab(2).Control(5)=   "ChkGoki(42)"
         Tab(2).Control(6)=   "ChkGoki(41)"
         Tab(2).Control(7)=   "ChkGoki(40)"
         Tab(2).Control(8)=   "ChkGoki(39)"
         Tab(2).Control(9)=   "ChkGoki(38)"
         Tab(2).Control(10)=   "ChkGoki(37)"
         Tab(2).Control(11)=   "ChkGoki(36)"
         Tab(2).Control(12)=   "ChkGoki(35)"
         Tab(2).Control(13)=   "ChkGoki(34)"
         Tab(2).Control(14)=   "ChkGoki(33)"
         Tab(2).Control(15)=   "ChkGoki(32)"
         Tab(2).ControlCount=   16
         TabCaption(3)   =   "  "
         TabPicture(3)   =   "自改保守SW設定クリア画面.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "ChkGoki(63)"
         Tab(3).Control(1)=   "ChkGoki(62)"
         Tab(3).Control(2)=   "ChkGoki(61)"
         Tab(3).Control(3)=   "ChkGoki(60)"
         Tab(3).Control(4)=   "ChkGoki(59)"
         Tab(3).Control(5)=   "ChkGoki(58)"
         Tab(3).Control(6)=   "ChkGoki(57)"
         Tab(3).Control(7)=   "ChkGoki(56)"
         Tab(3).Control(8)=   "ChkGoki(55)"
         Tab(3).Control(9)=   "ChkGoki(54)"
         Tab(3).Control(10)=   "ChkGoki(53)"
         Tab(3).Control(11)=   "ChkGoki(52)"
         Tab(3).Control(12)=   "ChkGoki(51)"
         Tab(3).Control(13)=   "ChkGoki(50)"
         Tab(3).Control(14)=   "ChkGoki(49)"
         Tab(3).Control(15)=   "ChkGoki(48)"
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "  "
         TabPicture(4)   =   "自改保守SW設定クリア画面.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "ChkGoki(79)"
         Tab(4).Control(1)=   "ChkGoki(78)"
         Tab(4).Control(2)=   "ChkGoki(77)"
         Tab(4).Control(3)=   "ChkGoki(76)"
         Tab(4).Control(4)=   "ChkGoki(75)"
         Tab(4).Control(5)=   "ChkGoki(74)"
         Tab(4).Control(6)=   "ChkGoki(73)"
         Tab(4).Control(7)=   "ChkGoki(72)"
         Tab(4).Control(8)=   "ChkGoki(71)"
         Tab(4).Control(9)=   "ChkGoki(70)"
         Tab(4).Control(10)=   "ChkGoki(69)"
         Tab(4).Control(11)=   "ChkGoki(68)"
         Tab(4).Control(12)=   "ChkGoki(67)"
         Tab(4).Control(13)=   "ChkGoki(66)"
         Tab(4).Control(14)=   "ChkGoki(65)"
         Tab(4).Control(15)=   "ChkGoki(64)"
         Tab(4).ControlCount=   16
         TabCaption(5)   =   "  "
         TabPicture(5)   =   "自改保守SW設定クリア画面.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "ChkGoki(95)"
         Tab(5).Control(1)=   "ChkGoki(94)"
         Tab(5).Control(2)=   "ChkGoki(93)"
         Tab(5).Control(3)=   "ChkGoki(92)"
         Tab(5).Control(4)=   "ChkGoki(91)"
         Tab(5).Control(5)=   "ChkGoki(90)"
         Tab(5).Control(6)=   "ChkGoki(89)"
         Tab(5).Control(7)=   "ChkGoki(88)"
         Tab(5).Control(8)=   "ChkGoki(87)"
         Tab(5).Control(9)=   "ChkGoki(86)"
         Tab(5).Control(10)=   "ChkGoki(85)"
         Tab(5).Control(11)=   "ChkGoki(84)"
         Tab(5).Control(12)=   "ChkGoki(83)"
         Tab(5).Control(13)=   "ChkGoki(82)"
         Tab(5).Control(14)=   "ChkGoki(81)"
         Tab(5).Control(15)=   "ChkGoki(80)"
         Tab(5).ControlCount=   16
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   105
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   104
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   103
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   6360
            TabIndex        =   102
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   101
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   100
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   4200
            TabIndex        =   99
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   6360
            TabIndex        =   98
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   97
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   2160
            TabIndex        =   96
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   4200
            TabIndex        =   95
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   6360
            TabIndex        =   94
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   93
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   2160
            TabIndex        =   92
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   4200
            TabIndex        =   91
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "Z9号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   6360
            TabIndex        =   90
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   -74880
            TabIndex        =   89
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   -72840
            TabIndex        =   88
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   -70800
            TabIndex        =   87
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   -68640
            TabIndex        =   86
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   -74880
            TabIndex        =   85
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   -72840
            TabIndex        =   84
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   -70800
            TabIndex        =   83
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   -68640
            TabIndex        =   82
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   -74880
            TabIndex        =   81
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   -72840
            TabIndex        =   80
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   -70800
            TabIndex        =   79
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   -68640
            TabIndex        =   78
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   -74880
            TabIndex        =   77
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   -72840
            TabIndex        =   76
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   -70800
            TabIndex        =   75
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   -68640
            TabIndex        =   74
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   -74880
            TabIndex        =   73
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   -72840
            TabIndex        =   72
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   -70800
            TabIndex        =   71
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   -68640
            TabIndex        =   70
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   -74880
            TabIndex        =   69
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   -72840
            TabIndex        =   68
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   -70800
            TabIndex        =   67
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   -68640
            TabIndex        =   66
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   -74880
            TabIndex        =   65
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   -72840
            TabIndex        =   64
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   -70800
            TabIndex        =   63
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   43
            Left            =   -68640
            TabIndex        =   62
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   44
            Left            =   -74880
            TabIndex        =   61
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   45
            Left            =   -72840
            TabIndex        =   60
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   46
            Left            =   -70800
            TabIndex        =   59
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   47
            Left            =   -68640
            TabIndex        =   58
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   48
            Left            =   -74880
            TabIndex        =   57
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   49
            Left            =   -72840
            TabIndex        =   56
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   50
            Left            =   -70800
            TabIndex        =   55
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   51
            Left            =   -68640
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   52
            Left            =   -74880
            TabIndex        =   53
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   53
            Left            =   -72840
            TabIndex        =   52
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   54
            Left            =   -70800
            TabIndex        =   51
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   55
            Left            =   -68640
            TabIndex        =   50
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   56
            Left            =   -74880
            TabIndex        =   49
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   57
            Left            =   -72840
            TabIndex        =   48
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   58
            Left            =   -70800
            TabIndex        =   47
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   59
            Left            =   -68640
            TabIndex        =   46
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   60
            Left            =   -74880
            TabIndex        =   45
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   61
            Left            =   -72840
            TabIndex        =   44
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   62
            Left            =   -70800
            TabIndex        =   43
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   63
            Left            =   -68640
            TabIndex        =   42
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   64
            Left            =   -74880
            TabIndex        =   41
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   65
            Left            =   -72840
            TabIndex        =   40
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   66
            Left            =   -70800
            TabIndex        =   39
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   67
            Left            =   -68640
            TabIndex        =   38
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   68
            Left            =   -74880
            TabIndex        =   37
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   69
            Left            =   -72840
            TabIndex        =   36
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   70
            Left            =   -70800
            TabIndex        =   35
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   71
            Left            =   -68640
            TabIndex        =   34
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   72
            Left            =   -74880
            TabIndex        =   33
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   73
            Left            =   -72840
            TabIndex        =   32
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   74
            Left            =   -70800
            TabIndex        =   31
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   75
            Left            =   -68640
            TabIndex        =   30
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   76
            Left            =   -74880
            TabIndex        =   29
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   77
            Left            =   -72840
            TabIndex        =   28
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   78
            Left            =   -70800
            TabIndex        =   27
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   79
            Left            =   -68640
            TabIndex        =   26
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   80
            Left            =   -74880
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   81
            Left            =   -72840
            TabIndex        =   24
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   82
            Left            =   -70800
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   83
            Left            =   -68640
            TabIndex        =   22
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   84
            Left            =   -74880
            TabIndex        =   21
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   85
            Left            =   -72840
            TabIndex        =   20
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   86
            Left            =   -70800
            TabIndex        =   19
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   87
            Left            =   -68640
            TabIndex        =   18
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   88
            Left            =   -74880
            TabIndex        =   17
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   89
            Left            =   -72840
            TabIndex        =   16
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   90
            Left            =   -70800
            TabIndex        =   15
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   91
            Left            =   -68640
            TabIndex        =   14
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   92
            Left            =   -74880
            TabIndex        =   13
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   93
            Left            =   -72840
            TabIndex        =   12
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   94
            Left            =   -70800
            TabIndex        =   11
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkGoki 
            Caption         =   "１２３４５号機"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   95
            Left            =   -68640
            TabIndex        =   10
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         Caption         =   "統合監視盤内に保存している改札機保守ＳＷ設定をクリアします。"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   11295
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  データ収集・出力    画面へ戻る"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "改札機保守SW設定クリア"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmHoshuSwClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmHoshuSwClear.frm
'//  パッケージ名：自改保守SW設定クリア画面
'//
'//  概要：自改保守SW設定クリア画面
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//                 フェーズ２対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//                 【フェーズ２対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

Private mintStatus(31) As Integer       'EG20 V2.1.0.1 ADD 【フェーズ２対応】

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 自改保守SW設定クリア画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    On Error Resume Next
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 自改保守SW設定クリア画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    On Error Resume Next
    
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信タイマ、タイムアップ処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmHoshuSwClear.Caption, False
        pfFormActive (frmHoshuSwClear.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 自改保守SW設定クリア画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-24   CODED   BY [TCC] M.Matsumoto
'//                 【統合No53対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim iCnt As Integer
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    Dim iCnt2 As Integer
    Dim intIndex As Integer
    Dim strCorner1 As String
    Dim strCorner2 As String
    'EG20 V2.1.0.1 ADD END
   
    On Error Resume Next
   
    '「自改保守SW設定クリア画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_SW_CLEAR_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'    OptGateSet(0).Value = True
'    Call OptGateSet_Click(0)
'    Call CmdGokiSelect_Click(0)
'
'     For iCnt = 0 To ChkGoki.UBound
'         gClear_Gouki(iCnt) = CLEAR_FLAG.NOT_CLEAR
'     Next
    'EG20 V2.1.0.1 DEL END
          
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    For iCnt = 0 To UBound(mintStatus)
        gClear_Gouki(iCnt) = CLEAR_FLAG.NOT_CLEAR
    Next
    
    '号機情報取得
    Call gsGetGateInfo
    Call gsGetCornerName
    
    'タブ数を設置コーナ数とする
    tabCorner.Tab = 0
    
    '収集状態初期化
    Erase mintStatus
    
    For iCnt = 0 To UBound(gblnCornerSet)
    
        '設定ありのコーナを活性にする
        If gblnCornerSet(iCnt) = True Then
            'コーナー名称表示
            strCorner1 = MidB(gstrCornerName(iCnt), 1, 12)
            strCorner2 = MidB(gstrCornerName(iCnt), 13, 24)
            tabCorner.TabCaption(iCnt) = strCorner1 & vbCrLf & strCorner2
            
        End If
    
    Next iCnt
    
    '設置コーナ数分ループ
    For iCnt = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(iCnt) = False Then
            tabCorner.TabVisible(iCnt) = False
        End If

        '最大号機数分ループ
        For iCnt2 = 0 To 15
            intIndex = (iCnt * 16) + iCnt2
            ChkGoki(intIndex).Visible = False
            ChkGoki(intIndex).Tag = "0"
        Next
        
        For iCnt2 = 0 To 15
            intIndex = (iCnt * 16) + (gudtSettiCorner(iCnt).intGokiNo(iCnt2) - 1)
            If gudtSettiCorner(iCnt).intGokiNo(iCnt2) > 0 Then
                ChkGoki(intIndex).Caption = gudtSettiCorner(iCnt).strDispGoki(iCnt2) + "号機"
                'Tagに対応する号機番号を記録（1～32号機）
                ChkGoki(intIndex).Tag = CStr(gudtSettiCorner(iCnt).intGateNo(iCnt2))
'                mintStatus(gudtSettiCorner(iCnt).intGateNo(iCnt2) - 1) = CHECKBOX_OFF      'EG20 V5.4.0.1 DEL 【統合No53対応】
                mintStatus(gudtSettiCorner(iCnt).intGateNo(iCnt2) - 1) = CHECKBOX_ON        'EG20 V5.4.0.1 ADD 【統合No53対応】
                ChkGoki(intIndex).Visible = True
'                ChkGoki(intIndex).Value = CHECKBOX_OFF         'EG20 V5.4.0.1 DEL 【統合No53対応】
                ChkGoki(intIndex).Value = CHECKBOX_ON           'EG20 V5.4.0.1 ADD 【統合No53対応】
            End If
        Next iCnt2
        
    Next iCnt
    'EG20 V2.1.0.1 ADD END
    
    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : CmdClear_Click
'//  機能名称  : 「クリア実行」釦押下
'//  機能概要  : 自改保守SW設定のクリア処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()
    Dim iRet As Integer
    Dim iCnt As Integer
        
    iRet = MsgBox("指定した号機の自改保守SWデータを削除します。" & vbCrLf & "よろしいですか？", _
            vbQuestion + vbOKCancel, "クリア確認")
    If iRet = vbOK Then
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'      If OptGateSet(0).Value = True Then
'         'ラジオ釦：全号機選択時→削除対象全号機
'          For iCnt = 0 To ChkGoki.UBound
'              gClear_Gouki(iCnt) = CLEAR_FLAG.TARGET_CLEAR
'          Next
'       Else
'         'ラジオ釦：指定号機のみ選択時→削除対象指定号機
'          For iCnt = 0 To ChkGoki.UBound
'             If ChkGoki(iCnt).Value = 1 Then
'              gClear_Gouki(iCnt) = CLEAR_FLAG.TARGET_CLEAR
'             End If
'          Next
'       End If
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        'ラジオ釦：指定号機のみ選択時→削除対象指定号機
        For iCnt = 0 To UBound(mintStatus)
            If mintStatus(iCnt) = 1 Then
                gClear_Gouki(iCnt) = CLEAR_FLAG.TARGET_CLEAR
            Else
                gClear_Gouki(iCnt) = CLEAR_FLAG.NOT_CLEAR
            End If
        Next
        'EG20 V2.1.0.1 ADD END
        
       'クリア中画面を表示する。
       frmHoshuClear.Show vbModal
    End If
End Sub

'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : CmdGokiSelect_Click
''//  機能名称  : 「全号機選択」「全号機解除」」釦押下
''//  機能概要  : 指定号機部の状態を全号機選択状態/解除状態に更新する。
''//
''//              型        名称      意味
''//  引数      : Integer　Index　　[IN]インデックス値
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Sub CmdGokiSelect_Click(Index As Integer)
'
'    Dim iLoopCnt As Integer
'
'    If Index = 0 Then
'        '「全号機選択」釦押下時
'        For iLoopCnt = 0 To ChkGoki.UBound
'            ChkGoki(iLoopCnt).Value = 1
'        Next
'
'    Else
'        '「全号機解除」釦押下時
'        For iLoopCnt = 0 To ChkGoki.UBound
'            ChkGoki(iLoopCnt).Value = 0
'        Next
'
'    End If
'End Sub
'EG20 V2.1.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    '「自改保守SW設定クリア画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_SW_CLEAR_GAMEN_END, 0)
    
    Unload Me
End Sub

'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : chkLogGouki_Click
'//  機能名称  : 指定号機チェックボックスクリック時処理
'//  機能概要  : 内部変数のON/OFFを切り替える
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]チェックボックスインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub ChkGoki_Click(Index As Integer)

    Dim intGoki As Integer
    
    On Error Resume Next
    
    intGoki = CInt(ChkGoki(Index).Tag) - 1
    
    mintStatus(intGoki) = ChkGoki(Index).Value
    
End Sub

''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : OptGateSet_Click
''//  機能名称  : ラジオ釦「全号機」「指定号機のみ」選択時
''//  機能概要  : ラジオ釦にそってい、指定号機の選択不可/可への状態更新する。
''//
''//              型        名称      意味
''//  引数      : Integer　Index　　[IN]インデックス値
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Sub OptGateSet_Click(Index As Integer)
'
'    Dim iLoopCnt As Integer
'
'    If Index = 0 Then
'        'ラジオ釦：全号機
'        CmdGokiSelect(0).Enabled = False
'        CmdGokiSelect(1).Enabled = False
'        FramGoki.Enabled = False
'
'        For iLoopCnt = 0 To ChkGoki.UBound
'            ChkGoki(iLoopCnt).Enabled = False
'        Next
'
'    Else
'        'ラジオ釦：指定号機のみ
'        CmdGokiSelect(0).Enabled = True
'        CmdGokiSelect(1).Enabled = True
'        FramGoki.Enabled = True
'
'        For iLoopCnt = 0 To ChkGoki.UBound
'            ChkGoki(iLoopCnt).Enabled = True
'        Next
'    End If
'End Sub
'EG20 V2.1.0.1 DEL END

'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdZHisentaku_Click
'//  機能名称  : 全コーナ全号機非選択ボタン押下処理
'//  機能概要  : すべての号機を非選択状態にする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZHisentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To ChkGoki.UBound
        ChkGoki(intLoop).Value = CHECKBOX_OFF
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdZSentaku_Click
'//  機能名称  : 全コーナ全号機選択ボタン押下処理
'//  機能概要  : すべての号機を選択状態にする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZSentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To ChkGoki.UBound
        ChkGoki(intLoop).Value = CHECKBOX_ON
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdHHisentaku_Click
'//  機能名称  : 表示コーナ全号機非選択ボタン押下処理
'//  機能概要  : すべての号機を非選択状態にする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdHHisentaku_Click()

    Dim intLoop As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next

    intStIndex = tabCorner.Tab * 16
    intEdIndex = intStIndex + 15
    
    For intLoop = intStIndex To intEdIndex
        ChkGoki(intLoop).Value = CHECKBOX_OFF
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdHSentaku_Click
'//  機能名称  : 表示コーナ全号機選択ボタン押下処理
'//  機能概要  : すべての号機を選択状態にする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdHSentaku_Click()

    Dim intLoop As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next

    intStIndex = tabCorner.Tab * 16
    intEdIndex = intStIndex + 15
    
    For intLoop = intStIndex To intEdIndex
        ChkGoki(intLoop).Value = CHECKBOX_ON
    Next intLoop
    
End Sub

'EG20 V2.1.0.1 ADD END
