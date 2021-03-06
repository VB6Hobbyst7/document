VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConectSts 
   BorderStyle     =   0  'なし
   Caption         =   "通信切断"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9600
      Top             =   5640
   End
   Begin VB.CommandButton cmdDataUp 
      Caption         =   "表示更新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9120
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "通信確認・表示 画面へ戻る"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9120
      TabIndex        =   1
      Top             =   7755
      Width           =   2600
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   15000
      Width           =   2895
   End
   Begin TabDlg.SSTab tabConect 
      Height          =   8620
      Left            =   0
      TabIndex        =   3
      Top             =   380
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   15214
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   706
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "改札機"
      TabPicture(0)   =   "通信切断画面.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tabCorner"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ICM"
      TabPicture(1)   =   "通信切断画面.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabIcmCorner"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "上位機器"
      TabPicture(2)   =   "通信切断画面.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabKikiCorner"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "操作卓"
      TabPicture(3)   =   "通信切断画面.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "コーナ一覧"
         ForeColor       =   &H00FF0000&
         Height          =   5760
         Left            =   -74640
         TabIndex        =   180
         Top             =   840
         Width           =   7335
         Begin VB.Frame fraALLSitei 
            Caption         =   "一括指定"
            ForeColor       =   &H00FF0000&
            Height          =   1095
            Index           =   12
            Left            =   240
            TabIndex        =   181
            Top             =   4440
            Width           =   6735
            Begin VB.CommandButton cmdInOutTaku 
               Caption         =   "全コーナ切離"
               Height          =   495
               Index           =   1
               Left            =   2535
               TabIndex        =   182
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton cmdInOutTaku 
               Caption         =   "全コーナ接続"
               Height          =   495
               Index           =   0
               Left            =   360
               TabIndex        =   183
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "コーナ名称"
            Height          =   4020
            Left            =   240
            TabIndex        =   184
            Top             =   360
            Width           =   4575
            Begin VB.CheckBox chkTaku 
               BackColor       =   &H0080FF80&
               Caption         =   "接続"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   3600
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   185
               Top             =   360
               Width           =   800
            End
            Begin VB.CheckBox chkTaku 
               BackColor       =   &H0080FF80&
               Caption         =   "接続"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   1
               Left            =   3600
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   186
               Top             =   960
               Width           =   800
            End
            Begin VB.CheckBox chkTaku 
               BackColor       =   &H000000FF&
               Caption         =   "切離"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   2
               Left            =   3600
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   187
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Width           =   800
            End
            Begin VB.CheckBox chkTaku 
               BackColor       =   &H000000FF&
               Caption         =   "切離"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   3
               Left            =   3600
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   188
               Top             =   2160
               Value           =   1  'ﾁｪｯｸ
               Width           =   800
            End
            Begin VB.CheckBox chkTaku 
               BackColor       =   &H000000FF&
               Caption         =   "切離"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   4
               Left            =   3600
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   189
               Top             =   2760
               Value           =   1  'ﾁｪｯｸ
               Width           =   800
            End
            Begin VB.CheckBox chkTaku 
               BackColor       =   &H000000FF&
               Caption         =   "切離"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   5
               Left            =   3600
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   190
               Top             =   3360
               Value           =   1  'ﾁｪｯｸ
               Width           =   800
            End
            Begin VB.Label LblTaku 
               Caption         =   "○○○○○○○○○○○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   191
               Top             =   480
               Width           =   3135
            End
            Begin VB.Label LblTaku 
               Caption         =   "○○○○○○○○○○○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   192
               Top             =   1080
               Width           =   3135
            End
            Begin VB.Label LblTaku 
               Caption         =   "○○○○○○○○○○○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   193
               Top             =   1680
               Width           =   3135
            End
            Begin VB.Label LblTaku 
               Caption         =   "○○○○○○○○○○○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   194
               Top             =   2280
               Width           =   3135
            End
            Begin VB.Label LblTaku 
               Caption         =   "○○○○○○○○○○○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   195
               Top             =   2880
               Width           =   3135
            End
            Begin VB.Label LblTaku 
               Caption         =   "○○○○○○○○○○○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   196
               Top             =   3480
               Width           =   3135
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "通信状態"
            Height          =   4020
            Left            =   4920
            TabIndex        =   197
            Top             =   360
            Width           =   2055
            Begin VB.Label lblTakuSts 
               Caption         =   "○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   360
               TabIndex        =   198
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label lblTakuSts 
               Caption         =   "○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   360
               TabIndex        =   199
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblTakuSts 
               Caption         =   "○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   360
               TabIndex        =   200
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label lblTakuSts 
               Caption         =   "○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   360
               TabIndex        =   201
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label lblTakuSts 
               Caption         =   "○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   360
               TabIndex        =   202
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label lblTakuSts 
               Caption         =   "○○"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   360
               TabIndex        =   203
               Top             =   3480
               Width           =   975
            End
         End
      End
      Begin TabDlg.SSTab tabCorner 
         Height          =   7530
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   13282
         _Version        =   393216
         Tabs            =   6
         Tab             =   1
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
         TabCaption(0)   =   " ○○○○○○ ○○○○○○"
         TabPicture(0)   =   "通信切断画面.frx":0070
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraConerGouki(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   " ○○○○○○ ○○○○○○"
         TabPicture(1)   =   "通信切断画面.frx":008C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraConerGouki(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " ○○○○○○ ○○○○○○"
         TabPicture(2)   =   "通信切断画面.frx":00A8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraConerGouki(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   " ○○○○○○ ○○○○○○"
         TabPicture(3)   =   "通信切断画面.frx":00C4
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraConerGouki(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   " ○○○○○○ ○○○○○○"
         TabPicture(4)   =   "通信切断画面.frx":00E0
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "fraConerGouki(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   " ○○○○○○ ○○○○○○"
         TabPicture(5)   =   "通信切断画面.frx":00FC
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "fraConerGouki(5)"
         Tab(5).ControlCount=   1
         Begin VB.Frame fraConerGouki 
            Caption         =   "改札機号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   5
            Left            =   -74835
            TabIndex        =   526
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   5
               Left            =   330
               TabIndex        =   593
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   10
                  Left            =   360
                  TabIndex        =   595
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   11
                  Left            =   2535
                  TabIndex        =   594
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame fraConectSts 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   5
               Left            =   330
               TabIndex        =   560
               Top             =   540
               Width           =   7650
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   95
                  Left            =   6450
                  TabIndex        =   592
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   94
                  Left            =   5550
                  TabIndex        =   591
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   93
                  Left            =   4650
                  TabIndex        =   590
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   92
                  Left            =   3750
                  TabIndex        =   589
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   91
                  Left            =   2850
                  TabIndex        =   588
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   90
                  Left            =   1950
                  TabIndex        =   587
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   89
                  Left            =   1050
                  TabIndex        =   586
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   88
                  Left            =   150
                  TabIndex        =   585
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   87
                  Left            =   6450
                  TabIndex        =   584
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   86
                  Left            =   5550
                  TabIndex        =   583
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   85
                  Left            =   4650
                  TabIndex        =   582
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   84
                  Left            =   3750
                  TabIndex        =   581
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   83
                  Left            =   2850
                  TabIndex        =   580
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   82
                  Left            =   1950
                  TabIndex        =   579
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   81
                  Left            =   1050
                  TabIndex        =   578
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   80
                  Left            =   150
                  TabIndex        =   577
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   95
                  Left            =   6450
                  TabIndex        =   576
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   94
                  Left            =   5550
                  TabIndex        =   575
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   93
                  Left            =   4650
                  TabIndex        =   574
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   92
                  Left            =   3750
                  TabIndex        =   573
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   91
                  Left            =   2850
                  TabIndex        =   572
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   90
                  Left            =   1950
                  TabIndex        =   571
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   89
                  Left            =   1050
                  TabIndex        =   570
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   88
                  Left            =   150
                  TabIndex        =   569
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   87
                  Left            =   6450
                  TabIndex        =   568
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   86
                  Left            =   5550
                  TabIndex        =   567
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   85
                  Left            =   4650
                  TabIndex        =   566
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   84
                  Left            =   3750
                  TabIndex        =   565
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   83
                  Left            =   2850
                  TabIndex        =   564
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   82
                  Left            =   1950
                  TabIndex        =   563
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   81
                  Left            =   1050
                  TabIndex        =   562
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   80
                  Left            =   150
                  TabIndex        =   561
                  Top             =   360
                  Width           =   900
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   5
               Left            =   330
               TabIndex        =   527
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   95
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   543
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   94
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   542
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   93
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   541
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   92
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   540
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   91
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   539
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   90
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   538
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   89
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   537
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   88
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   536
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   87
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   535
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   86
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   534
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   85
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   533
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   84
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   532
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   83
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   531
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   82
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   530
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   81
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   529
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   80
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   528
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   80
                  Left            =   150
                  TabIndex        =   559
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   81
                  Left            =   1050
                  TabIndex        =   558
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   82
                  Left            =   1950
                  TabIndex        =   557
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   83
                  Left            =   2850
                  TabIndex        =   556
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   84
                  Left            =   3750
                  TabIndex        =   555
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   85
                  Left            =   4650
                  TabIndex        =   554
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   86
                  Left            =   5550
                  TabIndex        =   553
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   87
                  Left            =   6450
                  TabIndex        =   552
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   88
                  Left            =   150
                  TabIndex        =   551
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   89
                  Left            =   1050
                  TabIndex        =   550
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   90
                  Left            =   1950
                  TabIndex        =   549
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   91
                  Left            =   2850
                  TabIndex        =   548
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   92
                  Left            =   3750
                  TabIndex        =   547
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   93
                  Left            =   4650
                  TabIndex        =   546
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   94
                  Left            =   5550
                  TabIndex        =   545
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   95
                  Left            =   6450
                  TabIndex        =   544
                  Top             =   1185
                  Width           =   900
               End
            End
         End
         Begin VB.Frame fraConerGouki 
            Caption         =   "改札機号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   4
            Left            =   -74835
            TabIndex        =   456
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   4
               Left            =   330
               TabIndex        =   523
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   8
                  Left            =   360
                  TabIndex        =   525
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   9
                  Left            =   2535
                  TabIndex        =   524
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame fraConectSts 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   4
               Left            =   330
               TabIndex        =   490
               Top             =   540
               Width           =   7650
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   79
                  Left            =   6450
                  TabIndex        =   522
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   78
                  Left            =   5550
                  TabIndex        =   521
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   77
                  Left            =   4650
                  TabIndex        =   520
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   76
                  Left            =   3750
                  TabIndex        =   519
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   75
                  Left            =   2850
                  TabIndex        =   518
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   74
                  Left            =   1950
                  TabIndex        =   517
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   73
                  Left            =   1050
                  TabIndex        =   516
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   72
                  Left            =   150
                  TabIndex        =   515
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   71
                  Left            =   6450
                  TabIndex        =   514
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   70
                  Left            =   5550
                  TabIndex        =   513
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   69
                  Left            =   4650
                  TabIndex        =   512
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   68
                  Left            =   3750
                  TabIndex        =   511
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   67
                  Left            =   2850
                  TabIndex        =   510
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   66
                  Left            =   1950
                  TabIndex        =   509
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   65
                  Left            =   1050
                  TabIndex        =   508
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   64
                  Left            =   150
                  TabIndex        =   507
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   79
                  Left            =   6450
                  TabIndex        =   506
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   78
                  Left            =   5550
                  TabIndex        =   505
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   77
                  Left            =   4650
                  TabIndex        =   504
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   76
                  Left            =   3750
                  TabIndex        =   503
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   75
                  Left            =   2850
                  TabIndex        =   502
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   74
                  Left            =   1950
                  TabIndex        =   501
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   73
                  Left            =   1050
                  TabIndex        =   500
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   72
                  Left            =   150
                  TabIndex        =   499
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   71
                  Left            =   6450
                  TabIndex        =   498
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   70
                  Left            =   5550
                  TabIndex        =   497
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   69
                  Left            =   4650
                  TabIndex        =   496
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   68
                  Left            =   3750
                  TabIndex        =   495
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   67
                  Left            =   2850
                  TabIndex        =   494
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   66
                  Left            =   1950
                  TabIndex        =   493
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   65
                  Left            =   1050
                  TabIndex        =   492
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   64
                  Left            =   150
                  TabIndex        =   491
                  Top             =   360
                  Width           =   900
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   4
               Left            =   330
               TabIndex        =   457
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   79
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   473
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   78
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   472
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   77
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   471
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   76
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   470
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   75
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   469
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   74
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   468
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   73
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   467
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   72
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   466
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   71
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   465
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   70
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   464
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   69
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   463
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   68
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   462
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   67
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   461
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   66
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   460
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   65
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   459
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   64
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   458
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   64
                  Left            =   150
                  TabIndex        =   489
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   65
                  Left            =   1050
                  TabIndex        =   488
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   66
                  Left            =   1950
                  TabIndex        =   487
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   67
                  Left            =   2850
                  TabIndex        =   486
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   68
                  Left            =   3750
                  TabIndex        =   485
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   69
                  Left            =   4650
                  TabIndex        =   484
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   70
                  Left            =   5550
                  TabIndex        =   483
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   71
                  Left            =   6450
                  TabIndex        =   482
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   72
                  Left            =   150
                  TabIndex        =   481
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   73
                  Left            =   1050
                  TabIndex        =   480
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   74
                  Left            =   1950
                  TabIndex        =   479
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   75
                  Left            =   2850
                  TabIndex        =   478
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   76
                  Left            =   3750
                  TabIndex        =   477
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   77
                  Left            =   4650
                  TabIndex        =   476
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   78
                  Left            =   5550
                  TabIndex        =   475
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   79
                  Left            =   6450
                  TabIndex        =   474
                  Top             =   1185
                  Width           =   900
               End
            End
         End
         Begin VB.Frame fraConerGouki 
            Caption         =   "改札機号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   3
            Left            =   -74835
            TabIndex        =   386
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   3
               Left            =   330
               TabIndex        =   423
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   48
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   439
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   49
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   438
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   50
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   437
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   51
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   436
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   52
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   435
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   53
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   434
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   54
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   433
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   55
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   432
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   56
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   431
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   57
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   430
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   58
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   429
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   59
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   428
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   60
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   427
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   61
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   426
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   62
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   425
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   63
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   424
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   63
                  Left            =   6450
                  TabIndex        =   455
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   62
                  Left            =   5550
                  TabIndex        =   454
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   61
                  Left            =   4650
                  TabIndex        =   453
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   60
                  Left            =   3750
                  TabIndex        =   452
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   59
                  Left            =   2850
                  TabIndex        =   451
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   58
                  Left            =   1950
                  TabIndex        =   450
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   57
                  Left            =   1050
                  TabIndex        =   449
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   56
                  Left            =   150
                  TabIndex        =   448
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   55
                  Left            =   6450
                  TabIndex        =   447
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   54
                  Left            =   5550
                  TabIndex        =   446
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   4650
                  TabIndex        =   445
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   3750
                  TabIndex        =   444
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   2850
                  TabIndex        =   443
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   1950
                  TabIndex        =   442
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   1050
                  TabIndex        =   441
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   150
                  TabIndex        =   440
                  Top             =   345
                  Width           =   900
               End
            End
            Begin VB.Frame fraConectSts 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   3
               Left            =   330
               TabIndex        =   390
               Top             =   540
               Width           =   7650
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   150
                  TabIndex        =   422
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   1050
                  TabIndex        =   421
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   1950
                  TabIndex        =   420
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   2850
                  TabIndex        =   419
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   3750
                  TabIndex        =   418
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   4650
                  TabIndex        =   417
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   54
                  Left            =   5550
                  TabIndex        =   416
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   55
                  Left            =   6450
                  TabIndex        =   415
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   56
                  Left            =   150
                  TabIndex        =   414
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   57
                  Left            =   1050
                  TabIndex        =   413
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   58
                  Left            =   1950
                  TabIndex        =   412
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   59
                  Left            =   2850
                  TabIndex        =   411
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   60
                  Left            =   3750
                  TabIndex        =   410
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   61
                  Left            =   4650
                  TabIndex        =   409
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   62
                  Left            =   5550
                  TabIndex        =   408
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   63
                  Left            =   6450
                  TabIndex        =   407
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   150
                  TabIndex        =   406
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   1050
                  TabIndex        =   405
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   1950
                  TabIndex        =   404
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   2850
                  TabIndex        =   403
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   3750
                  TabIndex        =   402
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   4650
                  TabIndex        =   401
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   54
                  Left            =   5550
                  TabIndex        =   400
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   55
                  Left            =   6450
                  TabIndex        =   399
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   56
                  Left            =   150
                  TabIndex        =   398
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   57
                  Left            =   1050
                  TabIndex        =   397
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   58
                  Left            =   1950
                  TabIndex        =   396
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   59
                  Left            =   2850
                  TabIndex        =   395
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   60
                  Left            =   3750
                  TabIndex        =   394
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   61
                  Left            =   4650
                  TabIndex        =   393
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   62
                  Left            =   5550
                  TabIndex        =   392
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   63
                  Left            =   6450
                  TabIndex        =   391
                  Top             =   1560
                  Width           =   900
               End
            End
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   3
               Left            =   330
               TabIndex        =   387
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   7
                  Left            =   2535
                  TabIndex        =   389
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   6
                  Left            =   360
                  TabIndex        =   388
                  Top             =   360
                  Width           =   1815
               End
            End
         End
         Begin VB.Frame fraConerGouki 
            Caption         =   "改札機号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   2
            Left            =   -74835
            TabIndex        =   316
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   2
               Left            =   330
               TabIndex        =   383
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   4
                  Left            =   360
                  TabIndex        =   385
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   5
                  Left            =   2535
                  TabIndex        =   384
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame fraConectSts 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   2
               Left            =   330
               TabIndex        =   350
               Top             =   540
               Width           =   7650
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   6450
                  TabIndex        =   382
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   46
                  Left            =   5550
                  TabIndex        =   381
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   45
                  Left            =   4650
                  TabIndex        =   380
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   44
                  Left            =   3750
                  TabIndex        =   379
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   43
                  Left            =   2850
                  TabIndex        =   378
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   42
                  Left            =   1950
                  TabIndex        =   377
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   41
                  Left            =   1050
                  TabIndex        =   376
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   40
                  Left            =   150
                  TabIndex        =   375
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   39
                  Left            =   6450
                  TabIndex        =   374
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   38
                  Left            =   5550
                  TabIndex        =   373
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   37
                  Left            =   4650
                  TabIndex        =   372
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   36
                  Left            =   3750
                  TabIndex        =   371
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   35
                  Left            =   2850
                  TabIndex        =   370
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   34
                  Left            =   1950
                  TabIndex        =   369
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   1050
                  TabIndex        =   368
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   150
                  TabIndex        =   367
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   6450
                  TabIndex        =   366
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   46
                  Left            =   5550
                  TabIndex        =   365
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   45
                  Left            =   4650
                  TabIndex        =   364
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   44
                  Left            =   3750
                  TabIndex        =   363
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   43
                  Left            =   2850
                  TabIndex        =   362
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   42
                  Left            =   1950
                  TabIndex        =   361
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   41
                  Left            =   1050
                  TabIndex        =   360
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   40
                  Left            =   150
                  TabIndex        =   359
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   39
                  Left            =   6450
                  TabIndex        =   358
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   38
                  Left            =   5550
                  TabIndex        =   357
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   37
                  Left            =   4650
                  TabIndex        =   356
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   36
                  Left            =   3750
                  TabIndex        =   355
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   35
                  Left            =   2850
                  TabIndex        =   354
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   34
                  Left            =   1950
                  TabIndex        =   353
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   1050
                  TabIndex        =   352
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   150
                  TabIndex        =   351
                  Top             =   360
                  Width           =   900
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   2
               Left            =   330
               TabIndex        =   317
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   47
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   333
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   46
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   332
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   45
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   331
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   44
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   330
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   43
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   329
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   42
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   328
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   41
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   327
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   40
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   326
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   39
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   325
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   38
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   324
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   37
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   323
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   36
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   322
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   35
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   321
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   34
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   320
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   33
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   319
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   32
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   318
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   150
                  TabIndex        =   349
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   1050
                  TabIndex        =   348
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   34
                  Left            =   1950
                  TabIndex        =   347
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   35
                  Left            =   2850
                  TabIndex        =   346
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   36
                  Left            =   3750
                  TabIndex        =   345
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   37
                  Left            =   4650
                  TabIndex        =   344
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   38
                  Left            =   5550
                  TabIndex        =   343
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   39
                  Left            =   6450
                  TabIndex        =   342
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   40
                  Left            =   150
                  TabIndex        =   341
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   41
                  Left            =   1050
                  TabIndex        =   340
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   42
                  Left            =   1950
                  TabIndex        =   339
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   43
                  Left            =   2850
                  TabIndex        =   338
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   44
                  Left            =   3750
                  TabIndex        =   337
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   45
                  Left            =   4650
                  TabIndex        =   336
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   46
                  Left            =   5550
                  TabIndex        =   335
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   6450
                  TabIndex        =   334
                  Top             =   1185
                  Width           =   900
               End
            End
         End
         Begin VB.Frame fraConerGouki 
            Caption         =   "改札機号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   1
            Left            =   165
            TabIndex        =   75
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   1
               Left            =   330
               TabIndex        =   112
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   16
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   128
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   17
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   127
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   18
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   126
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   19
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   125
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   20
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   124
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   21
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   123
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   22
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   122
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   23
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   121
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   24
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   120
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   25
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   119
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   26
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   118
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   27
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   117
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   28
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   116
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   29
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   115
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   30
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   114
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   31
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   113
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   6450
                  TabIndex        =   144
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   30
                  Left            =   5550
                  TabIndex        =   143
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   4650
                  TabIndex        =   142
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   28
                  Left            =   3750
                  TabIndex        =   141
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   27
                  Left            =   2850
                  TabIndex        =   140
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   26
                  Left            =   1950
                  TabIndex        =   139
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   25
                  Left            =   1050
                  TabIndex        =   138
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   24
                  Left            =   150
                  TabIndex        =   137
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   136
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   22
                  Left            =   5550
                  TabIndex        =   135
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   21
                  Left            =   4650
                  TabIndex        =   134
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   20
                  Left            =   3750
                  TabIndex        =   133
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   2850
                  TabIndex        =   132
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   1950
                  TabIndex        =   131
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   1050
                  TabIndex        =   130
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   150
                  TabIndex        =   129
                  Top             =   345
                  Width           =   900
               End
            End
            Begin VB.Frame fraConectSts 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   1
               Left            =   330
               TabIndex        =   79
               Top             =   540
               Width           =   7650
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   150
                  TabIndex        =   111
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   1050
                  TabIndex        =   110
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   1950
                  TabIndex        =   109
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   2850
                  TabIndex        =   108
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   20
                  Left            =   3750
                  TabIndex        =   107
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   21
                  Left            =   4650
                  TabIndex        =   106
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   22
                  Left            =   5550
                  TabIndex        =   105
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   104
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   24
                  Left            =   150
                  TabIndex        =   103
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   25
                  Left            =   1050
                  TabIndex        =   102
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   26
                  Left            =   1950
                  TabIndex        =   101
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   27
                  Left            =   2850
                  TabIndex        =   100
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   28
                  Left            =   3750
                  TabIndex        =   99
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   4650
                  TabIndex        =   98
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   30
                  Left            =   5550
                  TabIndex        =   97
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   6450
                  TabIndex        =   96
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   150
                  TabIndex        =   95
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   1050
                  TabIndex        =   94
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   1950
                  TabIndex        =   93
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   2850
                  TabIndex        =   92
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   20
                  Left            =   3750
                  TabIndex        =   91
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   21
                  Left            =   4650
                  TabIndex        =   90
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   22
                  Left            =   5550
                  TabIndex        =   89
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   88
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   24
                  Left            =   150
                  TabIndex        =   87
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   25
                  Left            =   1050
                  TabIndex        =   86
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   26
                  Left            =   1950
                  TabIndex        =   85
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   27
                  Left            =   2850
                  TabIndex        =   84
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   28
                  Left            =   3750
                  TabIndex        =   83
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   4650
                  TabIndex        =   82
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   30
                  Left            =   5550
                  TabIndex        =   81
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   6450
                  TabIndex        =   80
                  Top             =   1560
                  Width           =   900
               End
            End
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   1
               Left            =   330
               TabIndex        =   76
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   3
                  Left            =   2535
                  TabIndex        =   78
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   2
                  Left            =   360
                  TabIndex        =   77
                  Top             =   360
                  Width           =   1815
               End
            End
         End
         Begin VB.Frame fraConerGouki 
            Caption         =   "改札機号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   0
            Left            =   -74835
            TabIndex        =   5
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   0
               Left            =   330
               TabIndex        =   42
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   0
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   58
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   1
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   57
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   2
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   56
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   3
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   55
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   4
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   54
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   5
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   53
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   6
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   52
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   7
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   51
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   8
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   50
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   9
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   49
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   10
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   48
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   11
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   47
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   12
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   46
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   13
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   45
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   14
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   44
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkJikai 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   15
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   43
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   6450
                  TabIndex        =   74
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   5550
                  TabIndex        =   73
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   4650
                  TabIndex        =   72
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   3750
                  TabIndex        =   71
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   2850
                  TabIndex        =   70
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   1950
                  TabIndex        =   69
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   1050
                  TabIndex        =   68
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   150
                  TabIndex        =   67
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   6450
                  TabIndex        =   66
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   5550
                  TabIndex        =   65
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   4650
                  TabIndex        =   64
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   3750
                  TabIndex        =   63
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2850
                  TabIndex        =   62
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1950
                  TabIndex        =   61
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   60
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   59
                  Top             =   345
                  Width           =   900
               End
            End
            Begin VB.Frame fraConectSts 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   0
               Left            =   330
               TabIndex        =   9
               Top             =   540
               Width           =   7650
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   41
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   40
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1950
                  TabIndex        =   39
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2850
                  TabIndex        =   38
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   3750
                  TabIndex        =   37
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   4650
                  TabIndex        =   36
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   5550
                  TabIndex        =   35
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   6450
                  TabIndex        =   34
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   150
                  TabIndex        =   33
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   1050
                  TabIndex        =   32
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   1950
                  TabIndex        =   31
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   2850
                  TabIndex        =   30
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   3750
                  TabIndex        =   29
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   4650
                  TabIndex        =   28
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   5550
                  TabIndex        =   27
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   6450
                  TabIndex        =   26
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   25
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   24
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1950
                  TabIndex        =   23
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2850
                  TabIndex        =   22
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   3750
                  TabIndex        =   21
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   4650
                  TabIndex        =   20
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   5550
                  TabIndex        =   19
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   6450
                  TabIndex        =   18
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   150
                  TabIndex        =   17
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   1050
                  TabIndex        =   16
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   1950
                  TabIndex        =   15
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   2850
                  TabIndex        =   14
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   3750
                  TabIndex        =   13
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   4650
                  TabIndex        =   12
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   5550
                  TabIndex        =   11
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   6450
                  TabIndex        =   10
                  Top             =   1560
                  Width           =   900
               End
            End
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   0
               Left            =   330
               TabIndex        =   6
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   1
                  Left            =   2535
                  TabIndex        =   8
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutJikai 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   0
                  Left            =   360
                  TabIndex        =   7
                  Top             =   360
                  Width           =   1815
               End
            End
         End
      End
      Begin TabDlg.SSTab tabIcmCorner 
         Height          =   7530
         Left            =   120
         TabIndex        =   145
         Top             =   600
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   13282
         _Version        =   393216
         Tabs            =   6
         Tab             =   5
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
         TabCaption(0)   =   " ○○○○○○ ○○○○○○"
         TabPicture(0)   =   "通信切断画面.frx":0118
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraICMGouki(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   " ○○○○○○ ○○○○○○"
         TabPicture(1)   =   "通信切断画面.frx":0134
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraICMGouki(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " ○○○○○○ ○○○○○○"
         TabPicture(2)   =   "通信切断画面.frx":0150
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraICMGouki(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   " ○○○○○○ ○○○○○○"
         TabPicture(3)   =   "通信切断画面.frx":016C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraICMGouki(10)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   " ○○○○○○ ○○○○○○"
         TabPicture(4)   =   "通信切断画面.frx":0188
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "fraICMGouki(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   " ○○○○○○ ○○○○○○"
         TabPicture(5)   =   "通信切断画面.frx":01A4
         Tab(5).ControlEnabled=   -1  'True
         Tab(5).Control(0)=   "fraICMGouki(5)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).ControlCount=   1
         Begin VB.Frame fraICMGouki 
            Caption         =   "ICM号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   5
            Left            =   165
            TabIndex        =   656
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   11
               Left            =   330
               TabIndex        =   204
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   11
                  Left            =   2535
                  TabIndex        =   748
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   10
                  Left            =   360
                  TabIndex        =   747
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   5
               Left            =   330
               TabIndex        =   658
               Top             =   540
               Width           =   7650
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   95
                  Left            =   6450
                  TabIndex        =   738
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   94
                  Left            =   5550
                  TabIndex        =   737
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   93
                  Left            =   4650
                  TabIndex        =   736
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   92
                  Left            =   3750
                  TabIndex        =   735
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   91
                  Left            =   2850
                  TabIndex        =   734
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   90
                  Left            =   1950
                  TabIndex        =   733
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   89
                  Left            =   1050
                  TabIndex        =   732
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   88
                  Left            =   150
                  TabIndex        =   731
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   87
                  Left            =   6450
                  TabIndex        =   730
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   86
                  Left            =   5550
                  TabIndex        =   729
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   85
                  Left            =   4650
                  TabIndex        =   728
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   84
                  Left            =   3750
                  TabIndex        =   727
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   83
                  Left            =   2850
                  TabIndex        =   726
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   82
                  Left            =   1950
                  TabIndex        =   725
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   81
                  Left            =   1050
                  TabIndex        =   724
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   80
                  Left            =   150
                  TabIndex        =   723
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   95
                  Left            =   6450
                  TabIndex        =   674
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   94
                  Left            =   5550
                  TabIndex        =   673
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   93
                  Left            =   4650
                  TabIndex        =   672
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   92
                  Left            =   3750
                  TabIndex        =   671
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   91
                  Left            =   2850
                  TabIndex        =   670
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   90
                  Left            =   1950
                  TabIndex        =   669
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   89
                  Left            =   1050
                  TabIndex        =   668
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   88
                  Left            =   150
                  TabIndex        =   667
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   87
                  Left            =   6450
                  TabIndex        =   666
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   86
                  Left            =   5550
                  TabIndex        =   665
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   85
                  Left            =   4650
                  TabIndex        =   664
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   84
                  Left            =   3750
                  TabIndex        =   663
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   83
                  Left            =   2850
                  TabIndex        =   662
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   82
                  Left            =   1950
                  TabIndex        =   661
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   81
                  Left            =   1050
                  TabIndex        =   660
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   80
                  Left            =   150
                  TabIndex        =   659
                  Top             =   360
                  Width           =   900
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   12
               Left            =   330
               TabIndex        =   657
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   95
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   908
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   94
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   907
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   93
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   906
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   92
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   905
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   91
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   904
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   90
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   903
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   89
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   902
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   88
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   901
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   87
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   900
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   86
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   899
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   85
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   898
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   84
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   897
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   83
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   896
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   82
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   895
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   81
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   894
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   80
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   893
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   95
                  Left            =   6450
                  TabIndex        =   892
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   94
                  Left            =   5550
                  TabIndex        =   891
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   93
                  Left            =   4650
                  TabIndex        =   890
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   92
                  Left            =   3750
                  TabIndex        =   889
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   91
                  Left            =   2850
                  TabIndex        =   888
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   90
                  Left            =   1950
                  TabIndex        =   887
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   89
                  Left            =   1050
                  TabIndex        =   886
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   88
                  Left            =   150
                  TabIndex        =   885
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   87
                  Left            =   6450
                  TabIndex        =   884
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   86
                  Left            =   5550
                  TabIndex        =   883
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   85
                  Left            =   4650
                  TabIndex        =   882
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   84
                  Left            =   3750
                  TabIndex        =   881
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   83
                  Left            =   2850
                  TabIndex        =   880
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   82
                  Left            =   1950
                  TabIndex        =   879
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   81
                  Left            =   1050
                  TabIndex        =   878
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   80
                  Left            =   150
                  TabIndex        =   877
                  Top             =   345
                  Width           =   900
               End
            End
         End
         Begin VB.Frame fraICMGouki 
            Caption         =   "ICM号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   4
            Left            =   -74835
            TabIndex        =   636
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   10
               Left            =   330
               TabIndex        =   655
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   9
                  Left            =   2535
                  TabIndex        =   746
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   8
                  Left            =   360
                  TabIndex        =   745
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   4
               Left            =   330
               TabIndex        =   638
               Top             =   540
               Width           =   7650
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   79
                  Left            =   6450
                  TabIndex        =   722
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   78
                  Left            =   5550
                  TabIndex        =   721
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   77
                  Left            =   4650
                  TabIndex        =   720
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   76
                  Left            =   3750
                  TabIndex        =   719
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   75
                  Left            =   2850
                  TabIndex        =   718
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   74
                  Left            =   1950
                  TabIndex        =   717
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   73
                  Left            =   1050
                  TabIndex        =   716
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   72
                  Left            =   150
                  TabIndex        =   715
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   71
                  Left            =   6450
                  TabIndex        =   714
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   70
                  Left            =   5550
                  TabIndex        =   713
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   69
                  Left            =   4650
                  TabIndex        =   712
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   68
                  Left            =   3750
                  TabIndex        =   711
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   67
                  Left            =   2850
                  TabIndex        =   710
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   66
                  Left            =   1950
                  TabIndex        =   709
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   65
                  Left            =   1050
                  TabIndex        =   708
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   64
                  Left            =   150
                  TabIndex        =   707
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   79
                  Left            =   6450
                  TabIndex        =   654
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   78
                  Left            =   5550
                  TabIndex        =   653
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   77
                  Left            =   4650
                  TabIndex        =   652
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   76
                  Left            =   3750
                  TabIndex        =   651
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   75
                  Left            =   2850
                  TabIndex        =   650
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   74
                  Left            =   1950
                  TabIndex        =   649
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   73
                  Left            =   1050
                  TabIndex        =   648
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   72
                  Left            =   150
                  TabIndex        =   647
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   71
                  Left            =   6450
                  TabIndex        =   646
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   70
                  Left            =   5550
                  TabIndex        =   645
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   69
                  Left            =   4650
                  TabIndex        =   644
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   68
                  Left            =   3750
                  TabIndex        =   643
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   67
                  Left            =   2850
                  TabIndex        =   642
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   66
                  Left            =   1950
                  TabIndex        =   641
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   65
                  Left            =   1050
                  TabIndex        =   640
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   64
                  Left            =   150
                  TabIndex        =   639
                  Top             =   360
                  Width           =   900
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   11
               Left            =   330
               TabIndex        =   637
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   79
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   876
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   78
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   875
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   77
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   874
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   76
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   873
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   75
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   872
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   74
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   871
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   73
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   870
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   72
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   869
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   71
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   868
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   70
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   867
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   69
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   866
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   68
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   865
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   67
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   864
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   66
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   863
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   65
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   862
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   64
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   861
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   79
                  Left            =   6450
                  TabIndex        =   860
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   78
                  Left            =   5550
                  TabIndex        =   859
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   77
                  Left            =   4650
                  TabIndex        =   858
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   76
                  Left            =   3750
                  TabIndex        =   857
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   75
                  Left            =   2850
                  TabIndex        =   856
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   74
                  Left            =   1950
                  TabIndex        =   855
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   73
                  Left            =   1050
                  TabIndex        =   854
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   72
                  Left            =   150
                  TabIndex        =   853
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   71
                  Left            =   6450
                  TabIndex        =   852
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   70
                  Left            =   5550
                  TabIndex        =   851
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   69
                  Left            =   4650
                  TabIndex        =   850
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   68
                  Left            =   3750
                  TabIndex        =   849
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   67
                  Left            =   2850
                  TabIndex        =   848
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   66
                  Left            =   1950
                  TabIndex        =   847
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   65
                  Left            =   1050
                  TabIndex        =   846
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   64
                  Left            =   150
                  TabIndex        =   845
                  Top             =   345
                  Width           =   900
               End
            End
         End
         Begin VB.Frame fraICMGouki 
            Caption         =   "ICM号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   10
            Left            =   -74835
            TabIndex        =   616
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   9
               Left            =   330
               TabIndex        =   635
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   7
                  Left            =   2535
                  TabIndex        =   744
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   6
                  Left            =   360
                  TabIndex        =   743
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   3
               Left            =   330
               TabIndex        =   618
               Top             =   540
               Width           =   7650
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   63
                  Left            =   6450
                  TabIndex        =   706
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   62
                  Left            =   5550
                  TabIndex        =   705
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   61
                  Left            =   4650
                  TabIndex        =   704
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   60
                  Left            =   3750
                  TabIndex        =   703
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   59
                  Left            =   2850
                  TabIndex        =   702
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   58
                  Left            =   1950
                  TabIndex        =   701
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   57
                  Left            =   1050
                  TabIndex        =   700
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   56
                  Left            =   150
                  TabIndex        =   699
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   55
                  Left            =   6450
                  TabIndex        =   698
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   54
                  Left            =   5550
                  TabIndex        =   697
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   4650
                  TabIndex        =   696
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   3750
                  TabIndex        =   695
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   2850
                  TabIndex        =   694
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   1950
                  TabIndex        =   693
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   1050
                  TabIndex        =   692
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   150
                  TabIndex        =   691
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   63
                  Left            =   6450
                  TabIndex        =   634
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   62
                  Left            =   5550
                  TabIndex        =   633
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   61
                  Left            =   4650
                  TabIndex        =   632
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   60
                  Left            =   3750
                  TabIndex        =   631
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   59
                  Left            =   2850
                  TabIndex        =   630
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   58
                  Left            =   1950
                  TabIndex        =   629
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   57
                  Left            =   1050
                  TabIndex        =   628
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   56
                  Left            =   150
                  TabIndex        =   627
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   55
                  Left            =   6450
                  TabIndex        =   626
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   54
                  Left            =   5550
                  TabIndex        =   625
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   4650
                  TabIndex        =   624
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   3750
                  TabIndex        =   623
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   2850
                  TabIndex        =   622
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   1950
                  TabIndex        =   621
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   1050
                  TabIndex        =   620
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   150
                  TabIndex        =   619
                  Top             =   360
                  Width           =   900
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   10
               Left            =   330
               TabIndex        =   617
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   63
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   844
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   62
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   843
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   61
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   842
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   60
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   841
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   59
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   840
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   58
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   839
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   57
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   838
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   56
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   837
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   55
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   836
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   54
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   835
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   53
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   834
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   52
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   833
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   51
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   832
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   50
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   831
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   49
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   830
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   48
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   829
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   63
                  Left            =   6450
                  TabIndex        =   828
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   62
                  Left            =   5550
                  TabIndex        =   827
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   61
                  Left            =   4650
                  TabIndex        =   826
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   60
                  Left            =   3750
                  TabIndex        =   825
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   59
                  Left            =   2850
                  TabIndex        =   824
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   58
                  Left            =   1950
                  TabIndex        =   823
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   57
                  Left            =   1050
                  TabIndex        =   822
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   56
                  Left            =   150
                  TabIndex        =   821
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   55
                  Left            =   6450
                  TabIndex        =   820
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   54
                  Left            =   5550
                  TabIndex        =   819
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   4650
                  TabIndex        =   818
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   3750
                  TabIndex        =   817
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   2850
                  TabIndex        =   816
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   1950
                  TabIndex        =   815
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   1050
                  TabIndex        =   814
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   150
                  TabIndex        =   813
                  Top             =   345
                  Width           =   900
               End
            End
         End
         Begin VB.Frame fraICMGouki 
            Caption         =   "ICM号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   2
            Left            =   -74835
            TabIndex        =   596
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   9
               Left            =   330
               TabIndex        =   615
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   47
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   812
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   46
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   811
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   45
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   810
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   44
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   809
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   43
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   808
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   42
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   807
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   41
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   806
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   40
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   805
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   39
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   804
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   38
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   803
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   37
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   802
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   36
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   801
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   35
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   800
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   34
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   799
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   33
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   798
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   32
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   797
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   6450
                  TabIndex        =   796
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   46
                  Left            =   5550
                  TabIndex        =   795
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   45
                  Left            =   4650
                  TabIndex        =   794
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   44
                  Left            =   3750
                  TabIndex        =   793
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   43
                  Left            =   2850
                  TabIndex        =   792
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   42
                  Left            =   1950
                  TabIndex        =   791
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   41
                  Left            =   1050
                  TabIndex        =   790
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   40
                  Left            =   150
                  TabIndex        =   789
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   39
                  Left            =   6450
                  TabIndex        =   788
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   38
                  Left            =   5550
                  TabIndex        =   787
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   37
                  Left            =   4650
                  TabIndex        =   786
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   36
                  Left            =   3750
                  TabIndex        =   785
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   35
                  Left            =   2850
                  TabIndex        =   784
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   34
                  Left            =   1950
                  TabIndex        =   783
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   1050
                  TabIndex        =   782
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   150
                  TabIndex        =   781
                  Top             =   345
                  Width           =   900
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   2
               Left            =   330
               TabIndex        =   598
               Top             =   540
               Width           =   7650
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   6450
                  TabIndex        =   690
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   46
                  Left            =   5550
                  TabIndex        =   689
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   45
                  Left            =   4650
                  TabIndex        =   688
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   44
                  Left            =   3750
                  TabIndex        =   687
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   43
                  Left            =   2850
                  TabIndex        =   686
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   42
                  Left            =   1950
                  TabIndex        =   685
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   41
                  Left            =   1050
                  TabIndex        =   684
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   40
                  Left            =   150
                  TabIndex        =   683
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   39
                  Left            =   6450
                  TabIndex        =   682
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   38
                  Left            =   5550
                  TabIndex        =   681
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   37
                  Left            =   4650
                  TabIndex        =   680
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   36
                  Left            =   3750
                  TabIndex        =   679
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   35
                  Left            =   2850
                  TabIndex        =   678
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   34
                  Left            =   1950
                  TabIndex        =   677
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   1050
                  TabIndex        =   676
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   150
                  TabIndex        =   675
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   150
                  TabIndex        =   614
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   1050
                  TabIndex        =   613
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   34
                  Left            =   1950
                  TabIndex        =   612
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   35
                  Left            =   2850
                  TabIndex        =   611
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   36
                  Left            =   3750
                  TabIndex        =   610
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   37
                  Left            =   4650
                  TabIndex        =   609
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   38
                  Left            =   5550
                  TabIndex        =   608
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   39
                  Left            =   6450
                  TabIndex        =   607
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   40
                  Left            =   150
                  TabIndex        =   606
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   41
                  Left            =   1050
                  TabIndex        =   605
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   42
                  Left            =   1950
                  TabIndex        =   604
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   43
                  Left            =   2850
                  TabIndex        =   603
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   44
                  Left            =   3750
                  TabIndex        =   602
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   45
                  Left            =   4650
                  TabIndex        =   601
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   46
                  Left            =   5550
                  TabIndex        =   600
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   6450
                  TabIndex        =   599
                  Top             =   1200
                  Width           =   900
               End
            End
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   8
               Left            =   330
               TabIndex        =   597
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   5
                  Left            =   2535
                  TabIndex        =   742
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   4
                  Left            =   360
                  TabIndex        =   741
                  Top             =   360
                  Width           =   1815
               End
            End
         End
         Begin VB.Frame fraICMGouki 
            Caption         =   "ICM号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   1
            Left            =   -74835
            TabIndex        =   280
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   19
               Left            =   330
               TabIndex        =   315
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   31
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   780
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   30
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   779
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   29
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   778
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   28
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   777
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   27
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   776
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   26
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   775
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   25
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   774
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   24
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   773
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   23
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   772
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   22
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   771
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   21
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   770
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   20
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   769
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   19
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   768
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   18
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   767
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   17
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   766
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   16
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   765
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   6450
                  TabIndex        =   764
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   30
                  Left            =   5550
                  TabIndex        =   763
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   4650
                  TabIndex        =   762
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   28
                  Left            =   3750
                  TabIndex        =   761
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   27
                  Left            =   2850
                  TabIndex        =   760
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   26
                  Left            =   1950
                  TabIndex        =   759
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   25
                  Left            =   1050
                  TabIndex        =   758
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   24
                  Left            =   150
                  TabIndex        =   757
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   756
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   22
                  Left            =   5550
                  TabIndex        =   755
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   21
                  Left            =   4650
                  TabIndex        =   754
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   20
                  Left            =   3750
                  TabIndex        =   753
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   2850
                  TabIndex        =   752
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   1950
                  TabIndex        =   751
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   1050
                  TabIndex        =   750
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   150
                  TabIndex        =   749
                  Top             =   345
                  Width           =   900
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   1
               Left            =   330
               TabIndex        =   282
               Top             =   540
               Width           =   7650
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   150
                  TabIndex        =   314
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   1050
                  TabIndex        =   313
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   1950
                  TabIndex        =   312
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   2850
                  TabIndex        =   311
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   20
                  Left            =   3750
                  TabIndex        =   310
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   21
                  Left            =   4650
                  TabIndex        =   309
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   22
                  Left            =   5550
                  TabIndex        =   308
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   307
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   24
                  Left            =   150
                  TabIndex        =   306
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   25
                  Left            =   1050
                  TabIndex        =   305
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   26
                  Left            =   1950
                  TabIndex        =   304
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   27
                  Left            =   2850
                  TabIndex        =   303
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   28
                  Left            =   3750
                  TabIndex        =   302
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   4650
                  TabIndex        =   301
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   30
                  Left            =   5550
                  TabIndex        =   300
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   6450
                  TabIndex        =   299
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   150
                  TabIndex        =   298
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   1050
                  TabIndex        =   297
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   1950
                  TabIndex        =   296
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   2850
                  TabIndex        =   295
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   20
                  Left            =   3750
                  TabIndex        =   294
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   21
                  Left            =   4650
                  TabIndex        =   293
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   22
                  Left            =   5550
                  TabIndex        =   292
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   291
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   24
                  Left            =   150
                  TabIndex        =   290
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   25
                  Left            =   1050
                  TabIndex        =   289
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   26
                  Left            =   1950
                  TabIndex        =   288
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   27
                  Left            =   2850
                  TabIndex        =   287
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   28
                  Left            =   3750
                  TabIndex        =   286
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   4650
                  TabIndex        =   285
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   30
                  Left            =   5550
                  TabIndex        =   284
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   6450
                  TabIndex        =   283
                  Top             =   1560
                  Width           =   900
               End
            End
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   7
               Left            =   330
               TabIndex        =   281
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   3
                  Left            =   2535
                  TabIndex        =   740
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   2
                  Left            =   360
                  TabIndex        =   739
                  Top             =   360
                  Width           =   1815
               End
            End
         End
         Begin VB.Frame fraICMGouki 
            Caption         =   "ICM号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   0
            Left            =   -74835
            TabIndex        =   210
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   6
               Left            =   330
               TabIndex        =   277
               Top             =   2775
               Width           =   7650
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   0
                  Left            =   360
                  TabIndex        =   279
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton cmdInOutICM 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   1
                  Left            =   2535
                  TabIndex        =   278
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "通信状態"
               ForeColor       =   &H00FF0000&
               Height          =   1980
               Index           =   0
               Left            =   330
               TabIndex        =   244
               Top             =   540
               Width           =   7650
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   6450
                  TabIndex        =   276
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   5550
                  TabIndex        =   275
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   4650
                  TabIndex        =   274
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   3750
                  TabIndex        =   273
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   2850
                  TabIndex        =   272
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   1950
                  TabIndex        =   271
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   1050
                  TabIndex        =   270
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   150
                  TabIndex        =   269
                  Top             =   1560
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   6450
                  TabIndex        =   268
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   5550
                  TabIndex        =   267
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   4650
                  TabIndex        =   266
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   3750
                  TabIndex        =   265
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2850
                  TabIndex        =   264
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1950
                  TabIndex        =   263
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   262
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGoukiConectSts 
                  Alignment       =   2  '中央揃え
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   261
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   6450
                  TabIndex        =   260
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   5550
                  TabIndex        =   259
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   4650
                  TabIndex        =   258
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   3750
                  TabIndex        =   257
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   2850
                  TabIndex        =   256
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   1950
                  TabIndex        =   255
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   1050
                  TabIndex        =   254
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   150
                  TabIndex        =   253
                  Top             =   1200
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   6450
                  TabIndex        =   252
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   5550
                  TabIndex        =   251
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   4650
                  TabIndex        =   250
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   3750
                  TabIndex        =   249
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2850
                  TabIndex        =   248
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1950
                  TabIndex        =   247
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   246
                  Top             =   360
                  Width           =   900
               End
               Begin VB.Label lblICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   245
                  Top             =   360
                  Width           =   900
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   14
               Left            =   330
               TabIndex        =   211
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   15
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   227
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   14
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   226
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   13
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   225
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   12
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   224
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   11
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   223
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   10
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   222
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   9
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   221
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   8
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   220
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   7
                  Left            =   6500
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   219
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   6
                  Left            =   5600
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   218
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   5
                  Left            =   4700
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   217
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   4
                  Left            =   3800
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   216
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   3
                  Left            =   2900
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   215
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   2
                  Left            =   2000
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   214
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   1
                  Left            =   1100
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   213
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICM 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   0
                  Left            =   200
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   212
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   243
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   242
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1950
                  TabIndex        =   241
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2850
                  TabIndex        =   240
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   3750
                  TabIndex        =   239
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   4650
                  TabIndex        =   238
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   5550
                  TabIndex        =   237
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   6450
                  TabIndex        =   236
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   150
                  TabIndex        =   235
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   1050
                  TabIndex        =   234
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   1950
                  TabIndex        =   233
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   2850
                  TabIndex        =   232
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   3750
                  TabIndex        =   231
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   4650
                  TabIndex        =   230
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   5550
                  TabIndex        =   229
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label lblTargetICMGouki 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   6450
                  TabIndex        =   228
                  Top             =   1185
                  Width           =   900
               End
            End
         End
         Begin VB.Frame fraICMGouki 
            Caption         =   "判定IC-M号機"
            ForeColor       =   &H00FF0000&
            Height          =   6525
            Index           =   3
            Left            =   -74835
            TabIndex        =   146
            Top             =   735
            Width           =   8295
            Begin VB.Frame fraALLSitei 
               Caption         =   "一括指定"
               ForeColor       =   &H00FF0000&
               Height          =   1095
               Index           =   30
               Left            =   9000
               TabIndex        =   207
               Top             =   0
               Width           =   7650
               Begin VB.CommandButton CmdIcmSetteing 
                  Caption         =   "全号機接続"
                  Height          =   495
                  Index           =   7
                  Left            =   360
                  TabIndex        =   209
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.CommandButton CmdIcmSetteing 
                  Caption         =   "全号機切離"
                  Height          =   495
                  Index           =   6
                  Left            =   2535
                  TabIndex        =   208
                  Top             =   360
                  Width           =   1815
               End
            End
            Begin VB.Frame fraLogGouki 
               Caption         =   "指定号機"
               ForeColor       =   &H00FF0000&
               Height          =   2160
               Index           =   30
               Left            =   345
               TabIndex        =   147
               Top             =   4110
               Width           =   7650
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   63
                  Left            =   6480
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   163
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   62
                  Left            =   5580
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   162
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   61
                  Left            =   4680
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   161
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   60
                  Left            =   3780
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   160
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   59
                  Left            =   2880
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   159
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   58
                  Left            =   1980
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   158
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   57
                  Left            =   1080
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   157
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   56
                  Left            =   180
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   156
                  Top             =   1455
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   55
                  Left            =   6480
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   155
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   54
                  Left            =   5580
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   154
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   53
                  Left            =   4680
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   153
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   52
                  Left            =   3780
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   152
                  Top             =   615
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   51
                  Left            =   2880
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   151
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   50
                  Left            =   1980
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   150
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   49
                  Left            =   1080
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   149
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkICMGouki 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   48
                  Left            =   180
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   148
                  Top             =   615
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   63
                  Left            =   150
                  TabIndex        =   179
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   62
                  Left            =   1050
                  TabIndex        =   178
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   61
                  Left            =   1950
                  TabIndex        =   177
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   60
                  Left            =   2850
                  TabIndex        =   176
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   59
                  Left            =   3750
                  TabIndex        =   175
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   58
                  Left            =   4650
                  TabIndex        =   174
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   57
                  Left            =   5550
                  TabIndex        =   173
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   56
                  Left            =   6450
                  TabIndex        =   172
                  Top             =   345
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   55
                  Left            =   150
                  TabIndex        =   171
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   54
                  Left            =   1050
                  TabIndex        =   170
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   1950
                  TabIndex        =   169
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   2850
                  TabIndex        =   168
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   3750
                  TabIndex        =   167
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   4650
                  TabIndex        =   166
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   5550
                  TabIndex        =   165
                  Top             =   1185
                  Width           =   900
               End
               Begin VB.Label LblGokiNo 
                  Alignment       =   2  '中央揃え
                  Caption         =   "Z9"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   6450
                  TabIndex        =   164
                  Top             =   1185
                  Width           =   900
               End
            End
         End
      End
      Begin TabDlg.SSTab tabKikiCorner 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   205
         Top             =   600
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   13573
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   794
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "統合監視盤"
         TabPicture(0)   =   "通信切断画面.frx":01C0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraOver(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "操作卓"
         TabPicture(1)   =   "通信切断画面.frx":01DC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraOver(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fraOver 
            Caption         =   "上位機器一覧"
            ForeColor       =   &H00FF0000&
            Height          =   6840
            Index           =   1
            Left            =   -74760
            TabIndex        =   942
            Top             =   600
            Width           =   8175
            Begin VB.Frame Frame5 
               Caption         =   "上位機器名称"
               Height          =   6300
               Index           =   1
               Left            =   240
               TabIndex        =   944
               Top             =   360
               Width           =   6495
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   19
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   964
                  Top             =   5760
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   18
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   963
                  Top             =   5160
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   17
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   962
                  Top             =   4560
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   16
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   961
                  Top             =   3960
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   15
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   960
                  Top             =   3360
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   14
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   959
                  Top             =   2760
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   13
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   958
                  Top             =   2160
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   12
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   957
                  Top             =   1560
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   11
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   956
                  Top             =   960
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   10
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   955
                  Top             =   360
                  Width           =   800
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   240
                  TabIndex        =   954
                  Top             =   5880
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   240
                  TabIndex        =   953
                  Top             =   5280
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   240
                  TabIndex        =   952
                  Top             =   4680
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   240
                  TabIndex        =   951
                  Top             =   4080
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   240
                  TabIndex        =   950
                  Top             =   3480
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   240
                  TabIndex        =   949
                  Top             =   2880
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   240
                  TabIndex        =   948
                  Top             =   2280
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   240
                  TabIndex        =   947
                  Top             =   1680
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   240
                  TabIndex        =   946
                  Top             =   1080
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "データ集計機（○○○○○○○○○○○○）"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   240
                  TabIndex        =   945
                  Top             =   480
                  Width           =   5055
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "通信状態"
               Height          =   6300
               Index           =   1
               Left            =   6840
               TabIndex        =   943
               Top             =   360
               Width           =   1215
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   10
                  Left            =   360
                  TabIndex        =   974
                  Top             =   480
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   19
                  Left            =   360
                  TabIndex        =   973
                  Top             =   5880
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   18
                  Left            =   360
                  TabIndex        =   972
                  Top             =   5280
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   17
                  Left            =   360
                  TabIndex        =   971
                  Top             =   4680
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   16
                  Left            =   360
                  TabIndex        =   970
                  Top             =   4080
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   15
                  Left            =   360
                  TabIndex        =   969
                  Top             =   3480
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   14
                  Left            =   360
                  TabIndex        =   968
                  Top             =   2880
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   13
                  Left            =   360
                  TabIndex        =   967
                  Top             =   2280
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   12
                  Left            =   360
                  TabIndex        =   966
                  Top             =   1680
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   11
                  Left            =   360
                  TabIndex        =   965
                  Top             =   1080
                  Width           =   700
               End
            End
         End
         Begin VB.Frame fraOver 
            Caption         =   "上位機器一覧"
            ForeColor       =   &H00FF0000&
            Height          =   6840
            Index           =   0
            Left            =   240
            TabIndex        =   909
            Top             =   600
            Width           =   8175
            Begin VB.Frame Frame5 
               Caption         =   "上位機器名称"
               Height          =   6300
               Index           =   0
               Left            =   240
               TabIndex        =   921
               Top             =   360
               Width           =   6495
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   0
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   931
                  Top             =   360
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H0080FF80&
                  Caption         =   "接続"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   1
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   930
                  Top             =   960
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   2
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   929
                  Top             =   1560
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   3
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   928
                  Top             =   2160
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   4
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   927
                  Top             =   2760
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   5
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   926
                  Top             =   3360
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   6
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   925
                  Top             =   3960
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   7
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   924
                  Top             =   4560
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   8
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   923
                  Top             =   5160
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.CheckBox chkKIKI 
                  BackColor       =   &H000000FF&
                  Caption         =   "切離"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   9
                  Left            =   5520
                  Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
                  TabIndex        =   922
                  Top             =   5760
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   800
               End
               Begin VB.Label KikiName 
                  Caption         =   "データ集計機（○○○○○○○○○○○○）"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   941
                  Top             =   480
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   940
                  Top             =   1080
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   939
                  Top             =   1680
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   938
                  Top             =   2280
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   937
                  Top             =   2880
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   240
                  TabIndex        =   936
                  Top             =   3480
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   240
                  TabIndex        =   935
                  Top             =   4080
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   240
                  TabIndex        =   934
                  Top             =   4680
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   240
                  TabIndex        =   933
                  Top             =   5280
                  Width           =   5055
               End
               Begin VB.Label KikiName 
                  Caption         =   "○○○○○○○○○○○○○○○○○○○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   240
                  TabIndex        =   932
                  Top             =   5880
                  Width           =   5055
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "通信状態"
               Height          =   6300
               Index           =   0
               Left            =   6840
               TabIndex        =   910
               Top             =   360
               Width           =   1215
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   0
                  Left            =   360
                  TabIndex        =   920
                  Top             =   480
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   1
                  Left            =   360
                  TabIndex        =   919
                  Top             =   1080
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   2
                  Left            =   360
                  TabIndex        =   918
                  Top             =   1680
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   3
                  Left            =   360
                  TabIndex        =   917
                  Top             =   2280
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   4
                  Left            =   360
                  TabIndex        =   916
                  Top             =   2880
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   5
                  Left            =   360
                  TabIndex        =   915
                  Top             =   3480
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   6
                  Left            =   360
                  TabIndex        =   914
                  Top             =   4080
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   7
                  Left            =   360
                  TabIndex        =   913
                  Top             =   4680
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   8
                  Left            =   360
                  TabIndex        =   912
                  Top             =   5280
                  Width           =   700
               End
               Begin VB.Label lblOverSts 
                  Caption         =   "○○"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   9
                  Left            =   360
                  TabIndex        =   911
                  Top             =   5880
                  Width           =   700
               End
            End
         End
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "通信接続・切断"
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
      TabIndex        =   206
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmConectSts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmConectSts.frm
'//  パッケージ名：通信接続・切断画面
'//
'//  概要：通信接続・切断画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.17.0.1) 2010-01-05   REVISED BY [TCC] S.Terao
'//                 不具合修正
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.1.0.1) 2011-11-16  CODED BY  [TCC] M.Matsumoto
'//                 EG20フェーズ３対応【非常通信断SW対応】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-237】
'//     REVISIONS :(EG20 V6.8.0.1) 2012-08-28 REVISED BY  [TCC] H.Sugimoto
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const CONECTSTS_NORMAL = 0             '正常
Private Const CONECTSTS_ERROR = 1              '異常
Private Const CONECTSTS_END = 2                '切離
Private Const GET_CONECTSTS_ERROR = 3          'ブランク

Private Const CONECT_NORMAL = "正常"
Private Const CONECT_ERROR = "異常"
Private Const CONECT_END = "切離"
Private Const GET_CONECT_ERROR = " "

Private Const MN_MAIL_INTERVAL = 1000           'メールタイマのインターバル値

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'*****************************************************************************
'*      ループカウンタ定義値
'*****************************************************************************
Private Const CONECT_HANTEI_ICM_CONTROLMAX = 95  ' 判定ＩＣ−Ｍ 最大コントロール数（インデックス）
Private Const CONECT_JIKAI_CONTROLMAX = 95       ' 改札機号機 最大コントロール数（インデックス）
Private Const CONECT_KIKI_CONTROLMAX = 19        ' 上位機器 最大コントロール数（インデックス）
Private Const CONECT_TAKU_CONTROLMAX = 5         ' 操作卓 最大コントロール数（インデックス）
Private Const CONTROL_CORNERMAX = 16             ' １コーナあたりのコントロール最大数（自改、ＩＣＭ）
Private Const CONTROL_KIKICORNERMAX = 10         ' １コーナあたりのコントロール最大数（上位機器）

Private Const CONECT_KIKI_TABKANSHI = 1        ' 上位機器タブ：統合監視盤
Private Const CONECT_KIKI_TABTAKU = 2          ' 上位機器タブ：操作卓

' EG20 V2.1.0.1[Mainte_03_01] 追加終了

'EG20 V3.1.0.1【非常通信断SW対応】ADD START
Private Const KINKYU_SW_OFF = 0                 '緊急通信断SW：OFF
Private Const KINKYU_SW_ON = 1                  '緊急通信断SW：ON
'EG20 V3.1.0.1 ADD END

' Private iSendID(0 To 9) As Integer              '送信先ID                 ' EG20 V2.1.0.1[Mainte_03_01] 削除
Private iSendID(0 To CONECT_KIKI_CONTROLMAX) As Integer     ' 送信先ID      ' EG20 V2.1.0.1[Mainte_03_01] 追加
' Private iJikaiType(0 To 17) As Integer          '自改タイプ(0：未設置。1：EGR。2：NEG)    ' EG20 V2.1.0.1[Mainte_03_01] 削除
Private iUpDataFlag As Integer                  '表示更新釦押下フラグ
Private iALLGoukiFlag As Integer                '全号機釦押下フラグ
Private iShokiFlag As Integer                   '初期処理フラグ
Private iCancelFlag As Integer                  'キャンセル釦押下フラグ
Private iMailRcvFlag As Integer                 'メール受信フラグ

Private udtMail As MAIL_CONECT_CMD    '通信設定要求CMD

'V1.4.0.1 ADD START
'【処理対象タブ】
Private Const JIKAI = 0               '自改タブ
Private Const ICM = 1                 '判定IC-Mタブ
Private Const KIKI = 2                '上位機器タブ
Private Const TAKU = 3                '操作卓タブ                   ' EG20 V2.1.0.1[Mainte_03_01] 追加
'【通信設定ステータス】
Private Const CONECT_SETU = 1         '接続
Private Const CONECT_DAN = 0          '切断
Private Const JIKAI_CONECT_SETU = 0   '接続                         ' EG20 V2.1.0.1[Mainte_03_01] 追加
Private Const JIKAI_CONECT_DAN = 1    '切断                         ' EG20 V2.1.0.1[Mainte_03_01] 追加
Private Const IDU_CONECT_DAN = 1      '切断
Private Const IDU_CONECT_SETU = 0     '接続
Private Const TAKU_CONECT_SETU = 0    '接続                         ' EG20 V2.1.0.1[Mainte_03_01] 追加
Private Const TAKU_CONECT_DAN = 1     '切断                         ' EG20 V2.1.0.1[Mainte_03_01] 追加
'【IDUアプリ設定ファイル更新ID定義】
Private Const HANTEI_ICM_ID = 26      '判定IC-M通信設定エリアID
Private Const ID_SVR_ID = 34          'IDサーバ通信設定エリアID
Private iKansiAreaId(0 To 9) As Long  '監視エリアID
'【IDU関連設定値保持エリア】
'Private iICM_Sts(0 To 15) As Integer  'ICM号機別設定値             ' EG20 V2.1.0.1[Mainte_03_01] 削除
Private iICM_Sts(0 To 31) As Integer  'ICM号機別設定値              ' EG20 V2.1.0.1[Mainte_03_01] 追加
Private iIDSVR_Sts As Integer         'IDサーバー設定値
Private sBottom_Sts As String         '押下釦ステータス
Private Const ZEN_SITEI = "全号機"    '全号機一括ステータス
Private iSend_Mail As Integer         'メール送信異常による処理かどうかのフラグ
Private Const MAIL_ERROR = 1          'メール送信異常
Private Const MAIL_OK = 0             'メール送信正常
'V1.4.0.1 ADD END

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
' 自動改札機、判定ICM設置構成
Private Type JIKAI_BUTTON_INFO
    bStatus As Boolean              ' 設定有無（TRUE:有り,FALSE:無し）
    nDisGoukiNo As Integer          ' 表示号機
    nDispCornerNo As Integer        ' コーナー番号(表示)
    nCornerNo As Integer            ' コーナー番号(論理)
    nCornerGoukiNo As Integer       ' コーナー別論理号機
    nControlNo As Integer           ' コントロール番号
    nKanshiNo As Integer            ' 監視状態番号
End Type
' 上位機器設定構成
Private Type TRANS_BUTTON_INFO
    bStatus As Boolean              ' 設定有無（TRUE:有り,FALSE:無し）
    sGetInf As String               ' 画面表示用名称
    iAreaID As Integer              ' 対象外部機器上位機器通信状態エリアID
    iSendID As Integer              ' プロセス名（送信機種種別を設定する）
    iKansiId As Integer             ' 監視設定ファイルのエリアID

    nCornerNo As Integer            ' タブ番号（統合監視盤、操作卓）
    nCornerGoukiNo As Integer       ' タブ別論理番号
    nControlNo As Integer           ' コントロール番号
    nIniListNo As Integer           ' 外部機器リスト番号
    nRonriType As Integer           ' 論理タイプ
    nCorner As Integer              ' コーナ番号
End Type

Private gJikaiButtonInfo(0 To MAX_GATE_NO) As JIKAI_BUTTON_INFO
Private gIcmButtonInfo(0 To MAX_GATE_NO) As JIKAI_BUTTON_INFO

Private gTransButtonInfo(0 To CONECT_KIKI_CONTROLMAX) As TRANS_BUTTON_INFO

' EG20 V2.1.0.1[Mainte_03_01] 追加終了

Private mintICMKinkyuSW As Integer  ' 緊急通信断SW      'EG20 V3.1.0.1【非常通信断SW対応】ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 通信接続・切断画面(アクティブ時)
'//  機能概要  : メール受信用、タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
    'タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 通信接続・切断画面(ディアクティブ時)
'//  機能概要  : メール受信用、タイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'タイマを停止する
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 通信接続・切断画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer    'カウンター
    
    On Error Resume Next
    
    '配置設定
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '押下フラグ初期化
    iUpDataFlag = 0
    iALLGoukiFlag = 0
    iCancelFlag = 0
  
    'V2.3.0.1 ADD START
    'IDU縮退チェック
    psIDUCheck
    'V2.3.0.1 ADD END

    '自改タブ
'    For i = CNT_MIN To CONECT_JIKAI_MAX            ' EG20 V2.1.0.1[Mainte_03_01] 削除
    For i = CNT_MIN To CONECT_JIKAI_CONTROLMAX      ' EG20 V2.1.0.1[Mainte_03_01] 追加
     lblGouki(i).Visible = False
     lblGoukiConectSts(i).Visible = False
     lblTargetGouki(i).Visible = False
     chkJikai(i).Visible = False
    Next
    
    '判定IC-Mタブ
'    For i = CNT_MIN To CONECT_HANTEI_ICM_MAX       ' EG20 V2.1.0.1[Mainte_03_01] 削除
    For i = CNT_MIN To CONECT_HANTEI_ICM_CONTROLMAX ' EG20 V2.1.0.1[Mainte_03_01] 追加
     lblICMGouki(i).Visible = False
     lblICMGoukiConectSts(i).Visible = False
     lblTargetICMGouki(i).Visible = False
     chkICM(i).Visible = False
    Next
    
    '上位機器タブ
'    For i = CNT_MIN To CONECT_KIKI_MAX             ' EG20 V2.1.0.1[Mainte_03_01] 削除
    For i = CNT_MIN To CONECT_KIKI_CONTROLMAX       ' EG20 V2.1.0.1[Mainte_03_01] 追加
     KikiName(i).Visible = False
     chkKIKI(i).Visible = False
     lblOverSts(i).Visible = False
    Next

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ' 操作卓タブ
    For i = CNT_MIN To CONECT_TAKU_CONTROLMAX
     LblTaku(i).Visible = False
     chkTaku(i).Visible = False
     lblTakuSts(i).Visible = False
    Next
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

    '自改タブ表示処理
    iShokiFlag = 1
    Call InitCornerTab                              ' EG20 V2.1.0.1[Mainte_03_01] 追加

    psJikaiConectSts
    
    '判定IC-Mタブ表示処理
    psICMConectSts

    '上位機器タブ表示処理
    pfGetKiKiConectSts
    
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    
    ' 操作卓タブ表示処理
    psTakuConectSts
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
    
    '「通信接続・切断画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_STAEND_GAMEN_START, 0)
   
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'    If gStrCurrentForm = sFormName_EJVer Then
'         cmdCancel.Caption = " EG-R自動改札機   バージョン管理 画面へ戻る"
'    ElseIf gStrCurrentForm = sFormName_NJVer Then
'         cmdCancel.Caption = " NEG自動改札機    バージョン管理 画面へ戻る"
'    Else
'         '何もしない
'    End If
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    If gStrCurrentForm = sFormName_GateVerUpdate Then
         cmdCancel.Caption = "自動改札機バージョン一括更新 画面へ戻る"
    ElseIf gStrCurrentForm = sFormName_EJVer Then
         cmdCancel.Caption = "自動改札機バージョン管理 画面へ戻る"
    Else
        cmdCancel.Caption = "通信確認・表示 画面へ戻る"
    End If
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
   
    tabConect.Tab = 0
    iShokiFlag = 0
 
    'V2.3.0.1 ADD START
    '通信接続・切断画面 押下不可処理(判定IC-M)
    If pbIDUSts = 1 Then
      tabConect.TabEnabled(1) = False
    End If
    'V2.3.0.1 ADD END

   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdDataUp_Click
'//  機能名称  : 「表示更新」釦押下時処理
'//  機能概要  : 画面を最新情報にて表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdDataUp_Click()
  
  On Error Resume Next

  iUpDataFlag = 1
  
   '「通信接続・切断画面：表示更新釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)

  '自改タブ表示処理
  psJikaiConectSts
    
  '判定IC-Mタブ表示処理
  psICMConectSts
    
  '上位機器タブ表示処理
  pfGetKiKiConectSts
  
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
  ' 操作卓タブ表示処理
  psTakuConectSts
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
  
  iUpDataFlag = 0

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdCancel_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
   On Error Resume Next
   
   '「通信接続・切断画面：消去」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_STAEND_GAMEN_END, 0)
 
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdJikai_Conect_Click
'//  機能名称  : 「接続」「切離」釦押下時処理
'//  機能概要  : 釦名称処理を行う。
'//              自改部：
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub chkJikai_Click(Index As Integer)
    Dim bRet As Boolean               'メール送信戻り値
    Dim iResponse As Integer          'メッセージの戻り値
    Dim iCnt As Integer               'カウンター
    Dim lSts As Long                  'ステータス値
    Dim bFlag As Boolean              '受信メールフラグ
    Dim lngErrCode As Long            'エラーコード
    Dim nInfoIndex As Integer         ' 保存情報インデックス    ' EG20 V2.1.0.1 追加
    On Error Resume Next
    
    bFlag = True                      'V1.4.0.1　ADD
    
    If iUpDataFlag <> 0 Or iALLGoukiFlag <> 0 Or _
       iShokiFlag = 1 Or iCancelFlag = 1 Or iMailRcvFlag = 1 Then
       iCancelFlag = 0
       Exit Sub
    End If
     
    '画面をロックする。
    SetEnableFalse (0)

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ' 押下釦に対応した自動改札機構成を検索
    nInfoIndex = 0
    For iCnt = 0 To MAX_GATE_NO - 1
        If gJikaiButtonInfo(iCnt).nControlNo = Index Then
            nInfoIndex = iCnt
            Exit For
        End If
    Next
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

     If chkJikai(Index).Value = 0 Then
         '「通信接続・切断画面：切離→接続 設定」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMEND_TOSTA_BUTTOM, 0)

         '切離→接続
         '「通信接続確認」ポップアップ画面表示
         iResponse = MsgBox("指定した外部機器との通信接続を開始します。よろしいですか？", _
                             vbOKCancel + vbQuestion, _
                             "通信接続確認")
         If iResponse = vbOK Then
            '通信設定要求CMD(自改,接続,対象号機)を監マプロセスに送信する
            chkJikai(Index).Caption = "接続"
            chkJikai(Index).BackColor = CONECT_ON
            'ヘッダ部共通作成処理
            SendMailHeader
            udtMail.dwRequestKIKI = ML_DT_JIKAI
            udtMail.dwRequestConectType = ML_REQUEST_CONECT
            For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
              udtMail.dwGouki(iCnt) = ML_TARGET_OFF
            Next
'            udtMail.dwGouki(Index) = ML_TARGET_ON                                  ' EG20 V2.1.0.1[Mainte_03_01] 削除
            udtMail.dwGouki(gJikaiButtonInfo(nInfoIndex).nKanshiNo - 1) = ML_TARGET_ON ' EG20 V2.1.0.1[Mainte_03_01] 追加
           
           'V1.4.0.1 ADD START
           If CheckAppStart(PROC_KANRI) = 0 Then
              '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
'              bRet = pfSetSettei(JIKAI, Index, CONECT_SETU, IdGate.JIKAI_CONECT_SETTEI)    ' EG20 V2.1.0.1[Mainte_03_01] 削除
              bRet = pfSetSettei(JIKAI, gJikaiButtonInfo(nInfoIndex).nKanshiNo - 1, _
                                 JIKAI_CONECT_SETU, IdGate.JIKAI_CONECT_SETTEI)                   ' EG20 V2.1.0.1[Mainte_03_01] 追加
              bFlag = False
              GoTo Error_Click
           End If
           'V1.4.0.1 ADD END

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを表示する
            Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

            bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
            If False = bRet Then
               '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
               '送信異常時：画面ロック解除
               GoTo Error_Click
            End If
               '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
               Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
         Else
            '「キャンセル」釦押下時
            GoTo Error_Click
         End If
     End If
    
     If chkJikai(Index).Value = 1 Then
        '「通信接続・切断画面：切離→接続 設定」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMSTA_TOEND_BUTTOM, 0)
 
        '「通信切断確認」ポップアップ画面表示
        'V1.8.0.1 DEL START
'        iResponse = MsgBox("指定した外部機器との通信切断を開始します。よろしいですか？", _
'                           vbOKCancel + vbQuestion, _
'                           "通信切断確認")
        'V1.8.0.1 DEL END
        'V1.8.0.1 ADD START
        iResponse = MsgBox("指定した外部機器との通信を切断します。よろしいですか？", _
                           vbOKCancel + vbQuestion, _
                           "通信切断確認")
        'V1.8.0.1 ADD END
        If iResponse = vbOK Then
           '通信設定要求CMD(自改,接続,対象号機)を監マプロセスに送信する
           chkJikai(Index).Caption = "切離"
           chkJikai(Index).BackColor = CONECT_OFF
           'ヘッダ部共通作成処理
           SendMailHeader
           udtMail.dwRequestKIKI = ML_DT_JIKAI
           udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
           For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
             udtMail.dwGouki(iCnt) = ML_TARGET_OFF
           Next
'           udtMail.dwGouki(Index) = ML_TARGET_ON                                   ' EG20 V2.1.0.1[Mainte_03_01] 削除
           udtMail.dwGouki(gJikaiButtonInfo(nInfoIndex).nKanshiNo - 1) = ML_TARGET_ON ' EG20 V2.1.0.1[Mainte_03_01] 追加
     
           'V1.4.0.1 ADD START
           If CheckAppStart(PROC_KANRI) = 0 Then
              '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
'              bRet = pfSetSettei(JIKAI, Index, CONECT_DAN, IdGate.JIKAI_CONECT_SETTEI)     ' EG20 V2.1.0.1[Mainte_03_01] 削除
              bRet = pfSetSettei(JIKAI, gJikaiButtonInfo(nInfoIndex).nKanshiNo - 1, _
                                 JIKAI_CONECT_DAN, IdGate.JIKAI_CONECT_SETTEI)              ' EG20 V2.1.0.1[Mainte_03_01] 追加
              bFlag = False
              GoTo Error_Click
           End If
           'V1.4.0.1 ADD END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを表示する
            Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           
             bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
             If False = bRet Then
                '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
                lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
                '送信異常時：画面ロック解除
                GoTo Error_Click
             End If
                '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
        Else
           '「キャンセル」釦押下時
           GoTo Error_Click
        End If
    End If
Exit Sub

'キャンセル釦押下、又は送信異常時
Error_Click:
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
  
  If bFlag = True Then                  'V1.4.0.1 ADD
     If chkJikai(Index).Value = 0 Then
        '「キャンセル」釦押下時
        SetEnableTrue (0)
        iCancelFlag = 1
        chkJikai(Index).Caption = "切離"
        chkJikai(Index).BackColor = CONECT_OFF
        chkJikai(Index).Value = 1
        Exit Sub
    End If
   If chkJikai(Index).Value = 1 Then
      SetEnableTrue (0)
      iCancelFlag = 1
      chkJikai(Index).Caption = "接続"
      chkJikai(Index).BackColor = CONECT_ON
      chkJikai(Index).Value = 0
      Exit Sub
   End If
'V1.4.0.1 ADD START
  Else
   SetEnableTrue (0)
   iShokiFlag = 1
   psJikaiConectSts
   iShokiFlag = 0
  End If
'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdInOutJikai_Click
'//  機能名称  : 「全号機接続」「全号機切離」釦押下時処理
'//  機能概要  : 釦名称処理を行う。
'//              自改部：
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V1.2.1.0) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-237】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdInOutJikai_Click(Index As Integer)
  Dim iCnt As Integer               'カウンター
  Dim iResponse As Integer          'メッセージボックスの戻り値
  Dim bRet As Boolean               'メール送信戻り値
  Dim lngErrCode As Long            'エラーコード
  Dim bytWork()   As Byte
  Dim i As Integer
  Dim bFlag As Boolean              'エラーフラグ処理　'V1.4.0.1 ADD

  Dim bInOutStatus As Boolean       ' 押下した釦（TRUE:接続,FALSE:切断）        ' EG20 V1.2.1.0[Mainte_03_01] 追加
  Erase bytWork
  
  On Error Resume Next
 
  '画面をロックする。
  SetEnableFalse (0)

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    Select Case Index
    Case 0, 2, 4, 6, 8, 10
        bInOutStatus = True     ' 全号機接続
    Case Else
        bInOutStatus = False    ' 全号機切断
    End Select
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

'  If Index = 0 Then            ' EG20 V2.1.0.1[Mainte_03_01] 削除
  If bInOutStatus = True Then   ' EG20 V2.1.0.1[Mainte_03_01] 追加
     '「通信接続・切断画面：全号機接続釦押下」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_ALLGOUKI_STA_BUTTOM, 0)
 
     '「通信接続確認」ポップアップ画面表示
     iResponse = MsgBox("指定した外部機器との通信接続を開始します。よろしいですか？", _
                         vbOKCancel + vbQuestion, _
                         "通信接続確認")
     If iResponse = vbOK Then
        '「全号機接続」釦押下時
        iALLGoukiFlag = 1
        '通信設定要求CMD(自改,接続)を監マプロセスに送信する
        'ヘッダ部共通作成処理
        SendMailHeader
        udtMail.dwRequestKIKI = ML_DT_JIKAI
        udtMail.dwRequestConectType = ML_REQUEST_CONECT
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'           If chkEGR.Value = 1 And iJikaiType(iCnt) = GATE_JISEDAI Then
'              chkJikai(iCnt).Caption = "接続"
'              chkJikai(iCnt).BackColor = CONECT_ON
'              chkJikai(iCnt).Value = 0
'              udtMail.dwGouki(iCnt) = ML_TARGET_ON
'           ElseIf chkNEG.Value = 1 And iJikaiType(iCnt) = GATE_NGATE Then
'              chkJikai(iCnt).Caption = "接続"
'              chkJikai(iCnt).BackColor = CONECT_ON
'              chkJikai(iCnt).Value = 0
'              udtMail.dwGouki(iCnt) = ML_TARGET_ON
'           ElseIf chkNEG.Value = 0 And chkEGR.Value = 0 Then
'               'NEG自改/EG-R自改チェック無し時
'                SetEnableTrue (0)
'                psJikaiConectSts
'                Exit Sub
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'           If gJikaiButtonInfo(iCnt).bStatus = True Then               ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
           ' 設定されている号機、かつ選択されたコーナに対して設定
           If gJikaiButtonInfo(iCnt).bStatus = True And _
               gJikaiButtonInfo(iCnt).nCornerNo = tabCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
              chkJikai(gJikaiButtonInfo(iCnt).nControlNo).Caption = "接続"
              chkJikai(gJikaiButtonInfo(iCnt).nControlNo).BackColor = CONECT_ON
              chkJikai(gJikaiButtonInfo(iCnt).nControlNo).Value = 0
              udtMail.dwGouki(gJikaiButtonInfo(iCnt).nKanshiNo - 1) = ML_TARGET_ON
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
           Else
'              udtMail.dwGouki(iCnt) = ML_TARGET_OFF                            ' EG20 V2.1.0.1[Mainte_03_01] 削除
              udtMail.dwGouki(gJikaiButtonInfo(iCnt).nKanshiNo - 1) = ML_TARGET_OFF ' EG20 V2.1.0.1[Mainte_03_01] 追加
           End If
        Next
        
        'V1.4.0.1 ADD START
        If CheckAppStart(PROC_KANRI) = 0 Then
           '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
           For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'              '号機別釦が表示されているかどうかチェック(設定がないものに対して変更を行わないように)
'              If chkJikai(iCnt).Visible = True Then
'                 'EG-R自改のみ更新
'                 If chkEGR.Value = 1 And iJikaiType(iCnt) = GATE_JISEDAI Then
'                    bRet = pfSetSettei(JIKAI, iCnt, CONECT_SETU, IdGate.JIKAI_CONECT_SETTEI)
'                 'NEG自改のみ更新
'                 ElseIf chkNEG.Value = 1 And iJikaiType(iCnt) = GATE_NGATE Then
'                    bRet = pfSetSettei(JIKAI, iCnt, CONECT_SETU, IdGate.JIKAI_CONECT_SETTEI)
'                 End If
'              End If
'              Index = Index + 1
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'              If gJikaiButtonInfo(iCnt).bStatus = True Then               ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
              ' 設定されている号機、かつ選択されたコーナに対して設定
              If gJikaiButtonInfo(iCnt).bStatus = True And _
                    gJikaiButtonInfo(iCnt).nCornerNo = tabCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
                  bRet = pfSetSettei(JIKAI, gJikaiButtonInfo(iCnt).nKanshiNo - 1, _
                                     JIKAI_CONECT_SETU, IdGate.JIKAI_CONECT_SETTEI)
              End If
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
           Next
           bFlag = False
           GoTo Error_Click
        End If
        'V1.4.0.1 ADD END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
        If False = bRet Then
           '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
           lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
           Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
           GoTo Error_Click
           Exit Sub
        End If
         '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
     Else
        '「キャンセル」釦押下時
         GoTo Error_Click
     End If
 Else
    '「通信接続・切断画面：全号機切離釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_ALLGOUKI_END_BUTTOM, 0)
 
    '「通信切断確認」ポップアップ画面表示
   'V1.8.0.1 DEL START
'    iResponse = MsgBox("指定した外部機器との通信切断を開始します。よろしいですか？", _
'                        vbOKCancel + vbQuestion, _
'                        "通信切断確認")
   'V1.8.0.1 DEL END
   'V1.8.0.1 ADD START
     iResponse = MsgBox("指定した外部機器との通信を切断します。よろしいですか？", _
                         vbOKCancel + vbQuestion, _
                         "通信切断確認")
   'V1.8.0.1 ADD END
     If iResponse = vbOK Then
        iALLGoukiFlag = 1
        '通信設定要求CMD(自改,切断)を監マプロセスに送信する
        'ヘッダ部共通作成処理
        SendMailHeader
        udtMail.dwRequestKIKI = ML_DT_JIKAI
        udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
        '「全号機切離」釦押下時
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'            If chkEGR.Value = 1 And iJikaiType(iCnt) = GATE_JISEDAI Then
'               chkJikai(iCnt).Caption = "切離"
'               chkJikai(iCnt).BackColor = CONECT_OFF
'               chkJikai(iCnt).Value = 1
'               udtMail.dwGouki(iCnt) = ML_TARGET_ON
'            ElseIf chkNEG.Value = 1 And iJikaiType(iCnt) = GATE_NGATE Then
'               chkJikai(iCnt).Caption = "切離"
'               chkJikai(iCnt).BackColor = CONECT_OFF
'               chkJikai(iCnt).Value = 1
'               udtMail.dwGouki(iCnt) = ML_TARGET_ON
'            ElseIf chkNEG.Value = 0 And chkEGR.Value = 0 Then
'               'NEG自改/EG-R自改チェック無し時
'                SetEnableTrue (0)
'                psJikaiConectSts
'                Exit Sub
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'           If gJikaiButtonInfo(iCnt).bStatus = True Then               ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
           ' 設定されている号機、かつ選択されたコーナに対して設定
           If gJikaiButtonInfo(iCnt).bStatus = True And _
                gJikaiButtonInfo(iCnt).nCornerNo = tabCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
               chkJikai(gJikaiButtonInfo(iCnt).nControlNo).Caption = "切離"
               chkJikai(gJikaiButtonInfo(iCnt).nControlNo).BackColor = CONECT_OFF
               chkJikai(gJikaiButtonInfo(iCnt).nControlNo).Value = 1
               udtMail.dwGouki(gJikaiButtonInfo(iCnt).nKanshiNo - 1) = ML_TARGET_ON
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
            Else
'               udtMail.dwGouki(iCnt) = ML_TARGET_OFF                               ' EG20 V2.1.0.1[Mainte_03_01] 削除
               udtMail.dwGouki(gJikaiButtonInfo(iCnt).nKanshiNo - 1) = ML_TARGET_OFF  ' EG20 V2.1.0.1[Mainte_03_01] 追加
            End If
        Next
        
        'V1.4.0.1 ADD START
        If CheckAppStart(PROC_KANRI) = 0 Then
           '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
           For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'               '号機別釦が表示されているかどうかチェック(設定がないものに対して変更を行わないように)
'               If chkJikai(iCnt).Visible = True Then
'                  'EG-R自改のみ更新
'                  If chkEGR.Value = 1 And iJikaiType(iCnt) = GATE_JISEDAI Then
'                     bRet = pfSetSettei(JIKAI, iCnt, CONECT_DAN, IdGate.JIKAI_CONECT_SETTEI)
'                  'NEG自改のみ更新
'                  ElseIf chkNEG.Value = 1 And iJikaiType(iCnt) = GATE_NGATE Then
'                     bRet = pfSetSettei(JIKAI, iCnt, CONECT_DAN, IdGate.JIKAI_CONECT_SETTEI)
'                  End If
'               End If
'               Index = Index + 1
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'               If gJikaiButtonInfo(iCnt).bStatus = True Then               ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
              ' 設定されている号機、かつ選択されたコーナに対して設定
              If gJikaiButtonInfo(iCnt).bStatus = True And _
                    gJikaiButtonInfo(iCnt).nCornerNo = tabCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
                  bRet = pfSetSettei(JIKAI, gJikaiButtonInfo(iCnt).nKanshiNo - 1, _
                                     JIKAI_CONECT_DAN, IdGate.JIKAI_CONECT_SETTEI)
              End If
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
           Next
           bFlag = False
           GoTo Error_Click
        End If
        'V1.4.0.1 ADD END

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
        If False = bRet Then
           '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
           lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
           Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
           GoTo Error_Click
           Exit Sub
        End If
          '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
     Else
       '「キャンセル」釦押下時
        GoTo Error_Click
     End If
 End If
Exit Sub

'キャンセル釦押下、又は送信異常時
Error_Click:

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
 
 If bFlag = True Then                  'V1.4.0.1 ADD
  If Index = 0 Then
     '「キャンセル」釦押下時
     SetEnableTrue (0)
     psJikaiConectSts
     Exit Sub
  Else
     SetEnableTrue (0)
     psJikaiConectSts
     Exit Sub
  End If
'V1.4.0.1 ADD START
Else
    SetEnableTrue (0)
    psJikaiConectSts
    iALLGoukiFlag = 0
End If
'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : chkICM_Click
'//  機能名称  : 「接続」「切離」チェックボックス釦押下時処理
'//  機能概要  : 各チェックボックス処理を行う。
'//              判定IC-M部：
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下チェックボックスインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub chkICM_Click(Index As Integer)
    Dim bRet As Boolean               'メール送信戻り値
    Dim iResponse As Integer          'メッセージの戻り値
    Dim iCnt As Integer               'カウンター
    Dim lSts As Long                  'ステータス値
    Dim bFlag As Boolean              '受信メールフラグ
    Dim lngErrCode As Long            'エラーコード
    Dim nInfoIndex As Integer         ' 保存情報インデックス    ' EG20 V2.1.0.1[Mainte_03_01] 追加
    
    On Error Resume Next
     
   bFlag = True                      'V1.4.0.1　ADD
  
   If iUpDataFlag <> 0 Or iALLGoukiFlag <> 0 Or _
       iShokiFlag = 1 Or iCancelFlag = 1 Or iMailRcvFlag = 1 Then
       iCancelFlag = 0
       Exit Sub
    End If
  
    '画面をロックする。
    SetEnableFalse (1)

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ' 押下釦に対応した自動改札機構成を検索
    nInfoIndex = 0
    For iCnt = 0 To MAX_GATE_NO - 1
        If gIcmButtonInfo(iCnt).nControlNo = Index Then
            nInfoIndex = iCnt
            Exit For
        End If
    Next
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

    If chkICM(Index).Value = 0 Then
        '「通信接続・切断画面：切離→接続 設定」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMEND_TOSTA_BUTTOM, 0)

        '切離→接続
        '「通信接続確認」ポップアップ画面表示
        iResponse = MsgBox("指定した外部機器との通信接続を開始します。よろしいですか？", _
                            vbOKCancel + vbQuestion, _
                            "通信接続確認")
        If iResponse = vbOK Then
           '通信設定要求CMD(ICM,接続,対象号機)をID制に送信する
           chkICM(Index).Caption = "接続"
           chkICM(Index).BackColor = CONECT_ON
           'ヘッダ部共通作成処理
           SendMailHeader
           udtMail.dwRequestKIKI = ML_DT_ICM
           udtMail.dwRequestConectType = ML_REQUEST_CONECT
           For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX
             udtMail.dwGouki(iCnt) = ML_TARGET_OFF
           Next
'            udtMail.dwGouki(Index) = ML_TARGET_ON                                  ' EG20 V2.1.0.1[Mainte_03_01] 削除
            udtMail.dwGouki(gIcmButtonInfo(nInfoIndex).nKanshiNo - 1) = ML_TARGET_ON  ' EG20 V2.1.0.1[Mainte_03_01] 追加
           'V1.4.0.1 ADD START
           'ICM設定保持
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'           iICM_Sts(Index) = CONECT_SETU
'           sBottom_Sts = CStr(Index)
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
           iICM_Sts(gIcmButtonInfo(nInfoIndex).nKanshiNo - 1) = CONECT_SETU
           sBottom_Sts = CStr(gIcmButtonInfo(nInfoIndex).nKanshiNo - 1)
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
           
           If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_PC) = 0 Then
              '監視盤未起動、IDU未起動時(処理対象タブ,号機番号,設定値,設定ID)
              iSend_Mail = MAIL_OK
'              bRet = pfSetSettei(ICM, Index, CONECT_SETU, IdGate.HANTEI_ICM_CONECT_SETTEI)     ' EG20 V2.1.0.1[Mainte_03_01] 削除
              bRet = pfSetSettei(ICM, gIcmButtonInfo(nInfoIndex).nKanshiNo - 1, _
                                 CONECT_SETU, IdGate.HANTEI_ICM_CONECT_SETTEI)                  ' EG20 V2.1.0.1[Mainte_03_01] 追加
              bFlag = False
              GoTo Error_Click
           End If
           'V1.4.0.1 ADD END
           
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
           'プログレスバーを表示する
           Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           
           bRet = DssSendMail(MAIL_SLOT_IDSEI, MlSize.CONECT_CMD, udtMail.mlHeader)
           If False = bRet Then
              '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
              lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
              GoTo Error_Click
           End If
              '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
        Else
           '「キャンセル」釦押下時
           GoTo Error_Click
         End If
     End If
     
     If chkICM(Index).Value = 1 Then
        '「通信接続・切断画面：切離→接続 設定」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMSTA_TOEND_BUTTOM, 0)
 
        '「通信切断確認」ポップアップ画面表示
        'V1.8.0.1 DEL START
'        iResponse = MsgBox("指定した外部機器との通信切断を開始します。よろしいですか？", _
'                           vbOKCancel + vbQuestion, _
'                           "通信切断確認")
        'V1.8.0.1 DEL END
        'V1.8.0.1 ADD START
        iResponse = MsgBox("指定した外部機器との通信を切断します。よろしいですか？", _
                           vbOKCancel + vbQuestion, _
                           "通信切断確認")
        'V1.8.0.1 ADD END
        If iResponse = vbOK Then
           '「OK」釦押下時
           chkICM(Index).Caption = "切離"
           chkICM(Index).BackColor = CONECT_OFF
           '通信設定要求CMD(ICM,切離,対象号機)をID制に送信する
           'ヘッダ部共通作成処理
           SendMailHeader
           udtMail.dwRequestKIKI = ML_DT_ICM
           udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
           For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX
             udtMail.dwGouki(iCnt) = ML_TARGET_OFF
           Next
'           udtMail.dwGouki(Index) = ML_TARGET_ON                                   ' EG20 V2.1.0.1[Mainte_03_01] 削除
           udtMail.dwGouki(gIcmButtonInfo(nInfoIndex).nKanshiNo - 1) = ML_TARGET_ON   ' EG20 V2.1.0.1[Mainte_03_01] 追加
           
           'V1.4.0.1 ADD START
           'ICM設定保持
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'           iICM_Sts(Index) = CONECT_DAN
'           sBottom_Sts = CStr(Index)
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
           iICM_Sts(gIcmButtonInfo(nInfoIndex).nKanshiNo - 1) = CONECT_DAN
           sBottom_Sts = CStr(gIcmButtonInfo(nInfoIndex).nKanshiNo - 1)
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

           If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_PC) = 0 Then
              '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
              iSend_Mail = MAIL_OK
'              bRet = pfSetSettei(ICM, Index, CONECT_DAN, IdGate.HANTEI_ICM_CONECT_SETTEI)      ' EG20 V2.1.0.1[Mainte_03_01] 削除
              bRet = pfSetSettei(ICM, gIcmButtonInfo(nInfoIndex).nKanshiNo - 1, _
                                 CONECT_DAN, IdGate.HANTEI_ICM_CONECT_SETTEI)                   ' EG20 V2.1.0.1[Mainte_03_01] 追加
              bFlag = False
              GoTo Error_Click
           End If
           'V1.4.0.1 ADD END
           
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
           'プログレスバーを表示する
           Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           
           bRet = DssSendMail(MAIL_SLOT_IDSEI, MlSize.CONECT_CMD, udtMail.mlHeader)
           If False = bRet Then
              '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
              lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
              '送信異常時：画面ロック解除
              GoTo Error_Click
           End If
              '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
        Else
           '「キャンセル」釦押下時
           GoTo Error_Click
        End If
      End If
Exit Sub

'キャンセル釦押下、又は送信異常時
Error_Click:

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

If bFlag = True Then                  'V1.4.0.1 ADD
     If chkICM(Index).Value = 0 Then
        '「キャンセル」釦押下時
        SetEnableTrue (1)
        iCancelFlag = 1
        chkICM(Index).Caption = "切離"
        chkICM(Index).BackColor = CONECT_OFF
        chkICM(Index).Value = 1
        Exit Sub
    End If
   If chkICM(Index).Value = 1 Then
      SetEnableTrue (1)
      iCancelFlag = 1
      chkICM(Index).Caption = "接続"
      chkICM(Index).BackColor = CONECT_ON
      chkICM(Index).Value = 0
      Exit Sub
   End If
'V1.4.0.1 ADD START
Else
  SetEnableTrue (1)
  iShokiFlag = 1
  psICMConectSts
  iShokiFlag = 0
End If
'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdInOutICM_Click
'//  機能名称  : 「全号機接続」「全号機切離」釦押下時処理
'//  機能概要  : 釦名称処理を行う。
'//              判定IC-M部：
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-237】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdInOutICM_Click(Index As Integer)
  Dim iCnt As Integer       'カウンター
  Dim iResponse As Integer  'メッセージボックスの戻り値
  Dim bRet As Boolean               'メール送信戻り値
  Dim lngErrCode As Long            'エラーコード
  Dim bFlag As Boolean              'エラーフラグ処理　'V1.4.0.1 ADD
  Dim bInOutStatus As Boolean       ' 押下した釦（TRUE:接続,FALSE:切断）        ' EG20 V2.1.0.1[Mainte_03_01]追加
  
  On Error Resume Next

  '画面をロックする。
  SetEnableFalse (1)

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    Select Case Index
    Case 0, 2, 4, 6, 8, 10
        bInOutStatus = True     ' 全号機接続
    Case Else
        bInOutStatus = False    ' 全号機切断
    End Select
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

'  If Index = 0 Then            ' EG20 V1.1.1.1 削除
  If bInOutStatus = True Then   ' EG20 V1.1.1.1 追加
     '「通信接続確認」ポップアップ画面表示
     iResponse = MsgBox("指定した外部機器との通信接続を開始します。よろしいですか？", _
                        vbOKCancel + vbQuestion, _
                        "通信接続確認")
     If iResponse = vbOK Then
        iALLGoukiFlag = 1
       '「全号機接続」釦押下時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'       For iCnt = 0 To 15
'           chkICM(iCnt).Caption = "接続"
'           chkICM(iCnt).BackColor = CONECT_ON
'           chkICM(iCnt).Value = 0
'       Next
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
       '通信設定要求CMD(判定IC-M,接続)を監マプロセスに送信する
       'ヘッダ部共通作成処理
       SendMailHeader
       udtMail.dwRequestKIKI = ML_DT_ICM
       udtMail.dwRequestConectType = ML_REQUEST_CONECT
       For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'          If gIcmButtonInfo(iCnt).bStatus = True Then                 ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
           ' 設定されている号機、かつ選択されたコーナに対して設定
          If gIcmButtonInfo(iCnt).bStatus = True And _
               gIcmButtonInfo(iCnt).nCornerNo = tabIcmCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
              chkICM(gIcmButtonInfo(iCnt).nControlNo).Caption = "接続"
              chkICM(gIcmButtonInfo(iCnt).nControlNo).BackColor = CONECT_ON
              chkICM(gIcmButtonInfo(iCnt).nControlNo).Value = 0
' EG20 V3.3.0.1【結合TR-237】追加開始
              udtMail.dwGouki(iCnt) = ML_TARGET_ON
          Else
              udtMail.dwGouki(iCnt) = ML_TARGET_OFF
' EG20 V3.3.0.1【結合TR-237】追加終了
          End If
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
' EG20 V3.3.0.1【結合TR-237】削除開始
'          udtMail.dwGouki(iCnt) = ML_TARGET_ON
' EG20 V3.3.0.1【結合TR-237】削除終了
          'V1.4.0.1 ADD START
          'ICM設定保持
          iICM_Sts(iCnt) = CONECT_SETU
          sBottom_Sts = ZEN_SITEI
          'V1.4.0.1 ADD END
       Next
       
       'V1.4.0.1 ADD START
        If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_PC) = 0 Then
          '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
          iSend_Mail = MAIL_OK
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'          For iCnt = 0 To 15
'            If chkICM(iCnt).Visible = True Then
'               bRet = pfSetSettei(ICM, iCnt, CONECT_SETU, IdGate.HANTEI_ICM_CONECT_SETTEI)
'            End If
'            Index = Index + 1
'          Next
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
          For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX
'              If gIcmButtonInfo(iCnt).bStatus = True Then             ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
              ' 設定されている号機、かつ選択されたコーナに対して設定
              If gIcmButtonInfo(iCnt).bStatus = True And _
                   gIcmButtonInfo(iCnt).nCornerNo = tabIcmCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
                  bRet = pfSetSettei(ICM, gIcmButtonInfo(iCnt).nKanshiNo - 1, _
                                     CONECT_SETU, IdGate.HANTEI_ICM_CONECT_SETTEI)
              End If
          Next
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
          bFlag = False
          GoTo Error_Click
       End If
       'V1.4.0.1 ADD END

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
       'プログレスバーを表示する
       Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           
       bRet = DssSendMail(MAIL_SLOT_IDSEI, MlSize.CONECT_CMD, udtMail.mlHeader)
       If False = bRet Then
         '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
         lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
         Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
         GoTo Error_Click:
      End If
    Else
      GoTo Error_Click:
    End If
Else
   '「通信切断確認」ポップアップ画面表示
   'V1.8.0.1 DEL START
'   iResponse = MsgBox("指定した外部機器との通信切断を開始します。よろしいですか？", _
'                      vbOKCancel + vbQuestion, _
'                      "通信切断確認")
   'V1.8.0.1 DEL END
   'V1.8.0.1 ADD START
    iResponse = MsgBox("指定した外部機器との通信を切断します。よろしいですか？", _
                        vbOKCancel + vbQuestion, _
                        "通信切断確認")
   'V1.8.0.1 ADD END
   If iResponse = vbOK Then
      iALLGoukiFlag = 1
    '「全号機切離」釦押下時
'     For iCnt = 0 To 15
'       chkICM(iCnt).Caption = "切離"
'       chkICM(iCnt).BackColor = CONECT_OFF
'       chkICM(iCnt).Value = 1
'     Next
     '通信設定要求CMD(判定IC-M,切断)を監マプロセスに送信する
     'ヘッダ部共通作成処理
     SendMailHeader
     udtMail.dwRequestKIKI = ML_DT_ICM
     udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
     For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'          If gIcmButtonInfo(iCnt).bStatus = True Then                 ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
           ' 設定されている号機、かつ選択されたコーナに対して設定
          If gIcmButtonInfo(iCnt).bStatus = True And _
               gIcmButtonInfo(iCnt).nCornerNo = tabIcmCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
              chkICM(gIcmButtonInfo(iCnt).nControlNo).Caption = "切離"
              chkICM(gIcmButtonInfo(iCnt).nControlNo).BackColor = CONECT_OFF
              chkICM(gIcmButtonInfo(iCnt).nControlNo).Value = 1
' EG20 V3.3.0.1【結合TR-237】追加開始
              udtMail.dwGouki(iCnt) = ML_TARGET_ON
          Else
              udtMail.dwGouki(iCnt) = ML_TARGET_OFF
' EG20 V3.3.0.1【結合TR-237】追加終了
          End If
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
' EG20 V3.3.0.1【結合TR-237】削除開始
'         udtMail.dwGouki(iCnt) = ML_TARGET_ON
' EG20 V3.3.0.1【結合TR-237】削除終了
         'V1.4.0.1 ADD START
         'ICM設定保持
         iICM_Sts(iCnt) = CONECT_DAN
         sBottom_Sts = ZEN_SITEI
         'V1.4.0.1 ADD END
     Next
     
     'V1.4.0.1 ADD START
     If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_PC) = 0 Then
       '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
       iSend_Mail = MAIL_OK
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'       For iCnt = 0 To 15
'         If chkICM(iCnt).Visible = True Then
'            bRet = pfSetSettei(ICM, iCnt, CONECT_DAN, IdGate.HANTEI_ICM_CONECT_SETTEI)
'         End If
'         Index = Index + 1
'       Next
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
       For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX
'           If gIcmButtonInfo(iCnt).bStatus = True Then             ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
           ' 設定されている号機、かつ選択されたコーナに対して設定
           If gIcmButtonInfo(iCnt).bStatus = True And _
                gIcmButtonInfo(iCnt).nCornerNo = tabIcmCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
               bRet = pfSetSettei(ICM, gIcmButtonInfo(iCnt).nKanshiNo - 1, _
                                  CONECT_DAN, IdGate.HANTEI_ICM_CONECT_SETTEI)
           End If
       Next
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
       bFlag = False
       GoTo Error_Click
     End If
     'V1.4.0.1 ADD END
     
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
     'プログレスバーを表示する
     Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
     
     bRet = DssSendMail(MAIL_SLOT_IDSEI, MlSize.CONECT_CMD, udtMail.mlHeader)
     If False = bRet Then
        '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
        GoTo Error_Click:
     End If
   Else
      GoTo Error_Click:
   End If
End If
  
  '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
  Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
Exit Sub

'キャンセル釦押下、又は送信異常時
Error_Click:

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
 
 If bFlag = True Then                  'V1.4.0.1 ADD
   If Index = 0 Then
     '「キャンセル」釦押下時
     SetEnableTrue (1)
     psICMConectSts
     Exit Sub
   Else
     SetEnableTrue (1)
     psICMConectSts
     Exit Sub
   End If
'V1.4.0.1 ADD START
Else
    SetEnableTrue (1)
    psICMConectSts
    iALLGoukiFlag = 0
End If
'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : chkKIKI_Click
'//  機能名称  : 「接続」「切離」釦押下時処理
'//  機能概要  : 釦名称処理を行う。
'//              上位機器部：
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub chkKIKI_Click(Index As Integer)
  Dim iCnt As Integer              'カウンター
  Dim bRet As Boolean              'メール送信戻り値
  Dim iResponse As Integer
  Dim lngErrCode As Long           'エラーコード
  Dim sProcId As String            'プロセスＩＤ
  'V1.4.0.1 ADD START
  Dim iSetSts As Integer           '設定ファイル設定値
  Dim bFlag As Boolean
  'V1.4.0.1 ADD END
  Dim nInfoIndex As Integer        ' 保存情報インデックス    ' EG20 V2.1.0.1[Mainte_03_01] 追加
  Dim lKanshiWork As Long          ' ワークエリア            ' EG20 V2.1.0.1[Mainte_03_01] 追加
  
  On Error Resume Next
    
  If iUpDataFlag <> 0 Or iALLGoukiFlag <> 0 Or _
      iShokiFlag = 1 Or iCancelFlag = 1 Or iMailRcvFlag = 1 Then
      iCancelFlag = 0
      Exit Sub
   End If
  
  '画面をロックする。
   SetEnableFalse (2)

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ' 押下釦に対応した上位機器構成を検索
    nInfoIndex = 0
    For iCnt = 0 To CONECT_KIKI_CONTROLMAX
        If gTransButtonInfo(iCnt).nControlNo = Index Then
            nInfoIndex = iCnt
            Exit For
        End If
    Next
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

   If chkKIKI(Index).Value = 0 Then
      chkKIKI(Index).Caption = "接続"
      chkKIKI(Index).BackColor = CONECT_ON
      '「通信接続確認」ポップアップ画面表示
      iResponse = MsgBox("指定した外部機器との通信接続を開始します。よろしいですか？", _
                          vbOKCancel + vbQuestion, _
                          "通信接続確認")
      '「通信接続・切断画面：切離→接続 設定」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMEND_TOSTA_BUTTOM, 0)
      udtMail.dwRequestConectType = ML_REQUEST_CONECT  'V1.3.0.1 ADD
      iSetSts = CONECT_SETU 'V1.4.0.1 ADD
   ElseIf chkKIKI(Index).Value = 1 Then
       chkKIKI(Index).Caption = "切離"
       chkKIKI(Index).BackColor = CONECT_OFF
       '「通信切断確認」ポップアップ画面表示
       'V1.8.0.1 DEL START
'      iResponse = MsgBox("指定した外部機器との通信切断を開始します。よろしいですか？", _
'                         vbOKCancel + vbQuestion, _
'                         "通信切断確認")
       'V1.8.0.1 DEL END
       'V1.8.0.1 ADD START
       iResponse = MsgBox("指定した外部機器との通信を切断します。よろしいですか？", _
                          vbOKCancel + vbQuestion, _
                          "通信切断確認")
       'V1.8.0.1 ADD END
       '「通信接続・切断画面：接続→切離 設定」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMSTA_TOEND_BUTTOM, 0)
       udtMail.dwRequestConectType = ML_REQUEST_SETUDAN  'V1.3.0.1 ADD
       iSetSts = CONECT_DAN 'V1.4.0.1 ADD
   End If
       
   If iResponse = vbOK Then
      '通信設定要求CMD(対象ID)を監マプロセスに送信する
      'ヘッダ部共通作成処理
      SendMailHeader
'      udtMail.dwRequestKIKI = iSendId(Index)                       ' EG20 V2.1.0.1[Mainte_03_01] 削除
      udtMail.dwRequestKIKI = gTransButtonInfo(nInfoIndex).iSendID  ' EG20 V2.1.0.1[Mainte_03_01] 追加
'      udtMail.dwRequestConectType = ML_REQUEST_CONECT  'V1.3.0.1 DEL
      'V1.3.0.1 ADD START
      If udtMail.dwRequestKIKI = ML_DT_ICSVR Then
          sProcId = MAIL_SLOT_IDSEI
         'V1.4.0.1 ADD START
         'IDサーバーM設定保持
         If iSetSts = CONECT_DAN Then
            iIDSVR_Sts = CONECT_DAN
            iSetSts = CONECT_DAN
         Else
            iIDSVR_Sts = CONECT_SETU
            iSetSts = CONECT_SETU
         End If
         'V1.4.0.1 ADD END
      Else
          sProcId = MAIL_SLOT_KANMA
      End If
      'V1.3.0.1 ADD END
      For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
          udtMail.dwGouki(iCnt) = ML_TARGET_OFF
      Next
      udtMail.dwGouki(gTransButtonInfo(nInfoIndex).nCorner - 1) = ML_TARGET_ON  ' EG20 V3.0.0.2[Mainte_03_01] 追加
      
     'V1.4.0.1 ADD START
      If sProcId = MAIL_SLOT_KANMA Then
         'IDサーバー以外の上位機器時は管理のみチェック
         If CheckAppStart(PROC_KANRI) = 0 Then
            '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
'            bRet = pfSetSettei(KIKI, 0, iSetSts, iKansiAreaId(Index))      ' EG20 V2.1.0.1[Mainte_03_01] 削除
            lKanshiWork = gTransButtonInfo(nInfoIndex).iKansiId             ' EG20 V2.1.0.1[Mainte_03_01] 追加
' EG20 V3.0.0.2[Mainte_03_01] 追加開始
            If gTransButtonInfo(nInfoIndex).nRonriType = 0 Then
                ' 接続／切断の論理が逆の場合
                If iSetSts = CONECT_SETU Then
                    iSetSts = gTransButtonInfo(nInfoIndex).nRonriType
                Else
                    iSetSts = gTransButtonInfo(nInfoIndex).nRonriType + 1
                End If
            End If
' EG20 V3.0.0.2[Mainte_03_01] 追加終了
            bRet = pfSetSettei(KIKI, 0, iSetSts, lKanshiWork)               ' EG20 V2.1.0.1[Mainte_03_01] 追加
            bFlag = False
            GoTo Error_Click
         End If
      ElseIf sProcId = MAIL_SLOT_IDSEI Then
         'IDサーバー上位機器時は管理、IDUPCをチェック
          If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_PC) = 0 Then
             iSend_Mail = MAIL_OK
             '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
'             bRet = pfSetSettei(KIKI, 0, iSetSts, iKansiAreaId(Index))      ' EG20 V2.1.0.1[Mainte_03_01] 削除
             lKanshiWork = gTransButtonInfo(nInfoIndex).iKansiId             ' EG20 V2.1.0.1[Mainte_03_01] 追加
             bRet = pfSetSettei(KIKI, 0, iSetSts, lKanshiWork)               ' EG20 V2.1.0.1[Mainte_03_01] 追加
             bFlag = False
             GoTo Error_Click
          End If
      End If
      'V1.4.0.1 ADD END
     
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
      'プログレスバーを表示する
      Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
     
'     bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader) 'V1.3.0.1 DEL
      bRet = DssSendMail(sProcId, MlSize.CONECT_CMD, udtMail.mlHeader) 'V1.3.0.1 ADD
      If False = bRet Then
         
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
          'プログレスバーを消去する
          Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
         
         '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
         lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
          Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
          'メール送信異常：ロック解除
          SetEnableTrue (2)
           iCancelFlag = 1
          If chkKIKI(Index).Value = 0 Then
             chkKIKI(Index).Value = 1
             chkKIKI(Index).Caption = "切離"
             chkKIKI(Index).BackColor = CONECT_OFF
          ElseIf chkKIKI(Index).Value = 1 Then
             chkKIKI(Index).Value = 0
             chkKIKI(Index).Caption = "接続"
             chkKIKI(Index).BackColor = CONECT_ON
          End If
          Exit Sub
      Else
         '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
      End If
   Else
     'ロック解除
      SetEnableTrue (2)
      iCancelFlag = 1
      If chkKIKI(Index).Value = 0 Then
          chkKIKI(Index).Value = 1
          chkKIKI(Index).Caption = "切離"
          chkKIKI(Index).BackColor = CONECT_OFF
      ElseIf chkKIKI(Index).Value = 1 Then
          chkKIKI(Index).Value = 0
          chkKIKI(Index).Caption = "接続"
          chkKIKI(Index).BackColor = CONECT_ON
      End If
      Exit Sub
   End If

'V1.4.0.1 ADD START
Exit Sub
'キャンセル釦押下、又は送信異常時
Error_Click:

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

If bFlag = True Then
   iCancelFlag = 1
  Exit Sub
Else
    SetEnableTrue (2)
    iShokiFlag = 1
    pfGetKiKiConectSts
    iShokiFlag = 0
    Exit Sub
End If
'V1.4.0.1 ADD END

 End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ時処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.8.0.1) 2012-08-28 REVISED BY  [TCC] H.Sugimoto
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    Dim lLen As Long                    'メールサイズ
    Dim uMail As MAIL_CONECT_RES        'メール
    Dim bRet As Boolean                 '戻り値　'V1.4.0.1　ADD
    
    On Error Resume Next

    'メール受信
    lLen = fDssMailReadConect(plMSlot_MN, uMail)
    If lLen > 0 Then                            '受信正常の時

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        If uMail.mlHeader.dwId = ML_ID_CONECT_RES Then
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        End If
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
      
      'V1.4.0.1 ADD START
      If uMail.dwConectSts = 0 Then
       iSend_Mail = MAIL_OK
      'V1.4.0.1 ADD END
       If uMail.dwRequestKIKI = 0 Then
          iMailRcvFlag = 1
          '機器種別=自改の場合
          '自改タブ表示処理
           psJikaiConectSts
          SetEnableTrue (0)
          iMailRcvFlag = 0
       ElseIf uMail.dwRequestKIKI = 1 Then
          iMailRcvFlag = 1
          '機器種別=判定IC-Mの場合
          '判定IC-Mタブ表示処理
           psICMConectSts
          SetEnableTrue (1)
          iMailRcvFlag = 0
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
       ElseIf uMail.dwRequestKIKI = ML_DT_TAKU Then
          iMailRcvFlag = 1
          ' 機器種別=操作卓の場合
          ' 操作卓タブ表示処理
          psTakuConectSts
          SetEnableTrue (3)
          iMailRcvFlag = 0
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
       Else
          iMailRcvFlag = 1
          '上位機器タブ表示処理
          pfGetKiKiConectSts
          SetEnableTrue (2)
          iMailRcvFlag = 0
       End If
     'V1.4.0.1 ADD START
     Else
       iMailRcvFlag = 1
       iSend_Mail = MAIL_ERROR
       pfSetUpData (uMail.dwRequestKIKI)
       iMailRcvFlag = 0
     End If
     'V1.4.0.1 ADD END
       
       'V1.3.0.1 ADD START
       If uMail.mlHeader.dwId = ML_ID_HOSHU_ACTIVE_REQ Then
          '保守画面アクティブ表示の場合
          '「保守画面アクティブ表示要求受信正常」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
          AppActivate frmConectSts.Caption, False
          pfFormActive (frmConectSts.hwnd)
       
' EG20 V6.8.0.1 ADD START
       ElseIf uMail.mlHeader.dwId = ML_ID_PROEND_ORD Then
           'プロセス終了指示の場合
           '「プロセス終了指示受信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            '強制終了処理を行う
            pfAbortProc
            Exit Sub       '処理を終了する
' EG20 V6.8.0.1 ADD END
       End If
       'V1.3.0.1 ADD END
       
       iALLGoukiFlag = 0
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psJikaiConectSts
'//  機能名称  : 自改タブ表示処理
'//  機能概要  : 自改タブの画面表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psJikaiConectSts()
    Dim sKeyName As String
    Dim sGateData As String * CONECT_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim iJikaiSts As Integer    '通信状態
    Dim iFlag As Integer        '釦押下可/押下不可

    Dim nControlNo As Integer   ' 処理コントロール釦番号    EG20 V2.1.0.1[Mainte_03_01]追加

    On Error Resume Next
  
    iJikaiSts = CONECTSTS_ERROR
    
   '自動改札機情報取得
'    For i = CNT_MIN To MAX_OVER_NO                 ' EG20 V2.1.0.1[Mainte_03_01] 削除
    For i = CNT_MIN To MAX_GATE_NO                  ' EG20 V2.1.0.1[Mainte_03_01] 追加

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        ' 釦情報の初期化
        gJikaiButtonInfo(i).bStatus = False          ' 設定有無（TRUE:有り,FALSE:無し）
        gJikaiButtonInfo(i).nDisGoukiNo = 0          ' 表示号機
        gJikaiButtonInfo(i).nDispCornerNo = 0        ' コーナー番号(表示)
        gJikaiButtonInfo(i).nCornerNo = 0            ' コーナー番号(論理)
        gJikaiButtonInfo(i).nCornerGoukiNo = 0       ' コーナー別論理号機
        gJikaiButtonInfo(i).nControlNo = 0           ' コントロール番号
        gJikaiButtonInfo(i).nKanshiNo = 0            ' 監視状態番号
        nControlNo = 0
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

        sKeyName = "gate" & Format(i, "00")
        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                       sKeyName, _
                                       DEFAILT, sGateData, Len(sGateData), _
                                       PATH_GATE_FILE)
      If iRet <> 0 Then
        If Len(sGateData) <> 0 Then
            'データの取得
            ReDim sFData(15)
            iFCnt = 1
            
            For iFLoop = 1 To Len(sGateData)
                If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
                    iFLoop2 = iFLoop
                    Do
                        iFLoop2 = iFLoop2 + 1
                        If iFLoop2 > Len(sGateData) Then
                            sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                            iFCnt = iFCnt + 1
                            If iFCnt >= 16 Then
                                Exit For
                            End If
                            iFLoop = iFLoop2
                            Exit Do
                        End If
                        
                        If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
                            sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                            iFCnt = iFCnt + 1
                            If iFCnt >= 16 Then
                                Exit For
                            End If
                            iFLoop = iFLoop2
                            Exit Do
                        End If
                    Loop
                End If
            Next
               
            '機種タイプによって表示を行う。
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'            'E：EG-R自動改札機。N：NEG自動改札機。＊：表示しない。
'            If Trim(sFData(4)) = EGR Then
'              '通信状態表示
'              lblGouki(i - 1).Visible = True
'              lblGouki(i - 1).Caption = Format(i, "0")
'              lblGouki(i - 1).ForeColor = CONECT_EGR
'              lblGoukiConectSts(i - 1).Visible = True
'              '指定号機表示
'              lblTargetGouki(i - 1).Caption = Format(i, "0")
'              lblTargetGouki(i - 1).Visible = True
'              lblTargetGouki(i - 1).ForeColor = CONECT_EGR
'              '指定号機釦表示
'              chkJikai(i - 1).Visible = True
'              iJikaiType(i - 1) = GATE_JISEDAI
'            End If
'            If Trim(sFData(4)) = NEG Then
'              '通信状態表示
'              lblGouki(i - 1).Visible = True
'              lblGouki(i - 1).Caption = Format(i, "0")
'              lblGoukiConectSts(i - 1).Visible = True
'              '指定号機表示
'              lblTargetGouki(i - 1).Caption = Format(i, "0")
'              lblTargetGouki(i - 1).Visible = True
'              '指定号機釦表示
'              chkJikai(i - 1).Visible = True
'              iJikaiType(i - 1) = GATE_NGATE
'            End If
' EG20 V2.1.0.1[Mainte_03_01] 削除終了

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
            If Trim(sFData(4)) = EG20 Then

              gJikaiButtonInfo(i - 1).bStatus = True
              gJikaiButtonInfo(i - 1).nKanshiNo = i
              gJikaiButtonInfo(i - 1).nDisGoukiNo = CInt(Trim(sFData(1)))
              gJikaiButtonInfo(i - 1).nDispCornerNo = CInt(Trim(sFData(3)))
              gJikaiButtonInfo(i - 1).nCornerNo = CInt(Trim(sFData(12)))
              gJikaiButtonInfo(i - 1).nCornerGoukiNo = CInt(Trim(sFData(13)))

              nControlNo = GetButtonNo(gJikaiButtonInfo(i - 1).nCornerNo, _
                                        gJikaiButtonInfo(i - 1).nCornerGoukiNo)

              gJikaiButtonInfo(i - 1).nControlNo = nControlNo

              '通信状態表示
              lblGouki(nControlNo).Visible = True
              lblGouki(nControlNo).Caption = Format(gJikaiButtonInfo(i - 1).nDisGoukiNo, "0")
              lblGoukiConectSts(nControlNo).Visible = True
              '指定号機表示
              lblTargetGouki(nControlNo).Caption = Format(gJikaiButtonInfo(i - 1).nDisGoukiNo, "0")
              lblTargetGouki(nControlNo).Visible = True
              '指定号機釦表示
              chkJikai(nControlNo).Visible = True
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
            
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'            If Trim(sFData(4)) = MISETI Then
'               '処理を行わない。
'               iJikaiType(i - 1) = GATE_NASI
'            End If
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
            
            '監視盤起動有無チェック
            If CheckAppStart(PROC_KANRI) <> 0 Then
               '監視盤起動有り時
               '通信状態取得処理
                pfGetjikaiConectSts iJikaiSts, i
                If iJikaiSts = CONECTSTS_NORMAL Then
                    '正常時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblGoukiConectSts(i - 1).Caption = CONECT_NORMAL
'                    lblGoukiConectSts(i - 1).ForeColor = CONECT_OK
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblGoukiConectSts(nControlNo).Caption = CONECT_NORMAL
                    lblGoukiConectSts(nControlNo).ForeColor = CONECT_OK
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                ElseIf iJikaiSts = CONECTSTS_ERROR Then
                    '異常時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblGoukiConectSts(i - 1).Caption = CONECT_ERROR
'                    lblGoukiConectSts(i - 1).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblGoukiConectSts(nControlNo).Caption = CONECT_ERROR
                    lblGoukiConectSts(nControlNo).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                ElseIf iJikaiSts = CONECTSTS_END Then
                    '切離時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblGoukiConectSts(i - 1).Caption = CONECT_END
'                    lblGoukiConectSts(i - 1).ForeColor = CONECT_CUT
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblGoukiConectSts(nControlNo).Caption = CONECT_END
                    lblGoukiConectSts(nControlNo).ForeColor = CONECT_CUT
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                Else
                  '「通信接続・切断画面：エリア・ファイル参照異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_API, CONECT_AREA_FILE_NOTACCESS_ERROR, 0)
                    '上記以外：状態取得異常
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblGoukiConectSts(i - 1).Caption = GET_CONECT_ERROR
'                    chkJikai(i - 1).Visible = False
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblGoukiConectSts(nControlNo).Caption = GET_CONECT_ERROR
                    chkJikai(nControlNo).Visible = False
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                End If
           Else
              '監視盤起動無し時
              'V1.4.0.1 DEL START
'              chkJikai(i - 1).Enabled = False '押下不可
'              cmdInOutJikai(0).Enabled = False
'              cmdInOutJikai(1).Enabled = False
              'V1.4.0.1 DEL END
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'              lblGoukiConectSts(i - 1).Caption = CONECT_ERROR
'              lblGoukiConectSts(i - 1).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
              lblGoukiConectSts(nControlNo).Caption = CONECT_ERROR
              lblGoukiConectSts(nControlNo).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
              'V1.4.0.1 DEL START
'              chkEGR.Enabled = False
'              chkNEG.Enabled = False
              'V1.4.0.1 DEL END
            End If
            
            '号機別釦情報取得
             pfGetJikaiSts iJikaiSts, i
             If iJikaiSts = CONECTSTS_ERROR Then
                '接続の場合
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                 chkJikai(i - 1).Value = 0
'                 chkJikai(i - 1).Caption = "接続"
'                 chkJikai(i - 1).BackColor = CONECT_ON
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                 chkJikai(nControlNo).Value = 0
                 chkJikai(nControlNo).Caption = "接続"
                 chkJikai(nControlNo).BackColor = CONECT_ON
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
             ElseIf iJikaiSts = CONECTSTS_END Then
                '切離の場合
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                 chkJikai(i - 1).Value = 1
'                 chkJikai(i - 1).Caption = "切離"
'                 chkJikai(i - 1).BackColor = CONECT_OFF
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                 chkJikai(nControlNo).Value = 1
                 chkJikai(nControlNo).Caption = "切離"
                 chkJikai(nControlNo).BackColor = CONECT_OFF
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
             ElseIf iJikaiSts = GET_CONECTSTS_ERROR Then
                 '号機別取得異常時は非表示。
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                 chkJikai(i - 1).Visible = False '押下不可
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                 chkJikai(nControlNo).Visible = False '押下不可
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
             End If
           End If       ' EG20 V2.1.0.1[Mainte_03_01] 追加
       End If
    End If
   Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetjikaiConectSts
'//  機能名称  : 自改タブ表示処理(監視盤起動有りのためエリア参照可能)
'//  機能概要  : 自改タブの通信状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetjikaiConectSts(iJikaiSts As Integer, iGouki As Integer)
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim iDataArea_Tuushin As Integer        '自改通信状態
    Dim strMutexName    As String           'ミューテックス名
    Dim lngMuHandle     As Long             '排他処理用ハンドル

    On Error Resume Next
    
    Set Idinf_JikaiTuushin = New IdInfProc             '自改通信状態エリア
    '参照(自改通信状態)エリア名を設定
    Idinf_JikaiTuushin.ProcMode = DATA_ID.Data_Id_JikaiTuushinJyotai  '自改通信状態エリア
    Idinf_JikaiTuushin.IdOpen
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       iJikaiSts = GET_CONECTSTS_ERROR
       Set Idinf_JikaiTuushin = Nothing              '自改通信状態エリア
       Exit Function
    End If
     
    '参照(自改通信状態)エリアをＬＯＣＫする。
    Idinf_JikaiTuushin.IdLock
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iJikaiSts = GET_CONECTSTS_ERROR
       Idinf_JikaiTuushin.IdFree
       Set Idinf_JikaiTuushin = Nothing              '自改通信状態エリア
       Exit Function
    End If
    
    'エリアの内容を読み込む。
    Idinf_JikaiTuushin.id = IdGateComSts.GATE_COM
    Idinf_JikaiTuushin.GetJikai_Tuusin iGouki - 1
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iJikaiSts = GET_CONECTSTS_ERROR
       Idinf_JikaiTuushin.IdFree
       Set Idinf_JikaiTuushin = Nothing              '自改通信状態エリア
       Exit Function
    End If
        
    iDataArea_Tuushin = CInt(Idinf_JikaiTuushin.DataArea(iGouki - 1))
    Idinf_JikaiTuushin.IdFree
    Set Idinf_JikaiTuushin = Nothing              '自改通信状態エリア
    
  If iDataArea_Tuushin <> IdGateCom.GATE_COM_CONNECT_NORMAL Then
     '自改通信状態が正常以外の場合以下を行う。
     Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
     '自改設定エリアをオープンする。
     Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
     Idinf_JikaiSettei.IdOpen
     If Idinf_JikaiSettei.Errsts <> 0 Then
      'データ参照異常時はブランク表示設定を行う。
      iJikaiSts = GET_CONECTSTS_ERROR
      Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
      Exit Function
     End If
    
      '自改設定エリアをＬＯＣＫする。
      Idinf_JikaiSettei.IdLock
      If Idinf_JikaiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iJikaiSts = GET_CONECTSTS_ERROR
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
        Exit Function
      End If
     
      'エリアの内容を読み込む。
      Idinf_JikaiSettei.id = IdGate.JIKAI_CONECT_SETTEI
      Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
      If Idinf_JikaiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iJikaiSts = GET_CONECTSTS_ERROR
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
        Exit Function
      End If
      
      '設定内容を取得
      iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
      Select Case iAreaSts
'         Case 1                                         ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応削除
         Case 0                                          ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応追加
           '接続
           iJikaiSts = CONECTSTS_ERROR
           Idinf_JikaiSettei.IdFree
           Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
           Exit Function
'         Case 0                                         ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応削除
         Case 1                                          ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応追加
           iJikaiSts = CONECTSTS_END
           Idinf_JikaiSettei.IdFree
           Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
           Exit Function
        End Select
  End If
   
   '状態：正常
   iJikaiSts = CONECTSTS_NORMAL
   Idinf_JikaiSettei.IdFree
   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetJikaiSts
'//  機能名称  : 自改タブ表示処理(監視盤起動有無対応参照)
'//  機能概要  : 自改タブの号機別釦状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-18   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応(監視盤未起動時でも設定変更可)
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetJikaiSts(iJikaiSts As Integer, iGouki As Integer)
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255     As GATE_INFO        '読込み用エリア（255設定用）
    Dim lngSts          As Long             'ヒットエリアID
    Dim lngLoop1        As Long             'カウンター
    Dim lngHandle       As Long             'ハンドル
    Dim FileName        As String           'ファイル有無チェック
    Dim lngRet          As Long             '戻り値
    Dim bRet            As Boolean          '読み込み結果戻り値
    Dim sSetteiFile     As String           'ファイルパス　'V1.4.0.1　ADD
    
    On Error Resume Next
'V1.4.0.1 DEL START
'    '自改設定ファイル有無
'    FileName = Dir(G_SETTEI_FILE)
'    If FileName = "" Then
'       '無ければ参照不可のため参照異常
'       iJikaiSts = GET_CONECTSTS_ERROR
'       Exit Function
'    End If
'V1.4.0.1 DEL END
'V1.4.0.1 ADD START
   '自改設定ファイル有無
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '自改設定ファイルがない場合
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '自改設定ファイルがある場合
       sSetteiFile = G_SETTEI_FILE
    End If
'V1.4.0.1 ADD END

    '監視盤起動有無チェック
    If CheckAppStart(PROC_KANRI) = 0 Then
        
        '自改設定ファイルをオープン
'        lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 DEL
        lngHandle = CreateFile(sSetteiFile, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
           'オープン異常時は参照不可のため参照異常
           iJikaiSts = GET_CONECTSTS_ERROR
           Exit Function
        End If
        
        '自改設定ファイル読み込み
        For lngLoop1 = 0 To iGouki - 1
            bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        Next
        
        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
        lngSts = SerchId(udtAreaR255, IdGate.JIKAI_CONECT_SETTEI)
        If lngSts >= 0 Then
           'IDが有った場合
           iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
        Else
          ' 該当ＩＤ無しの場合参照異常
          iJikaiSts = GET_CONECTSTS_ERROR
          Exit Function
        End If
        
        Select Case iAreaSts
'           Case 1                                       ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応削除
           Case 0                                        ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応追加
             '接続
              iJikaiSts = CONECTSTS_ERROR
              Exit Function
'           Case 0                                       ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応削除
           Case 1                                        ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応追加
              iJikaiSts = CONECTSTS_END
              Exit Function
        End Select
    Else
     
         Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
         '自改設定エリアをオープンする。
          Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
          Idinf_JikaiSettei.IdOpen
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
             iJikaiSts = GET_CONECTSTS_ERROR
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
          End If
             
          '自改設定エリアをＬＯＣＫする。
          Idinf_JikaiSettei.IdLock
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
             iJikaiSts = GET_CONECTSTS_ERROR
             Idinf_JikaiSettei.IdFree
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
           End If
              
           'エリアの内容を読み込む。
            Idinf_JikaiSettei.id = IdGate.JIKAI_CONECT_SETTEI
            Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
            If Idinf_JikaiSettei.Errsts <> 0 Then
               'データ参照異常時はブランク表示設定を行う。
                iJikaiSts = GET_CONECTSTS_ERROR
                Idinf_JikaiSettei.IdFree
                Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                Exit Function
            End If
               
            '設定内容を取得
             iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
             Select Case iAreaSts
'                 Case 1                                       ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応削除
                 Case 0                                        ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応追加
                  '接続
                   iJikaiSts = CONECTSTS_ERROR
                   Idinf_JikaiSettei.IdFree
                   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                   Exit Function
'                 Case 0                                       ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応削除
                 Case 1                                        ' EG20 V2.1.0.1 【Mainte_03_01】設定値変更対応追加
                   iJikaiSts = CONECTSTS_END
                   Idinf_JikaiSettei.IdFree
                   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                   Exit Function
             End Select
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SerchId
'//  機能名称  : ＩＤ検索処理(全タブ専用)
'//  機能概要  : ＩＤ検索を行う。
'//
'//              型        名称        意味
'//  引数      : GATE_INFO udtArea255 [IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : Long　　　         　[OUT]　0以上：正常。-1以下：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SerchId(udtArea255 As GATE_INFO, lngID As Long) As Long

    Dim lngIndex As Long                '検索用インデックス
    Dim lngMin As Long                  '最小インデックス
    Dim lngMax As Long                  '最大インデックス
    Dim lngChkIndex As Long             '該当インデックス
    Dim lngWorkId   As Long             '標準ＩＤ

    On Error Resume Next
    
    '初期化
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '検索開始
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             'ＩＤ取り出し
        If lngID = lngWorkId Then                                  '同じ？
            lngChkIndex = lngIndex                                  'データ取り出し後、検索終了
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngID < lngID) Then         'データが予備か小さい
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    SerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ChgData
'//  機能名称  : データ変換処理処理
'//  機能概要  : データ変換処理処理を行う。
'//
'//              型        名称        意味
'//  引数      : ID_FMT 　DataArea 　[IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : String　　　        [OUT]　vbNullstring以外：正常。vbNullString    ：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function ChgData(DataArea As ID_FMT) As String

    Dim lngloop As Long
    Dim lngWork As Long
    Dim lngErrsts As Long

    On Error GoTo ChgDataErr
    
    lngErrsts = IdInfErr.OK
    
    Select Case DataArea.intType
    Case ID_TYPE.Flag   '状態
        If (DataArea.bytDATA(0) <> 255) Then
            ChgData = str$(DataArea.bytDATA(0))
            
        Else
            ChgData = "-1"                      '値が不定ならー１セット
            
        End If
            
    Case ID_TYPE.Count  '回数
        lngWork = 0                              '初期化
        For lngloop = 3 To 0 Step -1
            lngWork = lngWork * 256 + DataArea.bytDATA(lngloop)
        Next lngloop
                        
        ChgData = str$(lngWork)
    
    Case ID_TYPE.Date_Type, ID_TYPE.time_type '日付、時刻
        ChgData = StrConv(DataArea.bytDATA, vbUnicode)
        
    Case Else
        ChgData = vbNullString
        lngErrsts = IdInfErr.ID_TYPE_MISS
        Exit Function

    End Select
    
    Exit Function
    
ChgDataErr:
        ChgData = vbNullString
        lngErrsts = IdInfErr.PROC_ERR
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetKansiJyotai
'//  機能名称  : 監視状態ファイル取得処理
'//  機能概要  : 監視状態ファイル取得処理を行う。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V3.1.0.1) 2011-11-17  CODED BY  [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetKansiJyotai() As Boolean

    On Error Resume Next
    
    pfGetKansiJyotai = False     '初期値
    
    '監視盤起動有無チェック
    If CheckAppStart(PROC_KANRI) <> 0 Then
    '起動ありの場合
     
        Set Idinf_KansiJyotai = New IdInfProc              '監視装置状態エリア
        '共有エリアオープン
        Idinf_KansiJyotai.ProcMode = DATA_ID.Data_Id_KansiJyotai    '監視装置状態エリア
        Idinf_KansiJyotai.IdOpen
        If Idinf_KansiJyotai.Errsts <> 0 Then
           '「監視状態画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
           Exit Function
        End If
        
        '監視状態エリアをＬＯＣＫする。
        Idinf_KansiJyotai.IdLock
        If Idinf_KansiJyotai.Errsts <> 0 Then
            'データ参照異常時:異常
            Idinf_KansiJyotai.IdFree
            '「監視状態画面：エリア・ファイル参照異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
            Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
            Exit Function
        End If
    
        '監視状態エリアIDを設定
        Idinf_KansiJyotai.id = K_STS_KINKYU_DAN_SW
        Idinf_KansiJyotai.IdGet
        If Idinf_KansiJyotai.Errsts <> 0 Then
            'データ参照異常時はブランク表示設定を行う。
            Idinf_KansiJyotai.IdFree
            '「監視状態画面：エリア・ファイル参照異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
            Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
            Exit Function
        End If
    
        mintICMKinkyuSW = Idinf_KansiJyotai.DataArea(0)   '設定内容
      
        Idinf_KansiJyotai.IdFree
        Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
        
    '起動なしの場合、緊急通信断SWはOFFとして処理する
    Else
        mintICMKinkyuSW = KINKYU_SW_OFF
    End If
    
    pfGetKansiJyotai = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psICMConectSts
'//  機能名称  : 判定IC-Mタブ表示処理
'//  機能概要  : 判定IC-Mタブの画面表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(1.17.0.1) 2010-01-05   REVISED BY [TCC] S.Terao
'//                 不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.1.0.1) 2011-11-16  CODED BY  [TCC] M.Matsumoto
'//                 EG20フェーズ３対応【非常通信断SW対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psICMConectSts()
    Dim sKeyName As String
'    Dim sGateData As String * CONECT_GATE_SIZE    '１行分ファイル内容取得用 'V1.17.0.1 DEL
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim iICMSts As Integer          '通信状態
    Dim sIcmData As String * IDU_LOG_SIZE     '１行分ファイル内容取得用 'V1.17.0.1 ADD

    Dim nControlNo As Integer   ' 処理コントロール釦番号    EG20 V2.1.0.1[Mainte_03_01]追加

    On Error Resume Next

    iICMSts = CONECTSTS_ERROR
    
    'EG20 V3.1.0.1【非常通信断SW対応】ADD START
    If pfGetKansiJyotai = False Then
        mintICMKinkyuSW = KINKYU_SW_OFF
    End If
    
    '緊急通信断SWがONの場合、全号機接続・切断ボタンを非活性にする
    For iFLoop = cmdInOutICM.LBound To cmdInOutICM.UBound
        If mintICMKinkyuSW = KINKYU_SW_ON Then
            cmdInOutICM(iFLoop).Enabled = False
        Else
            cmdInOutICM(iFLoop).Enabled = True
        End If
    Next iFLoop
    'EG20 V3.1.0.1 ADD END
    
   '自動改札機情報取得
'V1.17.0.1 DEL START
'    For i = CNT_MIN To CONECT_HANTEI_ICM_MAX + 1
'        sKeyName = "gate" & Format(i, "00")
'        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
'                                       sKeyName, _
'                                       DEFAILT, sGateData, Len(sGateData), _
'                                       PATH_GATE_FILE)
'      If iRet <> 0 Then
'        If Len(sGateData) <> 0 Then
'            'データの取得
'            ReDim sFData(15)
'            iFCnt = 1
'
'            For iFLoop = 1 To Len(sGateData)
'                If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
'                    iFLoop2 = iFLoop
'                    Do
'                        iFLoop2 = iFLoop2 + 1
'                        If iFLoop2 > Len(sGateData) Then
'                            sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
'                            iFCnt = iFCnt + 1
'                            If iFCnt >= 16 Then
'                                Exit For
'                            End If
'                            iFLoop = iFLoop2
'                            Exit Do
'                        End If
'
'                        If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
'                            sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
'                            iFCnt = iFCnt + 1
'                            If iFCnt >= 16 Then
'                                Exit For
'                            End If
'                            iFLoop = iFLoop2
'                            Exit Do
'                        End If
'                    Loop
'                End If
'            Next
'V1.17.0.1 DEL END
'V1.17.0.1 ADD START
  'ICM情報取得
'    For i = CNT_MIN To CONECT_HANTEI_ICM_MAX + 1   ' EG20 V2.1.0.1[Mainte_03_01] 削除
    For i = CNT_MIN To MAX_GATE_NO                  ' EG20 V2.1.0.1[Mainte_03_01] 追加

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        ' 釦情報の初期化
        gIcmButtonInfo(i).bStatus = False           ' 設定有無（TRUE:有り,FALSE:無し）
        gIcmButtonInfo(i).nDisGoukiNo = 0           ' 表示号機
        gIcmButtonInfo(i).nDispCornerNo = 0         ' コーナー番号(表示)
        gIcmButtonInfo(i).nCornerNo = 0             ' コーナー番号(論理)
        gIcmButtonInfo(i).nCornerGoukiNo = 0        ' コーナー別論理号機
        gIcmButtonInfo(i).nControlNo = 0            ' コントロール番号
        gIcmButtonInfo(i).nKanshiNo = 0             ' 監視状態番号
        nControlNo = 0
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

      sKeyName = "icm" & Format(i, "00")
      iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ICM, _
                                     sKeyName, _
                                     DEFAILT, sIcmData, Len(sIcmData), _
                                     PATH_IDU_APP & IDU_ICM_FILE)
      If iRet <> 0 Then
         'データの取得
         ReDim sFData(14)
         iFCnt = 1
        
         For iFLoop = 1 To Len(sIcmData)
             If Mid(sIcmData, iFLoop, 1) <> " " Or Mid(sIcmData, iFLoop, 1) <> "," Then
                iFLoop2 = iFLoop
                Do
                 iFLoop2 = iFLoop2 + 1
                 If iFLoop2 > Len(sIcmData) Then
                    sFData(iFCnt) = Mid(sIcmData, iFLoop, iFLoop2 - iFLoop)
                    iFCnt = iFCnt + 1
                    If iFCnt >= 15 Then
                       Exit For
                    End If
                    iFLoop = iFLoop2
                    Exit Do
                 End If
                    
                 If Mid(sIcmData, iFLoop2, 1) = " " Or Mid(sIcmData, iFLoop2, 1) = "," Then
                    sFData(iFCnt) = Mid(sIcmData, iFLoop, iFLoop2 - iFLoop)
                    If Len(Trim(sFData(iFCnt))) <> 0 Then
                       iFCnt = iFCnt + 1
                       If iFCnt >= 15 Then
                          Exit For
                       End If
                    End If
                    iFLoop = iFLoop2
                    Exit Do
                 End If
                Loop
             End If
         Next
'V1.17.0.1 ADD END
            '判定IC-Mのアドレスチェックを行う。
            '0:表示しない。0以外：判定ICモジュール。
           ' If Trim(sFData(14)) <> 0 Then      'V1.17.0.1 DEL
            If Trim(sFData(5)) <> MISETI Then  'V1.17.0.1 ADD

' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'              '通信状態表示
'              lblICMGouki(i - 1).Visible = True
'              lblICMGouki(i - 1).Caption = Format(i, "0")
'              lblICMGoukiConectSts(i - 1).Visible = True
'              '指定号機表示
'              lblTargetICMGouki(i - 1).Caption = Format(i, "0")
'              lblTargetICMGouki(i - 1).Visible = True
'              '指定号機釦表示
'              chkICM(i - 1).Visible = True
' EG20 V2.1.0.1[Mainte_03_01] 削除終了

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
              gIcmButtonInfo(i - 1).bStatus = True
              gIcmButtonInfo(i - 1).nKanshiNo = i
              gIcmButtonInfo(i - 1).nDisGoukiNo = CInt(Trim(sFData(1)))
              gIcmButtonInfo(i - 1).nDispCornerNo = CInt(Trim(sFData(2)))
              gIcmButtonInfo(i - 1).nCornerNo = CInt(Trim(sFData(3)))
              gIcmButtonInfo(i - 1).nCornerGoukiNo = CInt(Trim(sFData(4)))

              nControlNo = GetButtonNo(gIcmButtonInfo(i - 1).nCornerNo, _
                                        gIcmButtonInfo(i - 1).nCornerGoukiNo)

              gIcmButtonInfo(i - 1).nControlNo = nControlNo

              '通信状態表示
              lblICMGouki(nControlNo).Visible = True
              lblICMGouki(nControlNo).Caption = Format(gIcmButtonInfo(i - 1).nDisGoukiNo, "0")
              lblICMGoukiConectSts(nControlNo).Visible = True
              '指定号機表示
              lblTargetICMGouki(nControlNo).Caption = Format(gIcmButtonInfo(i - 1).nDisGoukiNo, "0")
              lblTargetICMGouki(nControlNo).Visible = True
              '指定号機釦表示
              chkICM(nControlNo).Visible = True

' EG20 V2.1.0.1[Mainte_03_01] 追加終了

'            End If         ' EG20 V2.1.0.1[Mainte_03_01] 削除
            
           'ID中継ユニット起動有無チェック
            If CheckAppStart(PROCESS_IDU_PC) <> 0 Then
             'ID中継ユニット起動有り時
                '通信状態取得処理
                pfGetICMConectSts iICMSts, i
                If iICMSts = CONECTSTS_NORMAL Then
                    '正常時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblICMGoukiConectSts(i - 1).Caption = CONECT_NORMAL
'                    lblICMGoukiConectSts(i - 1).ForeColor = CONECT_OK
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblICMGoukiConectSts(nControlNo).Caption = CONECT_NORMAL
                    lblICMGoukiConectSts(nControlNo).ForeColor = CONECT_OK
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                ElseIf iICMSts = CONECTSTS_ERROR Then
                    '異常時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblICMGoukiConectSts(i - 1).Caption = CONECT_ERROR
'                    lblICMGoukiConectSts(i - 1).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblICMGoukiConectSts(nControlNo).Caption = CONECT_ERROR
                    lblICMGoukiConectSts(nControlNo).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                ElseIf iICMSts = CONECTSTS_END Then
                    '切離時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblICMGoukiConectSts(i - 1).Caption = CONECT_END
'                    lblICMGoukiConectSts(i - 1).ForeColor = CONECT_CUT
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblICMGoukiConectSts(nControlNo).Caption = CONECT_END
                    lblICMGoukiConectSts(nControlNo).ForeColor = CONECT_CUT
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                Else
                  '「通信接続・切断画面：エリア・ファイル参照異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_API, CONECT_AREA_FILE_NOTACCESS_ERROR, 0)
                    '上記以外：状態取得異常
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                    lblICMGoukiConectSts(i - 1).Caption = GET_CONECT_ERROR
'                    chkICM(i - 1).Visible = False
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    lblICMGoukiConectSts(nControlNo).Caption = GET_CONECT_ERROR
                    chkICM(nControlNo).Visible = False
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                End If
            Else
             'ID中継ユニット起動無し時
             'V1.4.0.1 DEL START
'              chkICM(i - 1).Enabled = False '押下不可
'              cmdInOutICM(0).Enabled = False
'              cmdInOutICM(1).Enabled = False
             'V1.4.0.1 DEL END
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'              lblICMGoukiConectSts(i - 1).Caption = CONECT_ERROR
'              lblICMGoukiConectSts(i - 1).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
              lblICMGoukiConectSts(nControlNo).Caption = CONECT_ERROR
              lblICMGoukiConectSts(nControlNo).ForeColor = CONECT_ERROR_COLOR
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
            End If
                         
            '号機別状態取得
             pfGetICMSts iICMSts, i
             If iICMSts = CONECTSTS_ERROR Then
                '接続の場合
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                 chkICM(i - 1).Value = 0
'                 chkICM(i - 1).Caption = "接続"
'                 chkICM(i - 1).BackColor = CONECT_ON
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                 chkICM(nControlNo).Value = 0
                 chkICM(nControlNo).Caption = "接続"
                 chkICM(nControlNo).BackColor = CONECT_ON
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
             ElseIf iICMSts = CONECTSTS_END Then
                '切離の場合
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                 chkICM(i - 1).Value = 1
'                 chkICM(i - 1).Caption = "切離"
'                 chkICM(i - 1).BackColor = CONECT_OFF
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                 chkICM(nControlNo).Value = 1
                 chkICM(nControlNo).Caption = "切離"
                 chkICM(nControlNo).BackColor = CONECT_OFF
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
             ElseIf iICMSts = GET_CONECTSTS_ERROR Then
                 '状態取得異常時は非表示
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                 chkICM(i - 1).Visible = False
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                 chkICM(nControlNo).Visible = False
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

             End If
            'EG20 V3.1.0.1【非常通信断SW対応】ADD START
            '緊急通信断SWがONの場合、指定号機ボタンを非活性にする
            If mintICMKinkyuSW = KINKYU_SW_ON Then
                chkICM(nControlNo).Enabled = False
            Else
                chkICM(nControlNo).Enabled = True
            End If
            DoEvents
            'EG20 V3.1.0.1 ADD END
'       End If   'V1.17.0.1 DEL
        End If         ' EG20 V2.1.0.1[Mainte_03_01] 追加
    End If
  Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetICMConectSts
'//  機能名称  : 判定IC-Mタブ表示処理
'//  機能概要  : 判定IC-Mタブの通信状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iICMSts　 [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-10   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.1.0.1) 2011-11-16  CODED BY  [TCC] M.Matsumoto
'//                 EG20フェーズ３対応【非常通信断SW対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetICMConectSts(iICMSts As Integer, iGouki As Integer)
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim iJikaiaArea_Jyotai As Integer       '自改状態エリア状態値
    Dim lSts            As Long             '関数戻り値
    Dim lngMuHandle     As Long             '排他処理用ハンドル
    Dim strMutexName As String
    
    On Error Resume Next
    
    strMutexName = "Mu_" & GGateStatus
    lngMuHandle = dllOpenMutex(strMutexName)            '排他処理(OPEN)
    If lngMuHandle = 0 Then
       'エリア参照不可のため、参照異常
       'iICMSts = GET_CONECTSTS_ERROR               'V1.4.0.1 DEL
       iICMSts = CONECTSTS_ERROR                    'V1.4.0.1 ADD
       Exit Function
    End If
  
    dllCloseHandle (lngMuHandle)                 '排他処理(CLOSE)
    
    Set Idinf_JikaiJyotai = New IdInfProc              '自改状態エリア
    '参照(自改状態)エリア名を設定
    Idinf_JikaiJyotai.ProcMode = DATA_ID.Data_Id_JkaiJyotai    '自改状態エリア
    Idinf_JikaiJyotai.IdOpen
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       iICMSts = GET_CONECTSTS_ERROR
        Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
     
    '参照(自改状態)エリアをＬＯＣＫする。
    Idinf_JikaiJyotai.IdLock
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iICMSts = GET_CONECTSTS_ERROR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
    
     'エリアの内容を読み込む。
    Idinf_JikaiJyotai.id = IdGate.HANTEI_ICM_CONECT_STS
    Idinf_JikaiJyotai.GetJikai_Sts iGouki - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iICMSts = GET_CONECTSTS_ERROR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
            
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGouki - 1))
    Idinf_JikaiJyotai.IdFree
    Set Idinf_JikaiJyotai = Nothing               '自改状態エリア

    'EG20 V3.1.0.1【非常通信断SW対応】ADD START
    If mintICMKinkyuSW = KINKYU_SW_ON Then
        iICMSts = CONECTSTS_END     '切離
        Exit Function
    End If
    'EG20 V3.1.0.1 ADD END
    
  If iJikaiaArea_Jyotai <> IdGateCom.GATE_COM_CONNECT_NORMAL Then
     '自改状態が正常以外の場合以下を行う。
      Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
     '自改設定エリアをオープンする。
      Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei    '自改設定エリア
      Idinf_JikaiSettei.IdOpen
      If Idinf_JikaiSettei.Errsts <> 0 Then
         iICMSts = GET_CONECTSTS_ERROR
         Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
         Exit Function
      End If
         
      '自改設定エリアをＬＯＣＫする。
      Idinf_JikaiSettei.IdLock
      If Idinf_JikaiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iICMSts = GET_CONECTSTS_ERROR
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
        Exit Function
      End If
     
      'エリアの内容を読み込む。
      Idinf_JikaiSettei.id = IdGate.HANTEI_ICM_CONECT_SETTEI
      Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
      If Idinf_JikaiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iICMSts = GET_CONECTSTS_ERROR
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
        Exit Function
      End If
       
      '設定内容を取得
      iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
      Select Case iAreaSts
         Case 1
           '接続
           iICMSts = CONECTSTS_ERROR
           Idinf_JikaiSettei.IdFree
           Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
           Exit Function
         Case 0
           iICMSts = CONECTSTS_END
           Idinf_JikaiSettei.IdFree
           Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
           Exit Function
        End Select
   End If
   
   '状態：正常
    iICMSts = CONECTSTS_NORMAL
    
   Idinf_JikaiSettei.IdFree
   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetICMSts
'//  機能名称  : 判定IC-Mタブ表示処理
'//  機能概要  : 判定IC-Mタブの号機別状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iICMSts　 [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-18   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応(監視盤未起動時でも設定変更可)
'//     REVISIONS :(EG20 V3.1.0.1) 2011-11-16  CODED BY  [TCC] M.Matsumoto
'//                 EG20フェーズ３対応【非常通信断SW対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetICMSts(iICMSts As Integer, iGouki As Integer)
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255 As GATE_INFO            '読込み用エリア（255設定用）
    Dim lngSts As Long
    Dim lngLoop1 As Long
    Dim lngHandle As Long
    Dim FileName As String
    Dim lngRet As Long
    Dim bRet As Boolean
    Dim sSetteiFile As String              'ファイルパス　’V1.4.0.1　ADD
    
    On Error Resume Next
'V1.4.0.1 DEL START
'    '自改設定ファイル有無
'    FileName = Dir(G_SETTEI_FILE)
'    If FileName = "" Then
'       '無ければ参照不可のため参照異常
'       iICMSts = GET_CONECTSTS_ERROR
'       Exit Function
'    End If
'V1.4.0.1 DEL END
'V1.4.0.1 ADD START
   '自改設定ファイル有無
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '自改設定ファイルがない場合
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '自改設定ファイルがある場合
       sSetteiFile = G_SETTEI_FILE
    End If
'V1.4.0.1 ADD END
    '監視盤起動有無チェック
    If CheckAppStart(PROC_KANRI) = 0 Then
    
        '自改設定ファイルをオープン
'        lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 DEL
        lngHandle = CreateFile(sSetteiFile, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)  'V1.4.0.1 ADD

        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
           'オープン異常時は参照不可のため参照異常
           iICMSts = GET_CONECTSTS_ERROR
           Exit Function
        End If
        
        '自改設定ファイル読み込み
        For lngLoop1 = 0 To iGouki - 1
            bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        Next
        
        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
        lngSts = SerchId(udtAreaR255, IdGate.HANTEI_ICM_CONECT_SETTEI)
        If lngSts >= 0 Then
           'IDが有った場合
           iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
        Else
          ' 該当ＩＤ無しの場合参照異常
          iICMSts = GET_CONECTSTS_ERROR
          Exit Function
        End If
        
        Select Case iAreaSts
           Case 1
             '接続
              iICMSts = CONECTSTS_ERROR
              Exit Function
           Case 0
              iICMSts = CONECTSTS_END
              Exit Function
        End Select
    Else
    
         Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
         '自改設定エリアをオープンする。
          Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
          Idinf_JikaiSettei.IdOpen
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
             iICMSts = GET_CONECTSTS_ERROR
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
          End If
             
          '自改設定エリアをＬＯＣＫする。
          Idinf_JikaiSettei.IdLock
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
             iICMSts = GET_CONECTSTS_ERROR
             Idinf_JikaiSettei.IdFree
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
           End If
              
           'エリアの内容を読み込む。
            Idinf_JikaiSettei.id = IdGate.HANTEI_ICM_CONECT_SETTEI
            Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
            If Idinf_JikaiSettei.Errsts <> 0 Then
               'データ参照異常時はブランク表示設定を行う。
                iICMSts = GET_CONECTSTS_ERROR
                Idinf_JikaiSettei.IdFree
                Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                Exit Function
            End If
               
            '設定内容を取得
             iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
            'EG20 V3.1.0.1【非常通信断SW対応】ADD START
            If mintICMKinkyuSW = KINKYU_SW_ON Then
                iICMSts = CONECTSTS_END
            Else
            'EG20 V3.1.0.1 ADD END
                Select Case iAreaSts
                    Case 1
                     '接続
                      iICMSts = CONECTSTS_ERROR
                      Idinf_JikaiSettei.IdFree
                      Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                      Exit Function
                    Case 0
                      iICMSts = CONECTSTS_END
                      Idinf_JikaiSettei.IdFree
                      Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                      Exit Function
                End Select
            End If          'EG20 V3.1.0.1【非常通信断SW対応】ADD
            
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetKiKiConectSts
'//  機能名称  : 上位機器タブ表示処理
'//  機能概要  : 上位機器タブの通信状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-26   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視盤未起動時、接続・切断設定可、又値保持
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-01  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub pfGetKiKiConectSts()
  Dim iCnt As Integer              'カウンター
  Dim sKey As String               'キー名
  Dim sGetInf As String * CONECT_KIKI_SIZE      '取得情報(表示名称)
  Dim lSts As Long                 'INI取得処理戻り値
  Dim iAreaID As Integer           '取得情報(エリアID)
  Dim i As Integer                 '前詰めのカウンター
  Dim iKikists As Integer          '外部機器通信状態
  Dim iFlag As Integer             '釦押下可/押下不可フラグ
  Dim iKansiId As Integer          '監視設定エリアID
  Dim iSyoriErrSts As Integer      '処理異常ステータス
  Dim nCorner As Integer           ' コーナ番号                     EG20 V5.2.0.1追加
  
    Dim nControlNo(1 To 2) As Integer   ' 処理コントロール釦番号    EG20 V2.1.0.1[Mainte_03_01]追加

  On Error Resume Next
  
  iKikists = CONECTSTS_ERROR
  i = 0

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    nControlNo(CONECT_KIKI_TABKANSHI) = 0
    nControlNo(CONECT_KIKI_TABTAKU) = 0
    For i = 0 To CONECT_KIKI_CONTROLMAX
        gTransButtonInfo(i).bStatus = False              ' 設定有無（TRUE:有り,FALSE:無し）
        gTransButtonInfo(i).sGetInf = ""                 ' 画面表示用名称
        gTransButtonInfo(i).iAreaID = 0                  ' 対象外部機器上位機器通信状態エリアID
        gTransButtonInfo(i).iSendID = 0                  ' プロセス名（送信機種種別を設定する）
        gTransButtonInfo(i).iKansiId = 0                 ' 監視設定ファイルのエリアID
        gTransButtonInfo(i).nCornerNo = 0                ' タブ番号（統合監視盤、操作卓）
        gTransButtonInfo(i).nCornerGoukiNo = 0           ' タブ別論理番号
        gTransButtonInfo(i).nControlNo = 0               ' コントロール番号
        gTransButtonInfo(i).nIniListNo = 0               ' 外部機器リスト番号
    Next i

  i = 0
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
    
    For iCnt = CONECT_KIKI_MIN To CONECT_KIKI_INI_MAX
      iSyoriErrSts = 0
      ' OUTKIKI_LIST.iniから表示用外部機器名称を取得する。
      sKey = PROFILE_KEY_KIKINAME & Format(iCnt, "00")
      lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                       sKey, _
                                       DEFAILT, _
                                       sGetInf, _
                                       Len(sGetInf), _
                                       OUTKIKI_LIST_FILE)
      If lSts = 0 Then
         '処理なし
         iSyoriErrSts = 1
      End If
      
      If iSyoriErrSts = 0 Then
         ' OUTKIKI_LIST.iniから上位通信エリアIDを取得する。
         sKey = PROFILE_KEY_KIKIAREA_NAME & Format(iCnt, "00")
         iAreaID = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                        sKey, _
                                        DEFAILT_Int, _
                                        OUTKIKI_LIST_FILE)
                                       
         If iAreaID = False Then
            '処理なし
             iSyoriErrSts = 1
         End If
      End If
      
      'V2.3.0.1 ADD START
      'IDU設置無しかつ現在の上位機器通信状態エリアがIDサーバで無い場合
      'または、IDU設置有りの場合は、以降の表示処理を行う。
      If (pbIDUSts = 1 And iAreaID <> IdKikiComSts.ID_SERVER_COM) Or _
         (pbIDUSts = 0) Then
      'V2.3.0.1 ADD END

         If iSyoriErrSts = 0 Then
            ' OUTKIKI_LIST.iniから送信先IDを取得する。
            sKey = PROFILE_KEY_PROCESS & Format(iCnt, "00")
            iSendID(iCnt - 1) = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                                     sKey, _
                                                     DEFAILT_Int, _
                                                     OUTKIKI_LIST_FILE)
            If iSendID(iCnt - 1) = False Then
               '処理なし
                iSyoriErrSts = 1
            End If
         End If
         
         If iSyoriErrSts = 0 Then
            ' OUTKIKI_LIST.iniから監視設定エリアIDを取得する。
            sKey = PROFILE_KEY_KANSI_ID & Format(iCnt, "00")
            iKansiId = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                           sKey, _
                                           DEFAILT_Int, _
                                           OUTKIKI_LIST_FILE)
            If iKansiId = False Then
              '処理なし
              iSyoriErrSts = 1
            Else

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                gTransButtonInfo(i).bStatus = True               ' 設定有無（TRUE:有り,FALSE:無し）
                gTransButtonInfo(i).sGetInf = _
                    Left(sGetInf, InStr(sGetInf, Chr(0)) - 1)    ' 画面表示用名称
                gTransButtonInfo(i).iAreaID = iAreaID            ' 対象外部機器上位機器通信状態エリアID
                gTransButtonInfo(i).iSendID = iSendID(iCnt - 1)  ' プロセス名（送信機種種別を設定する）
                gTransButtonInfo(i).iKansiId = iKansiId          ' 監視設定ファイルのエリアID
                
                Call psAddKikiCornerName(i)
' EG20 V5.2.0.1追加開始
                If gTransButtonInfo(i).nCorner <> 0 Then
                    nCorner = gTransButtonInfo(i).nCorner - 1
                    If gudtSettiCorner(nCorner).intGokiNum > 0 Then
                        iSyoriErrSts = 0
                    Else
                        iSyoriErrSts = 1
                    End If
                End If
' EG20 V5.2.0.1追加終了
                If iSyoriErrSts = 0 Then        ' EG20 V5.2.0.1条件追加
                    nControlNo(gTransButtonInfo(i).nCornerNo) = nControlNo(gTransButtonInfo(i).nCornerNo) + 1
                    gTransButtonInfo(i).nCornerGoukiNo = nControlNo(gTransButtonInfo(i).nCornerNo)  ' タブ別論理番号
                    
                    gTransButtonInfo(i).nControlNo = GetKikiButtonNo(gTransButtonInfo(i).nCornerNo, _
                                                            gTransButtonInfo(i).nCornerGoukiNo)     ' コントロール番号
                    gTransButtonInfo(i).nIniListNo = iCnt            ' 外部機器リスト番号
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
    
                    iKansiAreaId(iCnt - 1) = iKansiId   'V1.4.0.1 ADD
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                    i = i + 1
' EG20 V5.2.0.1追加開始
                Else
                    ' エリアを初期化
                    gTransButtonInfo(i).bStatus = False              ' 設定有無（TRUE:有り,FALSE:無し）
                    gTransButtonInfo(i).sGetInf = ""                 ' 画面表示用名称
                    gTransButtonInfo(i).iAreaID = 0                  ' 対象外部機器上位機器通信状態エリアID
                    gTransButtonInfo(i).iSendID = 0                  ' プロセス名（送信機種種別を設定する）
                    gTransButtonInfo(i).iKansiId = 0                 ' 監視設定ファイルのエリアID
                    gTransButtonInfo(i).nCornerNo = 0                ' タブ番号（統合監視盤、操作卓）
                    gTransButtonInfo(i).nCornerGoukiNo = 0           ' タブ別論理番号
                    gTransButtonInfo(i).nControlNo = 0               ' コントロール番号
                    gTransButtonInfo(i).nIniListNo = 0               ' 外部機器リスト番号
                End If
' EG20 V5.2.0.1追加終了
            End If
         End If
      End If
    Next iCnt

    For i = 0 To CONECT_KIKI_CONTROLMAX

        If gTransButtonInfo(i).bStatus = True Then

            iCnt = gTransButtonInfo(i).nControlNo                   ' コントロール番号
            KikiName(iCnt).Caption = gTransButtonInfo(i).sGetInf    ' 画面表示用名称
            KikiName(iCnt).Visible = True
            chkKIKI(iCnt).Visible = True
            lblOverSts(iCnt).Visible = True

            '監視盤起動有無チェック
            If CheckAppStart(PROC_KANRI) <> 0 Then
                
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'                pfKikiConectSts iKikists, gTransButtonInfo(i).iAreaID, gTransButtonInfo(i).iKansiId
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
                ' 操作卓タブ側の設定情報は自改情報から取得する
                ' 上記以外は上位機器通信状態
                If gTransButtonInfo(i).nCornerNo = CONECT_KIKI_TABTAKU Then
                
                    pfKikiConectStsTaku iKikists, i
                Else
                    pfKikiConectSts iKikists, gTransButtonInfo(i).iAreaID, gTransButtonInfo(i).iKansiId
                End If
                If gTransButtonInfo(i).nRonriType = 0 Then
                    ' 接続論理が逆の場合は結果を反転
                    Select Case iKikists
                    Case CONECTSTS_END
                        iKikists = CONECTSTS_ERROR
                    Case CONECTSTS_ERROR
                        iKikists = CONECTSTS_END
                    Case CONECTSTS_NORMAL
                    Case Else
                    End Select
                End If
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
                If iKikists = CONECTSTS_NORMAL Then
                    '正常時
                    lblOverSts(iCnt).Caption = CONECT_NORMAL
                    lblOverSts(iCnt).ForeColor = CONECT_OK
                ElseIf iKikists = CONECTSTS_ERROR Then
                    '異常時
                    lblOverSts(iCnt).Caption = CONECT_ERROR
                    lblOverSts(iCnt).ForeColor = CONECT_ERROR_COLOR
                ElseIf iKikists = CONECTSTS_END Then
                    '切離時
                    lblOverSts(iCnt).Caption = CONECT_END
                    lblOverSts(iCnt).ForeColor = CONECT_CUT
                Else
                    '上記以外：状態取得異常
                    lblOverSts(iCnt).Caption = GET_CONECT_ERROR
                    chkKIKI(iCnt).Visible = False
                    '「通信接続・切断画面：エリア・ファイル参照異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_API, CONECT_AREA_FILE_NOTACCESS_ERROR, 0)
                End If
            Else
                '監視盤未起動時
                'V1.4.0.1 DEL START
                'chkKIKI(i).Enabled = False
                'V1.4.0.1 DEL END
                lblOverSts(iCnt).Caption = CONECT_ERROR
                lblOverSts(iCnt).ForeColor = CONECT_ERROR_COLOR
            End If
              
            '上位機器別状態取得
            pfGetKiKiSts iKikists, gTransButtonInfo(i).iKansiId
            If gTransButtonInfo(i).nRonriType = 0 Then
                ' 接続論理が逆の場合は結果を反転
                Select Case iKikists
                Case CONECTSTS_END
                    iKikists = CONECTSTS_ERROR
                Case CONECTSTS_ERROR
                    iKikists = CONECTSTS_END
                Case CONECTSTS_NORMAL
                Case Else
                End Select
            End If
            If iKikists = CONECTSTS_ERROR Then
                '接続の場合
                chkKIKI(iCnt).Value = 0
                chkKIKI(iCnt).Caption = "接続"
                chkKIKI(iCnt).BackColor = CONECT_ON
            ElseIf iKikists = CONECTSTS_END Then
                '切離の場合
                chkKIKI(iCnt).Value = 1
                chkKIKI(iCnt).Caption = "切離"
                chkKIKI(iCnt).BackColor = CONECT_OFF
            ElseIf iKikists = GET_CONECTSTS_ERROR Then
                '状態取得異常時は非表示
                chkKIKI(iCnt).Visible = False
            End If
        End If
    Next i
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'               KikiName(i).Caption = sGetInf
'               KikiName(i).Visible = True
'               chkKIKI(i).Visible = True
'               lblOverSts(i).Visible = True
'
'               '監視盤起動有無チェック
'              If CheckAppStart(PROC_KANRI) <> 0 Then
'                  pfKikiConectSts iKikists, iAreaID, iKansiId
'                  If iKikists = CONECTSTS_NORMAL Then
'                     '正常時
'                     lblOverSts(i).Caption = CONECT_NORMAL
'                     lblOverSts(i).ForeColor = CONECT_OK
'                  ElseIf iKikists = CONECTSTS_ERROR Then
'                     '異常時
'                     lblOverSts(i).Caption = CONECT_ERROR
'                     lblOverSts(i).ForeColor = CONECT_ERROR_COLOR
'                  ElseIf iKikists = CONECTSTS_END Then
'                     '切離時
'                     lblOverSts(i).Caption = CONECT_END
'                     lblOverSts(i).ForeColor = CONECT_CUT
'                  Else
'                    '上記以外：状態取得異常
'                     lblOverSts(i).Caption = GET_CONECT_ERROR
'                     chkKIKI(i).Visible = False
'                     '「通信接続・切断画面：エリア・ファイル参照異常」ログ出力
'                     Call sLogTraceReq(LTYP_ERROR, L3AN_API, CONECT_AREA_FILE_NOTACCESS_ERROR, 0)
'                  End If
'              Else
'                  '監視盤未起動時
'                  'V1.4.0.1 DEL START
'                  'chkKIKI(i).Enabled = False
'                  'V1.4.0.1 DEL END
'                  lblOverSts(i).Caption = CONECT_ERROR
'                  lblOverSts(i).ForeColor = CONECT_ERROR_COLOR
'              End If
'
'              '上位機器別状態取得
'              pfGetKiKiSts iKikists, iKansiId
'              If iKikists = CONECTSTS_ERROR Then
'                 '接続の場合
'                  chkKIKI(i).Value = 0
'                  chkKIKI(i).Caption = "接続"
'                  chkKIKI(i).BackColor = CONECT_ON
'              ElseIf iKikists = CONECTSTS_END Then
'                 '切離の場合
'                  chkKIKI(i).Value = 1
'                  chkKIKI(i).Caption = "切離"
'                  chkKIKI(i).BackColor = CONECT_OFF
'              ElseIf iKikists = GET_CONECTSTS_ERROR Then
'                  '状態取得異常時は非表示
'                  chkKIKI(i).Visible = False
'              End If
'               i = i + 1
'            End If
'          End If
'      End If 'V2.3.0.1 ADD END
'   Next iCnt
' EG20 V2.1.0.1[Mainte_03_01] 削除終了

  End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfKikiConectSts
'//  機能名称  : 上位機器タブ表示処理
'//  機能概要  : 上位機器タブの通信状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iKikists　[OUT]表示ステータス
'//              Integer　iAreId  　[IN]上位機器通信状態エリアID
'//              Integer　iKansiId　[IN]監視設定ID
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfKikiConectSts(iKikists As Integer, iAreId As Integer, iKansiId As Integer)
    Dim iAreaSts As Integer     '監視設定状態値
    Dim iConectSts As Integer   '上位通信状態値
        
    On Error Resume Next
    
    'ＩＤ別情報操作クラスの生成
    Set Idinf_Jyoui = New IdInfProc                    '上位通信状態エリア
   '参照(上位機器通信状態)エリア名を設定
    Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui
    Idinf_Jyoui.IdOpen
    If Idinf_Jyoui.Errsts <> 0 Then
      'データ参照異常時はブランク表示設定にして処理終了。
      iKikists = GET_CONECTSTS_ERROR
      Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
      Exit Function
    End If
    
    '参照(上位機器通信状態)エリアをＬＯＣＫする。
    Idinf_Jyoui.IdLock
    If Idinf_Jyoui.Errsts <> 0 Then
      'データ参照異常時はブランク表示設定にして処理終了。
      iKikists = GET_CONECTSTS_ERROR
      Idinf_Jyoui.IdFree
      Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
      Exit Function
    End If
    
     'エリアの内容を読み込む。
    Idinf_Jyoui.id = iAreId
    Idinf_Jyoui.GetInf (CONECT)
    If Idinf_Jyoui.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定にして処理終了。
       iKikists = GET_CONECTSTS_ERROR
       Idinf_Jyoui.IdFree
       Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
       Exit Function
    End If
    
    iConectSts = CInt(Idinf_Jyoui.DataArea(0))
    Idinf_Jyoui.IdFree
    Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
    
  If iConectSts <> IdGateCom.GATE_COM_CONNECT_NORMAL Then
     '上位機器通信状態が正常以外の場合以下を行う。
      Set Idinf_KansiSettei = New IdInfProc              '監視装置設定エリア
     '共有エリアオープン
     Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei    '監視装置設定エリア
     Idinf_KansiSettei.IdOpen
     If Idinf_KansiSettei.Errsts <> 0 Then
        iKikists = GET_CONECTSTS_ERROR
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
        Exit Function
     End If
        
      '監視設定エリアをＬＯＣＫする。
      Idinf_KansiSettei.IdLock
      If Idinf_KansiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iKikists = GET_CONECTSTS_ERROR
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
        Exit Function
      End If
     
      '監視設定エリアIDを設定
      Idinf_KansiSettei.id = iKansiId
      Idinf_KansiSettei.IdGet
      If Idinf_KansiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iKikists = GET_CONECTSTS_ERROR
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
        Exit Function
      End If
 
      iAreaSts = Idinf_KansiSettei.DataArea(0)   '設定内容
      Select Case iAreaSts
         Case 1
          '接続
           iKikists = CONECTSTS_ERROR
           Idinf_KansiSettei.IdFree
           Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
           Exit Function
         Case 0
           iKikists = CONECTSTS_END
           Idinf_KansiSettei.IdFree
           Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
           Exit Function
      End Select
   End If
   
   '状態：正常
   iKikists = CONECTSTS_NORMAL
   Idinf_KansiSettei.IdFree
   Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pfKikiConectStsTaku
'//  機能名称  : 上位機器タブ表示処理（操作卓タブ用）
'//  機能概要  : 上位機器タブの通信状態取得処理を行う。（操作卓タブ用）
'//
'//              型        名称      意味
'//  引数      : Integer　iKikists　[OUT]表示ステータス
'//              Integer　iIndex  　[IN]上位機器設定構成インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：pfKikiConectSts流用
'///////////////////////////////////////////////////////////////////
Private Function pfKikiConectStsTaku(iKikists As Integer, iIndex As Integer)
    Dim iAreaSts As Integer                 ' 監視設定状態値
        
    Dim iJikaiaArea_Jyotai As Integer       ' 自改状態エリア状態値
    Dim lngMuHandle As Long                 ' 排他処理用ハンドル
    Dim strMutexName As String
        
    Dim iAreaID As Integer                  ' 通信状態エリアID
    Dim iKansiId As Integer                 ' 監視設定ファイルのエリアID
    Dim iGokiNo As Integer                  ' 自改状態の号機
        
    On Error Resume Next
    
    strMutexName = "Mu_" & GGateStatus
    lngMuHandle = dllOpenMutex(strMutexName)            '排他処理(OPEN)
    If lngMuHandle = 0 Then
       'エリア参照不可のため、参照異常
       'iICMSts = GET_CONECTSTS_ERROR               'V1.4.0.1 DEL
       iKikists = CONECTSTS_ERROR                    'V1.4.0.1 ADD
       Exit Function
    End If
  
    dllCloseHandle (lngMuHandle)                 '排他処理(CLOSE)
    
    ' 設定情報の取得
    iKansiId = gTransButtonInfo(iIndex).iKansiId
    iAreaID = gTransButtonInfo(iIndex).iAreaID
    iGokiNo = gTransButtonInfo(iIndex).nCornerGoukiNo
    
    Set Idinf_JikaiJyotai = New IdInfProc              '自改状態エリア
    '参照(自改状態)エリア名を設定
    Idinf_JikaiJyotai.ProcMode = DATA_ID.Data_Id_JkaiJyotai    '自改状態エリア
    Idinf_JikaiJyotai.IdOpen
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       iKikists = GET_CONECTSTS_ERROR
        Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
    
    '参照(自改状態)エリアをＬＯＣＫする。
    Idinf_JikaiJyotai.IdLock
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iKikists = GET_CONECTSTS_ERROR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
    
     'エリアの内容を読み込む。
    Idinf_JikaiJyotai.id = iAreaID
    Idinf_JikaiJyotai.GetJikai_Sts iGokiNo - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iKikists = GET_CONECTSTS_ERROR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
    
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGokiNo - 1))
    Idinf_JikaiJyotai.IdFree
    Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
    
    If iJikaiaArea_Jyotai <> IdGateCom.GATE_COM_CONNECT_NORMAL Then
     '上位機器通信状態が正常以外の場合以下を行う。
      Set Idinf_KansiSettei = New IdInfProc              '監視装置設定エリア
     '共有エリアオープン
     Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei    '監視装置設定エリア
     Idinf_KansiSettei.IdOpen
     If Idinf_KansiSettei.Errsts <> 0 Then
        iKikists = GET_CONECTSTS_ERROR
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
        Exit Function
     End If
        
      '監視設定エリアをＬＯＣＫする。
      Idinf_KansiSettei.IdLock
      If Idinf_KansiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iKikists = GET_CONECTSTS_ERROR
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
        Exit Function
      End If
     
      '監視設定エリアIDを設定
      Idinf_KansiSettei.id = iKansiId
      Idinf_KansiSettei.IdGet
      If Idinf_KansiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iKikists = GET_CONECTSTS_ERROR
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
        Exit Function
      End If
 
      iAreaSts = Idinf_KansiSettei.DataArea(0)   '設定内容
      Select Case iAreaSts
         Case 1
          '接続
           iKikists = CONECTSTS_ERROR
           Idinf_KansiSettei.IdFree
           Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
           Exit Function
         Case 0
           iKikists = CONECTSTS_END
           Idinf_KansiSettei.IdFree
           Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
           Exit Function
      End Select
   End If
   
   '状態：正常
   iKikists = CONECTSTS_NORMAL
   Idinf_KansiSettei.IdFree
   Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetKiKiSts
'//  機能名称  : 上位機器タブ表示処理
'//  機能概要  : 上位機器タブの通信状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iKikists　[OUT]表示ステータス
'//              Integer　iKansiId　[IN]監視設定ID
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-18   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応(監視盤未起動時でも設定変更可)
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetKiKiSts(iKikists As Integer, iKansiId As Integer)
    Dim iAreaSts As Integer     '監視設定状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255 As GATE_INFO            '読込み用エリア（255設定用）
    Dim lngSts As Long
    Dim lngLoop1 As Long
    Dim lngHandle As Long
    Dim FileName As String
    Dim lngRet As Long
    Dim bRet As Boolean
    Dim sSetteiFile As String               'ファイルパス　'V1.4.0.1 ADD
        
    On Error Resume Next
'V1.4.0.1 DEL START
'    '監視設定ファイル有無
'    FileName = Dir(K_SETTEI_FILE)
'    If FileName = "" Then
'       '無ければ参照不可のため参照異常
'       iKikists = GET_CONECTSTS_ERROR
'       Exit Function
'    End If
'V1.4.0.1 DEL END
'V1.4.0.1 ADD START
  '監視設定ファイル有無
    FileName = Dir(K_SETTEI_FILE)
    If FileName = "" Then
       '監視設定ファイルがない場合
       sSetteiFile = SHOKI_K_SETTEI_FILE
    Else
       '監視設定ファイルがある場合
       sSetteiFile = K_SETTEI_FILE
    End If
'V1.4.0.1 ADD END
    '監視盤起動有無チェック
    If CheckAppStart(PROC_KANRI) = 0 Then
        
        '監視設定ファイルをオープン
'       lngHandle = CreateFile(K_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)　’V1.4.0.1　DEL
        lngHandle = CreateFile(sSetteiFile, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)  'V1.4.0.1　ADD
        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
           'オープン異常時は参照不可のため参照異常
           iKikists = GET_CONECTSTS_ERROR
           Exit Function
        End If
        
        '監視設定ファイル読み込み
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)

        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
        lngSts = KansiSerchId(udtAreaR255, CLng(iKansiId))
        If lngSts >= 0 Then
           'IDが有った場合
           iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
        Else
          ' 該当ＩＤ無しの場合参照異常
          iKikists = GET_CONECTSTS_ERROR
          Exit Function
        End If
        
        Select Case iAreaSts
           Case 1
             '接続
              iKikists = CONECTSTS_ERROR
              Exit Function
           Case 0
              iKikists = CONECTSTS_END
              Exit Function
        End Select
    Else
        Set Idinf_KansiSettei = New IdInfProc              '監視装置設定エリア
        '共有エリアオープン
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei    '監視装置設定エリア
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
            iKikists = GET_CONECTSTS_ERROR
            Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
            Exit Function
        End If
        
        '監視設定エリアをＬＯＣＫする。
        Idinf_KansiSettei.IdLock
        If Idinf_KansiSettei.Errsts <> 0 Then
          'データ参照異常時はブランク表示設定を行う。
          iKikists = GET_CONECTSTS_ERROR
          Idinf_KansiSettei.IdFree
          Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
          Exit Function
        End If
    
        '監視設定エリアIDを設定
        Idinf_KansiSettei.id = iKansiId
        Idinf_KansiSettei.IdGet
        If Idinf_KansiSettei.Errsts <> 0 Then
          'データ参照異常時はブランク表示設定を行う。
          iKikists = GET_CONECTSTS_ERROR
          Idinf_KansiSettei.IdFree
          Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
          Exit Function
        End If

        iAreaSts = Idinf_KansiSettei.DataArea(0)   '設定内容
        Select Case iAreaSts
           Case 1
            '接続
             iKikists = CONECTSTS_ERROR
             Idinf_KansiSettei.IdFree
             Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
             Exit Function
           Case 0
             iKikists = CONECTSTS_END
             Idinf_KansiSettei.IdFree
             Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
             Exit Function
        End Select
        
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
   End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SerchId
'//  機能名称  : ＩＤ検索処理(全タブ専用)
'//  機能概要  : ＩＤ検索を行う。
'//
'//              型        名称        意味
'//  引数      : GATE_INFO udtArea255 [IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : Long　　　         　[OUT]　0以上：正常。-1以下：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function KansiSerchId(udtArea255 As GATE_INFO, lngID As Long) As Long

    Dim lngIndex As Long                '検索用インデックス
    Dim lngMin As Long                  '最小インデックス
    Dim lngMax As Long                  '最大インデックス
    Dim lngChkIndex As Long             '該当インデックス
    Dim lngWorkId   As Long             '標準ＩＤ

    On Error Resume Next
    
    '初期化
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '検索開始
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             'ＩＤ取り出し
        If lngID = lngWorkId Then                                  '同じ？
            lngChkIndex = lngIndex                                  'データ取り出し後、検索終了
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngID < lngID) Then         'データが予備か小さい
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    KansiSerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetEnableFalse
'//  機能名称  : 画面ロック処理
'//  機能概要  : 画面のロックをする。
'//
'//              型        名称      意味
'//  引数      : Integer  iFalseFlag [IN]対象タブ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse(iFalseFlag As Integer)
    Dim iCnt As Integer 'カウンター

    On Error Resume Next
    'タブをfalseにする。
    tabConect.Enabled = False
    
    If iFalseFlag = 0 Then
       '自改部：全号機接続、全号機切離
        cmdInOutJikai(0).Enabled = False
        cmdInOutJikai(1).Enabled = False
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        cmdInOutJikai(2).Enabled = False
        cmdInOutJikai(3).Enabled = False
        cmdInOutJikai(4).Enabled = False
        cmdInOutJikai(5).Enabled = False
        cmdInOutJikai(6).Enabled = False
        cmdInOutJikai(7).Enabled = False
        cmdInOutJikai(8).Enabled = False
        cmdInOutJikai(9).Enabled = False
        cmdInOutJikai(10).Enabled = False
        cmdInOutJikai(11).Enabled = False
        tabCorner.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
        '自改部：号機別釦
'        For iCnt = CNT_MIN To CONECT_JIKAI_MAX             ' EG20 V2.1.0.1[Mainte_03_01] 削除
        For iCnt = CNT_MIN To CONECT_JIKAI_CONTROLMAX       ' EG20 V2.1.0.1[Mainte_03_01] 追加
            chkJikai(iCnt).Enabled = False
        Next
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'        chkNEG.Enabled = False
'        chkEGR.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
    ElseIf iFalseFlag = 1 Then
       '判定IC-M部：全号機接続、全号機切離
        cmdInOutICM(0).Enabled = False
        cmdInOutICM(1).Enabled = False
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        cmdInOutICM(2).Enabled = False
        cmdInOutICM(3).Enabled = False
        cmdInOutICM(4).Enabled = False
        cmdInOutICM(5).Enabled = False
        cmdInOutICM(6).Enabled = False
        cmdInOutICM(7).Enabled = False
        cmdInOutICM(8).Enabled = False
        cmdInOutICM(9).Enabled = False
        cmdInOutICM(10).Enabled = False
        cmdInOutICM(11).Enabled = False
        tabIcmCorner.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
        '判定IC-M部：号機別釦
'        For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX        ' EG20 V2.1.0.1[Mainte_03_01] 削除
        For iCnt = CNT_MIN To CONECT_HANTEI_ICM_CONTROLMAX  ' EG20 V2.1.0.1[Mainte_03_01] 追加
            chkICM(iCnt).Enabled = False
        Next
    ElseIf iFalseFlag = 2 Then
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        tabKikiCorner.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
        '外部機器部：機器別釦
'        For iCnt = CNT_MIN To CONECT_KIKI_MAX              ' EG20 V2.1.0.1[Mainte_03_01] 削除
        For iCnt = CNT_MIN To CONECT_KIKI_CONTROLMAX        ' EG20 V2.1.0.1[Mainte_03_01] 追加
            chkKIKI(iCnt).Enabled = False
        Next
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ElseIf iFalseFlag = 3 Then
        For iCnt = CNT_MIN To CONECT_TAKU_CONTROLMAX
            chkTaku(iCnt).Enabled = False
        Next
        cmdInOutTaku(0).Enabled = False
        cmdInOutTaku(1).Enabled = False
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
    End If
    
    '表示更新釦
     cmdDataUp.Enabled = False
    
    'メニュー画面へ戻る釦
    cmdCancel.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetEnableTrue
'//  機能名称  : 画面ロック解除処理
'//  機能概要  : 画面のロックを解除する。
'//
'//              型        名称      意味
'//  引数      : Integer  iFalseFlag [IN]対象タブ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.1.0.1) 2011-11-16  CODED BY  [TCC] M.Matsumoto
'//                 EG20フェーズ３対応【非常通信断SW対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue(iFalseFlag As Integer)
    Dim iCnt As Integer 'カウンター

     On Error Resume Next
    'タブをtrueにする。
    tabConect.Enabled = True
    
    If iFalseFlag = 0 Then
        '自改部：全号機接続、全号機切離
        cmdInOutJikai(0).Enabled = True
        cmdInOutJikai(1).Enabled = True
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        cmdInOutJikai(2).Enabled = True
        cmdInOutJikai(3).Enabled = True
        cmdInOutJikai(4).Enabled = True
        cmdInOutJikai(5).Enabled = True
        cmdInOutJikai(6).Enabled = True
        cmdInOutJikai(7).Enabled = True
        cmdInOutJikai(8).Enabled = True
        cmdInOutJikai(9).Enabled = True
        cmdInOutJikai(10).Enabled = True
        cmdInOutJikai(11).Enabled = True
        tabCorner.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
        '自改部：号機別釦
'        For iCnt = CNT_MIN To CONECT_JIKAI_MAX             ' EG20 V2.1.0.1[Mainte_03_01] 削除
        For iCnt = CNT_MIN To CONECT_JIKAI_CONTROLMAX       ' EG20 V2.1.0.1[Mainte_03_01] 追加
            chkJikai(iCnt).Enabled = True
        Next
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'        chkNEG.Enabled = True
'        chkEGR.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
    ElseIf iFalseFlag = 1 Then
        
        If mintICMKinkyuSW = KINKYU_SW_OFF Then         'EG20 V3.1.0.1【非常通信断SW対応】ADD
            '判定IC-M部：全号機接続、全号機切離
            cmdInOutICM(0).Enabled = True
            cmdInOutICM(1).Enabled = True
    ' EG20 V2.1.0.1[Mainte_03_01] 追加開始
            cmdInOutICM(2).Enabled = True
            cmdInOutICM(3).Enabled = True
            cmdInOutICM(4).Enabled = True
            cmdInOutICM(5).Enabled = True
            cmdInOutICM(6).Enabled = True
            cmdInOutICM(7).Enabled = True
            cmdInOutICM(8).Enabled = True
            cmdInOutICM(9).Enabled = True
            cmdInOutICM(10).Enabled = True
            cmdInOutICM(11).Enabled = True
            tabIcmCorner.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
        '判定IC-M部：号機別釦
'        For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX        ' EG20 V2.1.0.1[Mainte_03_01] 削除
            For iCnt = CNT_MIN To CONECT_HANTEI_ICM_CONTROLMAX  ' EG20 V2.1.0.1[Mainte_03_01] 追加
                chkICM(iCnt).Enabled = True
            Next
        'EG20 V3.1.0.1【非常通信断SW対応】ADD START
        Else
            tabIcmCorner.Enabled = True
        End If
        'EG20 V3.1.0.1 ADD END
    ElseIf iFalseFlag = 2 Then
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        tabKikiCorner.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
         '外部機器部：機器別釦
'        For iCnt = CNT_MIN To CONECT_KIKI_MAX              ' EG20 V2.1.0.1[Mainte_03_01] 削除
        For iCnt = CNT_MIN To CONECT_KIKI_CONTROLMAX        ' EG20 V2.1.0.1[Mainte_03_01] 追加
            chkKIKI(iCnt).Enabled = True
        Next
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ElseIf iFalseFlag = 3 Then
        For iCnt = CNT_MIN To CONECT_TAKU_CONTROLMAX
            chkTaku(iCnt).Enabled = True
        Next
        cmdInOutTaku(0).Enabled = True
        cmdInOutTaku(1).Enabled = True
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
    End If
    
    '表示更新釦
     cmdDataUp.Enabled = True
    
    'メニュー画面へ戻る釦
    cmdCancel.Enabled = True
   
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SendMailHeader
'//  機能名称  : 送信メール作成処理
'//  機能概要  : 送信メール(ヘッダ部)作成を行う
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SendMailHeader()
    Dim bytWork()   As Byte
    Dim i As Integer
    
    Erase bytWork
    
      udtMail.mlHeader.dwId = ML_ID_CONECT_CMD
      udtMail.mlHeader.dwSize = MlSize.CONECT_CMD
      udtMail.mlHeader.dwProid = RHOSHU_ID
      udtMail.mlHeader.dwSubArea = 0
      
      bytWork = StrConv(MAIL_SLOT_HOSHU, vbFromUnicode)
      '動的配列の内容をログパラメータ構造体の静的配列に格納する。
      For i = 0 To UBound(bytWork)
        'Null値になったら処理を抜ける。
         If bytWork(i) = vbVEmpty Then Exit For
               
            udtMail.byMailName(i) = bytWork(i)
                
            '動的配列の最大要素になったら処理を抜ける
             If i = UBound(bytWork) Then Exit For
      Next
End Sub

'V1.4.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetSettei
'//  機能名称  : 自改タブ表示処理(監視盤未起動時)
'//  機能概要  : 自改タブの号機別釦状態取得処理を行う。
'//
'//              型        名称     　　 意味
'//  引数      : Integer　iTab      　[IN]表示処理時タブ
'//              Integer　iGouki  　　[IN]処理対象号機番号
'//              Integer　iConectType [IN]設定ステータス
'//              long     iId         [IN]設定ファイルID(外部機器)
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-26   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetSettei(iTab As Integer, iGouki As Integer, iConectType As Integer, iId As Long) As Boolean

    Dim bRet As Boolean
    Dim iIduAplId As Integer
    Dim iIduConect As Integer
    
    On Error Resume Next
    
    pfSetSettei = False
    bRet = True
       
   If JIKAI = iTab Then
      Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_G_SETTEIFILE_UPDATA, 0)
      '自改設定ファイル更新処理(設定値、号機番号、設定ID)
      bRet = pfSetJikaiSts(iConectType, iGouki + 1, iId)
    
    ElseIf ICM = iTab Then
       '判定IC-M設定ファイル更新処理(設定値、号機番号、設定ID)
       iIduAplId = HANTEI_ICM_ID
       If iConectType = CONECT_DAN Then
          iIduConect = IDU_CONECT_DAN
       Else
          iIduConect = IDU_CONECT_SETU
       End If
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_ICM_SETTEIFILE_UPDATA, 0)
       bRet = pfSetICMSts(iIduConect, iGouki + 1, iIduAplId)
       If bRet = True And iSend_Mail = MAIL_OK Then
          '自改設定ファイル更新処理(設定値、号機番号、設定ID)=監視盤未起動、IDU未起動時
          Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_G_SETTEIFILE_UPDATA, 0)
          bRet = pfSetJikaiSts(iConectType, iGouki + 1, iId)
       ElseIf bRet = True And iSend_Mail = MAIL_ERROR Then
          '自改設定エリア更新処理(設定値、号機番号、設定ID)=監視盤起動、IDU未起動時
          bRet = pfSetJikai_ICM(iConectType, iGouki + 1)
       End If
       '自改設定エリア/ファイルの更新処理が異常の時は、判定IC-M設定ファイルをリカバリ
       If bRet = False Then
         '判定IC-M設定ファイルリカバリ処理
         bRet = pfSetICMSts(iConectType, iGouki + 1, iIduAplId)
       End If
       
    ElseIf KIKI = iTab Then
      '設定ファイルID(外部機器)が、IDサーバーかどうかチェックする
      If iId = IdKansiSet.SET_ID_SVR_CONECT_SETTEI Then
        'IDサーバー時：ID中継ユニット設定ファイル更新処理(設定ID、設定ステータス)
        iIduAplId = ID_SVR_ID
        If iConectType = CONECT_DAN Then
           iIduConect = IDU_CONECT_DAN
        Else
           iIduConect = IDU_CONECT_SETU
        End If
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_IDU_SETTEIFILE_UPDATA, 0)
        bRet = pfSetIDSVRSts(iIduConect, iIduAplId)
      End If

      If bRet = True And iSend_Mail = MAIL_OK Then
         '監視設定ファイル更新処理(設定ID、設定ステータス)=監視盤未起動、IDUU未起動時
         Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_K_SETTEIFILE_UPDATA, 0)
         bRet = pfSetKansiSts(CInt(iId), iConectType)
      ElseIf bRet = True And iSend_Mail = MAIL_ERROR Then
         '監視設定エリア更新処理(設定ID、設定ステータス)=監視盤起動、IDU未起動時
         'IDサーバーのみ：エリア又はファイル
         bRet = pfSetKansi_IDSVR(iConectType)
      End If
      
      '監視設定ファイル/エリア更新処理が異常の時は、設定IDがIDサーバーのときID中継ユニット設定ファイルをリカバリ
      If bRet = False And iId = IdKansiSet.SET_ID_SVR_CONECT_SETTEI Then
        'ID中継ユニット設定ファイルリカバリ処理
         bRet = pfSetIDSVRSts(iConectType, iIduAplId)
      End If
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ElseIf TAKU = iTab Then
      Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_G_SETTEIFILE_UPDATA, 0)
      '自改設定ファイル更新処理(設定値、号機番号、設定ID)
      bRet = pfSetJikaiSts(iConectType, iGouki + 1, iId)
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
    End If
    
    pfSetSettei = bRet
    If bRet = False Then
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CONECT_SETTEIFILE_UPDATA_ERROR, 0)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetJikaiSts
'//  機能名称  : 自改設定ファイル更新処理
'//  機能概要  : 自改設定ファイル更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [IN]接続・切断タイプ
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-26   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetJikaiSts(iJikaiSts As Integer, iGouki As Integer, iJikaiID As Long) As Boolean

    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255     As GATE_INFO        '読込み用エリア（255設定用）
    Dim lngSts          As Long             'ヒットエリアID
    Dim lngLoop1        As Long             'カウンター
    Dim lngHandle       As Long             'ハンドル
    Dim FileName        As String           'ファイル有無チェック
    Dim lngRet          As Long             '戻り値
    Dim bRet            As Boolean          '読み込み結果戻り値
    Dim sSetteiFile     As String
    Dim udtAreaR255Work As GATE_INFO        '読込み用エリア（ポインタ移動用）
    
    On Error Resume Next
    
    '自改設定ファイル有無
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '自改設定ファイルがない場合
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '自改設定ファイルがある場合
       sSetteiFile = G_SETTEI_FILE
    End If
        
    '自改設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため更新異常
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        pfSetJikaiSts = False
        Exit Function
    End If
        
    '自改設定ファイル読み込み
    For lngLoop1 = 0 To iGouki - 1
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           'ハンドルのクローズ
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
           Call CloseHandle(lngHandle)
           pfSetJikaiSts = False
           Exit Function
        End If
    Next
    
    'ハンドルのクローズ
    Call CloseHandle(lngHandle)
    
    'ID検索
    lngSts = SerchId(udtAreaR255, iJikaiID)
    If lngSts >= 0 Then
       'IDが有った場合
       SetChgData udtAreaR255.GateInfo(lngSts), iJikaiSts   'データ設定
    Else
       ' 該当ＩＤ無しの場合更新異常
        pfSetJikaiSts = False
       Exit Function
    End If
      
    '自改設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため更新異常
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        pfSetJikaiSts = False
        Exit Function
    End If
     
    'ファイルポインタ移動のための読み込み
     For lngLoop1 = 1 To iGouki - 1
         bRet = ReadFile(lngHandle, udtAreaR255Work, LenB(udtAreaR255Work), lngRet, 0)
         If bRet = False Then
            'ハンドルのクローズ
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
            Call CloseHandle(lngHandle)
            pfSetJikaiSts = False
            Exit Function
         End If
     Next
    
    '自改設定ファイルに書き込む
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       'ハンドルのクローズ
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       Call CloseHandle(lngHandle)
       pfSetJikaiSts = False
       Exit Function
    End If
    
    'ハンドルのクローズ
     Call CloseHandle(lngHandle)

     pfSetJikaiSts = True
     
     Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_SETTEIFILE_UPDATA_OK, 0)
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetChgData
'//  機能名称  : データ変換処理処理
'//  機能概要  : データ変換処理処理を行う。
'//
'//              型        名称        意味
'//  引数      : ID_FMT 　DataArea 　[IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : String　　　        [OUT]　vbNullstring以外：正常。vbNullString    ：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SetChgData(DataArea As ID_FMT, iSts As Integer)
   
   On Error Resume Next

   DataArea.bytDATA(0) = iSts
  
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetKansiSts
'//  機能名称  : 監視設定ファイル更新処理
'//  機能概要  : 監視設定ファイル更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iKansiSts [IN]表示ステータス
'//              Integer　iKansiId　[IN]監視設定ID
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetKansiSts(iKansiId As Integer, iKansiSts As Integer) As Boolean

    Dim iAreaSts As Integer       '監視設定状態値
    Dim lSts            As Long   '関数戻り値
    Dim udtAreaR255 As GATE_INFO  '読込み用エリア（255設定用）
    Dim lngSts As Long
    Dim lngLoop1 As Long
    Dim lngHandle As Long
    Dim FileName As String
    Dim lngRet As Long
    Dim bRet As Boolean
    Dim sSetteiFile As String
        
    On Error Resume Next
       
    '監視設定ファイル有無
    FileName = Dir(K_SETTEI_FILE)
    If FileName = "" Then
       '監視設定ファイルがない場合
       sSetteiFile = SHOKI_K_SETTEI_FILE
    Else
       '監視設定ファイルがある場合
       sSetteiFile = K_SETTEI_FILE
    End If
    
    '監視設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)
    
    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
       pfSetKansiSts = False
       Exit Function
    End If
       
   '監視設定ファイル読み込み
    bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
       Call CloseHandle(lngHandle)
       pfSetKansiSts = False
       Exit Function
    End If
    
   'ハンドルのクローズ
    Call CloseHandle(lngHandle)
       
    'ID検索
     lngSts = KansiSerchId(udtAreaR255, CLng(iKansiId))
     If lngSts >= 0 Then
        'IDが有った場合
        SetChgData udtAreaR255.GateInfo(lngSts), iKansiSts         'データ変換
     Else
        ' 該当ＩＤ無しの場合参照異常
        pfSetKansiSts = False
        Exit Function
     End If
         
    '監視設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)
    
    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
       pfSetKansiSts = False
       Exit Function
    End If
       
    '監視設定ファイル書込み
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       'ハンドルのクローズ
       Call CloseHandle(lngHandle)
       pfSetKansiSts = False
       Exit Function
    End If
    
   'ハンドルのクローズ
     Call CloseHandle(lngHandle)
     
     pfSetKansiSts = True
     
     Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_SETTEIFILE_UPDATA_OK, 0)
 
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetICMSts
'//  機能名称  : 判定IC-M設定ファイル更新処理
'//  機能概要  : 判定IC-M設定ファイルの更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iICMSts　 [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetICMSts(iICMSts As Integer, iGouki As Integer, iId As Integer) As Boolean
    Dim udtAreaR255 As IDU_SETTEI_INFO            '読込み用エリア（255設定用）
    Dim udtAreaR255Work As IDU_SETTEI_INFO        '読込み用エリア（255設定用）
    Dim lngSts As Long
    Dim lngLoop1 As Long
    Dim lngHandle As Long
    Dim lngRet As Long
    Dim bRet As Boolean
    Dim FileName As String
    Dim sSetteiFile As String
           
    On Error Resume Next
        
    '判定IC-M設定ファイル有無
    FileName = Dir(PATH_IDU_APP & PATH_ICM_SETTEI)
    If FileName = "" Then
       '判定IC-M設定ファイルがない場合
       sSetteiFile = PATH_IDU_APP & PATH_SHOKI_ICM_SETTEI
    Else
       '判定IC-M設定ファイルがある場合
       sSetteiFile = PATH_IDU_APP & PATH_ICM_SETTEI
    End If
     
    '判定IC-M設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)
    
    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
       pfSetICMSts = False
       Exit Function
    End If
       
    '判定IC-M設定ファイル読み込み
     For lngLoop1 = 0 To iId - 1
         bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
         If bRet = False Then
            'ハンドルのクローズ
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
            Call CloseHandle(lngHandle)
            pfSetICMSts = False
            Exit Function
         End If
     Next
          
   'ハンドルのクローズ
     Call CloseHandle(lngHandle)

   '設定更新
   If iId = udtAreaR255.SetteiInfo(iGouki - 1).iId Then
         IDU_SetData udtAreaR255.SetteiInfo(iGouki - 1), iICMSts     'データ変換
   Else
      'ID無し
      pfSetICMSts = False
      Exit Function
   End If
    
    '判定IC-M設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)
    
    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
       pfSetICMSts = False
       Exit Function
    End If

    'ファイルポインタ移動のための読み込み
    '判定IC-M設定ファイル読み込み
     For lngLoop1 = 0 To iId - 2
         bRet = ReadFile(lngHandle, udtAreaR255Work, LenB(udtAreaR255Work), lngRet, 0)
         If bRet = False Then
            'ハンドルのクローズ
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
            Call CloseHandle(lngHandle)
            pfSetICMSts = False
            Exit Function
         End If
     Next
     
    '判定IC-M設定ファイル書込み
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       'ハンドルのクローズ
       Call CloseHandle(lngHandle)
       pfSetICMSts = False
       Exit Function
    End If
    
    'ハンドルのクローズ
    Call CloseHandle(lngHandle)
    
    pfSetICMSts = True

    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_SETTEIFILE_UPDATA_OK, 0)

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetIDSVRSts
'//  機能名称  : ID中継ユニット設定ファイル更新処理
'//  機能概要  : ID中継ユニット設定ファイルの更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iKikists　[OUT]表示ステータス
'//              Integer　iKansiId　[IN]監視設定ID
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetIDSVRSts(iKikists As Integer, iIduId As Integer) As Boolean
    Dim iAreaSts As Integer     '監視設定状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255 As IDU_ID_FMT            '読込み用エリア（255設定用）
    Dim udtAreaR255Work As IDU_ID_FMT            '読込み用エリア（255設定用）
    Dim lngSts As Long
    Dim lngLoop1 As Long
    Dim lngHandle As Long
    Dim FileName As String
    Dim lngRet As Long
    Dim bRet As Boolean
    Dim sSetteiFile As String
       
    On Error Resume Next
    
    pfSetIDSVRSts = False
    
    'ID中継ユニット設定ファイル有無
    FileName = Dir(PATH_IDU_APP & PATH_IDU_SETTEI)
    If FileName = "" Then
       'ID中継ユニット設定ファイルがない場合
       sSetteiFile = PATH_IDU_APP & PATH_SHOKI_IDU_SETTEI
    Else
       'ID中継ユニット設定ファイルがある場合
       sSetteiFile = PATH_IDU_APP & PATH_IDU_SETTEI
    End If
  
    'ID中継ユニット設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)
    
    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
       Exit Function
    End If
       
   'ID中継ユニット設定ファイル読み込み
    For lngLoop1 = 0 To iIduId - 1
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
           'ハンドルのクローズ
           Call CloseHandle(lngHandle)
           pfSetIDSVRSts = False
           Exit Function
        End If
    Next
    
   'ハンドルのクローズ
    Call CloseHandle(lngHandle)
       
    'ID検索
    If iIduId = udtAreaR255.iId Then
       IDU_SetData udtAreaR255, iKikists     'データ変換
    Else
      pfSetIDSVRSts = False
      Exit Function
    End If
           
    'ID中継ユニット設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)
    
    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
       Exit Function
    End If
       
    'ID中継ユニット設定ファイル読み込み
    For lngLoop1 = 0 To iIduId - 2
        bRet = ReadFile(lngHandle, udtAreaR255Work, LenB(udtAreaR255Work), lngRet, 0)
        If bRet = False Then
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
           'ハンドルのクローズ
           Call CloseHandle(lngHandle)
           pfSetIDSVRSts = False
           Exit Function
        End If
    Next
    
    'ID中継ユニット設定ファイル書込み
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       'ハンドルのクローズ
       Call CloseHandle(lngHandle)
       pfSetIDSVRSts = False
       Exit Function
    End If
   
   'ハンドルのクローズ
    Call CloseHandle(lngHandle)
    
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_SETTEIFILE_UPDATA_OK, 0)
    
    pfSetIDSVRSts = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : IDU_SetData
'//  機能名称  : データ変換処理処理
'//  機能概要  : データ変換処理処理を行う。
'//
'//              型        名称        意味
'//  引数      : ID_FMT 　DataArea 　[IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : String　　　        [OUT]　vbNullstring以外：正常。vbNullString    ：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub IDU_SetData(DataArea As IDU_ID_FMT, itest As Integer)
    
    On Error Resume Next
      
    DataArea.bytDATA(0) = itest
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetUpData
'//  機能名称  : IDU設定ファイル、自改設定ファイルの更新処理
'//  機能概要  : 監視盤有、IDU無状態のため、IDU更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Long　 　lngTab 　[IN]機器タイプ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-26   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-237】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetUpData(lngTab As Long) As Boolean
     Dim bRet As Boolean
     Dim iGouki As Integer
     Dim iCnt As Integer
      
     On Error Resume Next
 
     'メールの処理結果異常に処理を行うのは、IC-MとIDサーバーのみ
     If lngTab = 1 Then
        '機器種別：IC-M
        If ZEN_SITEI = sBottom_Sts Then
           '全号機接続・切断釦押下時
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'           For iCnt = 0 To 15
'              '号機別釦が表示されているかどうか
'              If chkICM(iCnt).Visible = True Then
'                 bRet = pfSetSettei(ICM, iCnt, iICM_Sts(iCnt), IdGate.HANTEI_ICM_CONECT_SETTEI)
'              End If
'           Next
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
           For iCnt = CNT_MIN To CONECT_HANTEI_ICM_MAX
'               If gIcmButtonInfo(iCnt).bStatus = True Then                 ' EG20 V3.3.0.1【結合TR-237】削除
' EG20 V3.3.0.1【結合TR-237】追加開始
               ' 設定されている号機、かつ選択されたコーナに対して設定
               If gIcmButtonInfo(iCnt).bStatus = True And _
                    gIcmButtonInfo(iCnt).nCornerNo = tabIcmCorner.Tab + 1 Then
' EG20 V3.3.0.1【結合TR-237】追加終了
                   bRet = pfSetSettei(ICM, gIcmButtonInfo(iCnt).nKanshiNo - 1, _
                                      iICM_Sts(iCnt), IdGate.HANTEI_ICM_CONECT_SETTEI)
               End If
           Next
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
        Else
           '号機別押下時
           iGouki = CInt(sBottom_Sts)
           bRet = pfSetSettei(ICM, iGouki, iICM_Sts(iGouki), IdGate.HANTEI_ICM_CONECT_SETTEI)
        End If
         SetEnableTrue (1)
         psICMConectSts
     ElseIf lngTab = 5 Then
        '機器種別：IDサーバ
         bRet = pfSetSettei(KIKI, 0, iIDSVR_Sts, IdKansiSet.SET_ID_SVR_CONECT_SETTEI)
         SetEnableTrue (2)
         pfGetKiKiConectSts
     End If
     pfSetUpData = True
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetJikai_ICM
'//  機能名称  : 自改設定エリアの判定IC−M設定値の更新処理
'//  機能概要  : 監視盤有、IDU無状態のため、IDU更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [IN]接続・切断タイプ
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-26   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetJikai_ICM(iJikaiSts As Integer, iGouki As Integer) As Boolean
    Dim iAreaSts As Integer
    Dim strMutexName As String
    Dim lngMuHandle As Long
        
    On Error Resume Next
    
    strMutexName = "Mu_" & GGateStatus
    lngMuHandle = dllOpenMutex(strMutexName)            '排他処理(OPEN)
    If lngMuHandle = 0 Then
       'エリア参照不可のため、参照異常
       pfSetJikai_ICM = False
       Exit Function
    End If
  
    dllCloseHandle (lngMuHandle)                 '排他処理(CLOSE)
    
    Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
    '自改設定エリアをオープンする。
    Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
    Idinf_JikaiSettei.IdOpen
    If Idinf_JikaiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       pfSetJikai_ICM = False
       Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
       Exit Function
    End If
             
    '自改設定エリアをＬＯＣＫする。
    Idinf_JikaiSettei.IdLock
    If Idinf_JikaiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       pfSetJikai_ICM = False
       Idinf_JikaiSettei.IdFree
       Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
       Exit Function
    End If
              
    'エリアの内容を読み込む。
    Idinf_JikaiSettei.id = IdGate.HANTEI_ICM_CONECT_SETTEI
    Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
    If Idinf_JikaiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       pfSetJikai_ICM = False
       Idinf_JikaiSettei.IdFree
       Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
       Exit Function
    End If
               
    '設定内容を取得
     iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
     If iJikaiSts <> iAreaSts Then
        'エリア値の差分があれば保守が更新
        Idinf_JikaiSettei.SetICM_Sts iGouki - 1, iJikaiSts
     End If
     
     Idinf_JikaiSettei.IdFree
     Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
    
    pfSetJikai_ICM = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetJikai_ICM
'//  機能名称  : 監視設定エリアのIDサーバー設定値の更新処理
'//  機能概要  : 監視盤有、IDU無状態のため、IDU更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [IN]接続・切断タイプ
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-26   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetKansi_IDSVR(iJikaiSts As Integer) As Boolean
    Dim iAreaSts As Integer
    Dim strMutexName As String
    Dim lngMuHandle As Long
        
    On Error Resume Next
    
    strMutexName = "Mu_" & GKansiStatus
    lngMuHandle = dllOpenMutex(strMutexName)            '排他処理(OPEN)
    If lngMuHandle = 0 Then
       'エリア参照不可のため、参照異常
       pfSetKansi_IDSVR = False
       Exit Function
    End If
  
    dllCloseHandle (lngMuHandle)                 '排他処理(CLOSE)
    
    Set Idinf_KansiSettei = New IdInfProc              '監視設定エリア
    '自改設定エリアをオープンする。
    Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
    Idinf_KansiSettei.IdOpen
    If Idinf_KansiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       pfSetKansi_IDSVR = False
       Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
       Exit Function
    End If
             
    '監視設定エリアをＬＯＣＫする。
    Idinf_KansiSettei.IdLock
    If Idinf_KansiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       pfSetKansi_IDSVR = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
       Exit Function
    End If
              
    'エリアの内容を読み込む。
    Idinf_KansiSettei.id = IdKansiSet.SET_ID_SVR_CONECT_SETTEI
    Idinf_KansiSettei.IdGet
    If Idinf_KansiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       pfSetKansi_IDSVR = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
       Exit Function
    End If
               
    '設定内容を取得
     iAreaSts = Idinf_KansiSettei.DataArea(0)
     If iJikaiSts <> iAreaSts Then
        'エリア値の差分があれば保守が更新
        Idinf_KansiSettei.SetIDSVR iJikaiSts
     End If
     
     Idinf_KansiSettei.IdFree
     Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
    
    pfSetKansi_IDSVR = True
    
End Function
'V1.4.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称   ： psTakuConectSts
'//  機能名称   ： 操作卓タブ表示処理
'//  機能概要   ： 操作卓タブにおける操作卓の通信状態取得処理を行う。
'//
'//                型        名称      意味
'//  引数       ： なし
'//
'//                型        値          意味
'//  戻り値     ： なし
'//
'/   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'/   REVISIONS  ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考       ：
'///////////////////////////////////////////////////////////////////
Private Sub psTakuConectSts()
  
    Dim iJikaiSts As Integer             ' 操作卓通信状態
    Dim intLoop As Integer              ' ループカウンター


    On Error Resume Next
  
    ' コーナ名称設定処理
    Call gsGetCornerName
   
    For intLoop = 0 To CONECT_TAKU_CONTROLMAX
      
        '設定ありのコーナを活性にする
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
            ' /////////////////////////////////////////////////
            ' // ラベル（コーナー名称表示）
            LblTaku(intLoop).Caption = gstrCornerName(intLoop)
            LblTaku(intLoop).Visible = True
            
            lblTakuSts(intLoop).Visible = True

            chkTaku(intLoop).Visible = True

            '監視盤起動有無チェック
            If CheckAppStart(PROC_KANRI) <> 0 Then
               '監視盤起動有り時
               '通信状態取得処理
                pfGetTakuConectSts iJikaiSts, (intLoop + 1)
                If iJikaiSts = CONECTSTS_NORMAL Then
                    '正常時
                    lblTakuSts(intLoop).Caption = CONECT_NORMAL
                    lblTakuSts(intLoop).ForeColor = CONECT_OK
                ElseIf iJikaiSts = CONECTSTS_ERROR Then
                    '異常時
                    lblTakuSts(intLoop).Caption = CONECT_ERROR
                    lblTakuSts(intLoop).ForeColor = CONECT_ERROR_COLOR
                ElseIf iJikaiSts = CONECTSTS_END Then
                    '切離時
                    lblTakuSts(intLoop).Caption = CONECT_END
                    lblTakuSts(intLoop).ForeColor = CONECT_CUT
                Else
                  '「通信接続・切断画面：エリア・ファイル参照異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_API, CONECT_AREA_FILE_NOTACCESS_ERROR, 0)
                    '上記以外：状態取得異常
                    lblTakuSts(intLoop).Caption = GET_CONECT_ERROR
                    chkTaku(intLoop).Visible = False
                End If
            Else
                lblTakuSts(intLoop).Caption = CONECT_ERROR
                lblTakuSts(intLoop).ForeColor = CONECT_ERROR_COLOR
            End If
            
            '号機別釦情報取得
            pfGetTakuSts iJikaiSts, (intLoop + 1)
            If iJikaiSts = CONECTSTS_ERROR Then
                '接続の場合
                chkTaku(intLoop).Value = 0
                chkTaku(intLoop).Caption = "接続"
                chkTaku(intLoop).BackColor = CONECT_ON
            ElseIf iJikaiSts = CONECTSTS_END Then
                '切離の場合
                chkTaku(intLoop).Value = 1
                chkTaku(intLoop).Caption = "切離"
                chkTaku(intLoop).BackColor = CONECT_OFF
            ElseIf iJikaiSts = GET_CONECTSTS_ERROR Then
                '号機別取得異常時は非表示。
                chkJikai(intLoop).Visible = False '押下不可
            End If

        End If
    Next intLoop
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称   ： pfGetTakuConectSts
'//  機能名称   ： 操作卓タブ表示処理(監視盤起動有りのためエリア参照可能)
'//  機能概要   ： 操作卓タブの通信状態取得処理を行う。
'//
'//                型        名称      意味
'//  引数       ： Integer　iJikaiSts [OUT]表示ステータス
'//                Integer　iGouki  　[IN]処理対象号機番号
'//
'//                型        値          意味
'//  戻り値     ： なし
'//
'/   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'/   REVISIONS  ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考       ：pfGetjikaiConectSts流用
'///////////////////////////////////////////////////////////////////
Private Function pfGetTakuConectSts(iJikaiSts As Integer, iGouki As Integer)
    Dim iAreaSts   As Integer               ' 自改設定ファイル状態値
    Dim iConectSts As Integer               ' 上位通信状態値
    Dim iAreaID As Integer                  ' 上位通信状態エリアＩＤ
    
    Dim iDataArea_Tuushin As Integer        '自改通信状態
    Dim strMutexName    As String           'ミューテックス名
    Dim lngMuHandle     As Long             '排他処理用ハンドル

    On Error Resume Next
    
    ' 上位エリアＩＤは号機順が原則
    iAreaID = IdKikiComSts.ID_TAKU1_COM + (iGouki - 1)
    
    'ＩＤ別情報操作クラスの生成
    Set Idinf_Jyoui = New IdInfProc                    '上位通信状態エリア
   '参照(上位機器通信状態)エリア名を設定
    Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui
    Idinf_Jyoui.IdOpen
    If Idinf_Jyoui.Errsts <> 0 Then
      'データ参照異常時はブランク表示設定にして処理終了。
      iJikaiSts = GET_CONECTSTS_ERROR
      Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
      Exit Function
    End If
    
    '参照(上位機器通信状態)エリアをＬＯＣＫする。
    Idinf_Jyoui.IdLock
    If Idinf_Jyoui.Errsts <> 0 Then
      'データ参照異常時はブランク表示設定にして処理終了。
      iJikaiSts = GET_CONECTSTS_ERROR
      Idinf_Jyoui.IdFree
      Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
      Exit Function
    End If
    
     'エリアの内容を読み込む。
    Idinf_Jyoui.id = iAreaID
    Idinf_Jyoui.GetInf (CONECT)
    If Idinf_Jyoui.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定にして処理終了。
       iJikaiSts = GET_CONECTSTS_ERROR
       Idinf_Jyoui.IdFree
       Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
       Exit Function
    End If
    
    iConectSts = CInt(Idinf_Jyoui.DataArea(0))
    Idinf_Jyoui.IdFree
    Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
    
  If iConectSts <> IdGateCom.GATE_COM_CONNECT_NORMAL Then
     '自改通信状態が正常以外の場合以下を行う。
     Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
     '自改設定エリアをオープンする。
     Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
     Idinf_JikaiSettei.IdOpen
     If Idinf_JikaiSettei.Errsts <> 0 Then
      'データ参照異常時はブランク表示設定を行う。
      iJikaiSts = GET_CONECTSTS_ERROR
      Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
      Exit Function
     End If
    
      '自改設定エリアをＬＯＣＫする。
      Idinf_JikaiSettei.IdLock
      If Idinf_JikaiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iJikaiSts = GET_CONECTSTS_ERROR
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
        Exit Function
      End If
     
      'エリアの内容を読み込む。
      Idinf_JikaiSettei.id = IdGate.TAKU_CONECT_SETTEI        ' 操作卓通信設定
      Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
      If Idinf_JikaiSettei.Errsts <> 0 Then
        'データ参照異常時はブランク表示設定を行う。
        iJikaiSts = GET_CONECTSTS_ERROR
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
        Exit Function
      End If
      
      '設定内容を取得
      iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
      Select Case iAreaSts
'         Case 1                    ' EG20 V3.0.0.2削除
         Case 0                     ' EG20 V3.0.0.2追加
           '接続
           iJikaiSts = CONECTSTS_ERROR
           Idinf_JikaiSettei.IdFree
           Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
           Exit Function
'         Case 0                    ' EG20 V3.0.0.2削除
         Case 1                     ' EG20 V3.0.0.2追加
           iJikaiSts = CONECTSTS_END
           Idinf_JikaiSettei.IdFree
           Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
           Exit Function
        End Select
  End If
   
   '状態：正常
   iJikaiSts = CONECTSTS_NORMAL
   Idinf_JikaiSettei.IdFree
   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称   ： pfGetTakuSts
'//  機能名称   ： 自改タブ表示処理(監視盤起動有無対応参照)
'//  機能概要   ： 自改タブの号機別釦状態取得処理を行う。
'//
'//                型        名称      意味
'//  引数       ： Integer　iJikaiSts [OUT]表示ステータス
'//                Integer　iGouki  　[IN]処理対象号機番号
'//
'//                型        値          意味
'//  戻り値     ： なし
'//
'/   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'/   REVISIONS  ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考       ：pfGetJikaiSts流用
'///////////////////////////////////////////////////////////////////
Private Function pfGetTakuSts(iJikaiSts As Integer, iGouki As Integer)
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255     As GATE_INFO        '読込み用エリア（255設定用）
    Dim lngSts          As Long             'ヒットエリアID
    Dim lngLoop1        As Long             'カウンター
    Dim lngHandle       As Long             'ハンドル
    Dim FileName        As String           'ファイル有無チェック
    Dim lngRet          As Long             '戻り値
    Dim bRet            As Boolean          '読み込み結果戻り値
    Dim sSetteiFile     As String           'ファイルパス　'V1.4.0.1　ADD
    
    On Error Resume Next
'V1.4.0.1 DEL START
'    '自改設定ファイル有無
'    FileName = Dir(G_SETTEI_FILE)
'    If FileName = "" Then
'       '無ければ参照不可のため参照異常
'       iJikaiSts = GET_CONECTSTS_ERROR
'       Exit Function
'    End If
'V1.4.0.1 DEL END
'V1.4.0.1 ADD START
   '自改設定ファイル有無
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '自改設定ファイルがない場合
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '自改設定ファイルがある場合
       sSetteiFile = G_SETTEI_FILE
    End If
'V1.4.0.1 ADD END

    '監視盤起動有無チェック
    If CheckAppStart(PROC_KANRI) = 0 Then
        
        '自改設定ファイルをオープン
'        lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 DEL
        lngHandle = CreateFile(sSetteiFile, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
           'オープン異常時は参照不可のため参照異常
           iJikaiSts = GET_CONECTSTS_ERROR
           Exit Function
        End If
        
        '自改設定ファイル読み込み
        For lngLoop1 = 0 To iGouki - 1
            bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        Next
        
        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
        lngSts = SerchId(udtAreaR255, IdGate.TAKU_CONECT_SETTEI)   ' 操作卓通信設定
        If lngSts >= 0 Then
           'IDが有った場合
           iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
        Else
          ' 該当ＩＤ無しの場合参照異常
          iJikaiSts = GET_CONECTSTS_ERROR
          Exit Function
        End If
        
        Select Case iAreaSts
           Case 1
             '接続
              iJikaiSts = CONECTSTS_ERROR
              Exit Function
           Case 0
              iJikaiSts = CONECTSTS_END
              Exit Function
        End Select
    Else
     
         Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
         '自改設定エリアをオープンする。
          Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
          Idinf_JikaiSettei.IdOpen
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
             iJikaiSts = GET_CONECTSTS_ERROR
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
          End If
             
          '自改設定エリアをＬＯＣＫする。
          Idinf_JikaiSettei.IdLock
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
             iJikaiSts = GET_CONECTSTS_ERROR
             Idinf_JikaiSettei.IdFree
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
           End If
              
           'エリアの内容を読み込む。
            Idinf_JikaiSettei.id = IdGate.TAKU_CONECT_SETTEI   ' 操作卓通信設定
            Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
            If Idinf_JikaiSettei.Errsts <> 0 Then
               'データ参照異常時はブランク表示設定を行う。
                iJikaiSts = GET_CONECTSTS_ERROR
                Idinf_JikaiSettei.IdFree
                Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                Exit Function
            End If
               
            '設定内容を取得
             iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
             Select Case iAreaSts
'                 Case 1                    ' EG20 V3.0.0.2削除
                 Case 0                     ' EG20 V3.0.0.2追加
                  '接続
                   iJikaiSts = CONECTSTS_ERROR
                   Idinf_JikaiSettei.IdFree
                   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                   Exit Function
'                 Case 0                    ' EG20 V3.0.0.2削除
                 Case 1                     ' EG20 V3.0.0.2追加
                   iJikaiSts = CONECTSTS_END
                   Idinf_JikaiSettei.IdFree
                   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                   Exit Function
             End Select
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : chkTaku_Click
'//  機能名称  : 「接続」「切離」釦押下時処理
'//  機能概要  : 釦名称処理を行う。
'//              操作卓部：
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                  EG20フェーズ１対応
'//                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'//   REVISIONS  ： (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：chkJikai_Click流用
'///////////////////////////////////////////////////////////////////
Private Sub chkTaku_Click(Index As Integer)
    Dim bRet As Boolean               'メール送信戻り値
    Dim iResponse As Integer          'メッセージの戻り値
    Dim iCnt As Integer               'カウンター
    Dim lSts As Long                  'ステータス値
    Dim bFlag As Boolean              '受信メールフラグ
    Dim lngErrCode As Long            'エラーコード
    Dim nInfoIndex As Integer         ' 保存情報インデックス    ' EG20 V1.1.1.1 追加
    On Error Resume Next
    
    bFlag = True                      'V1.4.0.1　ADD
    
    If iUpDataFlag <> 0 Or iALLGoukiFlag <> 0 Or _
       iShokiFlag = 1 Or iCancelFlag = 1 Or iMailRcvFlag = 1 Then
       iCancelFlag = 0
       Exit Sub
    End If
     
    '画面をロックする。
    SetEnableFalse (3)

     If chkTaku(Index).Value = 0 Then
         '「通信接続・切断画面：切離→接続 設定」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMEND_TOSTA_BUTTOM, 0)

         '切離→接続
         '「通信接続確認」ポップアップ画面表示
         iResponse = MsgBox("指定した外部機器との通信接続を開始します。よろしいですか？", _
                             vbOKCancel + vbQuestion, _
                             "通信接続確認")
         If iResponse = vbOK Then
            '通信設定要求CMD(自改,接続,対象号機)を監マプロセスに送信する
            chkTaku(Index).Caption = "接続"
            chkTaku(Index).BackColor = CONECT_ON
            'ヘッダ部共通作成処理
            SendMailHeader
            udtMail.dwRequestKIKI = ML_DT_TAKU
            udtMail.dwRequestConectType = ML_REQUEST_CONECT
            For iCnt = 0 To CONECT_TAKU_CONTROLMAX
              udtMail.dwGouki(iCnt) = ML_TARGET_OFF
            Next
            udtMail.dwGouki(Index) = ML_TARGET_ON
           
           If CheckAppStart(PROC_KANRI) = 0 Then
              '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
              bRet = pfSetSettei(TAKU, Index, TAKU_CONECT_SETU, IdGate.TAKU_CONECT_SETTEI)    ' 操作卓通信設定
              bFlag = False
              GoTo Error_Click
           End If

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを表示する
            Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

            bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
            If False = bRet Then
               '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
               '送信異常時：画面ロック解除
               GoTo Error_Click
            End If
               '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
               Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
         Else
            '「キャンセル」釦押下時
            GoTo Error_Click
         End If
     End If
    
     If chkTaku(Index).Value = 1 Then
        '「通信接続・切断画面：切離→接続 設定」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_FROMSTA_TOEND_BUTTOM, 0)
 
        '「通信切断確認」ポップアップ画面表示
        iResponse = MsgBox("指定した外部機器との通信を切断します。よろしいですか？", _
                           vbOKCancel + vbQuestion, _
                           "通信切断確認")
        If iResponse = vbOK Then
           '通信設定要求CMD(自改,接続,対象号機)を監マプロセスに送信する
           chkTaku(Index).Caption = "切離"
           chkTaku(Index).BackColor = CONECT_OFF
           'ヘッダ部共通作成処理
           SendMailHeader
           udtMail.dwRequestKIKI = ML_DT_TAKU
           udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
           For iCnt = 0 To CONECT_TAKU_CONTROLMAX
             udtMail.dwGouki(iCnt) = ML_TARGET_OFF
           Next
           udtMail.dwGouki(Index) = ML_TARGET_ON
     
           If CheckAppStart(PROC_KANRI) = 0 Then
              '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
              bRet = pfSetSettei(TAKU, Index, TAKU_CONECT_DAN, IdGate.TAKU_CONECT_SETTEI)     ' 操作卓通信設定
              bFlag = False
              GoTo Error_Click
           End If
           
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを表示する
            Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           
             bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
             If False = bRet Then
                '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
                lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
                '送信異常時：画面ロック解除
                GoTo Error_Click
             End If
                '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
        Else
           '「キャンセル」釦押下時
           GoTo Error_Click
        End If
    End If
Exit Sub

'キャンセル釦押下、又は送信異常時
Error_Click:

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
  
  If bFlag = True Then                  'V1.4.0.1 ADD
     If chkTaku(Index).Value = 0 Then
        '「キャンセル」釦押下時
        SetEnableTrue (3)
        iCancelFlag = 1
        chkTaku(Index).Caption = "切離"
        chkTaku(Index).BackColor = CONECT_OFF
        chkTaku(Index).Value = 1
        Exit Sub
    End If
   If chkTaku(Index).Value = 1 Then
      SetEnableTrue (3)
      iCancelFlag = 1
      chkTaku(Index).Caption = "接続"
      chkTaku(Index).BackColor = CONECT_ON
      chkTaku(Index).Value = 0
      Exit Sub
   End If
'V1.4.0.1 ADD START
  Else
   SetEnableTrue (3)
   iShokiFlag = 1
   psTakuConectSts
   iShokiFlag = 0
  End If
'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdInOutTaku_Click
'//  機能名称  : 「全号機接続」「全号機切離」釦押下時処理
'//  機能概要  : 釦名称処理を行う。
'//              操作卓部：
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'/   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS  ： (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：cmdInOutJikai_Click流用
'///////////////////////////////////////////////////////////////////
Private Sub cmdInOutTaku_Click(Index As Integer)
  Dim iCnt As Integer               'カウンター
  Dim iResponse As Integer          'メッセージボックスの戻り値
  Dim bRet As Boolean               'メール送信戻り値
  Dim lngErrCode As Long            'エラーコード
  Dim bytWork()   As Byte
  Dim i As Integer
  Dim bFlag As Boolean              'エラーフラグ処理　'V1.4.0.1 ADD

  Erase bytWork
  
  On Error Resume Next
 
  '画面をロックする。
  SetEnableFalse (3)


  If Index = 0 Then
     '「通信接続・切断画面：全号機接続釦押下」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_ALLGOUKI_STA_BUTTOM, 0)
 
     '「通信接続確認」ポップアップ画面表示
     iResponse = MsgBox("指定した外部機器との通信接続を開始します。よろしいですか？", _
                         vbOKCancel + vbQuestion, _
                         "通信接続確認")
     If iResponse = vbOK Then
        '「全号機接続」釦押下時
        iALLGoukiFlag = 1
        '通信設定要求CMD(自改,接続)を監マプロセスに送信する
        'ヘッダ部共通作成処理
        SendMailHeader
        udtMail.dwRequestKIKI = ML_DT_TAKU
        udtMail.dwRequestConectType = ML_REQUEST_CONECT
        For iCnt = 0 To CONECT_TAKU_CONTROLMAX
           If chkTaku(iCnt).Visible = True Then
              chkTaku(iCnt).Caption = "接続"
              chkTaku(iCnt).BackColor = CONECT_ON
              chkTaku(iCnt).Value = 0
              udtMail.dwGouki(iCnt) = ML_TARGET_ON
           Else
              udtMail.dwGouki(iCnt) = ML_TARGET_OFF
           End If
        Next
        
        If CheckAppStart(PROC_KANRI) = 0 Then
           '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
           For iCnt = 0 To CONECT_TAKU_CONTROLMAX
              If chkTaku(iCnt).Visible = True Then
                  bRet = pfSetSettei(TAKU, iCnt, _
                                     TAKU_CONECT_SETU, IdGate.TAKU_CONECT_SETTEI)       ' 操作卓通信設定
              End If
           Next
           bFlag = False
           GoTo Error_Click
        End If
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
        If False = bRet Then
           '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
           lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
           Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
           GoTo Error_Click
           Exit Sub
        End If
         '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
     Else
        '「キャンセル」釦押下時
         GoTo Error_Click
     End If
 Else
    '「通信接続・切断画面：全号機切離釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_ALLGOUKI_END_BUTTOM, 0)
 
    '「通信切断確認」ポップアップ画面表示
     iResponse = MsgBox("指定した外部機器との通信を切断します。よろしいですか？", _
                         vbOKCancel + vbQuestion, _
                         "通信切断確認")
     If iResponse = vbOK Then
        iALLGoukiFlag = 1
        '通信設定要求CMD(自改,切断)を監マプロセスに送信する
        'ヘッダ部共通作成処理
        SendMailHeader
        udtMail.dwRequestKIKI = ML_DT_TAKU
        udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
        '「全号機切離」釦押下時
        For iCnt = 0 To CONECT_TAKU_CONTROLMAX
           If chkTaku(iCnt).Visible = True Then
               chkTaku(iCnt).Caption = "切離"
               chkTaku(iCnt).BackColor = CONECT_OFF
               chkTaku(iCnt).Value = 1
               udtMail.dwGouki(iCnt) = ML_TARGET_ON
            Else
               udtMail.dwGouki(iCnt) = ML_TARGET_OFF
            End If
        Next
        
        If CheckAppStart(PROC_KANRI) = 0 Then
           '監視盤未起動時(処理対象タブ,号機番号,設定値,設定ID)
            For iCnt = 0 To CONECT_TAKU_CONTROLMAX
              If chkTaku(iCnt).Visible = True Then
                  bRet = pfSetSettei(TAKU, iCnt, _
                                     TAKU_CONECT_DAN, IdGate.TAKU_CONECT_SETTEI)    ' 操作卓通信設定
              End If
           Next
           bFlag = False
           GoTo Error_Click
        End If
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
        If False = bRet Then
           '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
           lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
           Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
           GoTo Error_Click
           Exit Sub
        End If
          '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
     Else
       '「キャンセル」釦押下時
        GoTo Error_Click
     End If
 End If
Exit Sub

'キャンセル釦押下、又は送信異常時
Error_Click:

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    If bFlag = True Then
        '「キャンセル」釦押下時
        SetEnableTrue (3)
        psTakuConectSts
    Else
        SetEnableTrue (3)
        psTakuConectSts
        iALLGoukiFlag = 0
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称   ： GetButtonNo
'//  機能名称   ： 自改タブ表示処理
'//  機能概要   ： 自改タブの画面表示処理を行う。
'//
'//                型        名称      意味
'//  引数       ： Integer   nCorner   論理コーナ
'//             ： Integer   nGokiNum  論理号機番号
'//
'//                型        値          意味
'//  戻り値     ： Integer   nControlNo  釦コントロール番号
'//
'/   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'/   REVISIONS  ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考       ：
'///////////////////////////////////////////////////////////////////
Private Function GetButtonNo(nCorner As Integer, nGokiNum As Integer) As Integer

    Dim nControlNo As Integer
    

    If nCorner < 1 Or nCorner > 6 Then
        GetButtonNo = 0
        Exit Function
    End If

    If nGokiNum < 1 Or nGokiNum > 16 Then
        GetButtonNo = 0
        Exit Function
    End If

    nControlNo = (CONTROL_CORNERMAX * (nCorner - 1)) + (nGokiNum - 1)
    GetButtonNo = nControlNo

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称   ： GetKikiButtonNo
'//  機能名称   ： 上位タブ表示処理
'//  機能概要   ： 上位タブの画面表示処理を行う。
'//
'//                型        名称      意味
'//  引数       ： Integer   nCorner   論理コーナ
'//             ： Integer   nGokiNum  論理号機番号
'//
'//                型        値          意味
'//  戻り値     ： Integer   nControlNo  釦コントロール番号
'//
'/   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'/   REVISIONS  ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考       ：
'///////////////////////////////////////////////////////////////////
Private Function GetKikiButtonNo(nCorner As Integer, nGokiNum As Integer) As Integer

    Dim nControlNo As Integer
    

    If nCorner < 1 Or nCorner > 2 Then
        GetKikiButtonNo = 0
        Exit Function
    End If

    If nGokiNum < 1 Or nGokiNum > 10 Then
        GetKikiButtonNo = 0
        Exit Function
    End If

    nControlNo = (CONTROL_KIKICORNERMAX * (nCorner - 1)) + (nGokiNum - 1)
    GetKikiButtonNo = nControlNo

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称   ： InitCornerTab
'//  機能名称   ： コーナタブ設定処理
'//  機能概要   ： コーナタブの設定処理
'//
'//                型        名称      意味
'//  引数       ： なし
'//
'//                型        値          意味
'//  戻り値     ： なし
'//
'/   ORIGINAL   ： (EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ１対応
'/                  EG20統合監視盤USDM対応番号【Mainte_03_01】
'/   REVISIONS  ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考       ：
'///////////////////////////////////////////////////////////////////
Private Sub InitCornerTab()

    Dim intLoop         As Integer          ' ループカウンタ
    Dim strCorner1 As String                ' 文字列格納エリア1
    Dim strCorner2 As String                ' 文字列格納エリア2
    
    On Error Resume Next
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 操作卓設定
    ' /////////////////////////////////////////////////////////////////////////
    ' コーナ名称設定処理
    Call gsGetCornerName
    
    For intLoop = 0 To CONECT_TAKU_CONTROLMAX
    
        '設定ありのコーナを活性にする
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
            strCorner1 = MidB(gstrCornerName(intLoop), 1, 12)
            strCorner2 = MidB(gstrCornerName(intLoop), 13, 24)
            ' /////////////////////////////////////////////////
            ' // ラベル（コーナー名称表示）
            tabCorner.TabCaption(intLoop) = strCorner1 & vbCrLf & strCorner2
            tabCorner.TabEnabled(intLoop) = True

            ' /////////////////////////////////////////////////
            ' // ラベル（コーナー名称表示）
            tabIcmCorner.TabCaption(intLoop) = strCorner1 & vbCrLf & strCorner2
            tabIcmCorner.TabEnabled(intLoop) = True

        Else
            tabCorner.TabCaption(intLoop) = ""
            tabCorner.TabEnabled(intLoop) = False
            tabIcmCorner.TabCaption(intLoop) = ""
            tabIcmCorner.TabEnabled(intLoop) = False
        End If
    Next intLoop
    
    ' タブ設定を一番左側に設定する
    tabCorner.Tab = 0
    tabIcmCorner.Tab = 0
    tabKikiCorner.Tab = 0

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psAddKikiCornerName
'//  機能名称  : 上位機器コーナ名称追加処理
'//  機能概要  : 上位機器名称に対してコーナ名称を付加する必要があれば追加する。
'//
'//              型        名称      意味
'//  引数      : Integer　 nIndex  　[IN]上位機器釦情報インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-01  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psAddKikiCornerName(nIndex As Integer)

    Dim nTabNo As Integer                   ' タブインデックス
    Dim nCorner As Integer                  ' コーナインデックス
    Dim szCornerName As String              ' コーナ名称
    Dim nNullIndex As Integer               ' 文字数ワーク
    Dim nRonriType As Integer               ' 論理タイプ（接続値）

    nTabNo = CONECT_KIKI_TABKANSHI          ' タブインデックスの初期値は統合監視盤
    nCorner = 0                             ' コーナ設定不要
    ' 1.対象外部機器上位機器通信状態エリアIDをチェックして
    '   接続対象を選別する。
    Select Case gTransButtonInfo(nIndex).iAreaID
    Case IdKikiComSts.ID_DESYU_COM                                       ' 1:デ集通信状態
        nCorner = 1
        nRonriType = 0
    Case IdKikiComSts.ID_DESYU2_COM                                      ' 9:デ集2通信状態
        nCorner = 2
        nRonriType = 0
    Case IdKikiComSts.ID_DESYU3_COM                                      ' 10:デ集3通信状態
        nCorner = 3
        nRonriType = 0
    Case IdKikiComSts.ID_DESYU4_COM                                      ' 11:デ集4通信状態
        nCorner = 4
        nRonriType = 0
    Case IdKikiComSts.ID_DESYU5_COM                                      ' 12:デ集5通信状態
        nCorner = 5
        nRonriType = 0
    Case IdKikiComSts.ID_DESYU6_COM                                      ' 13:デ集6通信状態
        nCorner = 6
        nRonriType = 0
    Case IdKikiComSts.ID_ENKAKU_COM                                      ' 2:遠隔通信状態
        nTabNo = CONECT_KIKI_TABTAKU        ' 操作卓設定
        nCorner = 1
        nRonriType = 0
    Case IdKikiComSts.ID_ENKAKU2_COM                                     ' 21:遠隔2通信状態（エリア定義なし）
        nTabNo = CONECT_KIKI_TABTAKU        ' 操作卓設定
        nCorner = 2
        nRonriType = 0
    Case IdKikiComSts.ID_ENKAKU3_COM                                     ' 22:遠隔3通信状態（エリア定義なし）
        nTabNo = CONECT_KIKI_TABTAKU        ' 操作卓設定
        nCorner = 3
        nRonriType = 0
    Case IdKikiComSts.ID_ENKAKU4_COM                                     ' 23:遠隔4通信状態（エリア定義なし）
        nTabNo = CONECT_KIKI_TABTAKU        ' 操作卓設定
        nCorner = 4
        nRonriType = 0
    Case IdKikiComSts.ID_ENKAKU5_COM                                     ' 24:遠隔5通信状態（エリア定義なし）
        nTabNo = CONECT_KIKI_TABTAKU        ' 操作卓設定
        nCorner = 5
        nRonriType = 0
    Case IdKikiComSts.ID_ENKAKU6_COM                                     ' 25:遠隔6通信状態（エリア定義なし）
        nTabNo = CONECT_KIKI_TABTAKU        ' 操作卓設定
        nCorner = 6
        nRonriType = 0
' EG20 V5.2.0.1追加開始
    Case IdKikiComSts.ID_TOMAS_COM                                       ' 4:ＴＭサーバ通信状態
        nRonriType = 0
' EG20 V5.2.0.1追加終了
    Case Else
        nRonriType = 1
    End Select

    gTransButtonInfo(nIndex).nCornerNo = nTabNo                      ' タブ番号（統合監視盤、操作卓）
    If nTabNo = CONECT_KIKI_TABTAKU Then
        ' 監視エリアＩＤを自改ＩＤに変更
        gTransButtonInfo(nIndex).iAreaID = IdGate.ENKAKUKIKI_JIKAIAREAID
    End If
    If nCorner <> 0 Then
        ' コーナ名称の付加
        nNullIndex = InStr(gstrCornerName(nCorner - 1), Chr(0))
        If nNullIndex <> 0 Then
            szCornerName = "（" & Left(gstrCornerName(nCorner - 1), nNullIndex - 1) & "）"
        Else
'            szCornerName = ""                                          ' EG20 V3.3.0.1削除
            szCornerName = "（" & gstrCornerName(nCorner - 1) & "）"    ' EG20 V3.3.0.1追加
        End If
        gTransButtonInfo(nIndex).sGetInf = gTransButtonInfo(nIndex).sGetInf + szCornerName
    End If
    gTransButtonInfo(nIndex).nRonriType = nRonriType                ' 論理タイプ
    gTransButtonInfo(nIndex).nCorner = nCorner                      ' コーナ番号

End Sub

