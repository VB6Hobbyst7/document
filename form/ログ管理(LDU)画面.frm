VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLDULogkanri 
   BorderStyle     =   0  'なし
   Caption         =   "                                                                  LDユーティリティログ管理"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9480
      Top             =   3240
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "媒体取外"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   135
      Top             =   6540
      Width           =   2600
   End
   Begin VB.TextBox txtDummy 
      Height          =   375
      Left            =   11400
      TabIndex        =   0
      Text            =   "Text11"
      Top             =   15420
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogHyouzi 
      Caption         =   "  ログ表示     (テキスト表示)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   75
      Top             =   540
      Width           =   2600
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "ログ媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   76
      Top             =   1980
      Width           =   2600
   End
   Begin VB.CommandButton cmdSqllog 
      Caption         =   "   SQLログ     媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   77
      Top             =   5400
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "  ログ管理    画面へ戻る"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   78
      Top             =   7920
      Width           =   2600
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8535
      Left            =   120
      TabIndex        =   79
      Top             =   420
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15055
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "表示ファイル指定"
      TabPicture(0)   =   "ログ管理(LDU)画面.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LstFile"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "optApp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optHoshu"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRefresh"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFile"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblStart"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblEnd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblSize"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "表示項目指定"
      TabPicture(1)   =   "ログ管理(LDU)画面.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frmMod"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmShubetu"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmKekka"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frmOpt"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frmHani"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "表示号機指定"
      TabPicture(2)   =   "ログ管理(LDU)画面.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdHHisentaku"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdHSentaku"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdZHisentaku"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdZSentaku"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "tabCorner"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
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
         Left            =   -68280
         TabIndex        =   236
         Top             =   1200
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
         Left            =   -70440
         TabIndex        =   235
         Top             =   1200
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
         Left            =   -72600
         TabIndex        =   234
         Top             =   1200
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
         Left            =   -74760
         TabIndex        =   233
         Top             =   1200
         Width           =   2000
      End
      Begin VB.ListBox LstFile 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5820
         Left            =   -74640
         MultiSelect     =   2  '拡張
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   8055
      End
      Begin VB.OptionButton optApp 
         Caption         =   "アプリケーションログ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74280
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optHoshu 
         Caption         =   "保守プログラムログ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74280
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ログ切替"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -69240
         TabIndex        =   3
         Top             =   600
         Width           =   2295
      End
      Begin VB.Frame frmHani 
         Caption         =   "表示範囲指定"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   85
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton optHaninasi 
            Caption         =   "範囲指定無"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optHaniari 
            Caption         =   "範囲指定有"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtStNen 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "9999"
            Top             =   210
            Width           =   615
         End
         Begin VB.TextBox txtStTuki 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtStZi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   6240
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtStHi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   5400
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtStFun 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   7080
            MaxLength       =   2
            TabIndex        =   11
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtEdNen 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   12
            Text            =   "9999"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtEdTuki 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtEdZi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   6240
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtEdHi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   5400
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtEdFun 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'ｵﾌ固定
            Left            =   7080
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblSt 
            Caption         =   "開始"
            Enabled         =   0   'False
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
            Left            =   2760
            TabIndex        =   97
            Top             =   270
            Width           =   495
         End
         Begin VB.Label lblStNen 
            Caption         =   "年"
            Enabled         =   0   'False
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
            Left            =   4200
            TabIndex        =   96
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStTuki 
            Caption         =   "月"
            Enabled         =   0   'False
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
            Left            =   5040
            TabIndex        =   95
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStHi 
            Caption         =   "日"
            Enabled         =   0   'False
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
            Left            =   5880
            TabIndex        =   94
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStZi 
            Caption         =   "時"
            Enabled         =   0   'False
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
            Left            =   6720
            TabIndex        =   93
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStFun 
            Caption         =   "分"
            Enabled         =   0   'False
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
            Left            =   7560
            TabIndex        =   92
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblEd 
            Caption         =   "終了"
            Enabled         =   0   'False
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
            Left            =   2760
            TabIndex        =   91
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblEdNen 
            Caption         =   "年"
            Enabled         =   0   'False
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
            Left            =   4200
            TabIndex        =   90
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdTuki 
            Caption         =   "月"
            Enabled         =   0   'False
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
            Left            =   5040
            TabIndex        =   89
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdHi 
            Caption         =   "日"
            Enabled         =   0   'False
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
            Left            =   5880
            TabIndex        =   88
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdZi 
            Caption         =   "時"
            Enabled         =   0   'False
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
            Left            =   6720
            TabIndex        =   87
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdFun 
            Caption         =   "分"
            Enabled         =   0   'False
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
            Left            =   7560
            TabIndex        =   86
            Top             =   660
            Width           =   255
         End
      End
      Begin VB.Frame frmOpt 
         Caption         =   "表示オプション"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   84
         Top             =   1500
         Width           =   2775
         Begin VB.OptionButton optShousai 
            Caption         =   "詳細表示"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optSam 
            Caption         =   "サマリー表示"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame frmKekka 
         Caption         =   "処理結果指定"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3120
         TabIndex        =   83
         Top             =   1500
         Width           =   2895
         Begin VB.CheckBox chkSeijou 
            Caption         =   "正常"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Value           =   1  'ﾁｪｯｸ
            Width           =   855
         End
         Begin VB.CheckBox chkIjou 
            Caption         =   "異常"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   20
            Top             =   240
            Value           =   1  'ﾁｪｯｸ
            Width           =   855
         End
         Begin VB.CheckBox chkReigai 
            Caption         =   "例外"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            TabIndex        =   21
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Width           =   855
         End
         Begin VB.CheckBox chkKeikoku 
            Caption         =   "警告"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   22
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Width           =   855
         End
      End
      Begin VB.Frame frmShubetu 
         Caption         =   "項目種別指定"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6120
         TabIndex        =   82
         Top             =   1500
         Width           =   2535
         Begin VB.CheckBox chkKey 
            Caption         =   "キー項目"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Value           =   1  'ﾁｪｯｸ
            Width           =   1335
         End
         Begin VB.CheckBox chkDeb 
            Caption         =   "デバッグ項目"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Width           =   1815
         End
      End
      Begin VB.Frame frmMod 
         Caption         =   "モジュール指定"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   240
         TabIndex        =   80
         Top             =   2520
         Width           =   8415
         Begin VB.CommandButton cmdModSen 
            Caption         =   "全て選択"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdModHi 
            Caption         =   "全て非選択"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   26
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame frmModMeisai 
            Height          =   5055
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   8175
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   133
               Top             =   4680
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   132
               Top             =   4440
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   131
               Top             =   4200
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   130
               Top             =   3960
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   129
               Top             =   3720
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   128
               Top             =   3480
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   127
               Top             =   3240
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   126
               Top             =   3000
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   125
               Top             =   2760
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   124
               Top             =   2520
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   123
               Top             =   2280
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   122
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   121
               Top             =   1800
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   120
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   119
               Top             =   1320
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   118
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   117
               Top             =   840
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   116
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   115
               Top             =   360
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   6360
               TabIndex        =   114
               Top             =   120
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   113
               Top             =   4680
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   112
               Top             =   4440
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   111
               Top             =   4200
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   110
               Top             =   3960
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   109
               Top             =   3720
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   108
               Top             =   3480
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   107
               Top             =   3240
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   106
               Top             =   3000
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   105
               Top             =   2760
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   104
               Top             =   2520
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   103
               Top             =   2280
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   102
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   74
               Top             =   1800
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   73
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   72
               Top             =   1320
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   71
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   70
               Top             =   840
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   69
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   68
               Top             =   360
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   4320
               TabIndex        =   67
               Top             =   120
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   66
               Top             =   4680
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   65
               Top             =   4440
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   64
               Top             =   4200
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   63
               Top             =   3960
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   62
               Top             =   3720
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   61
               Top             =   3480
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   60
               Top             =   3240
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2280
               TabIndex        =   59
               Top             =   3000
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   58
               Top             =   2760
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   57
               Top             =   2520
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   56
               Top             =   2280
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   55
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   54
               Top             =   1800
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   53
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   52
               Top             =   1320
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   51
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   50
               Top             =   840
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   49
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   48
               Top             =   360
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   2295
               TabIndex        =   47
               Top             =   120
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   375
               TabIndex        =   46
               Top             =   4680
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   375
               TabIndex        =   45
               Top             =   4440
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   375
               TabIndex        =   44
               Top             =   4200
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   375
               TabIndex        =   43
               Top             =   3960
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   42
               Top             =   3720
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   41
               Top             =   3480
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   40
               Top             =   3240
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   39
               Top             =   3000
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   38
               Top             =   2760
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   37
               Top             =   2520
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   36
               Top             =   2280
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   35
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   34
               Top             =   1800
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   33
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   32
               Top             =   1320
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   31
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   30
               Top             =   840
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   29
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   28
               Top             =   360
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
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
               Left            =   360
               TabIndex        =   27
               Top             =   120
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
         End
      End
      Begin TabDlg.SSTab tabCorner 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   136
         Top             =   2280
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
         TabPicture(0)   =   "ログ管理(LDU)画面.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkEGXGoki(15)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkEGXGoki(14)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkEGXGoki(13)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkEGXGoki(12)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkEGXGoki(11)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkEGXGoki(10)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkEGXGoki(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkEGXGoki(8)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkEGXGoki(7)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkEGXGoki(6)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkEGXGoki(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkEGXGoki(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkEGXGoki(3)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkEGXGoki(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkEGXGoki(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "chkEGXGoki(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "  "
         TabPicture(1)   =   "ログ管理(LDU)画面.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkEGXGoki(16)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "chkEGXGoki(17)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "chkEGXGoki(18)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "chkEGXGoki(19)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "chkEGXGoki(20)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "chkEGXGoki(21)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "chkEGXGoki(22)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "chkEGXGoki(23)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "chkEGXGoki(24)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "chkEGXGoki(25)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "chkEGXGoki(26)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "chkEGXGoki(27)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "chkEGXGoki(28)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "chkEGXGoki(29)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "chkEGXGoki(30)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "chkEGXGoki(31)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "  "
         TabPicture(2)   =   "ログ管理(LDU)画面.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkEGXGoki(32)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "chkEGXGoki(33)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "chkEGXGoki(34)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "chkEGXGoki(35)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "chkEGXGoki(36)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "chkEGXGoki(37)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "chkEGXGoki(38)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "chkEGXGoki(39)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "chkEGXGoki(40)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "chkEGXGoki(41)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "chkEGXGoki(42)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "chkEGXGoki(43)"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "chkEGXGoki(44)"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "chkEGXGoki(45)"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "chkEGXGoki(46)"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "chkEGXGoki(47)"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).ControlCount=   16
         TabCaption(3)   =   "  "
         TabPicture(3)   =   "ログ管理(LDU)画面.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkEGXGoki(48)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "chkEGXGoki(49)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "chkEGXGoki(50)"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "chkEGXGoki(51)"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "chkEGXGoki(52)"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "chkEGXGoki(53)"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "chkEGXGoki(54)"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "chkEGXGoki(55)"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "chkEGXGoki(56)"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "chkEGXGoki(57)"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "chkEGXGoki(58)"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).Control(11)=   "chkEGXGoki(59)"
         Tab(3).Control(11).Enabled=   0   'False
         Tab(3).Control(12)=   "chkEGXGoki(60)"
         Tab(3).Control(12).Enabled=   0   'False
         Tab(3).Control(13)=   "chkEGXGoki(61)"
         Tab(3).Control(13).Enabled=   0   'False
         Tab(3).Control(14)=   "chkEGXGoki(62)"
         Tab(3).Control(14).Enabled=   0   'False
         Tab(3).Control(15)=   "chkEGXGoki(63)"
         Tab(3).Control(15).Enabled=   0   'False
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "  "
         TabPicture(4)   =   "ログ管理(LDU)画面.frx":00C4
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chkEGXGoki(64)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "chkEGXGoki(65)"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "chkEGXGoki(66)"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "chkEGXGoki(67)"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "chkEGXGoki(68)"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "chkEGXGoki(69)"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "chkEGXGoki(70)"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).Control(7)=   "chkEGXGoki(71)"
         Tab(4).Control(7).Enabled=   0   'False
         Tab(4).Control(8)=   "chkEGXGoki(72)"
         Tab(4).Control(8).Enabled=   0   'False
         Tab(4).Control(9)=   "chkEGXGoki(73)"
         Tab(4).Control(9).Enabled=   0   'False
         Tab(4).Control(10)=   "chkEGXGoki(74)"
         Tab(4).Control(10).Enabled=   0   'False
         Tab(4).Control(11)=   "chkEGXGoki(75)"
         Tab(4).Control(11).Enabled=   0   'False
         Tab(4).Control(12)=   "chkEGXGoki(76)"
         Tab(4).Control(12).Enabled=   0   'False
         Tab(4).Control(13)=   "chkEGXGoki(77)"
         Tab(4).Control(13).Enabled=   0   'False
         Tab(4).Control(14)=   "chkEGXGoki(78)"
         Tab(4).Control(14).Enabled=   0   'False
         Tab(4).Control(15)=   "chkEGXGoki(79)"
         Tab(4).Control(15).Enabled=   0   'False
         Tab(4).ControlCount=   16
         TabCaption(5)   =   "  "
         TabPicture(5)   =   "ログ管理(LDU)画面.frx":00E0
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkEGXGoki(80)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "chkEGXGoki(81)"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "chkEGXGoki(82)"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).Control(3)=   "chkEGXGoki(83)"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "chkEGXGoki(84)"
         Tab(5).Control(4).Enabled=   0   'False
         Tab(5).Control(5)=   "chkEGXGoki(85)"
         Tab(5).Control(5).Enabled=   0   'False
         Tab(5).Control(6)=   "chkEGXGoki(86)"
         Tab(5).Control(6).Enabled=   0   'False
         Tab(5).Control(7)=   "chkEGXGoki(87)"
         Tab(5).Control(7).Enabled=   0   'False
         Tab(5).Control(8)=   "chkEGXGoki(88)"
         Tab(5).Control(8).Enabled=   0   'False
         Tab(5).Control(9)=   "chkEGXGoki(89)"
         Tab(5).Control(9).Enabled=   0   'False
         Tab(5).Control(10)=   "chkEGXGoki(90)"
         Tab(5).Control(10).Enabled=   0   'False
         Tab(5).Control(11)=   "chkEGXGoki(91)"
         Tab(5).Control(11).Enabled=   0   'False
         Tab(5).Control(12)=   "chkEGXGoki(92)"
         Tab(5).Control(12).Enabled=   0   'False
         Tab(5).Control(13)=   "chkEGXGoki(93)"
         Tab(5).Control(13).Enabled=   0   'False
         Tab(5).Control(14)=   "chkEGXGoki(94)"
         Tab(5).Control(14).Enabled=   0   'False
         Tab(5).Control(15)=   "chkEGXGoki(95)"
         Tab(5).Control(15).Enabled=   0   'False
         Tab(5).ControlCount=   16
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   232
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   231
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   230
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   229
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   228
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   227
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   226
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   225
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   224
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   223
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   222
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   221
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   220
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   219
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   218
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   217
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   216
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   215
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   214
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   213
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   212
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   211
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   210
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   209
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   208
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   207
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   206
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   205
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   204
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   203
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   202
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   201
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   200
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   199
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   198
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   197
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   196
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   195
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   194
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   193
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   192
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   191
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   190
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   189
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   188
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   187
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   186
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   185
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   184
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   183
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   182
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   181
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   180
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   179
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   178
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   177
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   176
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   175
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   174
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   173
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   172
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   171
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   170
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   169
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   168
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   167
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   166
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   165
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   164
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   163
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   162
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   161
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   160
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   159
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   158
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   157
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   156
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   155
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   154
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   153
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   152
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   151
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   150
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   149
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   148
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   147
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   146
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   145
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   144
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   143
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   142
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   141
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   140
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   139
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   138
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   137
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Label lblFile 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   101
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblStart 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ログ開始日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73080
         TabIndex        =   100
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblEnd 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ログ終了日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70560
         TabIndex        =   99
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblSize 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "サイズ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68040
         TabIndex        =   98
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0000C000&
      Caption         =   "LDUアプリケーションログ管理"
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
      TabIndex        =   134
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmLDULogkanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmLDULogKanri.frm
'//  パッケージ名：LDユーティリティログ管理画面
'//
'//  概要：LDユーティリティログ管理画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・LDユーティリティ、ログ管理画面(frmLogKanri.frm)を流用
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応
'//     REVISIONS :(1.9.0.1) 2009-09-10   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応 メッセージボックス修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02 REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応【監視D-122】
'//                   ・ログ媒体出力結果メッセージボックス文言変更
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-07 REVISED BY [TCC] M.Matsumoto
'//                 ＥＧ２０フェーズ２対応
'//     REVISIONS :(EG20 3.6.0.1) 2012-02-23   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ２ 残件回収
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

Public iNowChk1 As Integer
Public iNowChk2 As Integer

'アプリ・保守種別定義
Public mlngLogType        As Long

'///////////////////////////////////////////////////////////////////
'対象ファイルフルパス（複数ﾌｧｲﾙの時、ｽﾍﾟｰｽ1文字で区切る。）
'///////////////////////////////////////////////////////////////////
Private sObjectFiles As String   'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽで選択中のﾌｧｲﾙのﾌﾙﾊﾟｽ文字列
Private sObjectTopFile As String '同上、選択中の先頭（最旧）ﾌｧｲﾙ名。(12文字)。

'///////////////////////////////////////////////////////////////////
'ログ情報格納エリア
'///////////////////////////////////////////////////////////////////
Private Type LogFileData
    sPath As String                 'ログファイルのパス
    sName As String                 'ログファイル名
    dtFileDate As Date              '作成日付・時刻
    dtFileDate2 As Date              '作成日付・時刻
    lFileSize As Long               'ファイルサイズ
    bSelect As Boolean              '選択フラグ
End Type

Private uLogfileData() As LogFileData

'///////////////////////////////////////////////////////////////////
'モジュール情報格納エリア
'///////////////////////////////////////////////////////////////////
Private Type ModFileData
    sName As String             'モジュール名
    sDai  As String             '大項目
    sShou As String             '小項目
    sType As String             'モジュールタイプ
    iBit  As Integer            'ビット番号
End Type

Private uModFileData(79) As ModFileData
Private iModCnt As Integer

'///////////////////////////////////////////////////////////////////
'EGX情報格納エリア
'///////////////////////////////////////////////////////////////////
Private Type EgxFileData
    iRonri As Integer               '論理号機
    iHyozi As Integer               '表示号機
    iIndex As Integer               'chkCornerのINDEX
End Type

Private uEgxFileData(15) As EgxFileData
Private iEgxCnt As Integer

'///////////////////////////////////////////////////////////////////
'イベントログコピー用ワークファイル名フルパス
'///////////////////////////////////////////////////////////////////
Private SAVEFILE_SYS As String
Private SAVEFILE_SEC As String
Private SAVEFILE_APP As String
Private SAVEFILE_LOG As String

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias _
    "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
        lpSectorsPerCluster As Long, _
        lpBytesPerSector As Long, _
        lpNumberOfFreeClusters As Long, _
        lpTtoalNumberOfClusters As Long) As Long
        
Private mintStatus(31) As Integer       'EG20 V2.1.0.1 ADD 【フェーズ２対応】

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : CmdRemove_Click
'//  機能名称  : 「媒体取外」釦押下時処理
'//  機能概要  : 媒体の取り外しを行う。
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
Private Sub cmdRemove_Click()
    On Error Resume Next
   
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : LDユーティリティログ管理画面(アクティブ時)
'//  機能概要  : 画面の最前面表示を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    'V1.3.0.1 ADD START
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
    'V1.3.0.1 ADD END
End Sub

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : LDユーティリティログ管理画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : LDユーティリティログ管理画面(ロード時)
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
'//     REVISIONS :(1.9.0.1) 2009-09-10   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応 メッセージボックス修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim intModulesFileNo As Integer
    Dim sModules As String * LDU_LOG_SIZE    '１行分ファイル内容取得用
    Dim Cnt As Integer
    Dim iMozi As Integer
    Dim iKbn As Integer
    Dim iRet As Integer
    Dim sType As String * LDU_LOG_MOZI_TYPE_SIZE         '設置タイプ
    Dim sKeyName As String
    Dim str As String
    Dim iLoop As Integer
    Dim MyName As String
    Dim iErr As Integer
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    Dim Cnt2 As Integer
    Dim strCorner1 As String
    Dim strCorner2 As String
    Dim intIndex As Integer
    'EG20 V2.1.0.1 ADD END
                   
    'パス指定
    LDU_PROFILE_NAME_EGX = PATH_LDU_APP & LDU_EGX_FILE

    SAVEFILE_SYS = PATH_LDU_APP & PATH_LDU_WORK & "SysEvent.Evt"
    SAVEFILE_SEC = PATH_LDU_APP & PATH_LDU_WORK & "SecEvent.Evt"
    SAVEFILE_APP = PATH_LDU_APP & PATH_LDU_WORK & "AppEvent.Evt"
    SAVEFILE_LOG = PATH_LDU_APP & PATH_LDU_WORK & "drwtsn32.log"
    
    gStrCurrentForm = sFormName_LDULog
    
    cmdLogHyouzi.Caption = "ログ表示" & Chr(13) & "(テキスト表示）"
    cmdCancel.Caption = "ログ管理" & Chr(13) & "画面へ戻る"
    
    'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'    cmdEGXZSentaku.Caption = "改札機全号機" & Chr(13) & "選択"
'    cmdEGXZHisentaku.Caption = "改札機全号機" & Chr(13) & "非選択"
    'EG20 V2.1.0.1 DEL END

    '配置設定
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
         
    '初期化
    tabMain.Tab = 0
    
    LstFile.Clear
    
    txtStNen.Text = ""
    txtStTuki.Text = ""
    txtStHi.Text = ""
    txtStZi.Text = ""
    txtStFun.Text = ""
    txtEdNen.Text = ""
    txtEdTuki.Text = ""
    txtEdHi.Text = ""
    txtEdZi.Text = ""
    txtEdFun.Text = ""
    
   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
   
    'INIファイルよりの設定取得
    'モジュール指定取得
    On Error GoTo FileError
    iErr = 1
    
    'ファイル有無チェック
    MyName = Dir(PATH_LDU_APP & LDU_MODULES_FILE_FULLPASS, vbNormal)
    If MyName = "" Then
        'MsgBox "DICPLOG.INIの取得に失敗しました｡", vbCritical 'V1.9.0.1 DEL
        GoTo FileError
    End If
    
    Cnt = 0
    
    For Cnt = 0 To 79
        sKeyName = "ID" & Format(Cnt, "000")
        iRet = GetPrivateProfileString(LDU_PROFILE_SECTION_NAME_ID, _
                                       sKeyName, _
                                       DEFAILT, sModules, Len(sModules), _
                                       PATH_LDU_APP & LDU_MODULES_FILE_FULLPASS)

        iMozi = 1
        iKbn = 1
        Do
            If Mid(sModules, iMozi, 1) = "," Then
                Select Case iKbn
                    Case 1
                        uModFileData(Cnt).sName = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 2
                        uModFileData(Cnt).sDai = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 3
                        uModFileData(Cnt).sShou = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 4
                        uModFileData(Cnt).sType = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 5
                        uModFileData(Cnt).iBit = Left(sModules, iMozi - 1)
                        Exit Do
                End Select
            End If
            iMozi = iMozi + 1
            If iMozi > Len(sModules) Then
                Exit Do
            End If
        Loop

        If iKbn = 5 Then
            chkMod(Cnt).Visible = True
            chkMod(Cnt).Caption = uModFileData(Cnt).sName
            If LenB(StrConv(uModFileData(Cnt).sName, vbFromUnicode)) > 14 Then
                str = uModFileData(Cnt).sName
                For iLoop = 0 To Len(uModFileData(Cnt).sName)
                    str = Left(str, Len(str) - 1)
                    If LenB(StrConv(str, vbFromUnicode)) <= 14 Then
                        chkMod(Cnt).Caption = str
                        Exit For
                    End If
                Next
            End If
            If Int(uModFileData(Cnt).sShou) = 0 Then
                chkMod(Cnt).Left = chkMod(Cnt).Left - 240
            End If
            iModCnt = Cnt
        End If
    Next
    
    iErr = 2
    
'    ファイル有無チェック
    MyName = Dir(LDU_PROFILE_NAME_EGX, vbNormal)
    If MyName = "" Then
        'MsgBox "EGX.INIの取得に失敗しました｡", vbCritical 'V1.9.0.1 DEL
        iErr = 3
        GoTo FileError
    End If
    
    'EGX情報取得
    If False = GetCpuEgxInfo() Then
        GoTo FileError
    End If
   
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    Call gsGetCornerName_LDU
    
    'タブ数を設置コーナ数とする
    tabCorner.Tab = 0
    
    '収集状態初期化
    Erase mintStatus
    
    For Cnt = 0 To UBound(gblnCornerSet)
    
        '設定ありのコーナを活性にする
        If gblnCornerSet(Cnt) = True Then
            'コーナー名称表示
            strCorner1 = MidB(gstrCornerName(Cnt), 1, 12)
            strCorner2 = MidB(gstrCornerName(Cnt), 13, 12)
            tabCorner.TabCaption(Cnt) = strCorner1 & vbCrLf & strCorner2
            
        End If
    
    Next Cnt
    
    '設置コーナ数分ループ
    For Cnt = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(Cnt) = False Then
            tabCorner.TabVisible(Cnt) = False
        End If

        '最大号機数分ループ
        For Cnt2 = 0 To 15
            intIndex = (Cnt * 16) + Cnt2
            chkEGXGoki(intIndex).Visible = False
            chkEGXGoki(intIndex).Tag = "0"
        Next
        
        For Cnt2 = 0 To 15
            intIndex = (Cnt * 16) + (gudtSettiCorner(Cnt).intGokiNo(Cnt2) - 1)
            If gudtSettiCorner(Cnt).intGokiNo(Cnt2) > 0 Then
                chkEGXGoki(intIndex).Caption = gudtSettiCorner(Cnt).strDispGoki(Cnt2) + "号機"
                'Tagに対応する号機番号を記録（1～32号機）
                chkEGXGoki(intIndex).Tag = CStr(gudtSettiCorner(Cnt).intGateNo(Cnt2))
                mintStatus(gudtSettiCorner(Cnt).intGateNo(Cnt2) - 1) = CHECKBOX_ON
                chkEGXGoki(intIndex).Visible = True
                chkEGXGoki(intIndex).Value = CHECKBOX_ON
            End If
        Next Cnt2
        
    Next Cnt
    'EG20 V2.1.0.1 ADD END
       
    On Error GoTo OtherError
       
    iNowChk1 = 1
    iNowChk2 = 2
    
    'リストの初期表示
    If sSetListBox = False Then
        '「LDユーティリティログ管理：アプリログ表示異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
        LstFile.Clear
        MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
    End If

   '「LDユーティリティログ管理：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_LOG_KANRI_GAMEN_START, 0)

  Exit Sub
     
FileError:
    Select Case iErr
    Case 1:
        'MsgBox "CASE1", vbCritical 'V1.9.0.1 DEL
        '「LDユーティリティログ管理：INIファイル異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    Case 2:
        'MsgBox "CASE2", vbCritical  'V1.9.0.1 DEL
        '「LDユーティリティログ管理：INIファイル異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    Case 3:
        'MsgBox "CASE3", vbCritical  'V1.9.0.1 DEL
        '「LDユーティリティログ管理：INIファイル異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    Case 4:
        'MsgBox "CASE4", vbCritical  'V1.9.0.1 DEL
        '「LDユーティリティログ管理：INIファイル異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    End Select
    
    MsgBox "INIファイルの取得に失敗しました｡", vbCritical, "ファイル異常"

   Exit Sub
OtherError:
   '「LDユーティリティログ管理：アプリログ表示異常」ログ出力
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
   'リストボックスの初期化
   LstFile.Clear
   MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : GetCpuEgxInfo
'//  機能名称  : EGX.INI情報取得処理
'//  機能概要  : EGX.INIファイルより情報を取得する。
'//　　　　　　　表示ファイル指定部：初期処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    :Boolean             [OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function GetCpuEgxInfo() As Boolean

    Dim sKeyName As String                 'キー名
    Dim iRet As Integer                    'INIファイル取得処理戻り値
    Dim sEgxData As String * LDU_LOG_SIZE  '１行分ファイル内容取得用
    Dim i As Integer                       'ループカウンタ
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim intCnrIdx As Integer    'EG20 V2.1.0.1 ADD 【フェーズ２対応】

    On Error Resume Next

    iEgxCnt = -1

    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    Erase gudtDisp
    Erase gudtSettiCorner
    'EG20 V2.1.0.1 ADD END
    
    'EGX情報取得
'    For i = 1 To 16                'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    For i = 1 To MAX_GATE_NO        'EG20 V2.1.0.1 ADD 【フェーズ２対応】
       sKeyName = "egx" & Format(i, "00")
        iRet = GetPrivateProfileString(LDU_PROFILE_SECTION_NAME_EGX, _
                                       sKeyName, _
                                       DEFAILT, sEgxData, Len(sEgxData), _
                                       LDU_PROFILE_NAME_EGX)
        If iRet = 0 Then
            GetCpuEgxInfo = False
            Exit Function
        End If

        'データの取得
        ReDim sFData(14)
        iFCnt = 1

        For iFLoop = 1 To Len(sEgxData)
            If Mid(sEgxData, iFLoop, 1) <> " " Or Mid(sEgxData, iFLoop, 1) <> "," Then
                iFLoop2 = iFLoop
                Do
                    iFLoop2 = iFLoop2 + 1
                    If iFLoop2 > Len(sEgxData) Then
                        sFData(iFCnt) = Mid(sEgxData, iFLoop, iFLoop2 - iFLoop)
                        iFCnt = iFCnt + 1
                        If iFCnt >= 15 Then
                            Exit For
                        End If
                        iFLoop = iFLoop2
                        Exit Do
                    End If

                    If Mid(sEgxData, iFLoop2, 1) = " " Or Mid(sEgxData, iFLoop2, 1) = "," Then
                        sFData(iFCnt) = Mid(sEgxData, iFLoop, iFLoop2 - iFLoop)
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

        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
        '通路種別が未設置の時は処理せず
'        If Trim(sFData(5)) <> MISETI Then
'            iEgxCnt = iEgxCnt + 1
'            uEgxFileData(iEgxCnt).iIndex = i
'            uEgxFileData(iEgxCnt).iRonri = i
'            uEgxFileData(iEgxCnt).iHyozi = Trim(sFData(1))
'            chkEGXGoki(uEgxFileData(iEgxCnt).iIndex).Visible = True
'            chkEGXGoki(uEgxFileData(iEgxCnt).iIndex).Caption = uEgxFileData(iEgxCnt).iHyozi & "号機"
'        End If
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        '号機が未実装の時
        If sFData(5) = MISETI Then
            gudtDisp(i - 1).intJiso = JissouUmu.Mijissou                        '未実装
               
        '号機が実装されている時
        Else
            gudtDisp(i - 1).intJiso = JissouUmu.jissou                          '実装
            
            '表示号機
            gudtDisp(i - 1).strJikaiDipNo = sFData(1)
            '論理コーナー番号
            gudtDisp(i - 1).intRonriCorner = CInt(sFData(3))
            '論理コーナー号機番号
            gudtDisp(i - 1).intRonriCornerGoki = sFData(4)
            
            'コーナ別号機数をカウント
            intCnrIdx = gudtDisp(i - 1).intRonriCorner - 1
            gudtSettiCorner(intCnrIdx).intGokiNum = gudtSettiCorner(intCnrIdx).intGokiNum + 1
            '論理号機番号、表示号機番号を格納
            gudtSettiCorner(intCnrIdx).intGateNo(gudtSettiCorner(intCnrIdx).intGokiNum - 1) = i - 1 + 1
            gudtSettiCorner(intCnrIdx).intGokiNo(gudtSettiCorner(intCnrIdx).intGokiNum - 1) = gudtDisp(i - 1).intRonriCornerGoki
            gudtSettiCorner(intCnrIdx).strDispGoki(gudtSettiCorner(intCnrIdx).intGokiNum - 1) = gudtDisp(i - 1).strJikaiDipNo
        End If
        'EG20 V2.1.0.1 ADD END
  Next

  GetCpuEgxInfo = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdLogHyouzi_Click
'//  機能名称  : 「ログ表示(テキスト表示)」釦押下時処理
'//  機能概要  : 選択ファイルを、テキストにて出力表示する。
'//　　　　　　　表示ファイル指定部：「ログ表示(テキスト表示)」
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
Private Sub cmdLogHyouzi_Click()
    Dim bRet As Boolean
    Dim lRetVal As Double
    Dim sCommand As String
    Dim sWriteDir As String
    Dim iObjFileNo As Integer
    Dim sFileName As String
    Dim lngErrCode As Long
    
   '「LDユーティリティログ管理：ログ表示釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_TEXT_HYOUJI_BUTTOM, 0)

    'ログ検索データ正当姓チェック
    bRet = fLogSearchCheck
    If bRet = False Then
    'ログ検索データにエラーがある場合、処理終了
        Exit Sub
    End If

    'ログテキストファイルを書き込む
    bRet = fWriteLogtxt
    If bRet = True Then                                 'ログテキストファイルが正常に作成された場合
      '「LDユーティリティログ管理：ログテキストファイル作成正常」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CREATE_TEXT_HYOUJI, 0)
      'ファイルコピー
      sFileName = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
      sFileName = PATH_LDU_APP & PATH_LDU_WORK & "\\" & Left(sFileName, Len(sFileName) - 4) & ".txt"
      
      'ファイルオープン
      On Error GoTo FileError
      sCommand = MN_EXE_MEMO & sFileName              '実行コマンドを作成する
      lRetVal = Shell(sCommand, vbMaximizedFocus)     'ノートパッドを起動する
      AppActivate lRetVal, True                       'アクティブ（前面表示）にする
      SendKeys "{LEFT}", True
      On Error GoTo 0
       '「LDユーティリティログ管理：ログテキスト表示正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_OK, 0)
    Else
       '「LDユーティリティログ管理：出力データ作成異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_CREATE_TEXT_ERROR, lngErrCode)
       MsgBox "媒体出力するデータの作成に失敗しました。", vbCritical, "データ出力失敗"
    End If
    Exit Sub

FileError:
   '「LDユーティリティログ管理：ログテキスト表示処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdLog_Click
'//  機能名称  : 「ログ媒体出力」釦押下時処理
'//  機能概要  : 選択ファイルを、指定フォルダへ出力する。
'//　　　　　　　表示ファイル指定部：「ログ媒体出力」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02 REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応【監視D-122】
'//                   ・ログ媒体出力結果メッセージボックス文言変更
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20V5.13.0.1) 2012-06-06 REVISED BY [TCC] H.Sugimoto
'//                 【媒体出力フォルダ作成対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdLog_Click()
    Dim sWriteDir
    Dim sFileName As String
    Dim dFileSize As Double
    Dim MyPath As String
    Dim MyName As String
    Dim iRet As Integer
    Dim Sekuta As Long      'セクタ（クラスタ当り）
    Dim nByte As Long       'バイト数（セクタ当り）
    Dim Kurasuta As Long    'フリークラスタ数
    Dim Drive As Long       'ドライブのクラスタ数（合計）
    Dim FreeSpace As Double 'ディスクの空き容量
    Dim lngErrCode As Long  'エラーコード
    Dim objFso         As New FileSystemObject 'ファイルシステムオブジェクト 'V1.6.0.1 ADD
    Dim iFileCounter As Integer  '対象ﾌｧｲﾙ数カウンタ    ' EG20 V5.9.0.1【ログ選択上限対応】ADD

    Dim fso As FileSystemObject     'ファイルシステムオブジェクト       ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加
    Dim szDefLogFolder As String    ' 出力ログフォルダ                  ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加

    '「LDユーティリティログ管理画面：ログ媒体出力釦押下」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_OUTPUT_BUTTOM, 0)

    Dim bFrmShow As Boolean
    bFrmShow = False

    txtDummy.SetFocus

    On Error GoTo EVENTLOG_ERROR
    If iNowChk1 = 1 Then
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_APP
    Else
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_HOSHU
    End If

    'ファイル有無チェック
    Dim i
    Dim Chk
    Chk = False
    dFileSize = 0
    iFileCounter = 0                                                                            ' EG20 V5.9.0.1【ログ選択上限対応】ADD
    For i = 0 To LstFile.ListCount - 1
        If LstFile.Selected(i) Then
            Chk = True
            MyName = Trim(Left(LstFile.List(i), 12))
            MyName = Dir(MyPath & MyName, vbNormal)
            If MyName = "" Then ' ループを開始します。
                MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
                Exit Sub
            End If
            dFileSize = dFileSize + FileLen(MyPath & MyName)
            iFileCounter = iFileCounter + 1                                                     ' EG20 V5.9.0.1【ログ選択上限対応】ADD
        End If
    Next

    If Chk = False Then
        '表示ファイルが選択されていなければ、エラーメッセージを表示する
        MsgBox "項目指定に異常があります。指定した表示項目を確認してください。", _
               vbCritical, _
               "項目指定異常"
        Exit Sub
    End If

' EG20 V5.9.0.1【ログ選択上限対応】ADD START
    If iFileCounter > LOG_FILECNT_MAX Then
        ' 警告文言表示
        MsgBox "選択されたファイル数が上限を超えました。" _
               & Chr(vbKeyReturn) & "選択できるファイル数は[" & LOG_FILECNT_MAX & "]件までです。", _
               vbOKOnly + vbCritical, _
               "ファイル指定異常"
        Exit Sub
    End If
' EG20 V5.9.0.1【ログ選択上限対応】ADD END

DirSelect:
    'フォルダ指定ダイアログの表示
'    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")                         'V1.12.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.12.0.1 ADD

    If Len(sWriteDir) = 0 Then
        Exit Sub
    End If

    If UCase(Left(sWriteDir, 1)) = "A" Then
        iRet = MsgBox("ＦＤを挿入してください。", vbQuestion + vbOKCancel, "媒体準備確認")
        If iRet = vbOK Then
            frmLDULogkanri.Refresh

             'ディスク情報を取得
            iRet = GetDiskFreeSpace("a:\", Sekuta, nByte, Kurasuta, Drive)

            If Drive = 0 Then
                iRet = MsgBox(" FDが挿入されていません。", _
                    vbCritical, _
                    "指定媒体出力異常")
                GoTo DirSelect
            End If

            '空き容量を取得
            FreeSpace = Sekuta * nByte * Kurasuta
            If dFileSize > FreeSpace Then
               iRet = MsgBox("出力ファイルのサイズが指定媒体より大きいため出力できません。", _
                            vbCritical, _
                            "指定媒体出力異常")

                GoTo DirSelect
            End If
        Else
          Exit Sub
        End If

    End If
    
' EG20V5.9.0.1【試験場指摘事項No.10修正対応】削除開始
'    '処理番号格納（処理中）
'    glShoriNo = SHORI_NO.NO_MEDIA_OUT
'
'    Load frmSyorityu
'    frmSyorityu.lblLogMessage.Caption = "媒体出力中"
'    frmSyorityu.Caption = "媒体出力中"
'    frmSyorityu.Show vbModal
'    frmSyorityu.Refresh
' EG20V5.9.0.1【試験場指摘事項No.10修正対応】削除終了

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加開始
    szDefLogFolder = fncCreateLogFolder()
    If sWriteDir Like ("*" & szDefLogFolder & "\") = False Then
        ' フォルダが存在するかチェックする
        sWriteDir = sWriteDir & "\" & szDefLogFolder
        Set fso = New FileSystemObject
        If fso.FolderExists(sWriteDir) = False Then
            ' フォルダが存在しない場合は作成する
            fso.CreateFolder (sWriteDir)
        End If
        Set fso = Nothing
    End If
' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加終了

'V1.6.0.1 ADD START
    'コピー先フォルダパス作成(指定フォルダ￥LDULOG)
    sWriteDir = sWriteDir & "\" & LDU_LOGKANRI_LDULOG
    
    'ファイルシステムオブジェクト生成
    Set objFso = CreateObject("Scripting.FileSystemObject")

    'コピー先フォルダの有無確認
    If objFso.FolderExists(sWriteDir) = False Then
    
        'コピー先フォルダ作成
        objFso.CreateFolder (sWriteDir)
    
    End If
    
    'ファイルシステムオブジェクト解放
    Set objFso = Nothing
'V1.6.0.1 ADD END
    For i = 0 To LstFile.ListCount - 1
        If LstFile.Selected(i) Then
            MyName = Trim(Left(LstFile.List(i), 12))
            FileCopy MyPath & MyName, sWriteDir & "\" & MyName
        End If
    Next

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    If UCase(Left(sWriteDir, 1)) = "A" Then
        MsgBox "ＦＤ出力は正常終了しました。", vbInformation + vbOKOnly, "出力結果"
    Else
'EG20 V2.0.1.1【監視D-122】DEL START
'        MsgBox "ＨＤＤ内一時フォルダへの出力は正常終了しました。", vbInformation + vbOKOnly, "出力結果"
'EG20 V2.0.1.1【監視D-122】DEL END
'EG20 V2.0.1.1【監視D-122】ADD START
        MsgBox "正常終了しました。", vbInformation + vbOKOnly, "出力結果"
'EG20 V2.0.1.1【監視D-122】ADD END
    End If
    
    '「LDユーティリティログ管理ログ管理画面：ログ媒体出力処理正常」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_OUTPUT_OK, 0)

    Exit Sub

EVENTLOG_ERROR:
    'V1.6.0.1 ADD START
     'ファイルシステムオブジェクト解放
     Set objFso = Nothing
     '「LDユーティリティログ管理画面：フォルダ作成異常」
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_CREATE_LOGFOLDER_ERROR, 0)
    'V1.6.0.1 ADD END
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    If UCase(Left(sWriteDir, 1)) = "A" Then
        MsgBox "ＦＤ出力は異常終了しました。", vbCritical, "出力結果"
    Else
'EG20 V2.0.1.1【監視D-122】DEL START
'        MsgBox "ＨＤＤ内一時フォルダへの出力は異常終了しました。", vbCritical, "出力結果"
'EG20 V2.0.1.1【監視D-122】DEL END
'EG20 V2.0.1.1【監視D-122】ADD START
        MsgBox "異常終了しました。", vbCritical, "出力結果"
'EG20 V2.0.1.1【監視D-122】ADD END
    End If
    
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「LDユーティリティログ管理画面：ログ媒体出力処理異常」
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_OUTPUT_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdRefresh_Click
'//  機能名称  : 「ログ切替」釦押下時処理
'//  機能概要  : ログを最新の状態に更新する。
'//　　　　　　　表示ファイル指定部：「ログ切替」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdRefresh_Click()
    Dim bRet As Boolean                     'メール送信処理の戻り値
    Dim udtMail As IDU_LDU_LGCHGREQ_CMD     '画面表示要求
    Dim lngErrCode As Long                  'エラーコード
    Dim bFlag As Boolean                    'メール受信処理の戻り値
    Dim lId As Long                         'メールID
   
    On Error Resume Next
    
    LstFile.Clear

    '「LDユーティリティログ管理画面：ログ切替釦押下」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_CHANGE_BUTTOM, 0)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    'ログ切替要求メールをID制に送信する。
    udtMail.udtlHeader.dwId = ML_ID_IDU_LDU_LGCHGREQ_CMD
    udtMail.udtlHeader.dwSize = MlSize.IDU_LDU_LGCHGREQ_CMD
    udtMail.udtlHeader.dwProid = RHOSHU_ID
    udtMail.udtlHeader.dwSubArea = 0
    If iNowChk1 = 0 Then
        udtMail.dwLgch_Type = ML_DT_APL_LOG           ' アプリログ
    ElseIf iNowChk1 = 1 Then
        udtMail.dwLgch_Type = ML_DT_APL_LOG           ' アプリログ
    Else
        udtMail.dwLgch_Type = ML_DT_HOSHU_LOG         ' 保守ログ
    End If
    bRet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.udtlHeader)
    If bRet = False Then
       '「ログ切替要求CMD送信異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, LOG_CHANGE_CMD_SEND, lngErrCode)
       
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
       'プログレスバーを消去する
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       
       Exit Sub
    Else
       '「ログ切替要求CMD送信異常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, LOG_CHANGE_CMD_SEND, 0)
    End If
  
    'ログ切替要求RES受信
    bFlag = False
    Do Until bFlag = True
        'メール受信処理を行う
        lId = fMailRecieve()
        Select Case lId         'メールＩＤ
        '「プロセス終了指示」の場合
        Case ML_ID_PROEND_ORD
             '「プロセス終了指示受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
            '処理を終了する
            Exit Sub
        '「ログ切替要求RES」の場合
        Case ML_ID_IDU_LDU_LGCHGREQ_RES
            '「ログ切替要求RES受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, LOG_CHANGE_RES_RECV, 0)
            'ループを抜ける
            Exit Do
        Case Else
        End Select
        Sleep (MN_MAIL_INTERVAL)
    Loop
    If sSetListBox = False Then
        'リストボックスの初期化
        LstFile.Clear
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    Else
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    End If
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
   '「LDユーティリティログ管理画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_LOG_KANRI_GAMEN_END, 0)
    frmLogMenu.ZOrder
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optApp_Click
'//  機能名称  : ラジオ釦：アプリケーションログ選択時処理
'//  機能概要  : 表示を更新する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(3.5.0.1) 2012-02-17   REVISED BY [TCC] T.Furuya
'//                 EG20 フェイズ2 サイクル3 残件改修
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub optApp_Click()

    On Error GoTo Err_mgs

   '「LDユーティリティログ管理画面：アプリログ」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_APLLOG, 0)
    
    '選択されていたのが、保守プログラムログだった場合
    If iNowChk1 <> 1 Then
        
        '選択されているチェックを保持する
        iNowChk1 = 1
                              
        '非表示になっていた項目を表示させる
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'        frmMod.Visible = True
'        frmEGXGokiSentaku.Visible = True
'        cmdEGXZSentaku.Visible = True
'        cmdEGXZHisentaku.Visible = True
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        frmMod.Visible = True                           'V3.5.0.1 ADD
        tabCorner.Visible = True
        cmdZSentaku.Visible = True
        cmdZHisentaku.Visible = True
        cmdHSentaku.Visible = True
        cmdHHisentaku.Visible = True
        'EG20 V2.1.0.1 ADD END
        
        '表示を再読み込みする
        If sSetListBox = False Then
            '「LDユーティリティログ管理画面：アプリログ表示異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
            'リストボックスの初期化
            LstFile.Clear
            MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
        End If
    End If
    
   '「LDユーティリティログ画面：アプリログ表示正常」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_APLLOG_OK, 0)
 
   Exit Sub
    
Err_mgs:
    '「LDユーティリティログ管理画面：アプリログ表示異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
    
    'リストボックスの初期化
    LstFile.Clear
    MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optHoshu_Click
'//  機能名称  : ラジオ釦：保守プログラムログ選択時処理
'//  機能概要  : 表示を更新する。
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
Private Sub optHoshu_Click()
    
    On Error GoTo Err_mgs
    
   '「LDユーティリティログ管理画面：保守ログ」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_HOSHULOG, 0)
    
    '選択されていたのが、アプリケーションログだった場合
    If iNowChk1 <> 2 Then
            
        '選択されているチェックを保持する
        iNowChk1 = 2
        
        '表示になっている項目を非表示にする
        frmMod.Visible = False
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'        cmdEGXZSentaku.Visible = False
'        cmdEGXZHisentaku.Visible = False
'        frmEGXGokiSentaku.Visible = False
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        tabCorner.Visible = False
        cmdZSentaku.Visible = False
        cmdZHisentaku.Visible = False
        cmdHSentaku.Visible = False
        cmdHHisentaku.Visible = False
        'EG20 V2.1.0.1 ADD END
        
        '表示を再読み込みする
        If sSetListBox = False Then
            '「LDユーティリティログ管理画面：保守ログ表示異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_HODHULOG_ERROR, 0)
            'リストボックスの初期化
            LstFile.Clear
            MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
        End If
    End If
    
    '「LDユーティリティログ管理画面：保守ログ表示正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_HODHULOG_OK, 0)
    
    Exit Sub

Err_mgs:
    '「LDユーティリティログ管理画面：保守ログ表示異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_HODHULOG_ERROR, 0)
    'リストボックスの初期化
    LstFile.Clear
    MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optHaniari_Click
'//  機能名称  : ラジオ釦：表示範囲指定有選択時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub optHaniari_Click()
   
    '選択されていたのが、表示範囲指定無だった場合
    If iNowChk2 = 2 Then
    
        '開始と終了を入力可能にする
        lblSt.Enabled = True
        lblStNen.Enabled = True
        lblStTuki.Enabled = True
        lblStHi.Enabled = True
        lblStZi.Enabled = True
        lblStFun.Enabled = True
        
        lblEd.Enabled = True
        lblEdNen.Enabled = True
        lblEdTuki.Enabled = True
        lblEdHi.Enabled = True
        lblEdZi.Enabled = True
        lblEdFun.Enabled = True
        
        txtStNen.Enabled = True
        txtStTuki.Enabled = True
        txtStHi.Enabled = True
        txtStZi.Enabled = True
        txtStFun.Enabled = True
        
        txtEdNen.Enabled = True
        txtEdTuki.Enabled = True
        txtEdHi.Enabled = True
        txtEdZi.Enabled = True
        txtEdFun.Enabled = True
        
        '選択されているチェックを保持する
        iNowChk2 = 1
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optHaninasi_Click
'//  機能名称  : ラジオ釦：表示範囲指定無選択時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub optHaninasi_Click()
    
    '選択されていたのが、表示範囲指定有だった場合
    If iNowChk2 = 1 Then
        
        '開始と終了を入力可能にする
        lblSt.Enabled = False
        lblStNen.Enabled = False
        lblStTuki.Enabled = False
        lblStHi.Enabled = False
        lblStZi.Enabled = False
        lblStFun.Enabled = False
        
        lblEd.Enabled = False
        lblEdNen.Enabled = False
        lblEdTuki.Enabled = False
        lblEdHi.Enabled = False
        lblEdZi.Enabled = False
        lblEdFun.Enabled = False
        
        txtStNen.Enabled = False
        txtStTuki.Enabled = False
        txtStHi.Enabled = False
        txtStZi.Enabled = False
        txtStFun.Enabled = False
        
        txtEdNen.Enabled = False
        txtEdTuki.Enabled = False
        txtEdHi.Enabled = False
        txtEdZi.Enabled = False
        txtEdFun.Enabled = False
        
        '選択されているチェックを保持する
        iNowChk2 = 2
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdModSen_Click
'//  機能名称  : 「全て選択」釦押下時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示項目指定部：「モジュール指定」
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
Private Sub cmdModSen_Click()
    Dim iCnt As Integer
        
    For iCnt = CNT_MIN To iModCnt
        chkMod(iCnt).Value = 1
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdModHi_Click
'//  機能名称  : 「全て非選択」釦押下時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示項目指定部：「モジュール指定」
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
Private Sub cmdModHi_Click()
    Dim iCnt As Integer
    
    For iCnt = CNT_MIN To iModCnt
        chkMod(iCnt).Value = 0
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : chkMod_Click
'//  機能名称  : 各チェックボックス押下時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示項目指定部：「モジュール指定」
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
Private Sub chkMod_Click(Index As Integer)
    Dim iCnt As Integer
    Dim sDai As String
    Dim iChkType As Integer
        
   '小項目値チェックを行う。
    If Int(uModFileData(Index).sShou) = 0 Then
        '小項目値が、チェックボックス最大値の場合は処理終了
        If Index = iModCnt Then
            Exit Sub
        End If
        
       '初期値設定
        '対象分類に連なるもののインデックスを取得する。
        iCnt = Index + 1
        '大項目番号を取得する。
        sDai = uModFileData(Index).sDai
        '大項目のチェックボックス値を取得する。
        iChkType = chkMod(Index).Value
        Do
           '対象分類の大項目番号と、押下分類の大項目番号一致するかチェックする。
            If sDai = uModFileData(iCnt).sDai Then
               '一致した場合、押下分類のチェックボックス値を、反映する。
                chkMod(iCnt).Value = iChkType
            Else
                Exit Do
            End If
            '次の分類に進む。
            iCnt = iCnt + 1
            If iCnt > iModCnt Then
              'チェックボックスの最大値になった場合は処理終了
                Exit Sub
            End If
        Loop
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtStNen_LostFocus
'//  機能名称  : 開始年入力時処理
'//  機能概要  : 入力開始年正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub txtStNen_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Year", txtStNen.Text)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtStTuki_LostFocus
'//  機能名称  : 開始月入力時処理
'//  機能概要  : 入力開始月正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub txtStTuki_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Month", txtStTuki.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtStHi_LostFocus
'//  機能名称  : 開始日入力時処理
'//  機能概要  : 入力開始日正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub txtStHi_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Day", txtStHi.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtStZi_LostFocus
'//  機能名称  : 開始時入力時処理
'//  機能概要  : 入力開始時正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtStZi_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Hour", txtStZi.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtStNen.Text)) <> 0 And _
       Len(Trim(txtStTuki.Text)) <> 0 And _
       Len(Trim(txtStHi.Text)) <> 0 And _
       Len(Trim(txtStZi.Text)) = 0 Then
    
        iRet = MsgBox("表示範囲の開始に未入力の項目があります。", vbExclamation, "入力異常")
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtStFun_LostFocus
'//  機能名称  : 開始分入力時処理
'//  機能概要  : 入力開始分正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtStFun_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Minutes", txtStFun.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtStNen.Text)) <> 0 And _
       Len(Trim(txtStTuki.Text)) <> 0 And _
       Len(Trim(txtStHi.Text)) <> 0 And _
       Len(Trim(txtStZi.Text)) <> 0 And _
       Len(Trim(txtStFun.Text)) = 0 Then
       
        iRet = MsgBox("表示範囲の開始に未入力の項目があります。", vbExclamation, "入力異常")
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtEdNen_LostFocus
'//  機能名称  : 終了年入力時処理
'//  機能概要  : 入力終了年正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub txtEdNen_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Year", txtEdNen.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtEdTuki_LostFocus
'//  機能名称  : 終了月入力時処理
'//  機能概要  : 入力終了月正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub txtEdTuki_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Month", txtEdTuki.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtEdHi_LostFocus
'//  機能名称  : 終了日入力時処理
'//  機能概要  : 入力終了日正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
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
Private Sub txtEdHi_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Day", txtEdHi.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtedZi_LostFocus
'//  機能名称  : 終了時入力時処理
'//  機能概要  : 入力終了時正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtedZi_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Hour", txtEdZi.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtEdNen.Text)) <> 0 And _
       Len(Trim(txtEdTuki.Text)) <> 0 And _
       Len(Trim(txtEdHi.Text)) <> 0 And _
       Len(Trim(txtEdZi.Text)) = 0 Then
    
        iRet = MsgBox("表示範囲の終了に未入力の項目があります。", vbExclamation, "入力異常")
        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtedFun_LostFocus
'//  機能名称  : 終了分入力時処理
'//  機能概要  : 入力終了分正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtedFun_LostFocus()
    Dim iRet
    '整合性チェック
    iRet = TextTime_Check("Minutes", txtEdFun.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtEdNen.Text)) <> 0 And _
       Len(Trim(txtEdTuki.Text)) <> 0 And _
       Len(Trim(txtEdHi.Text)) <> 0 And _
       Len(Trim(txtEdZi.Text)) <> 0 And _
       Len(Trim(txtEdFun.Text)) = 0 Then
    
        iRet = MsgBox("表示範囲の終了に未入力の項目があります。", vbExclamation, "入力異常")
        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : TextTime_Check
'//  機能名称  : 開始/終了年月日時分入力正当性チェック時処理
'//  機能概要  : 入力された値の正当性チェックを行う。
'//　　　　　　　表示項目指定部：「表示範囲指定」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-05 REVISED BY [TCC] T.Koyama
'//                ＥＧ２０フェーズ２対応
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-14 REVISED BY [TCC] M.Matsumoto
'//                【統-341対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function TextTime_Check(sType As String, sTxt As String)
    Dim iChk As Integer
    Dim iRet As Integer
    Dim sChk As String
    
    Dim k As Integer                                'EG20 V2.0.1.1 ADD
    
    '戻り値に異常をセット
    TextTime_Check = False
        
    If Trim(sTxt) <> "" Then
    
        'EG20 V2.0.1.1 ADD START 【統-341対応】
        '入力された中に数値以外の文字が存在する場合は、エラー
        For k = 1 To Len(sTxt)
            If Not Mid(sTxt, k, 1) Like "[0-9]" Then
                iRet = MsgBox("入力された文字は数字ではありません。", vbExclamation, "入力異常")
                Exit Function
            End If
        Next k
        'EG20 V2.0.1.1 ADD END
            
        iChk = Val(sTxt)
        'EG20 V2.0.1.1 DEL START 【統-341対応】
'        If iChk = 0 And sType <> "Hour" And sType <> "Minutes" Then
'            iRet = MsgBox("入力された文字は数字ではありません。", vbExclamation, "入力異常")
'            Exit Function
'        Else
        'EG20 V2.0.1.1 DEL END
            '頭が０の時のチェック（年以外）
            If sType <> "Year" Then
                sChk = Left(sTxt, 1)
                If Len(sTxt) = 2 And sChk = "0" Then
                    sTxt = Right(sTxt, 1)
                End If
            End If
            
            'EG20 V2.0.1.1 DEL START 【統-341対応】
            '共通
'            If Len(Trim(str(iChk))) <> Len(sTxt) Then
'                iRet = MsgBox("入力された文字に数字以外のものが含まれています。", vbExclamation, "入力異常")
'                Exit Function
'            End If
            'EG20 V2.0.1.1 DEL END
            
            '範囲チェック
            Select Case sType
                Case "Year"
                    '年
'                    If iChk < 1980 Or iChk > 2079 Then         'EG20 V2.0.1.1 DEL
                    If iChk < 2000 Or iChk > 2037 Then          'EG20 V2.0.1.1 ADD
                        iRet = MsgBox("年指定の範囲を超えています。", vbExclamation, "入力異常")
                        Exit Function
                    End If
                Case "Month"
                    '月
                    If iChk < 1 Or iChk > 12 Then
                        iRet = MsgBox("月指定の範囲を超えています。", vbExclamation, "入力異常")
                        Exit Function
                    End If
                Case "Day"
                    '日
                    If iChk < 1 Or iChk > 31 Then
                        iRet = MsgBox("日指定の範囲を超えています。", vbExclamation, "入力異常")
                        Exit Function
                    End If
                Case "Hour"
                    '時
                    If iChk < 0 Or iChk > 23 Then
                        iRet = MsgBox("時間指定の範囲を超えています。", vbExclamation, "入力異常")
                        Exit Function
                    End If
                Case "Minutes"
                    '分
                    If iChk < 0 Or iChk > 59 Then
                        iRet = MsgBox("時間指定の範囲を超えています。", vbExclamation, "入力異常")
                        Exit Function
                    End If
            End Select
        End If
'    End If         'EG20 V2.0.1.1 DEL 【統-341対応】
    
    '戻り値に正常を返す
    TextTime_Check = True
    Exit Function
End Function

'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : cmdEGXZSentaku_Click
''//  機能名称  : 「改札機全号機選択」釦押下時処理
''//  機能概要  : 表示を更新する。
''//　　　　　　　表示号機指定部：
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Sub cmdEGXZSentaku_Click()
'    Dim iCnt As Integer
'
'    For iCnt = 1 To 18
'        chkEGXGoki(iCnt).Value = 1
'    Next
'End Sub
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : cmdEGXZHisentaku_Click
''//  機能名称  : 「改札機全号機非選択」釦押下時処理
''//  機能概要  : 表示を更新する。
''//　　　　　　　表示号機指定部：
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Sub cmdEGXZHisentaku_Click()
'    Dim iCnt As Integer
'
'    For iCnt = 1 To 18
'        chkEGXGoki(iCnt).Value = 0
'    Next
'End Sub
'EG20 V2.1.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSetListBox
'//  機能名称  : ログファイル登録処理
'//  機能概要  : ログファイルをリストボックスに登録する。
'//　　　　　　　表示ファイル指定部：初期処理
'//　　　　　　　　　　　　　　　　　「ログ切替」釦押下時処理
'//                                  ラジオ釦：アプリ/保守ログ選択時
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sSetListBox()
    Dim i As Integer
    Dim iCnt As Integer
    Dim strSQL As String
    Dim datWork As Date
    Dim sEntry As String            '編集文字列
    Dim bRet As Integer
    
    Dim strLogFilePath As String    'アプリor保守ログファイルパス
    Dim strLogListPath As String    '表示用ログリスト作成用ファイルパス
    Dim strLogFileName As String    'アプリor保守ログファイル名
    Dim strConnectString As String  '表示用ログリストCSVファイル接続用
    Dim intLogFileNo As Integer     'ログリストファイル番号指定用
    Dim strLogData() As String      'CSVから読み込んだログデータ格納用配列
    Dim strLineCount As String      'CSVファイル行数読込
    Dim j As Integer                'CSVファイル行数カウント用
    Dim lErrSts As Long
    Dim lngErrCode As Long          'エラーコード
    
    On Error GoTo Err_mgs
    
    sSetListBox = False
    
    On Error Resume Next            ' エラーのトラップを留保します。
    Err.Clear

    
    'ログ一覧のリストを作成する
    'アプリ、保守チェック
    Select Case iNowChk1
        Case 1 'アプリだった時の処理
            strLogFilePath = PATH_LDU_LOG & PATH_LDU_LOG_APP '引数：アプリログのファイルパス
            strLogListPath = PATH_LDU_APP & PATH_LDU_WORK & LDU_APL_LOG_LIST 'アプリログ一覧CSVのファイルパス
        Case 2 '保守ログだった時の処理
            strLogFilePath = PATH_LDU_LOG & PATH_LDU_LOG_HOSHU '引数：保守ログのファイルパス
            strLogListPath = PATH_LDU_APP & PATH_LDU_WORK & LDU_HOSHU_LOG_LIST '保守ログ一覧CSVのファイルパス
        Case Else
            Exit Function
    End Select

    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '// 保守専用関数:表示用ログ一覧リスト作成
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateLDULogInfo(lErrSts, strLogFilePath, strLogListPath)

    'ログリスト(CSV)作成成功
    If bRet Then
       '「LDユーティリティログ管理画面：ログリストファイル作成正常」
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LDU_LOG_KANRI_CREATE_LOGLISTFILE_OK, 0)
    'ログリスト(CSV)作成失敗
    Else
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       '「LDユーティリティログ管理画面：ログリストファイル作成異常」
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LDU_LOG_KANRI_CREATE_LOGLISTFILE_ERROR, 0)
        Exit Function
    End If

    'CSVファイルの有無確認
    If Len(Trim(Dir(strLogListPath))) = 0 Then
        Exit Function
    End If

    'CSVファイル番号を取得する。
    intLogFileNo = FreeFile

    'CSVファイルオープン
    Open strLogListPath For Input As #intLogFileNo


    'CSVファイル行数カウント（ファイル終端までループを繰り返す）
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1削除
        Do While Not EOF(intLogFileNo)                      ' EG20 V3.3.0.1追加
            Line Input #intLogFileNo, strLineCount
            j = j + 1
        Loop

    'CSVファイルクローズ
    Close #intLogFileNo

    'CSVファイル読込用配列（行数分）
    ReDim strLogData(j)

    'CSVファイル番号を取得する。
    intLogFileNo = FreeFile

    'CSVファイルオープン
    Open strLogListPath For Input As #intLogFileNo


    'リスト表示分読み込み（ファイル終端までループを繰り返す）
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1削除
        Do While Not EOF(intLogFileNo)                      ' EG20 V3.3.0.1追加
            Line Input #intLogFileNo, strLogData(i)
            i = i + 1
        Loop

    'CSVファイルクローズ
    Close #intLogFileNo

    iCnt = i


    'エラー処理
    On Error GoTo Err_mgs

    '「ログファイル」リストボックスをクリアする
    LstFile.Clear

    'ログファイル情報を編集する
    For i = 0 To iCnt - 1
        sEntry = Left(strLogData(i), 12)  'ファイル名
            sEntry = sEntry & "  " & Mid(strLogData(i), 14, 19) 'ログ開始年月日を表示

            sEntry = sEntry & "  " & Mid(strLogData(i), 34, 19) 'ログ終了年月日を表示

            sEntry = sEntry & "  " & Format(Mid(strLogData(i), 54), "@@@@@@@@") 'ログサイズを表示
        LstFile.AddItem sEntry
    Next

    If iCnt > 0 Then                'ログファイルが存在する
        LstFile.ListIndex = 0        '一行目にインデックスをセット
    End If

    sSetListBox = True

    Exit Function

Err_mgs:
    'アプリ、保守チェック
    Select Case iNowChk1
        Case 1
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            '「LDユーティリティログ管理画面：ファイルアクセス異常」
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
        Case 2
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            '「LDユーティリティログ管理画面：ファイルアクセス異常」
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fLogSearchCheck
'//  機能名称  : ログ検索データチェック処理
'//  機能概要  : ログ検索データの正当性チェックを行う。
'//　　　　　　　表示ファイル指定部：「ログ表示(テキスト表示)」釦押下時処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 3.6.0.1) 2012-02-23   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ２ 残件回収
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fLogSearchCheck() As Boolean
    Dim i As Integer            'カウンタ
    Dim j As Integer            'コントロール配列数
    Dim bFlag As Boolean        'フラグ
    Dim iSelectedLines As Integer 'リストボックスで選択中の行数
    Dim iChk As Integer
    Dim iChk2 As Integer
    Dim sChk As String
    Dim sFileName As String
    Dim sStAll As String
    Dim sEdAll As String
    Dim dStAll As Double
    Dim dEdAll As Double
    Dim sChkDate As String
    Dim sTxt As String
    Dim iRet

    fLogSearchCheck = False     '戻り値に初期値としてエラーをセット
   
    'ファイル選択数チェック
    iChk = 0
    For i = 0 To LstFile.ListCount - 1
        If LstFile.Selected(i) Then
            iChk = iChk + 1
        End If
    Next
    
    If iChk = 0 Then
        '表示ファイルが選択されていなければ、エラーメッセージを表示する
        MsgBox "項目指定に異常があります。指定した表示項目を確認してください。", _
               vbCritical, _
               "項目指定異常"
               Exit Function
    ElseIf iChk > 1 Then
        '複数ファイルが選択されていても、エラーメッセージを表示する
        MsgBox "複数ファイル指定、キャビネットファイル以外のファイル指定はできません。", _
               vbCritical, _
               "ファイル指定異常"
        Exit Function
    End If
    
    'ファイル名称取得
    sFileName = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
    
    '拡張子チェック
    If LCase(Right(sFileName, 3)) <> "idu" Then
        'キャビネットファイル以外が指定された場合、エラーメッセージを表示する
        MsgBox "複数ファイル指定、キャビネットファイル以外のファイル指定はできません。", _
               vbCritical, _
               "ファイル指定異常"
        Exit Function
    End If
    
    '処理結果指定
    '正常
    If chkSeijou.Value = 0 Then
        '異常
        If chkIjou.Value = 0 Then
            '例外
            If chkReigai.Value = 0 Then
                '警告
                If chkKeikoku.Value = 0 Then
                    MsgBox "項目指定に異常があります。指定した表示項目を確認してください。", vbCritical, "項目指定異常"
                    Exit Function
                End If
            End If
        End If
    End If
    
    '項目種別指定
    'キー項目
    If chkKey.Value = 0 Then
        'デバッグ項目
        If chkDeb.Value = 0 Then
            MsgBox "項目指定に異常があります。指定した表示項目を確認してください。", vbCritical, "項目指定異常"
            Exit Function
        End If
    End If
    
    
    '対象時刻の正当姓チェック
    '範囲指定がありの場合のみチェックする
    If optHaniari.Value = True Then
        '開始チェック
        'ログデータ対象時刻の正当姓チェック
        If Len(Trim(txtStNen.Text)) = 0 And _
           Len(Trim(txtStTuki.Text)) = 0 And _
           Len(Trim(txtStHi.Text)) = 0 And _
           Len(Trim(txtStZi.Text)) = 0 And _
           Len(Trim(txtStFun.Text)) = 0 Then
           
           '全て未入力なら0をセット
            sStAll = "0"
    
        ElseIf Len(Trim(txtStNen.Text)) = 0 Or _
           Len(Trim(txtStTuki.Text)) = 0 Or _
           Len(Trim(txtStHi.Text)) = 0 Then
           
           '開始時刻に未入力の項目があるなら
            MsgBox "表示範囲の開始に未入力の項目があります。", _
                   vbExclamation, _
                   "入力異常"               ' EG20 V3.6.0.1 ADD
'                   "時刻指定異常"          ' EG20 V3.6.0.1 DEL
            Exit Function
        
        Else
            '開始年チェック
            iRet = TextTime_Check("Year", txtStNen.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '開始月チェック
            iRet = TextTime_Check("Month", txtStTuki.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '開始日チェック
            iRet = TextTime_Check("Day", txtStHi.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '開始時チェック
            If Len(Trim(txtStZi.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtStZi.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtStNen.Text)) = 0 And _
                   Len(Trim(txtStTuki.Text)) = 0 And _
                   Len(Trim(txtStHi.Text)) = 0 Then
                    
                    txtStZi.Text = "00"
                Else
                    iRet = MsgBox("表示範囲の開始に未入力の項目があります。", vbExclamation, "入力異常")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
                iRet = TextTime_Check("Hour", txtStZi.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '開始分チェック
            If Len(Trim(txtStFun.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtStFun.Text = "00"
'EG20 V2.0.1.1 DEL START
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtStNen.Text)) = 0 And _
                   Len(Trim(txtStTuki.Text)) = 0 And _
                   Len(Trim(txtStHi.Text)) = 0 And _
                   Len(Trim(txtStZi.Text)) = 0 Then
            
                    txtStFun.Text = "00"
            
                Else
                    iRet = MsgBox("表示範囲の開始に未入力の項目があります。", vbExclamation, "入力異常")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
                iRet = TextTime_Check("Minutes", txtStFun.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '日付整合性チェック
            sChkDate = Format(txtStNen.Text, "0000") & "/" & _
                     Format(txtStTuki.Text, "00") & "/" & _
                     Format(txtStHi.Text, "00") & " " & _
                     Format(txtStZi.Text, "00") & ":" & _
                     Format(txtStFun.Text, "00")
            If IsDate(sChkDate) = False Then
                '日付指定が正しくありません
                MsgBox "日付の指定が異常です。", vbExclamation, "入力異常"
                Exit Function
            End If
            
            
            '終了年月のセット
            sStAll = Format(txtStNen.Text, "0000") & _
                     Format(txtStTuki.Text, "00") & _
                     Format(txtStHi.Text, "00") & _
                     Format(txtStZi.Text, "00") & _
                     Format(txtStFun.Text, "00")
        End If
         
        '終了チェック
        'ログデータ対象時刻の正当姓チェック
        If Len(Trim(txtEdNen.Text)) = 0 And _
           Len(Trim(txtEdTuki.Text)) = 0 And _
           Len(Trim(txtEdHi.Text)) = 0 And _
           Len(Trim(txtEdZi.Text)) = 0 And _
           Len(Trim(txtEdFun.Text)) = 0 Then
           
           '全て未入力ならMaxをセット
            sEdAll = "999999999999"
    
        ElseIf Len(Trim(txtEdNen.Text)) = 0 Or _
           Len(Trim(txtEdTuki.Text)) = 0 Or _
           Len(Trim(txtEdHi.Text)) = 0 Then
           
           '終了時刻に未入力の項目があるなら
            MsgBox "表示範囲の終了に未入力の項目があります。", _
                   vbExclamation, _
                   "入力異常"               ' EG20 V3.6.0.1 ADD
'                   "時刻指定異常"          ' EG20 V3.6.0.1 DEL
            Exit Function
        
        Else
            '終了年チェック
            iRet = TextTime_Check("Year", txtEdNen.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '終了月チェック
            iRet = TextTime_Check("Month", txtEdTuki.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '終了日チェック
            iRet = TextTime_Check("Day", txtEdHi.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '終了時チェック
            If Len(Trim(txtEdZi.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtEdZi.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtEdNen.Text)) = 0 And _
                   Len(Trim(txtEdTuki.Text)) = 0 And _
                   Len(Trim(txtEdHi.Text)) = 0 Then

                    txtEdZi.Text = "00"
                Else
                    iRet = MsgBox("表示範囲の終了に未入力の項目があります。", vbExclamation, "入力異常")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
                iRet = TextTime_Check("Hour", txtEdZi.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '終了分チェック
            If Len(Trim(txtEdFun.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtEdFun.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtEdNen.Text)) = 0 And _
                   Len(Trim(txtEdTuki.Text)) = 0 And _
                   Len(Trim(txtEdHi.Text)) = 0 And _
                   Len(Trim(txtEdZi.Text)) = 0 Then
                   
                    txtEdFun.Text = "00"
                
                Else
                    iRet = MsgBox("表示範囲の終了に未入力の項目があります。", vbExclamation, "入力異常")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END            Else
                iRet = TextTime_Check("Minutes", txtEdFun.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '日付整合性チェック
            sChkDate = Format(txtEdNen.Text, "0000") & "/" & _
                       Format(txtEdTuki.Text, "00") & "/" & _
                       Format(txtEdHi.Text, "00") & " " & _
                       Format(txtEdZi.Text, "00") & ":" & _
                       Format(txtEdFun.Text, "00")
            If IsDate(sChkDate) = False Then
                '日付指定が正しくありません
                MsgBox "日付の指定が異常です。", vbExclamation, "入力異常"
                Exit Function
            End If
            
            '終了年月のセット
            sEdAll = Format(txtEdNen.Text, "0000") & _
                     Format(txtEdTuki.Text, "00") & _
                     Format(txtEdHi.Text, "00") & _
                     Format(txtEdZi.Text, "00") & _
                     Format(txtEdFun.Text, "00")
        End If

        '開始、終了前後チェック
        dStAll = Val(sStAll)
        dEdAll = Val(sEdAll)
        If dStAll > dEdAll Then
            MsgBox "範囲指定の開始時刻が終了時刻より後に設定されています。", vbExclamation, "入力異常"
            Exit Function
        End If

    End If
    
    
    'アプリケーションログの場合のみチェックする
    Dim bFlg As Boolean
    bFlg = False
    If optApp.Value = True Then
        'モジュールﾁｪｯｸ
        For i = 0 To iModCnt
            'チェックがＯＮなら処理する
            If chkMod(i).Value = 1 And uModFileData(i).sType <> "" Then
                'フラグを立てる
                bFlg = True
            End If
        Next

        '一つも選択されていない時、エラーとする
        If bFlg = False Then
            MsgBox "項目指定に異常があります。指定した表示項目を確認してください。", vbCritical, "項目指定異常"
            Exit Function
        End If
    
        
        '号機ﾁｪｯｸ
        bFlg = False
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
        'EGXチェック
'        For i = 0 To iEgxCnt
'            'チェックがＯＮなら処理する
'            If chkEGXGoki(uEgxFileData(i).iIndex).Value = 1 Then
'                'フラグを立てる
'                bFlg = True
'            End If
'        Next
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        'EGXチェック
        For i = 0 To chkEGXGoki.UBound
            'チェックがＯＮなら処理する
            If chkEGXGoki(i).Visible = True Then
                If chkEGXGoki(i).Value = 1 Then
                    'フラグを立てる
                    bFlg = True
                End If
            End If
        Next
        'EG20 V2.1.0.1 ADD END
        
        '一つも選択されていない時、エラーとする
        If bFlg = False Then
            MsgBox "項目指定に異常があります。指定した表示項目を確認してください。", vbCritical, "項目指定異常"
            Exit Function
        End If
    End If
    
    fLogSearchCheck = True              '戻り値に正常をセット
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fWriteLogtxt
'//  機能名称  : ログテキストファイル書込み処理
'//  機能概要  : ログテキストファイル書き込みを行う。
'//　　　　　　　表示ファイル指定部：「ログ表示(テキスト表示)」釦押下時処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fWriteLogtxt() As Boolean

    Dim uLogConv As VB_LOG_DISP_SETTING 'ログ検索データ
    Dim bRet As Boolean                 '戻り値
    Dim sFileName As String
    Dim lId As Long                     'メールＩＤ
    Dim bFlag As Boolean                'フラグ
    Dim iResponse As Integer            'MsgBoxボタンコード
    Dim iStatus As Long
    Dim MyPath As String
    Dim MyName As String
    Dim lErr As Long

    On Error Resume Next
    
    fWriteLogtxt = False

    'ログ変換情報を作成する
    If sGetSearchData(uLogConv) = False Then
        Exit Function
    End If

    'ログテキストの作成
    If iNowChk1 = 1 Then
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_APP
    Else
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_HOSHU
    End If

    sObjectTopFile = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
    sFileName = MyPath & sObjectTopFile
    sObjectTopFile = Left(sObjectTopFile, Len(sObjectTopFile) - 4) & ".txt"

    'ファイル有無チェック
    MyName = Dir(sFileName, vbNormal)
    If MyName = "" Then ' ループを開始します。
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Exit Function
    End If
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '////////////////////////////////////////////////
    '保守専用関数：表示ログファイル作成処理
    '////////////////////////////////////////////////
    iStatus = dllCreateDispLogFile(lErr, sFileName, uLogConv, sObjectTopFile, PATH_LDU_APP)
    If iStatus = 1 Then    '正常のとき
        fWriteLogtxt = True
    Else                    'エラーのとき
        fWriteLogtxt = False
    End If

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sGetSearchData
'//  機能名称  : ログ変換情報作成処理
'//  機能概要  : ログトレース画面からログ変換情報を作成する。
'//　　　　　　　表示ファイル指定部：「ログ表示(テキスト表示)」釦押下時処理
'//
'//              型                  名称      意味
'//  引数      : VB_LOG_DISP_SETTING uLogConv  [OUT]ログ変換情報
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sGetSearchData(uLogConv As VB_LOG_DISP_SETTING)
    
    Dim i As Integer
    Dim ii As Integer
    Dim iBitCnt As Double
    Dim bModFlg1 As Boolean
    Dim bModFlg2 As Boolean
    Dim bEGXGokiFlg As Boolean
    Dim bCPUGokiFlg As Boolean

    On Error Resume Next
    
    sGetSearchData = False

    'ログ種別
    If iNowChk1 = 1 Then
        'アプリ選択
        uLogConv.LogType = 0
    Else
        '保守選択
        uLogConv.LogType = 1
    End If


    '範囲指定
    If iNowChk2 = 1 Then
        '範囲指定アリ
        uLogConv.TermType = 1
        uLogConv.StartTime = Format(txtStNen.Text, "0000") & _
                             Format(txtStTuki.Text, "00") & _
                             Format(txtStHi.Text, "00") & _
                             Format(txtStZi.Text, "00") & _
                             Format(txtStFun.Text, "00")
        uLogConv.EndTime = Format(txtEdNen.Text, "0000") & _
                           Format(txtEdTuki.Text, "00") & _
                           Format(txtEdHi.Text, "00") & _
                           Format(txtEdZi.Text, "00") & _
                           Format(txtEdFun.Text, "00")
        '開始が未入力の時、最小値をセットする
        If Len(Trim(uLogConv.StartTime)) = 0 Then
            uLogConv.StartTime = "198001010000"
        End If
        '終了が未入力の時、最大値をセットする
        If Len(Trim(uLogConv.EndTime)) = 0 Then
            uLogConv.EndTime = "207912312359"
        End If
    Else
        '範囲指定無し
        uLogConv.TermType = 0
        uLogConv.StartTime = ""
        uLogConv.EndTime = ""
    End If


    '表示オプション
    If optSam.Value = True Then
        'サマリー表示
        uLogConv.DispType = 0
    Else
        '詳細表示
        uLogConv.DispType = 1
    End If


    '処理結果指定
    uLogConv.ResultType = 0
    '正常
    If chkSeijou.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 1
    End If
    '異常
    If chkIjou.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 2
    End If
    '例外
    If chkReigai.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 8
    End If
    '警告
    If chkKeikoku.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 4
    End If


    '項目種別指定
    uLogConv.ItemType = 0
    'キー項目
    If chkKey.Value = 1 Then
        uLogConv.ItemType = uLogConv.ItemType + 1
    End If
    'デバッグ項目
    If chkDeb.Value = 1 Then
        uLogConv.ItemType = uLogConv.ItemType + 2
    End If


    'モジュール指定
    uLogConv.ModuleType1 = 0
    uLogConv.ModuleType2 = 0
    uLogConv.ModuleType3 = 0
    '表示号機指定
    uLogConv.Goki = 0

    'アプリケーションログの場合のみチェックする
    If optApp.Value = True Then

        bModFlg1 = False
        bModFlg2 = False

        '全件チェック
        For i = 0 To iModCnt
            'チェックがＯＮなら処理する
            If chkMod(i).Value = 1 And uModFileData(i).sType <> "" Then

                If uModFileData(i).iBit = 31 Then
                    '対応するモジュールタイプのフラグを立てる
                    If uModFileData(i).sType = 1 Then
                        bModFlg1 = True
                    Else
                        bModFlg2 = True
                    End If
                Else

                    'ビットカウント計算
                    iBitCnt = 1
                    If uModFileData(i).iBit <> 0 Then
                        For ii = 1 To uModFileData(i).iBit
                            iBitCnt = iBitCnt * 2
                        Next
                    End If

                    '対応するモジュールタイプに追加する
                    If uModFileData(i).sType = 1 Then
                        uLogConv.ModuleType1 = uLogConv.ModuleType1 + iBitCnt
                    ElseIf uModFileData(i).sType = 2 Then
                        uLogConv.ModuleType2 = uLogConv.ModuleType2 + iBitCnt
                    ElseIf uModFileData(i).sType = 3 Then
                        uLogConv.ModuleType3 = uLogConv.ModuleType3 + iBitCnt
                    End If
                End If
            End If
        Next

        If bModFlg1 = True Then
            uLogConv.ModuleType1 = -2147483648# + uLogConv.ModuleType1
        End If

        If bModFlg2 = True Then
            uLogConv.ModuleType2 = -2147483648# + uLogConv.ModuleType2
        End If

        '全件チェック(EGX)
        bEGXGokiFlg = False
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'        For i = 0 To iEgxCnt
'            'チェックがＯＮなら処理する
'            If chkEGXGoki(uEgxFileData(i).iIndex).Value = 1 Then
'                If uEgxFileData(i).iRonri = 16 Then
'                    'フラグを立てる
'                    bEGXGokiFlg = True
'                Else
'
'                    'ビットカウント計算
'                    iBitCnt = 1
'                    If uEgxFileData(i).iRonri <> 1 Then
'                        For ii = 1 To uEgxFileData(i).iRonri - 1
'                            iBitCnt = iBitCnt * 2
'                        Next
'                    End If
'
'                    '変数に追加する
'                    uLogConv.Goki = uLogConv.Goki + iBitCnt
'                End If
'            End If
'        Next
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        For i = 0 To UBound(mintStatus)
            'チェックがＯＮなら処理する
            If mintStatus(i) = 1 Then
                If i = MAX_GATE_NO - 1 Then
                    'フラグを立てる
                    bEGXGokiFlg = True
                Else

                    'ビットカウント計算
                    iBitCnt = 1
                    If i <> 0 Then
                        For ii = 1 To i
                            iBitCnt = iBitCnt * 2
                        Next
                    End If

                    '変数に追加する
                    uLogConv.Goki = uLogConv.Goki + iBitCnt
                End If
            End If
        Next
        'EG20 V2.1.0.1 DEL END

        If bEGXGokiFlg = True Then
'            uLogConv.Goki = -32768# + uLogConv.Goki            'EG20 V2.1.0.1 DEL 【フェーズ２対応】
            uLogConv.Goki = -2147483648# + uLogConv.Goki        'EG20 V2.1.0.1 ADD 【フェーズ２対応】
        End If
    End If

    sGetSearchData = True
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMailRecieve
'//  機能名称  : メール受信処理
'//  機能概要  : 保守メール・スロットからメールを受信する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Integer　　　　　　[OUT]メールID
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Function fMailRecieve() As Integer
    Dim lLen As Long                    'メールサイズ
    Dim uMail As ML_KYOTU_INF           'メール

    On Error Resume Next
    
    fMailRecieve = 0

    'メール受信
    lLen = DssMailRead(plMSlot_MN, uMail)
    If lLen > 0 Then                            '受信正常の時

      Select Case uMail.udtlHeader.dwId  'メールＩＤ
        Case ML_ID_PROEND_ORD
             '「プロセス終了指示」を受信した場合
             '「プロセス終了指示受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
             '強制終了処理を行う
             pfAbortProc
             '戻り値にメールＩＤをセット
             fMailRecieve = ML_ID_PROEND_ORD

        Case ML_ID_HOSHU_ACTIVE_REQ
             '保守画面アクティブ表示の場合
             '「保守画面アクティブ表示要求受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
'             AppActivate frmKansiLogKanri.Caption, False   ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
'             pfFormActive (frmIDULogkanri.hwnd)            ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
             AppActivate frmLDULogkanri.Caption, False      ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
             pfFormActive (frmLDULogkanri.hwnd)             ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
             fMailRecieve = ML_ID_HOSHU_ACTIVE_REQ

        Case ML_ID_IDU_LDU_LGCHGREQ_RES
             'ログ切替要求RESの場合
             '「ログ切替要求RES受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, LOG_CHANGE_RES_RECV, 0)
             fMailRecieve = ML_ID_IDU_LDU_LGCHGREQ_RES

        Case Else
        'メールＩＤ不正
          '「メールID不正」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Function

'V1.3.0.1 ADD START
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
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmLDULogkanri.Caption, False
        pfFormActive (frmLDULogkanri.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZHisentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkEGXGoki.UBound
        chkEGXGoki(intLoop).Value = CHECKBOX_OFF
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZSentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkEGXGoki.UBound
        chkEGXGoki(intLoop).Value = CHECKBOX_ON
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
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
        chkEGXGoki(intLoop).Value = CHECKBOX_OFF
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
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
        chkEGXGoki(intLoop).Value = CHECKBOX_ON
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : chkEGXGoki_Click
'//  機能名称  : 指定号機オプションボタンクリック時処理
'//  機能概要  : 内部変数のON/OFFを切り替える
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]オプションボタンインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-08   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub chkEGXGoki_Click(Index As Integer)

    Dim intGoki As Integer
    
    On Error Resume Next
    
    intGoki = CInt(chkEGXGoki(Index).Tag) - 1
    
    mintStatus(intGoki) = chkEGXGoki(Index).Value
    
End Sub

'EG20 V2.1.0.1 ADD END

