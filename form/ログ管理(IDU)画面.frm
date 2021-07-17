VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIDULogkanri 
   BorderStyle     =   0  'なし
   Caption         =   "                                                                  ＩＤ中継ユニットログ管理"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9360
      Top             =   7440
   End
   Begin VB.CommandButton cmdInstall 
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
      TabIndex        =   238
      Top             =   6480
      Width           =   2600
   End
   Begin VB.TextBox txtDummy 
      Height          =   375
      Left            =   11400
      TabIndex        =   0
      Text            =   "Text11"
      Top             =   15000
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogHyouzi 
      Caption         =   "   ログ表示    (テキスト表示)"
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
      TabIndex        =   176
      Top             =   480
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
      TabIndex        =   177
      Top             =   1680
      Width           =   2600
   End
   Begin VB.CommandButton cmdSyslog 
      Caption         =   " システムログ   媒体出力"
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
      TabIndex        =   178
      Top             =   2880
      Visible         =   0   'False
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
      TabIndex        =   179
      Top             =   4080
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.CommandButton cmdMemoridump 
      Caption         =   " メモリダンプ   媒体出力"
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
      TabIndex        =   180
      Top             =   5280
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
      TabIndex        =   181
      Top             =   7920
      Width           =   2600
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8535
      Left            =   120
      TabIndex        =   182
      Top             =   375
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
      TabPicture(0)   =   "ログ管理(IDU)画面.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblSize"
      Tab(0).Control(1)=   "lblEnd"
      Tab(0).Control(2)=   "lblStart"
      Tab(0).Control(3)=   "lblFile"
      Tab(0).Control(4)=   "cmdRefresh"
      Tab(0).Control(5)=   "optHoshu"
      Tab(0).Control(6)=   "optApp"
      Tab(0).Control(7)=   "LstFile"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "表示項目指定"
      TabPicture(1)   =   "ログ管理(IDU)画面.frx":001C
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
      TabPicture(2)   =   "ログ管理(IDU)画面.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabCorner"
      Tab(2).Control(1)=   "cmdHHisentaku"
      Tab(2).Control(2)=   "cmdHSentaku"
      Tab(2).Control(3)=   "cmdZHisentaku"
      Tab(2).Control(4)=   "cmdZSentaku"
      Tab(2).ControlCount=   5
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
         TabIndex        =   188
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
            Top             =   630
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
            Top             =   330
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
            TabIndex        =   200
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
            TabIndex        =   199
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
            TabIndex        =   198
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
            TabIndex        =   197
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
            TabIndex        =   196
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
            TabIndex        =   195
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
            TabIndex        =   194
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
            TabIndex        =   193
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
            TabIndex        =   192
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
            TabIndex        =   191
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
            TabIndex        =   190
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
            TabIndex        =   189
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
         TabIndex        =   187
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
            Top             =   540
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
         TabIndex        =   186
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
         TabIndex        =   185
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
         TabIndex        =   183
         Top             =   2520
         Width           =   8415
         Begin VB.CommandButton cmdModSen 
            Caption         =   "全て選択"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            Width           =   1335
         End
         Begin VB.CommandButton cmdModHi 
            Caption         =   "全て非選択"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.Frame frmModMeisai 
            Height          =   5055
            Left            =   120
            TabIndex        =   184
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
               TabIndex        =   236
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
               TabIndex        =   235
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
               TabIndex        =   234
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
               TabIndex        =   233
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
               TabIndex        =   232
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
               TabIndex        =   231
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
               TabIndex        =   230
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
               TabIndex        =   229
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
               TabIndex        =   228
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
               TabIndex        =   227
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
               TabIndex        =   226
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
               TabIndex        =   225
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
               TabIndex        =   224
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
               TabIndex        =   223
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
               TabIndex        =   222
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
               TabIndex        =   221
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
               TabIndex        =   220
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
               TabIndex        =   219
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
               TabIndex        =   218
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
               TabIndex        =   217
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
               TabIndex        =   216
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
               TabIndex        =   215
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
               TabIndex        =   214
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
               TabIndex        =   213
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
               TabIndex        =   212
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
               TabIndex        =   211
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
               TabIndex        =   210
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
               TabIndex        =   209
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
               TabIndex        =   208
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
               TabIndex        =   207
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
               TabIndex        =   206
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
               TabIndex        =   205
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
               Width           =   1890
            End
         End
      End
      Begin VB.CommandButton cmdZSentaku 
         Caption         =   "  全コーナ    全号機   選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -74760
         TabIndex        =   75
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdZHisentaku 
         Caption         =   "   全コーナ     全号機 非選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -72600
         TabIndex        =   76
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdHSentaku 
         Caption         =   "  表示コーナ   全号機 選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -70440
         TabIndex        =   77
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdHHisentaku 
         Caption         =   "  表示コーナ    全号機 非選択"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -68280
         TabIndex        =   78
         Top             =   840
         Width           =   2000
      End
      Begin TabDlg.SSTab tabCorner 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   79
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
         TabCaption(0)   =   " "
         TabPicture(0)   =   "ログ管理(IDU)画面.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkCorner(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkCorner(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkCorner(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkCorner(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkCorner(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkCorner(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkCorner(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkCorner(7)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkCorner(8)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkCorner(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkCorner(10)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkCorner(11)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkCorner(12)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkCorner(13)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkCorner(14)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "chkCorner(15)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "  "
         TabPicture(1)   =   "ログ管理(IDU)画面.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkCorner(16)"
         Tab(1).Control(1)=   "chkCorner(17)"
         Tab(1).Control(2)=   "chkCorner(18)"
         Tab(1).Control(3)=   "chkCorner(19)"
         Tab(1).Control(4)=   "chkCorner(20)"
         Tab(1).Control(5)=   "chkCorner(21)"
         Tab(1).Control(6)=   "chkCorner(22)"
         Tab(1).Control(7)=   "chkCorner(23)"
         Tab(1).Control(8)=   "chkCorner(24)"
         Tab(1).Control(9)=   "chkCorner(25)"
         Tab(1).Control(10)=   "chkCorner(26)"
         Tab(1).Control(11)=   "chkCorner(27)"
         Tab(1).Control(12)=   "chkCorner(28)"
         Tab(1).Control(13)=   "chkCorner(29)"
         Tab(1).Control(14)=   "chkCorner(30)"
         Tab(1).Control(15)=   "chkCorner(31)"
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "  "
         TabPicture(2)   =   "ログ管理(IDU)画面.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkCorner(32)"
         Tab(2).Control(1)=   "chkCorner(33)"
         Tab(2).Control(2)=   "chkCorner(34)"
         Tab(2).Control(3)=   "chkCorner(35)"
         Tab(2).Control(4)=   "chkCorner(36)"
         Tab(2).Control(5)=   "chkCorner(37)"
         Tab(2).Control(6)=   "chkCorner(38)"
         Tab(2).Control(7)=   "chkCorner(39)"
         Tab(2).Control(8)=   "chkCorner(40)"
         Tab(2).Control(9)=   "chkCorner(41)"
         Tab(2).Control(10)=   "chkCorner(42)"
         Tab(2).Control(11)=   "chkCorner(43)"
         Tab(2).Control(12)=   "chkCorner(44)"
         Tab(2).Control(13)=   "chkCorner(45)"
         Tab(2).Control(14)=   "chkCorner(46)"
         Tab(2).Control(15)=   "chkCorner(47)"
         Tab(2).ControlCount=   16
         TabCaption(3)   =   "  "
         TabPicture(3)   =   "ログ管理(IDU)画面.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkCorner(48)"
         Tab(3).Control(1)=   "chkCorner(49)"
         Tab(3).Control(2)=   "chkCorner(50)"
         Tab(3).Control(3)=   "chkCorner(51)"
         Tab(3).Control(4)=   "chkCorner(52)"
         Tab(3).Control(5)=   "chkCorner(53)"
         Tab(3).Control(6)=   "chkCorner(54)"
         Tab(3).Control(7)=   "chkCorner(55)"
         Tab(3).Control(8)=   "chkCorner(56)"
         Tab(3).Control(9)=   "chkCorner(57)"
         Tab(3).Control(10)=   "chkCorner(58)"
         Tab(3).Control(11)=   "chkCorner(59)"
         Tab(3).Control(12)=   "chkCorner(60)"
         Tab(3).Control(13)=   "chkCorner(61)"
         Tab(3).Control(14)=   "chkCorner(62)"
         Tab(3).Control(15)=   "chkCorner(63)"
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "  "
         TabPicture(4)   =   "ログ管理(IDU)画面.frx":00C4
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chkCorner(64)"
         Tab(4).Control(1)=   "chkCorner(65)"
         Tab(4).Control(2)=   "chkCorner(66)"
         Tab(4).Control(3)=   "chkCorner(67)"
         Tab(4).Control(4)=   "chkCorner(68)"
         Tab(4).Control(5)=   "chkCorner(69)"
         Tab(4).Control(6)=   "chkCorner(70)"
         Tab(4).Control(7)=   "chkCorner(71)"
         Tab(4).Control(8)=   "chkCorner(72)"
         Tab(4).Control(9)=   "chkCorner(73)"
         Tab(4).Control(10)=   "chkCorner(74)"
         Tab(4).Control(11)=   "chkCorner(75)"
         Tab(4).Control(12)=   "chkCorner(76)"
         Tab(4).Control(13)=   "chkCorner(77)"
         Tab(4).Control(14)=   "chkCorner(78)"
         Tab(4).Control(15)=   "chkCorner(79)"
         Tab(4).ControlCount=   16
         TabCaption(5)   =   "  "
         TabPicture(5)   =   "ログ管理(IDU)画面.frx":00E0
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkCorner(80)"
         Tab(5).Control(1)=   "chkCorner(81)"
         Tab(5).Control(2)=   "chkCorner(82)"
         Tab(5).Control(3)=   "chkCorner(83)"
         Tab(5).Control(4)=   "chkCorner(84)"
         Tab(5).Control(5)=   "chkCorner(85)"
         Tab(5).Control(6)=   "chkCorner(86)"
         Tab(5).Control(7)=   "chkCorner(87)"
         Tab(5).Control(8)=   "chkCorner(88)"
         Tab(5).Control(9)=   "chkCorner(89)"
         Tab(5).Control(10)=   "chkCorner(90)"
         Tab(5).Control(11)=   "chkCorner(91)"
         Tab(5).Control(12)=   "chkCorner(92)"
         Tab(5).Control(13)=   "chkCorner(93)"
         Tab(5).Control(14)=   "chkCorner(94)"
         Tab(5).Control(15)=   "chkCorner(95)"
         Tab(5).ControlCount=   16
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   175
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   174
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   173
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   172
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   171
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   170
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   169
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   168
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   167
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   166
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   165
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   164
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   163
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   162
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   161
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   160
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   159
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   158
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   157
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   155
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   154
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   153
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   152
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   151
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   150
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   149
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   148
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   147
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   146
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   145
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   144
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   143
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   142
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   141
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   140
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   139
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   138
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   137
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   136
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   135
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   134
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   133
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   132
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   131
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   130
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   129
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   128
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   127
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   126
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   125
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   124
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   123
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   122
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   121
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   120
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   119
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   118
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   117
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   116
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   115
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   114
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   113
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   112
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   111
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   110
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   109
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   108
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   107
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   106
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   105
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   104
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   103
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   102
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   101
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   100
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   99
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   98
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   97
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   96
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   15
            Left            =   6360
            TabIndex        =   95
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   14
            Left            =   4200
            TabIndex        =   94
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   13
            Left            =   2160
            TabIndex        =   93
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   12
            Left            =   120
            TabIndex        =   92
            Top             =   2040
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   11
            Left            =   6360
            TabIndex        =   91
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   10
            Left            =   4200
            TabIndex        =   90
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   9
            Left            =   2160
            TabIndex        =   89
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   8
            Left            =   120
            TabIndex        =   88
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   7
            Left            =   6360
            TabIndex        =   87
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   6
            Left            =   4200
            TabIndex        =   86
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   5
            Left            =   2160
            TabIndex        =   85
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   4
            Left            =   120
            TabIndex        =   84
            Top             =   1080
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   3
            Left            =   6360
            TabIndex        =   83
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   2
            Left            =   4200
            TabIndex        =   82
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   1
            Left            =   2160
            TabIndex        =   81
            Top             =   600
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   600
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
         TabIndex        =   204
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
         TabIndex        =   203
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
         TabIndex        =   202
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
         TabIndex        =   201
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C000&
      Caption         =   "IDUアプリケーションログ管理"
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
      TabIndex        =   237
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmIDULogkanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmIDULogKanri.frm
'//  パッケージ名：ID中継ユニットログ管理画面
'//
'//  概要：ID中継ユニットログ管理画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・ID中継ユニット、ログ管理画面(frmLogKanri.frm)を流用
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02   REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応【監視D-115】
'//                 　・処理結果メッセージボックスの文言変更
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

Public sYobidasi As String
Public iNowChk1 As Integer
Public iNowChk2 As Integer

'DB接続用
'アプリログ用
Private cnConn              As New ADODB.Connection     'Connection オブジェクトの定義
Private rsRecordSet         As New ADODB.Recordset      'RecordSet オブジェクトの定義
'アプリログテーブル取得用
Private gLogData() As typLogDataTable
'アプリログ保存管理DB
Private Type typLogDataTable
    sName As String
    sStTime As String
    sEdTime As String
    iSize As Long
End Type

'保守ログ用
Private cnConn2              As New ADODB.Connection     'Connection オブジェクトの定義

'///////////////////////////////////////////////////////////////////
'対象ファイルフルパス（複数ﾌｧｲﾙの時、ｽﾍﾟｰｽ1文字で区切る。）
'///////////////////////////////////////////////////////////////////
Private sObjectFiles As String      'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽで選択中のﾌｧｲﾙのﾌﾙﾊﾟｽ文字列
Private sObjectTopFile As String    '同上、選択中の先頭（最旧）ﾌｧｲﾙ名。(12文字)。

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
    sName As String                'モジュール名
    sDai  As String                '大項目
    sShou As String                '小項目
    sType As String                'モジュールタイプ
    iBit  As Integer               'ビット番号
End Type

Private uModFileData(79) As ModFileData
Private iModCnt As Integer

'///////////////////////////////////////////////////////////////////
'ICM情報格納エリア
'///////////////////////////////////////////////////////////////////
Private Type IcmFileData
    iRonri As Integer               '論理号機
    iHyozi As Integer               '表示号機
    iConer As Integer               'コーナー番号
    iIndex As Integer               'chkCornerのINDEX
End Type

Private uIcmFileData(31) As IcmFileData
Private iIcmCnt As Integer

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

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ID中継ユニットログ管理画面(アクティブ時)
'//  機能概要  : 最前面表示を行う。
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
Private Sub Form_Activate()
    pfFormActive (hwnd)
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : ID中継ユニットログ管理画面(ディアクティブ時)
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
'//  機能名称  : ID中継ユニットログ管理画面(ロード時)
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
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-25  CODED BY  [TCC] T.Koyama
'//                 EG20フェーズ２対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim intModulesFileNo As Integer
    Dim sModules As String * IDU_LOG_SIZE    '１行分ファイル内容取得用
    Dim Cnt As Integer
    Dim iMozi As Integer
    Dim iKbn As Integer
    Dim iRet As Integer
'    Dim sConer As String * IDU_LOG_CONER_SIZE 'ファイルチェンジツールの実行ファイル名（フルパス)   ' EG20 V3.6.0.1 DEL
    Dim sConer As String * 30                  'ファイルチェンジツールの実行ファイル名（フルパス)   ' EG20 V3.6.0.1 ADD
    Dim sType As String * IDU_LOG_TYPE        '設置タイプ
    Dim sIcmData As String * IDU_LOG_SIZE     '１行分ファイル内容取得用
    Dim i As Integer                          'ループ用
    Dim sKeyName As String
    Dim str As String
    Dim iLoop As Integer
    Dim MyName As String
    Dim iErr As Integer
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
' EG20 V3.6.0.1 ADD START
    Dim myLen As Long
    Dim strCodeTxt As String
    Dim strCorner As String
' EG20 V3.6.0.1 ADD END
    
    'パス指定
    IDU_PROFILE_NAME = PATH_IDU_APP & IDU_STATION_FILE
    IDU_PROFILE_NAME_ICM = PATH_IDU_APP & IDU_ICM_FILE
    
    gStrCurrentForm = sFormName_IDULog
     
    cmdCancel.Caption = "ログ管理" & Chr(13) & "画面へ戻る"
    cmdLogHyouzi.Caption = "ログ表示" & Chr(13) & "(テキスト表示）"
    cmdZSentaku.Caption = "全コーナ" & Chr(13) & "全号機　選択"
    cmdZHisentaku.Caption = "全コーナ" & Chr(13) & "全号機　非選択"
    cmdHSentaku.Caption = "表示コーナ" & Chr(13) & "全号機　選択"
    cmdHHisentaku.Caption = "表示コーナ" & Chr(13) & "全号機　非選択"
     
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
    
    For i = 0 To 5
        tabCorner.Tab = 5 - i
        tabCorner.Caption = ""
    Next

    'INIファイルよりの設定取得
    'モジュール指定取得
    On Error GoTo FileError
    iErr = 1
    
    'ファイル有無チェック
    MyName = Dir(PATH_IDU_APP & IDU_MODULES_FILE_FULLPASS, vbNormal)
    If MyName = "" Then
        GoTo FileError
    End If
    
    Cnt = 0
    
    For Cnt = 0 To 79
        sKeyName = "ID" & Format(Cnt, "000")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ID, _
                                       sKeyName, _
                                       DEFAILT, sModules, Len(sModules), _
                                       PATH_IDU_APP & IDU_MODULES_FILE_FULLPASS)
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
    
    'ファイル有無チェック
    MyName = Dir(IDU_PROFILE_NAME, vbNormal)
    If MyName = "" Then
        GoTo FileError
    End If
    
    MyName = Dir(IDU_PROFILE_NAME_ICM, vbNormal)
    If MyName = "" Then
        iErr = 3
        GoTo FileError
    End If
    
    
    'コーナー情報の取得
    '６コーナー目の名称取得
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER6, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER6, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 5
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '文字列を変換
        myLen = LenB(strCodeTxt)                        '半角換算のバイト数を取得
    
        If myLen <= 24 Then                             '指定の長さより短い場合
            strCorner = strCodeTxt

        Else
            '該当の文字列の方が長い場合、指定のバイトでカットする
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '漢字１バイト目で分断された場合の処理
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 5
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 5
        tabCorner.Caption = ""
    End If
    
    '５コーナー目の名称取得
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER5, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER5, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 4
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '文字列を変換
        myLen = LenB(strCodeTxt)                        '半角換算のバイト数を取得
    
        If myLen <= 24 Then                             '指定の長さより短い場合
            strCorner = strCodeTxt

        Else
            '該当の文字列の方が長い場合、指定のバイトでカットする
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '漢字１バイト目で分断された場合の処理
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 4
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 4
        tabCorner.Caption = ""
    End If
    
    '４コーナー目の名称取得
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER4, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER4, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 3
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '文字列を変換
        myLen = LenB(strCodeTxt)                        '半角換算のバイト数を取得
    
        If myLen <= 24 Then                             '指定の長さより短い場合
            strCorner = strCodeTxt

        Else
            '該当の文字列の方が長い場合、指定のバイトでカットする
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '漢字１バイト目で分断された場合の処理
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 3
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 3
        tabCorner.Caption = ""
    End If
     
    '３コーナー目の名称取得
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER3, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER3, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 2
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '文字列を変換
        myLen = LenB(strCodeTxt)                        '半角換算のバイト数を取得
    
        If myLen <= 24 Then                             '指定の長さより短い場合
            strCorner = strCodeTxt

        Else
            '該当の文字列の方が長い場合、指定のバイトでカットする
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '漢字１バイト目で分断された場合の処理
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 2
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 2
        tabCorner.Caption = ""
    End If
    
    '２コーナー目の名称取得
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER2, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER2, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 1
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '文字列を変換
        myLen = LenB(strCodeTxt)                        '半角換算のバイト数を取得
    
        If myLen <= 24 Then                             '指定の長さより短い場合
            strCorner = strCodeTxt

        Else
            '該当の文字列の方が長い場合、指定のバイトでカットする
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '漢字１バイト目で分断された場合の処理
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 1
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 1
        tabCorner.Caption = ""
    End If

    '１コーナー目の名称取得
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER1, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER1, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If

' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 0
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '文字列を変換
        myLen = LenB(strCodeTxt)                        '半角換算のバイト数を取得
    
        If myLen <= 24 Then                             '指定の長さより短い場合
            strCorner = strCodeTxt

        Else
            '該当の文字列の方が長い場合、指定のバイトでカットする
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '漢字１バイト目で分断された場合の処理
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 0
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 0
        tabCorner.Caption = ""
    End If

    iErr = 3

    iIcmCnt = -1
    'ICM情報取得
    For i = 1 To 32
        sKeyName = "icm" & Format(i, "00")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ICM, _
                                       sKeyName, _
                                       DEFAILT, sIcmData, Len(sIcmData), _
                                       IDU_PROFILE_NAME_ICM)
        If iRet = 0 Then
            GoTo FileError
        End If
                        
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
        
        '通路種別が未設置の時は処理せず
        If Trim(sFData(5)) <> "＊" Then
            iIcmCnt = iIcmCnt + 1
            uIcmFileData(iIcmCnt).iRonri = i
            uIcmFileData(iIcmCnt).iHyozi = Trim(sFData(1))
            uIcmFileData(iIcmCnt).iConer = Trim(sFData(3))
            uIcmFileData(iIcmCnt).iIndex = uIcmFileData(iIcmCnt).iConer * 16 - 16 + Int(Trim(sFData(4))) - 1
            chkCorner(uIcmFileData(iIcmCnt).iIndex).Visible = True
            chkCorner(uIcmFileData(iIcmCnt).iIndex).Caption = uIcmFileData(iIcmCnt).iHyozi & "号機"
        End If
     Next
    
    On Error GoTo OtherError
       
    'DB接続
    'アプリログ用
    cnConn.ConnectionString = "File Name=" & PATH_IDU_APP & PATH_IDU_APPLOG
    cnConn.Open
        
    On Error GoTo 0
    
    iNowChk1 = 1
    iNowChk2 = 2
    
    'リストの初期表示
    If sSetListBox = False Then
        '「ID中継ユニットログ管理：アプリログ表示異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
        'リストボックスの初期化
        LstFile.Clear
        MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
   End If
   
   '「ID中継ユニットログ管理画面：表示」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_LOG_KANRI_GAMEN_START, 0)
   
   
 Exit Sub
    
FileError:
    Select Case iErr
    Case 1:
       '「ID中継ユニットログ管理：INIファイル異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
     Case 2:
       '「ID中継ユニットログ管理：INIファイル異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
     Case 3:
       '「ID中継ユニットログ管理：INIファイル異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
     End Select
   MsgBox "INIファイルの取得に失敗しました｡", vbCritical, "ファイル異常"
   
   Exit Sub
OtherError:
   '「ID中継ユニットログ管理：アプリログ表示異常」ログ出力
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
   LstFile.Clear
   MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdLogHyouzi_Click
'//  機能名称  : 「ログ表示(テキスト表示）」釦押下時処理
'//  機能概要  : 選択ファイルを、テキストにて表示する。
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
    Dim lngErrCode As Long   'エラーコード

   '「ID中継ユニットログ管理：ログ表示釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_TEXT_HYOUJI_BUTTOM, 0)

    'ログ検索データ正当姓チェック
    bRet = fLogSearchCheck
    If bRet = False Then                                'ログ検索データにエラーがある場合、処理終了
        Exit Sub
    End If

    'ログテキストファイルを書き込む
    bRet = fWriteLogtxt
    If bRet = True Then                                 'ログテキストファイルが正常に作成された場合
        '「ID中継ユニットログ管理：ログテキストファイル作成正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CREATE_TEXT_HYOUJI, 0)
        'ファイルコピー
        sFileName = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
        sFileName = PATH_IDU_APP & PATH_IDU_WORK & "\\" & Left(sFileName, Len(sFileName) - 4) & ".txt"
        'ファイルオープン
        On Error GoTo FileError
        sCommand = MN_EXE_MEMO & sFileName              '実行コマンドを作成する
        lRetVal = Shell(sCommand, vbMaximizedFocus)     'ノートパッドを起動する
        AppActivate lRetVal, True                       'アクティブ（前面表示）にする
        SendKeys "{LEFT}", True
        On Error GoTo 0
        '「ID中継ユニットログ管理：ログテキスト表示正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_OK, 0)
    Else
        '「ID中継ユニットログ管理：出力データ作成異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_CREATE_TEXT_ERROR, lngErrCode)
       '「データ出力失敗」ポップアップ表示
       MsgBox "媒体出力するデータの作成に失敗しました。", vbCritical, "データ出力失敗"
    End If
    Exit Sub

FileError:
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   '「ID中継ユニットログ管理：ログテキスト表示処理異常」ログ出力
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdLog_Click
'//  機能名称  : 「ログ媒体出力」釦押下時処理
'//  機能概要  : 選択ファイルを、指定フォルダへ出力する。
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
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02   REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応【監視D-115】
'//                 　・処理結果メッセージボックスの文言変更
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

    '「ID中継ユニットログ管理画面：ログ媒体出力釦押下」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_OUTPUT_BUTTOM, 0)

    Dim bFrmShow As Boolean
    bFrmShow = False

    txtDummy.SetFocus

    On Error GoTo EVENTLOG_ERROR
    If iNowChk1 = 1 Then
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_APP
    Else
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_HOSHU
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
            frmIDULogkanri.Refresh
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
'   '処理番号格納（処理中）
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

'V1.6.0.1 ADD START

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
    
    'コピー先フォルダパス作成(指定フォルダ￥IDULOG)
    sWriteDir = sWriteDir & "\" & IDU_LOGKANRI_IDULOG
    
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
'EG20 V2.0.1.1【監視D-115】DEL START
'        MsgBox "ＨＤＤ内一時フォルダへの出力は正常終了しました。", vbInformation + vbOKOnly, "出力結果"
'EG20 V2.0.1.1【監視D-115】DEL END
'EG20 V2.0.1.1【監視D-115】ADDL START
        MsgBox "正常終了しました。", vbInformation + vbOKOnly, "出力結果"
'EG20 V2.0.1.1【監視D-115】ADDL END
    End If
    
    '「ID中継ユニットログ管理画面：ログ媒体出力処理正常」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_OUTPUT_OK, 0)

    Exit Sub

EVENTLOG_ERROR:
   'V1.6.0.1 ADD START
       'ファイルシステムオブジェクト解放
      Set objFso = Nothing
      '「ID中継ユニットログ管理画面：フォルダ作成異常」
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_CREATE_LOGFOLDER_ERROR, 0)
   'V1.6.0.1 ADD END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    If UCase(Left(sWriteDir, 1)) = "A" Then
        MsgBox "ＦＤ出力は異常終了しました。", vbCritical, "出力結果"
    Else
'EG20 V2.0.1.1【監視D-115】DEL START
'        MsgBox "ＨＤＤ内一時フォルダへの出力は異常終了しました。", vbCritical, "出力結果"
'EG20 V2.0.1.1【監視D-115】DEL END
'EG20 V2.0.1.1【監視D-115】ADDL START
        MsgBox "異常終了しました。", vbCritical, "出力結果"
'EG20 V2.0.1.1【監視D-115】ADDL END
    End If
    
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「ID中継ユニットログ管理画面：ログ媒体出力処理異常」
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
    Dim udtMail As IDU_LDU_LGCHGREQ_CMD     'ログ切替要求
    Dim lngErrCode As Long                  'エラーコード
    Dim bFlag As Boolean                    'メール受信フラグ
    Dim lId As Long                         'メールID

    On Error Resume Next

    LstFile.Clear

    '「ID中継ユニットログ管理画面：ログ切替釦押下」
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
    bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
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
'//  関数名称  : cmdInstall_Click
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
Private Sub cmdInstall_Click()
   On Error Resume Next
  
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)
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
   
   '「ID中継ユニットログ管理画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_LOG_KANRI_GAMEN_END, 0)
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub optApp_Click()

    On Error GoTo Err_mgs
    
   '「ID中継ユニットログ管理画面：アプリログ」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_APLLOG, 0)
    
    '選択されていたのが、保守プログラムログだった場合
    If iNowChk1 <> 1 Then
        'DBの接続を切り替える
        If iNowChk1 <> 0 Then
            If Not cnConn2 Is Nothing Then
                cnConn2.Close
            End If
        End If
        '接続をナシにする
        iNowChk1 = 0

        cnConn.ConnectionString = "File Name=" & PATH_IDU_APP & PATH_IDU_APPLOG
        cnConn.Open
        
        '選択されているチェックを保持する
        iNowChk1 = 1
        
        '非表示になっていた項目を表示させる
        frmMod.Visible = True
        cmdZSentaku.Visible = True
        cmdZHisentaku.Visible = True
        cmdHSentaku.Visible = True
        cmdHHisentaku.Visible = True
        tabCorner.Visible = True
        
        '表示を再読み込みする
        If sSetListBox = False Then
            '「ID中継ユニットログ管理：アプリログ表示異常」ログ出力
             Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
            'リストボックスの初期化
            LstFile.Clear
            MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
        End If
    End If
   
   '「ID中継ユニットログ管理：アプリログ表示正常」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_APLLOG_OK, 0)

   Exit Sub
    
Err_mgs:
   '「ID中継ユニットログ管理：アプリログ表示異常」ログ出力
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
    
   '「ID中継ユニットログ管理画面：保守ログ」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_HOSHULOG, 0)
    
    '選択されていたのが、アプリケーションログだった場合
    If iNowChk1 <> 2 Then
        'DBの接続を切り替える
        If iNowChk1 <> 0 Then
            If Not cnConn Is Nothing Then
                cnConn.Close
            End If
        End If
        '接続をナシにする
        iNowChk1 = 0

        cnConn2.ConnectionString = "File Name=" & PATH_IDU_APP & PATH_IDU_HOSHULOG
        cnConn2.Open
            
        '選択されているチェックを保持する
        iNowChk1 = 2
        
        '表示になっている項目を非表示にする
        frmMod.Visible = False
        cmdZSentaku.Visible = False
        cmdZHisentaku.Visible = False
        cmdHSentaku.Visible = False
        cmdHHisentaku.Visible = False
        tabCorner.Visible = False
        
        '表示を再読み込みする
        If sSetListBox = False Then
            '「ID中継ユニットログ管理：保守ログ表示異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_HODHULOG_ERROR, 0)
            'リストボックスの初期化
            LstFile.Clear
            MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
        End If
    End If
     
    '「ID中継ユニットログ管理：保守ログ表示正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_HODHULOG_OK, 0)
    
    Exit Sub

Err_mgs:
    '「ID中継ユニットログ管理：保守ログ表示異常」ログ出力
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
    
    For iCnt = 0 To iModCnt
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
        
    For iCnt = 0 To iModCnt
        chkMod(iCnt).Value = 0
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdModHi_Click
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
        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-14 REVISED BY [TCC] M.Matsumoto
'//                【統-336対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function TextTime_Check(sType As String, sTxt As String)
    Dim iChk As Integer
    Dim iRet As Integer
    Dim sChk As String
    
    Dim k As Integer                            'EG20 V2.0.1.1 ADD
    
    '戻り値に異常をセット
    TextTime_Check = False
        
    If Trim(sTxt) <> "" Then
        iChk = Val(sTxt)
        'EG20 V2.0.1.1 ADD START 【統-336対応】
        '入力された中に数値以外の文字が存在する場合は、エラー
        For k = 1 To Len(sTxt)
            If Not Mid(sTxt, k, 1) Like "[0-9]" Then
                iRet = MsgBox("入力された文字は数字ではありません。", vbExclamation, "入力異常")
                '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                Exit Function
            End If
        Next k
        'EG20 V2.0.1.1 ADD END
                    
        'EG20 V2.0.1.1 DEL START 【統-336対応】
'        If iChk = 0 And sType <> "Hour" And sType <> "Minutes" Then
'            iRet = MsgBox("入力された文字は数字ではありません。", vbExclamation, "入力異常")
'            '「ID中継ユニットログ管理：入力時刻異常」ログ出力
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
            
            'EG20 V2.0.1.1 DEL START 【統-336対応】
            '共通
'            If Len(Trim(str(iChk))) <> Len(sTxt) Then
'                iRet = MsgBox("入力された文字に数字以外のものが含まれています。", vbExclamation, "入力異常")
'                '「ID中継ユニットログ管理：入力時刻異常」ログ出力
'                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
                        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Month"
                    '月
                    If iChk < 1 Or iChk > 12 Then
                        iRet = MsgBox("月指定の範囲を超えています。", vbExclamation, "入力異常")
                        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Day"
                    '日
                    If iChk < 1 Or iChk > 31 Then
                        iRet = MsgBox("日指定の範囲を超えています。", vbExclamation, "入力異常")
                        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Hour"
                    '時
                    If iChk < 0 Or iChk > 23 Then
                        iRet = MsgBox("時間指定の範囲を超えています。", vbExclamation, "入力異常")
                        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Minutes"
                    '分
                    If iChk < 0 Or iChk > 59 Then
                        iRet = MsgBox("時間指定の範囲を超えています。", vbExclamation, "入力異常")
                        '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
            End Select
'        End If             'EG20 V2.0.1.1 DEL 【統-336対応】
    End If
    
    '戻り値に正常を返す
    TextTime_Check = True
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdZSentaku_Click
'//  機能名称  : 「全コーナー　全号機選択」釦押下時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示号機指定部：
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
Private Sub cmdZSentaku_Click()
    Dim iCnt As Integer
        
    For iCnt = 0 To 95
        chkCorner(iCnt).Value = 1
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdZHisentaku_Click
'//  機能名称  : 「全コーナー　全号機非選択」釦押下時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示号機指定部：
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
Private Sub cmdZHisentaku_Click()
    Dim iCnt As Integer
        
    For iCnt = 0 To 95
        chkCorner(iCnt).Value = 0
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdHSentaku_Click
'//  機能名称  : 「表示コーナー　全号機選択」釦押下時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示号機指定部：
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
Private Sub cmdHSentaku_Click()
    Dim iCnt As Integer
    Dim iMin As Integer
    Dim iMax As Integer
        
    '最小値、最大値取得
    iMin = tabCorner.Tab * 16
    iMax = tabCorner.Tab * 16 + 15
        For iCnt = iMin To iMax
            chkCorner(iCnt).Value = 1
        Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdHHisentaku_Click
'//  機能名称  : 「表示コーナー　全号機非選択」釦押下時処理
'//  機能概要  : 表示を更新する。
'//　　　　　　　表示号機指定部：
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
Private Sub cmdHHisentaku_Click()
    Dim iCnt As Integer
    Dim iMin As Integer
    Dim iMax As Integer
        
    '最小値、最大値取得
    iMin = tabCorner.Tab * 16
    iMax = tabCorner.Tab * 16 + 15
        For iCnt = iMin To iMax
            chkCorner(iCnt).Value = 0
        Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Unload
'//  機能名称  : 「メニュー画面へ戻る」釦押下時処理
'//  機能概要  : 　DB接続の解放を行う。
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
Private Sub Form_Unload(Cancel As Integer)
    If iNowChk1 = 1 Then
        If Not cnConn Is Nothing Then
            cnConn.Close
        End If
    End If
    If iNowChk1 = 2 Then
        If Not cnConn2 Is Nothing Then
            cnConn2.Close
        End If
    End If

    'RecordSet定義をメモリからの削除する
    Set rsRecordSet = Nothing
    'Connection定義をメモリから削除する
    Set cnConn = Nothing
    'Connection2定義をメモリから削除する
    Set cnConn2 = Nothing
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSetListBox
'//  機能名称  : ログファイルを登録
'//  機能概要  : ログファイルをリストボックスに登録する。
'//　　　　　　　表示ファイル指定部：初期処理
'//　　　　　　　　　　　　　　　　  表示ログラジオ釦押下時処理
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
Private Function sSetListBox()
    Dim i As Integer
    Dim iCnt As Integer
    Dim strSQL As String
    Dim datWork As Date
    Dim sEntry As String        '編集文字列

    On Error GoTo Err_mgs

    sSetListBox = False
    '論理号機のログ情報取得のＳＱＬ文
    strSQL = "Select LOG_NAME,LOG_START_TIME,LOG_END_TIME,LOG_SIZE" _
            & " from T_LOG"

    On Error Resume Next            ' エラーのトラップを留保します。
    Err.Clear

    'アプリ、保守チェック
    Select Case iNowChk1
        Case 1
            rsRecordSet.Open strSQL, cnConn
        Case 2
            rsRecordSet.Open strSQL, cnConn2
        Case Else
            Exit Function
    End Select

    'ＳＱＬ実行エラーだった場合
    If Err.Number <> 0 Then
        'レコードセットのクローズ
        rsRecordSet.Close

        GoTo Err_mgs
    End If
    i = 0
   'ログ情報を構造体配列(gtypLogData)に格納する
    Do While Not rsRecordSet.EOF
        ReDim Preserve gLogData(i)
        gLogData(i).sName = rsRecordSet!LOG_NAME
        gLogData(i).sStTime = Format(rsRecordSet!LOG_START_TIME, "yyyy/mm/dd hh:mm:ss")
        gLogData(i).sEdTime = Format(rsRecordSet!LOG_END_TIME, "yyyy/mm/dd hh:mm:ss")
        gLogData(i).iSize = rsRecordSet!LOG_SIZE

        rsRecordSet.MoveNext
        i = i + 1
    Loop
    iCnt = i

    'ＳＱＬ実行エラーだった場合
    If Err.Number <> 0 Then
        'レコードセットのクローズ
        rsRecordSet.Close

        GoTo Err_mgs
    End If

    'レコードセットのクローズ
    rsRecordSet.Close


    On Error GoTo Err_mgs
    '「ログファイル」リストボックスをクリアする
    LstFile.Clear

    'ログファイル情報を編集する
    For i = 0 To iCnt - 1
        sEntry = Left(gLogData(i).sName, 12)
        If Len(gLogData(i).sStTime) = 19 Then
            sEntry = sEntry & "  " & gLogData(i).sStTime
        Else
            sEntry = sEntry & "                     "
        End If

        If Len(gLogData(i).sEdTime) = 19 Then
            sEntry = sEntry & "  " & gLogData(i).sEdTime
        Else
            sEntry = sEntry & "                     "
        End If

        sEntry = sEntry & "  " & Format(gLogData(i).iSize, "@@@@@@@@")
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
            '「ID中継ユニットログ管理画面：DBアクセス異常」
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, IDU_LOG_KANRI_DB_ACCESS_ERROR, 0)
        Case 2
            '「ID中継ユニットログ管理画面：DBアクセス異常」
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, IDU_LOG_KANRI_DB_ACCESS_ERROR, 0)
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2005 All Right Reserved
'/
'/  関数名称  : fLogSearchCheck
'/  概要     : ログ検索データチェック
'/  説明     : ログ検索データの正当性をチェックする
'/  ﾊﾟﾗﾒｰﾀ   :
'/           :
'/
'/  ORIGINAL  ：(1.0.0.1) 2005-01-27  CODED BY  [TCC] T.Yashiro
'//     REVISIONS :(EG20 3.6.0.1) 2012-02-23   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ２ 残件回収
'/  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'/  備考：
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
   
    On Error Resume Next

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
                    '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
'                    txtStFun.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtStNen.Text)) = 0 And _
                   Len(Trim(txtStTuki.Text)) = 0 And _
                   Len(Trim(txtStHi.Text)) = 0 And _
                   Len(Trim(txtStZi.Text)) = 0 Then
            
                    txtStFun.Text = "00"
            
                Else
                    iRet = MsgBox("表示範囲の開始に未入力の項目があります。", vbExclamation, "入力異常")
                    '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
'                    txtEdZi.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtEdNen.Text)) = 0 And _
                   Len(Trim(txtEdTuki.Text)) = 0 And _
                   Len(Trim(txtEdHi.Text)) = 0 Then

                    txtEdZi.Text = "00"
                Else
                    iRet = MsgBox("表示範囲の終了に未入力の項目があります。", vbExclamation, "入力異常")
                    '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
                    '「ID中継ユニットログ管理：入力時刻異常」ログ出力
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
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
        For i = 0 To iIcmCnt
            'チェックがＯＮなら処理する
            If chkCorner(uIcmFileData(i).iIndex).Value = 1 Then
                'フラグを立てる
                bFlg = True
            End If
        Next
        
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
'//  機能名称  : ログテキストファイル書き込み処理
'//  機能概要  : ログファイルをテキストファイルに書き込む
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
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_APP
    Else
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_HOSHU
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
    iStatus = dllCreateDispLogFile(lErr, sFileName, uLogConv, sObjectTopFile, PATH_IDU_APP)
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
'//  機能概要  : ログトレース画面から、ログ変換情報を作成する。
'//　　　　　　　表示ファイル指定部：「ログ表示(テキスト表示)」釦押下時処理
'//
'//              型        　　　　　　名称      意味
'//  引数      : VB_LOG_DISP_SETTING　uLogConv　[OUT]ログ変換情報
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
    Dim bGokiFlg As Boolean
   
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

        '全件チェック
         bGokiFlg = False
        For i = 0 To iIcmCnt
            'チェックがＯＮなら処理する
            If chkCorner(uIcmFileData(i).iIndex).Value = 1 Then
                If uIcmFileData(i).iRonri = 32 Then
                    'フラグを立てる
                    bGokiFlg = True
                Else
                    'ビットカウント計算
                    iBitCnt = 1
                    If uIcmFileData(i).iRonri <> 1 Then
                        For ii = 1 To uIcmFileData(i).iRonri - 1
                            iBitCnt = iBitCnt * 2
                        Next
                    End If
                    '変数に追加する
                    uLogConv.Goki = uLogConv.Goki + iBitCnt
                End If
            End If
        Next

        If bGokiFlg = True Then
            uLogConv.Goki = -2147483648# + uLogConv.Goki
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
             AppActivate frmIDULogkanri.Caption, False      ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
             pfFormActive (frmIDULogkanri.hwnd)
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
        AppActivate frmIDULogkanri.Caption, False
        pfFormActive (frmIDULogkanri.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
