VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTakuLogKanri 
   BorderStyle     =   0  'なし
   Caption         =   "監視盤ログ管理"
   ClientHeight    =   9000
   ClientLeft      =   2445
   ClientTop       =   1395
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
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
   Begin VB.CommandButton cmdInstall 
      Caption         =   "媒体取外"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      TabIndex        =   36
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "  ログ表示    (テキスト表示)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   9720
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  ログ管理     画面へ戻る"
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
      Left            =   9720
      TabIndex        =   1
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Timer tmrMail 
      Left            =   9600
      Top             =   6840
   End
   Begin TabDlg.SSTab tabLog 
      Height          =   8620
      Left            =   0
      TabIndex        =   0
      Top             =   380
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   15214
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "表示ファイル指定"
      TabPicture(0)   =   "ログ管理(操作卓)画面.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdLogShushu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdLzhFileWrite"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLog(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdUpdateDisplay"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tabTakuCorner"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "表示項目指定"
      TabPicture(1)   =   "ログ管理(操作卓)画面.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraKoumoku(2)"
      Tab(1).Control(1)=   "fraKoumoku(1)"
      Tab(1).Control(2)=   "fraKoumoku(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "表示号機指定"
      TabPicture(2)   =   "ログ管理(操作卓)画面.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraGouki"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdLogShushu 
         Caption         =   "ログ収集"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -68400
         TabIndex        =   242
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CommandButton cmdLzhFileWrite 
         Caption         =   "  ログ圧縮    媒体出力"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -68400
         TabIndex        =   241
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ログ媒体出力"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   -68400
         TabIndex        =   240
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdUpdateDisplay 
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
         Height          =   975
         Left            =   -68400
         TabIndex        =   239
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Frame fraKoumoku 
         Caption         =   "分類"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6135
         Index           =   2
         Left            =   -74895
         TabIndex        =   29
         Top             =   2400
         Width           =   9135
         Begin VB.Frame fraLogBunnrui 
            Caption         =   "指定分類"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   5800
            Left            =   2400
            TabIndex        =   38
            Top             =   240
            Width           =   6615
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
               Left            =   4470
               TabIndex        =   100
               Top             =   5400
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
               Left            =   4470
               TabIndex        =   99
               Top             =   5160
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
               Left            =   4470
               TabIndex        =   98
               Top             =   4920
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
               Left            =   4470
               TabIndex        =   97
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
               Index           =   55
               Left            =   4470
               TabIndex        =   96
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
               Index           =   54
               Left            =   4470
               TabIndex        =   95
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
               Index           =   53
               Left            =   4470
               TabIndex        =   94
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
               Index           =   52
               Left            =   4470
               TabIndex        =   93
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
               Index           =   51
               Left            =   4470
               TabIndex        =   92
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
               Index           =   50
               Left            =   4470
               TabIndex        =   91
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
               Index           =   49
               Left            =   4470
               TabIndex        =   90
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
               Index           =   48
               Left            =   4470
               TabIndex        =   89
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
               Index           =   47
               Left            =   4470
               TabIndex        =   88
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
               Index           =   46
               Left            =   4470
               TabIndex        =   87
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
               Index           =   45
               Left            =   4470
               TabIndex        =   86
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
               Index           =   44
               Left            =   4470
               TabIndex        =   85
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
               Index           =   43
               Left            =   4470
               TabIndex        =   84
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
               Index           =   42
               Left            =   4470
               TabIndex        =   83
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
               Index           =   41
               Left            =   4470
               TabIndex        =   82
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
               Index           =   40
               Left            =   4470
               TabIndex        =   81
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
               Index           =   39
               Left            =   2295
               TabIndex        =   80
               Top             =   5400
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
               Left            =   2295
               TabIndex        =   79
               Top             =   5160
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
               Left            =   2295
               TabIndex        =   78
               Top             =   4920
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
               Left            =   2295
               TabIndex        =   77
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
               Index           =   35
               Left            =   2295
               TabIndex        =   76
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
               Index           =   34
               Left            =   2295
               TabIndex        =   75
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
               Index           =   33
               Left            =   2295
               TabIndex        =   74
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
               Index           =   32
               Left            =   2295
               TabIndex        =   73
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
               Index           =   31
               Left            =   2295
               TabIndex        =   72
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
               Index           =   30
               Left            =   2295
               TabIndex        =   71
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
               Index           =   29
               Left            =   2295
               TabIndex        =   70
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
               Index           =   28
               Left            =   2295
               TabIndex        =   69
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
               Index           =   27
               Left            =   2295
               TabIndex        =   68
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
               Index           =   26
               Left            =   2295
               TabIndex        =   67
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
               Index           =   25
               Left            =   2295
               TabIndex        =   66
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
               Index           =   24
               Left            =   2295
               TabIndex        =   65
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
               Index           =   23
               Left            =   2295
               TabIndex        =   64
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
               Index           =   22
               Left            =   2295
               TabIndex        =   63
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
               Index           =   21
               Left            =   2295
               TabIndex        =   62
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
               Index           =   20
               Left            =   2295
               TabIndex        =   61
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
               Index           =   19
               Left            =   120
               TabIndex        =   60
               Top             =   5400
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
               Left            =   120
               TabIndex        =   59
               Top             =   5160
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
               Left            =   120
               TabIndex        =   58
               Top             =   4920
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
               Left            =   120
               TabIndex        =   57
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
               Index           =   15
               Left            =   120
               TabIndex        =   56
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
               Index           =   14
               Left            =   120
               TabIndex        =   55
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
               Index           =   13
               Left            =   120
               TabIndex        =   54
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
               Index           =   12
               Left            =   120
               TabIndex        =   53
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
               Index           =   11
               Left            =   120
               TabIndex        =   52
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
               Index           =   10
               Left            =   120
               TabIndex        =   51
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
               Index           =   9
               Left            =   120
               TabIndex        =   50
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
               Index           =   8
               Left            =   120
               TabIndex        =   49
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
               Index           =   7
               Left            =   120
               TabIndex        =   48
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
               Index           =   6
               Left            =   120
               TabIndex        =   47
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
               Index           =   5
               Left            =   120
               TabIndex        =   46
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
               Index           =   4
               Left            =   120
               TabIndex        =   45
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
               Index           =   3
               Left            =   120
               TabIndex        =   44
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
               Index           =   2
               Left            =   120
               TabIndex        =   43
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
               Index           =   1
               Left            =   120
               TabIndex        =   42
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
               Index           =   0
               Left            =   120
               TabIndex        =   41
               Top             =   840
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.OptionButton optAll 
               Caption         =   "全て未選択"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   1
               Left            =   1920
               TabIndex        =   40
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton optAll 
               Caption         =   "全て選択"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   0
               Left            =   480
               TabIndex        =   39
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00FFFFFF&
               X1              =   0
               X2              =   6550
               Y1              =   720
               Y2              =   720
            End
         End
         Begin VB.OptionButton optLogBunrui 
            Caption         =   "指定分類のみ表示"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Value           =   -1  'True
            Width           =   2275
         End
         Begin VB.OptionButton optLogBunrui 
            Caption         =   "全ての分類を表示"
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
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   2275
         End
         Begin VB.Frame fraLogData 
            Caption         =   "表示行"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2295
            Left            =   120
            TabIndex        =   30
            Top             =   1800
            Width           =   2175
            Begin VB.OptionButton optLogData 
               Caption         =   "１行目のみ表示"
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
               Index           =   1
               Left            =   120
               TabIndex        =   32
               Top             =   1560
               Value           =   -1  'True
               Width           =   2000
            End
            Begin VB.OptionButton optLogData 
               Caption         =   "全行表示"
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
               Index           =   0
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label lblLogData 
               Caption         =   "1ｲﾍﾞﾝﾄが複数行のとき"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   33
               Top             =   480
               Width           =   1695
            End
         End
      End
      Begin VB.Frame fraKoumoku 
         Caption         =   "種別"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Index           =   1
         Left            =   -70800
         TabIndex        =   20
         Top             =   480
         Width           =   4935
         Begin VB.OptionButton optLogSyu 
            Caption         =   "指定種別のみ表示"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optLogSyu 
            Caption         =   "全ての種別を表示"
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
            Index           =   0
            Left            =   2640
            TabIndex        =   27
            Top             =   360
            Width           =   2250
         End
         Begin VB.Frame fraLogSyu 
            Caption         =   "指定種別"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1095
            Index           =   0
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   4095
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "正常"
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
               Index           =   0
               Left            =   480
               TabIndex        =   26
               Top             =   240
               Value           =   1  'ﾁｪｯｸ
               Width           =   1095
            End
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "異常"
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
               Index           =   1
               Left            =   1560
               TabIndex        =   25
               Top             =   240
               Value           =   1  'ﾁｪｯｸ
               Width           =   1095
            End
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "警告"
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
               Index           =   2
               Left            =   2640
               TabIndex        =   24
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox chkLogSyu 
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
               Index           =   3
               Left            =   480
               TabIndex        =   23
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "デバッグ"
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
               Index           =   4
               Left            =   1560
               TabIndex        =   22
               Top             =   600
               Width           =   1300
            End
         End
      End
      Begin VB.Frame fraKoumoku 
         Caption         =   "時刻"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Index           =   0
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   3855
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   10
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   9
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   840
            MaxLength       =   2
            TabIndex        =   8
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   7
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   6
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   840
            MaxLength       =   2
            TabIndex        =   5
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblLogTime 
            Caption         =   "まで"
            BeginProperty Font 
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
            Left            =   3000
            TabIndex        =   19
            Top             =   1260
            Width           =   615
         End
         Begin VB.Label lblLogTime 
            Caption         =   "分"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   2680
            TabIndex        =   18
            Top             =   1260
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "時"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1960
            TabIndex        =   17
            Top             =   1260
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1220
            TabIndex        =   16
            Top             =   1260
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "から"
            BeginProperty Font 
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
            Left            =   3000
            TabIndex        =   15
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblLogTime 
            Caption         =   "分"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2680
            TabIndex        =   14
            Top             =   900
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "時"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1960
            TabIndex        =   13
            Top             =   900
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1220
            TabIndex        =   12
            Top             =   900
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "ログデータ対象時刻"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame fraGouki 
         Caption         =   "自改号機"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   7455
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   8895
         Begin VB.CommandButton cmdZSentaku 
            Caption         =   "  全コーナ    全号機 選択"
            Enabled         =   0   'False
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
            Left            =   240
            TabIndex        =   201
            Top             =   480
            Width           =   2000
         End
         Begin VB.CommandButton cmdZHisentaku 
            Caption         =   "  全コーナ    全号機 非選択"
            Enabled         =   0   'False
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
            Left            =   2400
            TabIndex        =   200
            Top             =   480
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
            Left            =   4560
            TabIndex        =   199
            Top             =   480
            Width           =   2000
         End
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
            Left            =   6720
            TabIndex        =   198
            Top             =   480
            Width           =   2000
         End
         Begin TabDlg.SSTab tabCorner 
            Height          =   2535
            Left            =   120
            TabIndex        =   101
            Top             =   1440
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
            TabPicture(0)   =   "ログ管理(操作卓)画面.frx":0054
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "chkLogGouki(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "chkLogGouki(1)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "chkLogGouki(2)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "chkLogGouki(3)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "chkLogGouki(4)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "chkLogGouki(5)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "chkLogGouki(6)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "chkLogGouki(7)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "chkLogGouki(8)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "chkLogGouki(9)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "chkLogGouki(10)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "chkLogGouki(11)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "chkLogGouki(12)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "chkLogGouki(13)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "chkLogGouki(14)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "chkLogGouki(15)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).ControlCount=   16
            TabCaption(1)   =   "  "
            TabPicture(1)   =   "ログ管理(操作卓)画面.frx":0070
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "chkLogGouki(31)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "chkLogGouki(30)"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "chkLogGouki(29)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "chkLogGouki(28)"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "chkLogGouki(27)"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "chkLogGouki(26)"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "chkLogGouki(25)"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "chkLogGouki(24)"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "chkLogGouki(23)"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "chkLogGouki(22)"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "chkLogGouki(21)"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "chkLogGouki(20)"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "chkLogGouki(19)"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "chkLogGouki(18)"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "chkLogGouki(17)"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "chkLogGouki(16)"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).ControlCount=   16
            TabCaption(2)   =   "  "
            TabPicture(2)   =   "ログ管理(操作卓)画面.frx":008C
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "chkLogGouki(47)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "chkLogGouki(46)"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "chkLogGouki(45)"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "chkLogGouki(44)"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "chkLogGouki(43)"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "chkLogGouki(42)"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "chkLogGouki(41)"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).Control(7)=   "chkLogGouki(40)"
            Tab(2).Control(7).Enabled=   0   'False
            Tab(2).Control(8)=   "chkLogGouki(39)"
            Tab(2).Control(8).Enabled=   0   'False
            Tab(2).Control(9)=   "chkLogGouki(38)"
            Tab(2).Control(9).Enabled=   0   'False
            Tab(2).Control(10)=   "chkLogGouki(37)"
            Tab(2).Control(10).Enabled=   0   'False
            Tab(2).Control(11)=   "chkLogGouki(36)"
            Tab(2).Control(11).Enabled=   0   'False
            Tab(2).Control(12)=   "chkLogGouki(35)"
            Tab(2).Control(12).Enabled=   0   'False
            Tab(2).Control(13)=   "chkLogGouki(34)"
            Tab(2).Control(13).Enabled=   0   'False
            Tab(2).Control(14)=   "chkLogGouki(33)"
            Tab(2).Control(14).Enabled=   0   'False
            Tab(2).Control(15)=   "chkLogGouki(32)"
            Tab(2).Control(15).Enabled=   0   'False
            Tab(2).ControlCount=   16
            TabCaption(3)   =   "  "
            TabPicture(3)   =   "ログ管理(操作卓)画面.frx":00A8
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "chkLogGouki(63)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).Control(1)=   "chkLogGouki(62)"
            Tab(3).Control(1).Enabled=   0   'False
            Tab(3).Control(2)=   "chkLogGouki(61)"
            Tab(3).Control(2).Enabled=   0   'False
            Tab(3).Control(3)=   "chkLogGouki(60)"
            Tab(3).Control(3).Enabled=   0   'False
            Tab(3).Control(4)=   "chkLogGouki(59)"
            Tab(3).Control(4).Enabled=   0   'False
            Tab(3).Control(5)=   "chkLogGouki(58)"
            Tab(3).Control(5).Enabled=   0   'False
            Tab(3).Control(6)=   "chkLogGouki(57)"
            Tab(3).Control(6).Enabled=   0   'False
            Tab(3).Control(7)=   "chkLogGouki(56)"
            Tab(3).Control(7).Enabled=   0   'False
            Tab(3).Control(8)=   "chkLogGouki(55)"
            Tab(3).Control(8).Enabled=   0   'False
            Tab(3).Control(9)=   "chkLogGouki(54)"
            Tab(3).Control(9).Enabled=   0   'False
            Tab(3).Control(10)=   "chkLogGouki(53)"
            Tab(3).Control(10).Enabled=   0   'False
            Tab(3).Control(11)=   "chkLogGouki(52)"
            Tab(3).Control(11).Enabled=   0   'False
            Tab(3).Control(12)=   "chkLogGouki(51)"
            Tab(3).Control(12).Enabled=   0   'False
            Tab(3).Control(13)=   "chkLogGouki(50)"
            Tab(3).Control(13).Enabled=   0   'False
            Tab(3).Control(14)=   "chkLogGouki(49)"
            Tab(3).Control(14).Enabled=   0   'False
            Tab(3).Control(15)=   "chkLogGouki(48)"
            Tab(3).Control(15).Enabled=   0   'False
            Tab(3).ControlCount=   16
            TabCaption(4)   =   "  "
            TabPicture(4)   =   "ログ管理(操作卓)画面.frx":00C4
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "chkLogGouki(79)"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "chkLogGouki(78)"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).Control(2)=   "chkLogGouki(77)"
            Tab(4).Control(2).Enabled=   0   'False
            Tab(4).Control(3)=   "chkLogGouki(76)"
            Tab(4).Control(3).Enabled=   0   'False
            Tab(4).Control(4)=   "chkLogGouki(75)"
            Tab(4).Control(4).Enabled=   0   'False
            Tab(4).Control(5)=   "chkLogGouki(74)"
            Tab(4).Control(5).Enabled=   0   'False
            Tab(4).Control(6)=   "chkLogGouki(73)"
            Tab(4).Control(6).Enabled=   0   'False
            Tab(4).Control(7)=   "chkLogGouki(72)"
            Tab(4).Control(7).Enabled=   0   'False
            Tab(4).Control(8)=   "chkLogGouki(71)"
            Tab(4).Control(8).Enabled=   0   'False
            Tab(4).Control(9)=   "chkLogGouki(70)"
            Tab(4).Control(9).Enabled=   0   'False
            Tab(4).Control(10)=   "chkLogGouki(69)"
            Tab(4).Control(10).Enabled=   0   'False
            Tab(4).Control(11)=   "chkLogGouki(68)"
            Tab(4).Control(11).Enabled=   0   'False
            Tab(4).Control(12)=   "chkLogGouki(67)"
            Tab(4).Control(12).Enabled=   0   'False
            Tab(4).Control(13)=   "chkLogGouki(66)"
            Tab(4).Control(13).Enabled=   0   'False
            Tab(4).Control(14)=   "chkLogGouki(65)"
            Tab(4).Control(14).Enabled=   0   'False
            Tab(4).Control(15)=   "chkLogGouki(64)"
            Tab(4).Control(15).Enabled=   0   'False
            Tab(4).ControlCount=   16
            TabCaption(5)   =   "  "
            TabPicture(5)   =   "ログ管理(操作卓)画面.frx":00E0
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "chkLogGouki(95)"
            Tab(5).Control(0).Enabled=   0   'False
            Tab(5).Control(1)=   "chkLogGouki(94)"
            Tab(5).Control(1).Enabled=   0   'False
            Tab(5).Control(2)=   "chkLogGouki(93)"
            Tab(5).Control(2).Enabled=   0   'False
            Tab(5).Control(3)=   "chkLogGouki(92)"
            Tab(5).Control(3).Enabled=   0   'False
            Tab(5).Control(4)=   "chkLogGouki(91)"
            Tab(5).Control(4).Enabled=   0   'False
            Tab(5).Control(5)=   "chkLogGouki(90)"
            Tab(5).Control(5).Enabled=   0   'False
            Tab(5).Control(6)=   "chkLogGouki(89)"
            Tab(5).Control(6).Enabled=   0   'False
            Tab(5).Control(7)=   "chkLogGouki(88)"
            Tab(5).Control(7).Enabled=   0   'False
            Tab(5).Control(8)=   "chkLogGouki(87)"
            Tab(5).Control(8).Enabled=   0   'False
            Tab(5).Control(9)=   "chkLogGouki(86)"
            Tab(5).Control(9).Enabled=   0   'False
            Tab(5).Control(10)=   "chkLogGouki(85)"
            Tab(5).Control(10).Enabled=   0   'False
            Tab(5).Control(11)=   "chkLogGouki(84)"
            Tab(5).Control(11).Enabled=   0   'False
            Tab(5).Control(12)=   "chkLogGouki(83)"
            Tab(5).Control(12).Enabled=   0   'False
            Tab(5).Control(13)=   "chkLogGouki(82)"
            Tab(5).Control(13).Enabled=   0   'False
            Tab(5).Control(14)=   "chkLogGouki(81)"
            Tab(5).Control(14).Enabled=   0   'False
            Tab(5).Control(15)=   "chkLogGouki(80)"
            Tab(5).Control(15).Enabled=   0   'False
            Tab(5).ControlCount=   16
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   197
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   196
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   195
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   194
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   193
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   192
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   191
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   190
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   189
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   188
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   187
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   186
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   185
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   184
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   183
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   182
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   181
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   180
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   179
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   178
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   177
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   176
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   175
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   174
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   173
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   172
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   171
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   170
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   169
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   168
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   166
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   165
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   164
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   163
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   162
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   161
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   160
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   159
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   158
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   157
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   156
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   155
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   154
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   153
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   152
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   151
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   150
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   149
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   148
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   147
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   146
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   145
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   144
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   143
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   142
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   141
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   140
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   139
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   138
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   137
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   136
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   135
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   134
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   133
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   132
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   131
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   130
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   129
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   128
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   127
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   126
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   125
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   124
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   123
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   122
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   121
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   120
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   119
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   118
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   117
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   116
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   115
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   114
               Top             =   2040
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   113
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   112
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   111
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   110
               Top             =   1560
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   109
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   108
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   107
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   106
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   105
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   104
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   103
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
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
               TabIndex        =   102
               Top             =   600
               Value           =   1  'ﾁｪｯｸ
               Visible         =   0   'False
               Width           =   1815
            End
         End
      End
      Begin TabDlg.SSTab tabTakuCorner 
         Height          =   7815
         Left            =   -74640
         TabIndex        =   202
         Top             =   600
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   13785
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
         TabCaption(0)   =   " ○○○○○○ ○○○○○○"
         TabPicture(0)   =   "ログ管理(操作卓)画面.frx":00FC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraLogFile(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   " ○○○○○○ ○○○○○○"
         TabPicture(1)   =   "ログ管理(操作卓)画面.frx":0118
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(1)"
         Tab(1).Control(1)=   "fraLogFile(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   " ○○○○○○ ○○○○○○"
         TabPicture(2)   =   "ログ管理(操作卓)画面.frx":0134
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(2)"
         Tab(2).Control(1)=   "fraLogFile(2)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   " ○○○○○○ ○○○○○○"
         TabPicture(3)   =   "ログ管理(操作卓)画面.frx":0150
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1(3)"
         Tab(3).Control(1)=   "fraLogFile(3)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   " ○○○○○○ ○○○○○○"
         TabPicture(4)   =   "ログ管理(操作卓)画面.frx":016C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame1(4)"
         Tab(4).Control(1)=   "fraLogFile(4)"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   " ○○○○○○ ○○○○○○"
         TabPicture(5)   =   "ログ管理(操作卓)画面.frx":0188
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame1(5)"
         Tab(5).Control(1)=   "fraLogFile(5)"
         Tab(5).ControlCount=   2
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'なし
            Caption         =   "Frame1"
            Height          =   975
            Index           =   5
            Left            =   -74760
            TabIndex        =   258
            Top             =   600
            Width           =   5775
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
               Index           =   5
               Left            =   120
               TabIndex        =   260
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "画面操作ログ"
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
               Index           =   5
               Left            =   120
               TabIndex        =   259
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'なし
            Caption         =   "Frame1"
            Height          =   975
            Index           =   4
            Left            =   -74760
            TabIndex        =   255
            Top             =   600
            Width           =   5775
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
               Index           =   4
               Left            =   120
               TabIndex        =   257
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "画面操作ログ"
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
               Index           =   4
               Left            =   120
               TabIndex        =   256
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'なし
            Caption         =   "Frame1"
            Height          =   975
            Index           =   3
            Left            =   -74760
            TabIndex        =   252
            Top             =   600
            Width           =   5775
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
               Index           =   3
               Left            =   120
               TabIndex        =   254
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "画面操作ログ"
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
               Index           =   3
               Left            =   120
               TabIndex        =   253
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'なし
            Caption         =   "Frame1"
            Height          =   975
            Index           =   2
            Left            =   -74760
            TabIndex        =   249
            Top             =   600
            Width           =   5775
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
               Index           =   2
               Left            =   120
               TabIndex        =   251
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "画面操作ログ"
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
               Index           =   2
               Left            =   120
               TabIndex        =   250
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'なし
            Caption         =   "Frame1"
            Height          =   975
            Index           =   1
            Left            =   -74760
            TabIndex        =   246
            Top             =   600
            Width           =   5775
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
               Index           =   1
               Left            =   120
               TabIndex        =   248
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "画面操作ログ"
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
               Index           =   1
               Left            =   120
               TabIndex        =   247
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'なし
            Caption         =   "Frame1"
            Height          =   975
            Index           =   0
            Left            =   240
            TabIndex        =   243
            Top             =   600
            Width           =   5775
            Begin VB.OptionButton optHoshu 
               Caption         =   "画面操作ログ"
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
               Index           =   0
               Left            =   120
               TabIndex        =   245
               Top             =   480
               Width           =   2655
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
               Index           =   0
               Left            =   120
               TabIndex        =   244
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "監視盤ログファイル"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   0
            Left            =   120
            TabIndex        =   233
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   0
               Left            =   120
               MultiSelect     =   2  '拡張
               TabIndex        =   234
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "ファイル名"
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
               Index           =   0
               Left            =   120
               TabIndex        =   238
               Top             =   360
               Width           =   1680
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "最終書込年月日"
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
               Index           =   1
               Left            =   1800
               TabIndex        =   237
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   " 時：分"
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
               Index           =   2
               Left            =   3530
               TabIndex        =   236
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "サイズ "
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
               Index           =   3
               Left            =   4500
               TabIndex        =   235
               Top             =   360
               Width           =   1005
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "監視盤ログファイル"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   1
            Left            =   -74880
            TabIndex        =   227
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   1
               Left            =   120
               MultiSelect     =   2  '拡張
               TabIndex        =   228
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "サイズ "
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
               Index           =   4
               Left            =   4500
               TabIndex        =   232
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   " 時：分"
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
               Index           =   5
               Left            =   3530
               TabIndex        =   231
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "最終書込年月日"
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
               Index           =   6
               Left            =   1800
               TabIndex        =   230
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "ファイル名"
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
               Index           =   7
               Left            =   120
               TabIndex        =   229
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "監視盤ログファイル"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   2
            Left            =   -74880
            TabIndex        =   221
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   2
               Left            =   120
               MultiSelect     =   2  '拡張
               TabIndex        =   222
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "サイズ "
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
               Index           =   8
               Left            =   4500
               TabIndex        =   226
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   " 時：分"
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
               Index           =   9
               Left            =   3530
               TabIndex        =   225
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "最終書込年月日"
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
               Index           =   10
               Left            =   1800
               TabIndex        =   224
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "ファイル名"
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
               Index           =   11
               Left            =   120
               TabIndex        =   223
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "監視盤ログファイル"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   3
            Left            =   -74880
            TabIndex        =   215
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   3
               Left            =   120
               MultiSelect     =   2  '拡張
               TabIndex        =   216
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "サイズ "
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
               Index           =   12
               Left            =   4500
               TabIndex        =   220
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   " 時：分"
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
               Index           =   13
               Left            =   3530
               TabIndex        =   219
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "最終書込年月日"
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
               Index           =   14
               Left            =   1800
               TabIndex        =   218
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "ファイル名"
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
               Index           =   15
               Left            =   120
               TabIndex        =   217
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "監視盤ログファイル"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   4
            Left            =   -74880
            TabIndex        =   209
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   4
               Left            =   120
               MultiSelect     =   2  '拡張
               TabIndex        =   210
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "サイズ "
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
               Index           =   16
               Left            =   4500
               TabIndex        =   214
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   " 時：分"
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
               Index           =   17
               Left            =   3530
               TabIndex        =   213
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "最終書込年月日"
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
               Index           =   18
               Left            =   1800
               TabIndex        =   212
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "ファイル名"
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
               Index           =   19
               Left            =   120
               TabIndex        =   211
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "監視盤ログファイル"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   5
            Left            =   -74880
            TabIndex        =   203
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   5
               Left            =   120
               MultiSelect     =   2  '拡張
               TabIndex        =   204
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "サイズ "
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
               Index           =   20
               Left            =   4500
               TabIndex        =   208
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   " 時：分"
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
               Index           =   21
               Left            =   3530
               TabIndex        =   207
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "最終書込年月日"
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
               Index           =   22
               Left            =   1800
               TabIndex        =   206
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '中央揃え
               BorderStyle     =   1  '実線
               Caption         =   "ファイル名"
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
               Index           =   23
               Left            =   120
               TabIndex        =   205
               Top             =   360
               Width           =   1680
            End
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "操作卓ログ管理"
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
      TabIndex        =   37
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmTakuLogKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmTakuLogKanri.frm
'//  パッケージ名：操作卓ログ管理画面
'//
'//  概要：操作卓ログ管理画面
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//                 【フェーズ２対応】操作卓ログ管理画面を流用して新規作成
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-13   CODED   BY [TCC] M.Matsumoto
'//                 【統-350対応】
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.115修正対応】
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 【圧縮フォルダ指定対応】
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20V5.10.0.1) 2012-05-09 REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、フォルダを作成する
'//     REVISIONS :(X.X.X.X) 0000-00-00   CODED   BY [ ]
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'*****************************************************************************
'*      定数
'*****************************************************************************
Private Const MN_COLOR_BLACK = &H80000008
Private Const MN_COLOR_RED = &HFF&
Private Const MN_COLOR_WHITE = &H80000005
Private Const MN_COLOR_YELLOW = &HFFFF&

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

'*****************************************************************************
'*      ログ情報格納エリア
'*****************************************************************************
Private Type LogFileData
    sPath As String                 'ログファイルのパス
    sName As String                 'ログファイル名
    dtFileDate As Date              '作成日付・時刻
    lFileSize As Long               'ファイルサイズ
    bSelect As Boolean              '選択フラグ
End Type

Private uLogfileData() As LogFileData
'*****************************************************************************
'*      対象ファイルフルパス（複数ﾌｧｲﾙの時、ｽﾍﾟｰｽ1文字で区切る。）
'*****************************************************************************
Private sObjectFiles As String   'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽで選択中のﾌｧｲﾙのﾌﾙﾊﾟｽ文字列
Private sObjectTopFile As String '同上、選択中の先頭（最旧）ﾌｧｲﾙ名。

'*****************************************************************************
'*      イベントログコピー用ワークファイル名フルパス
'*****************************************************************************
Private Const SAVEFILE_SYS As String = PATH_WORK & "SysEvent.Evt"
Private Const SAVEFILE_SEC As String = PATH_WORK & "ScuEvent.Evt"
Private Const SAVEFILE_APP As String = PATH_WORK & "AppEvent.Evt"

'圧縮ファイル用
Private Type files
    sFileName(255) As String
End Type

'EG20 V3.6.0.1【03統合TR-No.115修正対応】削除開始
''EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'Private Const CAB_LOG_FILE As String = PATH_WORK & "KANSI_LOG_TMP.CAB"
'Private Const DAT_LOG_FILE As String = PATH_WORK & "KANSI_LOG_TMP.DAT"
''EG20 V2.1.0.1 ADD END   【フェーズ２対応】
'EG20 V3.6.0.1【03統合TR-No.115修正対応】削除終了
'EG20 V3.6.0.1【03統合TR-No.115修正対応】追加開始
Private Const CAB_LOG_FILE As String = PATH_WORK & "KLOGTEMP.CAB"
Private Const DAT_LOG_FILE As String = PATH_WORK & "KLOGTEMP.DAT"
'EG20 V3.6.0.1【03統合TR-No.115修正対応】追加終了

'*****************************************************************************
'*      モジュール情報格納エリア
'*****************************************************************************
Private Type ModFileData
    sName As String             'プロセス名
    iProces As Integer          'プロセスID
    iFuzokuId As Integer        '付属プロセスID
    iFuzokuCnt As Integer       '付属カウンタ
End Type
Private uModFileData(59) As ModFileData
Private iModCnt As Integer

Private Const ASRT_LOG = &H200         ' 10:ログ･トレース
Private Const ASRT_HOSYU = &H400       ' 11:保守画面設定
Private Const ASRT_SYUKEI = &H800      ' 12:集計                'REV(03.00)行追加。
Private Const ASRT_ALL = &H7FFFFFFF    '全分類ログ収集

'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
Private Const PATH_LOG_CORNER = "E:\\KANSI\\CORNER"
Private Const DIR_LOG_APL = "\\OPERATE_APL_LOG\\"
Private Const DIR_LOG_SOUSA = "\\OPERATE_SOUSA_LOG\\"

Private mintStatus(31) As Integer
'EG20 V2.1.0.1 ADD END


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdInstall_Click
'//  機能名称  : 「媒体取外」釦押下時処理
'//  機能概要  : 媒体の取外しを行う
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

Private Sub cmdLogShushu_Click()

    Dim intTabIdx As Integer
    Dim iResponse As Integer
    
    intTabIdx = tabTakuCorner.Tab

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SAVE, 0)
    
    '確認メッセージボックスを表示する。
    iResponse = MsgBox("操作卓ログデータを収集しますがよろしいですか？", _
                        vbOKCancel, "収集")

    '「キャンセル」ボタン押下処理は処理を終了する
    If iResponse = vbCancel Then Exit Sub
    
    'アプリログ／画面操作ログ種別を設定
    If optApp(intTabIdx).Value = True Then
        glnglogKind = LOG_COL_KIND.LOG_APP
    Else
        glnglogKind = LOG_COL_KIND.LOG_SOUSA
    End If
    '対象コーナを設定
    glngTargetCorner = intTabIdx + 1
    
    '処理中画面を表示する
    dlgLogShushuMessage.Show vbModal

    'ログ一覧再表示
    Call sSetListBox(intTabIdx)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 操作卓ログ管理画面(アクティブ時)
'//  機能概要  : メール受信用のタイマ起動
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
    Dim bRet As Boolean                 '戻り値
    Dim lId As Long                     'メールＩＤ
    Dim bFlag As Boolean                'フラグ
    Dim lngErrCode As Long              'エラーコード
    Dim udtMail As ML_KYOTU_INF         'バッファフラッシュ要求

    On Error Resume Next
    
    tmrMail.Enabled = True
            
   'バッファフラッシュ要求をログプロセスに送信する
    udtMail.udtlHeader.dwId = ML_ID_LGBUFF_REQ
    udtMail.udtlHeader.dwSize = MlSize.BUFF_FLUSH_CMD
    udtMail.udtlHeader.dwProid = RHOSHU_ID
    udtMail.udtlHeader.dwSubArea = 0
    bRet = DssSendMail(MAIL_SLOT_LOG, MlSize.BUFF_FLUSH_CMD, udtMail.udtlHeader)
    If bRet = False Then
       '「バッファフラッシュ要求送信異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, lngErrCode)
       Exit Sub
    Else
       '「バッファフラッシュ要求送信正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, 0)
    End If
  
    'バッファフラッシュ終了通知受信
    bFlag = False
    Do Until bFlag = True
        'メール受信処理を行う
        lId = fMailRecieve()
        Select Case lId         'メールＩＤ
        '「プロセス終了指示」の場合
        Case ML_ID_PROEND_ORD
             '「プロセス終了指示受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            '処理を終了する
            Exit Sub
        '「バッファフラッシュ終了通知」の場合
        Case ML_ID_LGBUFF_ANS
            '「バッファフラッシュ終了通知受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
            'ループを抜ける
            Exit Do
        Case Else
        End Select
        Sleep (MN_MAIL_INTERVAL)
    Loop
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 操作卓ログ管理画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ起動
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
   
   If blnCabfrmOpenFlg = True Then
        Call fnTsbCabCallDiverge
        Exit Sub
    End If
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 操作卓ログ管理画面(ロード時)
'//  機能概要  : メール受信用のタイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//                 【フェーズ２対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim iRet As Integer             '関数の戻り値
    Dim sKeyName As String          'INIファイルキー名
    Dim iMozi As Integer            '読み込み文字数
    Dim iKbn As Integer             '読み込んだ文字数
    Dim sIni_Data As String * 128   'INIファイルより1行分取得
    Dim iCnt As Integer             'INIファイルカウンタ
    Dim i As Integer                'カウンタ
    Dim j As Integer                'コントロール配列数
    Dim MyName As String            'INI有無チェック
    'EG20 V2.1.0.1 ADD START 【フェーズ２】
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intIndex As Integer
    Dim strCorner1 As String
    Dim strCorner2 As String
    Dim bySyoAssort As Byte             'ログ用小分類
    'EG20 V2.1.0.1 ADD END

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'EG20 V2.1.0.1 ADD START 【フェーズ２】
    '号機情報取得
    Call gsGetGateInfo
    Call gsGetCornerName
    
    'タブ数を設置コーナ数とする
    tabCorner.Tab = 0
    
    '収集状態初期化
    Erase mintStatus
    
    '内部ファイルエラーのトラップ
    On Error GoTo OtherError
    bySyoAssort = L3AN_FILE
    
    For intCount = 0 To UBound(gblnCornerSet)
    
        '設定ありのコーナを活性にする
        If gblnCornerSet(intCount) = True Then
            'コーナー名称表示
            strCorner1 = MidB(gstrCornerName(intCount), 1, 12)
            strCorner2 = MidB(gstrCornerName(intCount), 13, 24)
'            tabCorner.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
            tabCorner.TabCaption(intCount) = Empty
            tabTakuCorner.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
        End If
    
    Next intCount
    
    '設置コーナ数分ループ
    For intCount = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(intCount) = False Then
            tabCorner.TabVisible(intCount) = False
            tabTakuCorner.TabVisible(intCount) = False
            optApp(intCount).Value = True
        End If

        '最大号機数分ループ
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + intCount2
            chkLogGouki(intIndex).Visible = False
            chkLogGouki(intIndex).Tag = "0"
        Next
        
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + (gudtSettiCorner(intCount).intGokiNo(intCount2) - 1)
            If gudtSettiCorner(intCount).intGokiNo(intCount2) > 0 Then
                chkLogGouki(intIndex).Caption = gudtSettiCorner(intCount).strDispGoki(intCount2) + "号機"
                'Tagに対応する号機番号を記録（1～32号機）
                chkLogGouki(intIndex).Tag = CStr(gudtSettiCorner(intCount).intGateNo(intCount2))
                mintStatus(gudtSettiCorner(intCount).intGateNo(intCount2) - 1) = CHECKBOX_ON
                chkLogGouki(intIndex).Visible = True
                chkLogGouki(intIndex).Value = CHECKBOX_ON
            End If
        Next intCount2
        
    Next intCount
    'EG20 V2.1.0.1 ADD END
    
    '表示ファイル指定を登録する
'    sSetListBox            'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    For intCount = 0 To 5
        Call sSetListBox(intCount)
    Next
    'EG20 V2.1.0.1 ADD END
    
    'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
     
    'ファイル有無チェック
'    MyName = Dir(DISP_FILE, vbNormal)          'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    MyName = Dir(DISP_FILE_TAKU, vbNormal)      'EG20 V2.1.0.1 ADD 【フェーズ２対応】
    If MyName = "" Then
        GoTo FileError
    End If
    
    For iCnt = 0 To 59
        sKeyName = DISP_KEY_NAME & Format(iCnt, "00")
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'        iRet = GetPrivateProfileString(DISP_SECTION_NAME, _
'                                       sKeyName, _
'                                       DEFAILT, sIni_Data, Len(sIni_Data), _
'                                       DISP_FILE)
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        iRet = GetPrivateProfileString(DISP_SECTION_NAME, _
                                       sKeyName, _
                                       DEFAILT, sIni_Data, Len(sIni_Data), _
                                       DISP_FILE_TAKU)
        'EG20 V2.1.0.1 ADD END
        iMozi = 1
        iKbn = 1
        Do
           'モジュール情報格納エリアに1行分のデータを保持させる。
            If Mid(sIni_Data, iMozi, 1) = "," Then
                Select Case iKbn
                    Case 1
                        uModFileData(iCnt).sName = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 2
                        uModFileData(iCnt).iProces = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 3
                        uModFileData(iCnt).iFuzokuId = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 4
                        uModFileData(iCnt).iFuzokuCnt = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                End Select
            End If
            iMozi = iMozi + 1
            If iMozi > Len(sIni_Data) Then
                Exit Do
            End If
        Loop
        
        '1行分データの保持処理後、表示処理を行う。
        If iKbn = 4 Then
            chkMod(iCnt).Visible = True
            chkMod(iCnt).Caption = uModFileData(iCnt).sName
           If uModFileData(iCnt).iFuzokuId = 0 Then
              Select Case iCnt
                Case 0 To 19
                  '大分類扱い：分類カウンター0～19の場合
                  chkMod(iCnt).Left = 120
                Case 20 To 39
                  '大分類扱い：分類カウンター20～39の場合
                  chkMod(iCnt).Left = 2295
                Case 40 To 59
                  '大分類扱い：分類カウンター40～59の場合
                  chkMod(iCnt).Left = 4470
              End Select
          Else
              Select Case iCnt
                Case 0 To 19
                '中分類扱い：分類カウンター0～19の場合
                  chkMod(iCnt).Left = 330
                Case 20 To 39
                '中分類扱い：分類カウンター20～39の場合
                  chkMod(iCnt).Left = 2500
                Case 40 To 59
                '大分類扱い：分類カウンター40～59の場合
                  chkMod(iCnt).Left = 4670
             End Select
          End If
            iModCnt = iCnt
        End If
    Next
          
   '表示項目指定を初期化する
    optLogSyu(0).Value = True               'ラジオ釦：「全ての種別を表示」を有効
    j = chkLogSyu.UBound
    For i = 0 To j                          '種別分繰り返す
        chkLogSyu(i).Value = CHECKBOX_ON    '「？？種別」を有効にする
    Next
    
    optLogBunrui(0).Value = True            'ラジオ釦：「全ての分類を表示」を有効
    For i = 0 To iModCnt                    '分類分繰り返す
        If chkMod(i).Visible = True Then
            chkMod(i).Value = CHECKBOX_ON   '「？？分類」を有効にする
        End If
    Next

    optLogData(1).Value = True             '「１行目のみ表示」を有効にする

    'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'    '表示自改号機指定を初期化する
'    optLogGouki(0).Value = True            'ラジオ釦：「全自改」を有効
'    cmdChkAll.Enabled = False
'    cmdChkAllKai.Enabled = False
'
'    j = chkLogGouki.UBound
'    For i = 0 To j                         '号機分繰り返す
'        chkLogGouki(i).Value = CHECKBOX_ON '「？？号機」を有効にする
'        chkLogGouki(i).Enabled = False     '全号機押下不可
'    Next
    'EG20 V2.1.0.1 DEL END
   
   tabLog.Tab = 0
   tabTakuCorner.Tab = 0        'EG20 V2.1.0.1 ADD 【フェーズ２対応】
   
   Call tabTakuCorner_Click(0)  'EG20 V2.1.0.1 ADD 【統-350対応】
   
   '「操作卓ログ管理画面：表示」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_LOG_TAKU_GAMEN_START, 0)
   
   Exit Sub

FileError:
   '「操作卓ログ管理：INIファイル異常」ログ出力
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
 
   'INIファイル有無チェック異常時：「ファイル異常」ポップアップを表示
   MsgBox "INIファイルの取得に失敗しました｡", vbCritical, "ファイル異常"
   Exit Sub
   
OtherError:
  '「操作卓ログ管理：ログ表示異常」ログ出力
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, KANSI_LOG_KANRI_LOG_ERROR, 0)
  'リストボックスの初期化
'   lstLogFile.Clear        'EG20 V2.1.0.1 DEL 【フェーズ２対応】
  'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    For i = 0 To lstLogFile.UBound
        lstLogFile(i).Clear
    Next
  'EG20 V2.1.0.1 ADD END
   MsgBox "ログ一覧の取得に失敗しました。", vbCritical, "表示異常"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSetListBox
'//  機能名称  : ログファイル登録処理
'//  機能概要  : ログファイルをリストボックスに登録する。
'//              表示ファイル指定部：初期処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sSetListBox()                      'EG20 V2.1.0.1 DEL 【フェーズ２対応】
Private Sub sSetListBox(intTabIdx As Integer)   'EG20 V2.1.0.1 ADD 【フェーズ２対応】
    Dim i As Integer            'カウンタ
    Dim j As Integer            'カウンタ
    Dim iCnt As Integer         'ログファイル数
    Dim sEntry As String        '編集文字列
    Dim uLogData As LogFileData 'バージョン情報バッファ

    On Error Resume Next
    
    'ログファイル情報を取得する
'    iCnt = fGetLogfileInf()            'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    iCnt = fGetLogfileInf(intTabIdx)    'EG20 V2.1.0.1 ADD 【フェーズ２対応】

    'ファイル名でソートする
    For i = 0 To iCnt - 2
        For j = i + 1 To iCnt - 1
            'If uLogfileData(i).sName > uLogfileData(j).sName Then              'V1.7.0.1 DEL
            If UCase(uLogfileData(i).sName) > UCase(uLogfileData(j).sName) Then 'V1.7.0.1 ADD
                uLogData = uLogfileData(i)
                uLogfileData(i) = uLogfileData(j)
                uLogfileData(j) = uLogData
            End If
        Next j
    Next i

    '「ログファイル」リストボックスをクリアする
'    lstLogFile.Clear               'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    lstLogFile(intTabIdx).Clear     'EG20 V2.1.0.1 ADD 【フェーズ２対応】

    'ログファイル情報を編集する
    For i = 0 To iCnt - 1       'ログファイル数分繰り返す
        sEntry = Mid$(uLogfileData(i).sName & Space(14), 1, 14)
        sEntry = sEntry & "    " & Format(uLogfileData(i).dtFileDate, "yyyy/mm/dd  hh:nn")
        sEntry = sEntry & Format(uLogfileData(i).lFileSize, "@@@@@@@@@")
'        lstLogFile.AddItem sEntry       'リストボックスに追加する              'EG20 V2.1.0.1 DEL 【フェーズ２対応】
        lstLogFile(intTabIdx).AddItem sEntry       'リストボックスに追加する    'EG20 V2.1.0.1 ADD 【フェーズ２対応】
    Next
    If iCnt > 0 Then                    'ログファイルが存在する
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'        lstLogFile.ListIndex = 0        '一行目にインデックスをセット
'        lstLogFile.Selected(0) = True   '一行目を選択済にする
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
        lstLogFile(intTabIdx).ListIndex = 0        '一行目にインデックスをセット
        lstLogFile(intTabIdx).Selected(0) = True   '一行目を選択済にする
        'EG20 V2.1.0.1 DEL END
    End If

End Sub

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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZHisentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkLogGouki.UBound
        chkLogGouki(intLoop).Value = CHECKBOX_OFF
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZSentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkLogGouki.UBound
        chkLogGouki(intLoop).Value = CHECKBOX_ON
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
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
        chkLogGouki(intLoop).Value = CHECKBOX_OFF
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
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
        chkLogGouki(intLoop).Value = CHECKBOX_ON
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : chkLogGouki_Click
'//  機能名称  : 指定号機オプションボタンクリック時処理
'//  機能概要  : 内部変数のON/OFFを切り替える
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]オプションボタンインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-22   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub chkLogGouki_Click(Index As Integer)

    Dim intGoki As Integer
    
    On Error Resume Next
    
    intGoki = CInt(chkLogGouki(Index).Tag) - 1
    
    mintStatus(intGoki) = chkLogGouki(Index).Value
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : optApp_Click
'//  機能名称  : ログ区分オプションボタンクリック時処理
'//  機能概要  : ログの種類を切り替える
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]オプションボタンインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub optApp_Click(Index As Integer)

    Call sSetListBox(Index)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : optHoshu_Click
'//  機能名称  : ログ区分オプションボタンクリック時処理
'//  機能概要  : ログの種類を切り替える
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]オプションボタンインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub optHoshu_Click(Index As Integer)

    Call sSetListBox(Index)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdUpdateDisplay_Click
'//  機能名称  : 「表示更新」釦押下時処理
'//  機能概要  : ログファイルの表示リストの内容を最新状態に更新する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdUpdateDisplay_Click()

    Dim intTabIdx As Integer
    
    intTabIdx = tabTakuCorner.Tab

    Call sSetListBox(intTabIdx)
    
End Sub
'EG20 V2.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fGetLogfileInf
'//  機能名称  : ログファイル情報取得処理
'//  機能概要  : 全ログファイルの情報を取得する。
'//              表示ファイル指定部：ログファイル登録処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Integer　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Public Function fGetLogfileInf() As Integer                    'EG20 V2.1.0.1 DEL 【フェーズ２対応】
Public Function fGetLogfileInf(intIndex As Integer) As Integer  'EG20 V2.1.0.1 ADD 【フェーズ２対応】
    Dim MyPath As String       'フォルダ名
    Dim MyName As String       'ファイル名
    Dim iLogfileCnt As Integer 'カウンター
    Dim bSelectLogSousa As Boolean                      ' 操作ログ選択状態（TRUE:選択）     ' EG20 V3.0.0.2追加
    Dim bFileOK As Boolean                              ' ファイル検索結果                  ' EG20 V3.0.0.2追加

    On Error Resume Next
    
    'ログファイル数を初期化する
    iLogfileCnt = 0
    
    '保守画面操作ログファイルを検索する。
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    MyPath = PATH_LOG_CORNER & CStr(intIndex + 1)
    If optHoshu(intIndex).Value = True Then
        MyPath = MyPath & DIR_LOG_SOUSA
        bSelectLogSousa = True                          ' 操作ログ選択状態（選択）          ' EG20 V3.0.0.2追加
    Else
        MyPath = MyPath & DIR_LOG_APL                              ' パスを設定します。
        bSelectLogSousa = False                         ' 操作ログ選択状態（非選択）        ' EG20 V3.0.0.2追加
    End If      'EG20 V2.1.0.1 ADD 【フェーズ２対応】
    
    MyName = Dir(MyPath & HOSHULOG_FILE, vbNormal) ' 最初のディレクトリ名を返します。
    If MyName <> "" Then
      iLogfileCnt = iLogfileCnt + 1
      ReDim Preserve uLogfileData(iLogfileCnt)
      'ログファイル情報を格納する
      uLogfileData(iLogfileCnt - 1).sPath = MyPath
      uLogfileData(iLogfileCnt - 1).sName = HOSHULOG_FILE
      uLogfileData(iLogfileCnt - 1).dtFileDate = FileDateTime(MyPath & HOSHULOG_FILE)
      uLogfileData(iLogfileCnt - 1).lFileSize = FileLen(MyPath & HOSHULOG_FILE)
      uLogfileData(iLogfileCnt - 1).bSelect = False
    End If
    
    'ログトレースファイルを検索する。
'    MyPath = PATH_LOG                           ' パスを設定します。   'EG20 V2.1.0.1 DEL
'    MyName = Dir(MyPath & "L*.DAT", vbNormal)   ' 最初のディレクトリ名を返します。     'EG20 V2.1.0.1 DEL 【統-331対応】
'    MyName = Dir(MyPath & "L*.*", vbNormal)   ' 最初のディレクトリ名を返します。        'EG20 V2.1.0.1 ADD 【統-331対応】  EG20 V3.0.0.2 DEK
' EG20 V3.0.0.2追加開始
    ' 検索すべきファイル名を変更する。
    If bSelectLogSousa = True Then
        MyName = Dir(MyPath & "*.TXT", vbNormal)        ' 最初のディレクトリ名を返します。
    Else
        MyName = Dir(MyPath & "L*.*", vbNormal)         ' 最初のディレクトリ名を返します。
    End If
' EG20 V3.0.0.2追加終了
    Do While MyName <> ""                       ' ループを開始します。
        ' 現在のディレクトリと親ディレクトリは無視します。
        If MyName <> "." And MyName <> ".." Then
' EG20 V3.0.0.2追加開始
            ' オプションに応じて検索すべきファイルの条件を絞りこみする
            bFileOK = False
            If bSelectLogSousa = True Then
                ' 操作ログについては現状無条件
                bFileOK = True
            Else
                ' アプリケーションログ
                If Right(MyName, 3) = "IDU" Or Right(MyName, 3) = "DAT" Then        'EG20 V2.1.0.1 ADD 【統-331対応】
                    bFileOK = True
                End If
            End If
' EG20 V3.0.0.2追加終了
            If bFileOK = True Then
                ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
                If (GetAttr(MyPath & MyName) And vbDirectory) <> vbDirectory Then
                    iLogfileCnt = iLogfileCnt + 1
                    ReDim Preserve uLogfileData(iLogfileCnt)
    
                    'ログファイル情報を格納する
                    uLogfileData(iLogfileCnt - 1).sPath = MyPath
                    uLogfileData(iLogfileCnt - 1).sName = MyName
                    uLogfileData(iLogfileCnt - 1).dtFileDate = FileDateTime(MyPath & MyName)
                    uLogfileData(iLogfileCnt - 1).lFileSize = FileLen(MyPath & MyName)
                    uLogfileData(iLogfileCnt - 1).bSelect = False
    
                End If                      ' それを表示します。
            End If
        End If          'EG20 V2.1.0.1 ADD 【統-331対応】
        ' 次のディレクトリ名を返します。
        MyName = Dir
    Loop
    fGetLogfileInf = iLogfileCnt
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdLog_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称の処理を行う。
'//              「ログ表示(テキスト表示)」「ログ媒体出力」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ログファイル書込み先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20V5.10.0.1) 2012-05-09 REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、フォルダを作成する
'//     REVISIONS :(EG20V5.13.0.1) 2012-06-06 REVISED BY [TCC] H.Sugimoto
'//                 【媒体出力フォルダ作成対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdLog_Click(Index As Integer)
    Dim bRet As Boolean
    Dim lRetVal As Double
    Dim sCommand As String
    Dim sWriteDir As String    '書込みディレクトリ
    Dim iObjFileNo As Integer  '書込み対象ﾌｧｲﾙ数
    On Error GoTo ErrorHandle:
    Dim lngErrCode As Long     'エラーコード
    Dim fso As FileSystemObject     'ファイルシステムオブジェクト       ' EG20 V5.10.0.1【ログフォルダ作成対応】ADD
    Dim szDefLogFolder As String    ' 出力ログフォルダ                  ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加
    Dim szCornerFolder As String    ' コーナ                            ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加

    On Error Resume Next
 
 Select Case Index   'ボタンインデックス
   Case 0
     '「操作卓ログ管理画面：ログ表示釦押下」
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_TEXT_HYOUJI_BUTTOM, 0)
      'ログ検索データ正当姓チェック
      bRet = fLogSearchCheck
      If bRet = False Then    'ログ検索データにエラーがある場合
          Exit Sub            '処理を終了する
      End If

      'ログテキストファイルを書き込む
       bRet = fWriteLogtxt
       If bRet = True Then         'ログテキストファイルが正常に作成された場合
           sCommand = MN_EXE_MEMO & MN_LOG_FILE        '実行コマンドを作成する
           lRetVal = Shell(sCommand, vbMaximizedFocus) 'ノートパッドを起動する
           AppActivate lRetVal, True                   'アクティブ（前面表示）にする
           SendKeys "{LEFT}", True
          '「操作卓ログ管理画面：ログ表示処理正常」
           Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_OK, 0)
       Else
          lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
          '「操作卓ログ管理画面：ログ表示処理異常」
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_ERROR, lngErrCode)
       End If

    Case 1
       '「操作卓ログ管理画面：ログ媒体出力釦押下」
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_OUTPUT_BUTTOM, 0)

        'ログ検索データ正当姓チェック
        iObjFileNo = fLogSelectCheck
        If iObjFileNo <= 0 Then
            Exit Sub            '処理を終了する
' EG20 V5.9.0.1【ログ選択上限対応】ADD START
        ElseIf iObjFileNo > LOG_FILECNT_MAX Then
            ' 警告文言表示
            MsgBox "選択されたファイル数が上限を超えました。" _
                    & Chr(vbKeyReturn) & "選択できるファイル数は[" & LOG_FILECNT_MAX & "]件までです。", _
                    vbOKOnly + vbCritical, _
                    "ファイル指定異常"
            Exit Sub
' EG20 V5.9.0.1【ログ選択上限対応】ADD END
        End If
        ' 取出し先ディレクトリを選択する
'        sWriteDir = pfDirSelection("a:", "ログファイル書込み先のディレクトリ選択")     'V1.12.0.1 DEL
        'sWriteDir = pfDirSelection("H:", "ログファイル書込み先のディレクトリ選択")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
        sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
        'V1.5.0.1 DEL START
        'frmDir.Caption = "ログファイル書込み先のディレクトリ選択"
        'frmDir.Show 1
        'V1.5.0.1 DEL END
        If sWriteDir <> "" Then
        'ディレクトリが指定されれば、ログファイルを取出す

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
' EG20 V5.10.0.1【ログフォルダ作成対応】ADD START
            szCornerFolder = "OPERATE_LOG" & CStr(tabTakuCorner.Tab + 1)
'            If sWriteDir Like "*OPERATE_LOG\" = False Then                 ' EG20V5.13.0.1【媒体出力フォルダ作成対応】削除
            If sWriteDir Like ("*" & szCornerFolder & "\") = False Then     ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加
                ' フォルダが存在するかチェックする
'                sWriteDir = sWriteDir & "\" & "OPERATE_LOG"                ' EG20V5.13.0.1【媒体出力フォルダ作成対応】削除
                sWriteDir = sWriteDir & "\" & szCornerFolder                ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加
                Set fso = New FileSystemObject
                If fso.FolderExists(sWriteDir) = False Then
                    ' フォルダが存在しない場合は作成する
                    fso.CreateFolder (sWriteDir)
                End If
                Set fso = Nothing
            End If
' EG20 V5.10.0.1【ログフォルダ作成対応】ADD END
            sCopyLogFile sWriteDir, iObjFileNo
        End If
     Case Else
    
    End Select
ErrorHandle:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdLzhFileWrite_Click
'//  機能名称  : 「ログ圧縮媒体出力」釦押下時処理
'//  機能概要  : ログの圧縮媒体出力を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ログファイル圧縮書込み先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 「ログ圧縮媒体出力」ポップアップ画面を追加
'//                 フォルダ選択画面をOS仕様に変更
'//                  「ログ圧縮媒体出力」釦押下処理での保守ログ選択時ファイル名修正
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-24   REVISED BY [TCC] M.Matsumoto
'//                 【プリズミー統合-6対応】PASSLOG.TXTの出力に対応
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 【圧縮フォルダ指定対応】
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20V5.13.0.1) 2012-06-06 REVISED BY [TCC] H.Sugimoto
'//                 【媒体出力フォルダ作成対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdLzhFileWrite_Click()
    Dim sLzhDirName As String    '.LZHﾌｧｲﾙ格納ディレクトリ名
    Dim sLzhFileName As String   '.LZHﾌｧｲﾙ名
    Dim iObjFileNo As Integer    '圧縮対象ﾌｧｲﾙ数
    Dim nIndex As Integer        ' 文字数                    ' EG20 V5.6.0.1追加

    Dim fso As FileSystemObject     'ファイルシステムオブジェクト       ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加
    Dim szDefLogFolder As String    ' 出力ログフォルダ                  ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加
    Dim szCornerFolder As String    ' コーナ                            ' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加

    On Error Resume Next
    
    '「操作卓ログ管理画面：ログ圧縮媒体出力釦押下」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_LOG_LZH_OUTPUT_BUTTOM, 0)

    'リストボックスで、ファイルが指定されているかチェックする。
    iObjFileNo = fLogSelectCheck
    If iObjFileNo <= 0 Then       'ファイル指定されていなければ、処理終了
        Exit Sub
' EG20 V5.9.0.1【ログ選択上限対応】ADD START
    ElseIf iObjFileNo > LOG_FILECNT_MAX Then
        ' 警告文言表示
        MsgBox "選択されたファイル数が上限を超えました。" _
               & Chr(vbKeyReturn) & "選択できるファイル数は[" & LOG_FILECNT_MAX & "]件までです。", _
               vbOKOnly + vbCritical, _
               "ファイル指定異常"
        Exit Sub
' EG20 V5.9.0.1【ログ選択上限対応】ADD END
    End If
    
    'ディレクトリ選択画面を表示させ、圧縮ファイル格納ディレクトリ名を得る。（ﾃﾞﾌｫﾙﾄﾃﾞｨﾚｸﾄﾘ＝ＦＤ）
'    sLzhDirName = pfDirSelection("a:", "ログファイル圧縮書込み先のディレクトリ選択")   'V1.12.0.1 DEL
    'sLzhDirName = pfDirSelection("H:", "ログファイル圧縮書込み先のディレクトリ選択")    'V1.12.0.1 ADD  'V1.20.0.1 DEL
    sLzhDirName = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
    If sLzhDirName = "" Then Exit Sub  'ディレクトリが指定されなければ、処理終了
 
' EG20 V5.6.0.1【圧縮フォルダ指定対応】追加開始
    ' 出力フォルダに半角スペースが含まれている場合、圧縮で異常が発生してしまうため
    ' 圧縮前にチェックして異常を表示する。
    nIndex = InStr(sLzhDirName, " ")
    If nIndex <> 0 Then
        ' 警告ポップアップウィンドウを表示する。
        Call MsgBox(CABFOLDERSELECT_ERRORMESSAGE, vbCritical, CABFOLDERSELECT_ERRORTITLE)
        Exit Sub  'ディレクトリが指定されなければ、処理終了
    End If
' EG20 V5.6.0.1【圧縮フォルダ指定対応】追加終了

 'V1.20.0.1 ADD START
     '選択されたファイルがHOSHU_LOG.datか、L*.datかのチェックを行う。
     If sObjectTopFile = HOSHULOG_FILE Then
        sLzhFileName = Left$(sObjectTopFile, 9)
     'EG20 V5.4.0.1 ADD START 【プリズミー統合-6対応】
     '選択されたファイルがPASSLOG.txtの場合
     ElseIf sObjectTopFile = PASSLOG_FILE Then
        sLzhFileName = Left$(sObjectTopFile, 7)
     'EG20 V5.4.0.1 ADD END
     Else
 'V1.20.0.1 ADD END
        '１番目のファイル(拡張子を含まない８文字)を、.LZHファイル名用に取出す。
        sLzhFileName = Left$(sObjectTopFile, 8)
    
     End If  'V1.20.0.1 ADD
    
    '.LZHファイル名を完成する。
    If iObjFileNo >= 2 Then
        '複数選択なら、選択ファイル数を付加する。
        sLzhFileName = sLzhFileName & "." & CStr(iObjFileNo)
    End If
    
    '拡張子は、.CABである。
    sLzhFileName = sLzhFileName & ".CAB"

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

' EG20 V5.9.0.1【圧縮フォルダ数対応】追加開始
    ' 圧縮対象フォルダ（ワーク）へ選択したログをコピー
    If funcCopyFileTemporary(PATH_LOGOUTTMP, iObjFileNo, sObjectFiles) = False Then
        Call subDeleteFolder(PATH_LOGOUTTMP)
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        ' ログ圧縮媒体出力処理正常時：「ログ圧縮媒体出力」ポップアップを表示
        MsgBox "ログ圧縮媒体出力処理は異常終了しました。", _
                vbOKOnly + vbInformation, _
                "ログ圧縮媒体出力"
        Exit Sub
    End If
    sObjectFiles = PATH_LOGOUTTMP
' EG20 V5.9.0.1【圧縮フォルダ数対応】追加終了

' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加開始
    szDefLogFolder = fncCreateLogFolder()
    If sLzhDirName Like ("*" & szDefLogFolder & "\") = False Then
        ' フォルダが存在するかチェックする
        sLzhDirName = sLzhDirName & "\" & szDefLogFolder
        Set fso = New FileSystemObject
        If fso.FolderExists(sLzhDirName) = False Then
            ' フォルダが存在しない場合は作成する
            fso.CreateFolder (sLzhDirName)
        End If
        Set fso = Nothing
    End If
    
    szCornerFolder = "OPERATE_LOG" & CStr(tabTakuCorner.Tab + 1)
    If sLzhDirName Like ("*" & szCornerFolder & "\") = False Then
        ' フォルダが存在するかチェックする
        sLzhDirName = sLzhDirName & "\" & szCornerFolder
        Set fso = New FileSystemObject
        If fso.FolderExists(sLzhDirName) = False Then
            ' フォルダが存在しない場合は作成する
            fso.CreateFolder (sLzhDirName)
        End If
        Set fso = Nothing
        sLzhDirName = sLzhDirName & "\"
    End If
' EG20V5.13.0.1【媒体出力フォルダ作成対応】追加終了

    Call psCabReqest(CABREQEST.CAB_COMPRESSION, sLzhDirName & sLzhFileName, sObjectFiles)
    'V1.20.0.1 ADD START
    If (glngCabErrCd = 0) Then   '圧縮結果が正常(0)
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        ' ログ圧縮媒体出力処理正常時：「ログ圧縮媒体出力」ポップアップを表示
        MsgBox "ログ圧縮媒体出力処理は正常終了しました。", _
                vbOKOnly + vbInformation, _
                "ログ圧縮媒体出力"
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        Call subDeleteFolder(PATH_LOGOUTTMP)
        Exit Sub
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    End If
    'V1.20.0.1 ADD END

' EG20 V5.9.0.1【圧縮フォルダ数対応】追加開始
    Call subDeleteFolder(PATH_LOGOUTTMP)
' EG20 V5.9.0.1【圧縮フォルダ数対応】追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面に戻る」釦押下時処理
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
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next

    '「操作卓ログ管理画面：消去」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_LOG_TAKU_GAMEN_END, 0)
  
    '操作卓ログ管理画面を閉じる
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fLogSearchCheck
'//  機能名称  : ログ検索データチェック処理
'//  機能概要  : ログ検索データの正当性をチェックする。
'//　　　　　　　表示ファイル指定部：「ログ表示(テキスト表示)釦押下時
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//                 【フェーズ２対応】
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-26   CODED   BY [TCC] M.Matsumoto
'//                 【統合No55対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fLogSearchCheck() As Boolean
    Dim bRet As Boolean         '関数の戻り値
    Dim i As Integer            'カウンタ
    Dim j As Integer            'コントロール配列数
    Dim bFlag As Boolean        'フラグ
    Dim iSelectedLines As Integer 'リストボックスで選択中の行数

    On Error Resume Next
    
    fLogSearchCheck = False     '戻り値に初期値としてエラーをセット

    'リストボックスで選択中のﾌｧｲﾙのﾌﾙﾊﾟｽ文字列をsObjectFilesにセットする。選択中行数を得る。
    iSelectedLines = fSelectedFilesGet
    '表示ファイル指定のチェックを行う
    If iSelectedLines <= 0 Then
        '表示ファイル未選択時：「表示ファイル未選択」ポップアップを表示
        MsgBox "表示ファイルが選択されていません。" _
               & Chr(vbKeyReturn) & "選択してください。", _
               vbOKOnly + vbExclamation, _
               "操作卓ログ管理"
        Exit Function                   '処理を終了する
    ElseIf iSelectedLines >= 2 Then
        '複数ファイル選択時：「複数ファイル指定」ポップアップを表示
        MsgBox "複数ファイルが選択されています。" _
               & Chr(vbKeyReturn) & "一つだけ選択してください。", _
               vbOKOnly + vbExclamation, _
               "操作卓ログ管理"
        Exit Function                   '処理を終了する
    End If

    'ログデータ対象時刻の正当姓チェック
    bRet = fLogTimeCheck
    If bRet = False Then                'エラーがある時は処理を終了する。
        Exit Function
    End If

    '指定種別のチェックを行う
    If optLogSyu(1).Value = True Then   '指定種別を選択した時
        j = chkLogSyu.UBound
        bFlag = False
        For i = 0 To j                  '指定種別分繰り返す
            If chkLogSyu(i).Value = CHECKBOX_ON Then
                bFlag = True            '指定が一つでもあれば、チェック処理終了
                Exit For
            End If
        Next
        If bFlag = False Then
        '指定種別未選択時：「指定種別なし」ポップアップを表示
            MsgBox "指定種別が選択されていません。" _
                   & Chr(vbKeyReturn) & "選択してください。", _
                   vbOKOnly + vbExclamation, _
                   "操作卓ログ管理"
            Exit Function               '処理を終了する
        End If
    End If

    '指定分類のチェックを行う
'    If optLogBunrui(1).Value = True Then   '指定分類を選択した時       'EG20 V2.1.0.1 DEL 【フェーズ２】
    If optLogBunrui(1).Value = True And optApp(tabTakuCorner.Tab).Value = True Then   '指定分類を選択した時   'EG20 V2.1.0.1 ADD 【フェーズ２】
        bFlag = False
        For i = 0 To iModCnt             '指定分類分繰り返す
            If chkMod(i).Visible = True And _
               chkMod(i).Value = CHECKBOX_ON Then
                bFlag = True            '指定がひとつでもあれば、チェック処理終了
                Exit For
            End If
        Next
        If bFlag = False Then
        '指定分類未選択時：「指定分類なし」ポップアップを表示
            MsgBox "指定分類が選択されていません。" _
                   & Chr(vbKeyReturn) & "選択してください。", _
                   vbOKOnly + vbExclamation, _
                   "操作卓ログ管理"
            Exit Function               '処理を終了する
        End If
    End If

    '指定号機のチェックを行う
'    If optLogGouki(1).Value = True Then   '指定号機を選択した時    'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    If optApp(tabTakuCorner.Tab).Value = True Then    'アプリログを選択した時     'EG20 V2.1.0.1 ADD 【フェーズ２対応】
'        j = chkLogGouki.UBound     'EG20 V5.5.0.1 DEL 【統合No55対応】
        j = UBound(mintStatus)
        bFlag = False
        For i = 0 To j                 '指定号機分繰り返す
'            If chkLogGouki(i).Value = CHECKBOX_ON Then             'EG20 V2.1.0.1 DEL 【フェーズ２対応】
            'EG20 V5.5.0.1 DEL START 【統合No55対応】
'            If chkLogGouki(i).Visible = True And chkLogGouki(i).Value = CHECKBOX_ON Then    'EG20 V2.1.0.1 ADD 【フェーズ２対応】
            'EG20 V5.5.0.1 DEL END
            If mintStatus(i) = CHECKBOX_ON Then         'EG20 V5.5.0.1 ADD 【統合No55対応】
                bFlag = True            '指定が一つでもある場合、チェック処理終了
                Exit For
            End If
        Next
        If bFlag = False Then
        '指定号機未選択時：「指定号機なし」ポップアップ表示
            MsgBox "指定号機が選択されていません。" _
                   & Chr(vbKeyReturn) & "選択してください。", _
                   vbOKOnly + vbExclamation, _
                   "操作卓ログ管理"
            Exit Function               '処理を終了する
        End If
    End If

    fLogSearchCheck = True              '戻り値に正常をセット
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fSelectedFilesGet
'//  機能名称  : 選択ファイル取得処理
'//  機能概要  : 選択中のファイルのフルパスを取得する。
'//　　　　　　　表示ファイル指定部：ログ検索データチェック処理
'//　　　　　　　　　　　　　　　　　「ログ媒体出力」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Integer　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSelectedFilesGet() As Integer
    Dim iLine As Integer         'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行ｲﾝﾃﾞｯｸｽ
    Dim iMaxLine As Integer      'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数
    Dim sLineFile As String      'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽ指定行のﾌｧｲﾙ名
    Dim iFileCounter As Integer  '対象ﾌｧｲﾙ数カウンタ
    
    sObjectFiles = ""
    'リストボックス表示中の全行について以下を実施する。
'    iMaxLine = lstLogFile.ListCount  'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数を得る。    'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    iMaxLine = lstLogFile(tabTakuCorner.Tab).ListCount  'ﾛｸﾞﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数を得る。  'EG20 V2.1.0.1 ADD 【フェーズ２対応】
    
    iFileCounter = 0
    For iLine = 0 To iMaxLine - 1
'        If lstLogFile.Selected(iLine) = True Then      'EG20 V2.1.0.1 DEL 【フェーズ２対応】
        If lstLogFile(tabTakuCorner.Tab).Selected(iLine) = True Then    'EG20 V2.1.0.1 ADD 【フェーズ２対応】
        '選択された行ならば、該当行のファイル名をリストボックスから得る。
            'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'            sLineFile = Left$(lstLogFile.List(iLine), _
'                              InStr(lstLogFile.List(iLine), " ") - 1)
            'EG20 V2.1.0.1 DEL END
            'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
            sLineFile = Left$(lstLogFile(tabTakuCorner.Tab).List(iLine), _
                              InStr(lstLogFile(tabTakuCorner.Tab).List(iLine), " ") - 1)
            'EG20 V2.1.0.1 DEL END
            '対象ﾌｧｲﾙとしてﾌﾙﾊﾟｽを作成し、文字列として保存する。
'            sObjectFiles = sObjectFiles & PATH_LOG & sLineFile & " "       'EG20 V2.1.0.1 DEL
            'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
            If optHoshu(tabTakuCorner.Tab).Value = True Then
                sObjectFiles = sObjectFiles & PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_SOUSA & sLineFile & " "
            Else
                sObjectFiles = sObjectFiles & PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_APL & sLineFile & " "
            End If
            'EG20 V2.1.0.1 ADD END
            If iFileCounter = 0 Then
            '選択行中の先頭（最旧）ﾌｧｲﾙであれば、ﾌｧｲﾙ名（拡張子を含む12文字）を保存する。
                sObjectTopFile = sLineFile
            End If
            iFileCounter = iFileCounter + 1
        End If
    Next
    '選択中ﾌｧｲﾙの数を返す。
    fSelectedFilesGet = iFileCounter
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fLogTimeCheck
'//  機能名称  : ログ対象時刻チェック処理
'//  機能概要  : ログ対象時刻の正当性チェックを行う。
'//　　　　　　　表示ファイル指定部：ログ検索データチェック処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fLogTimeCheck() As Boolean
    Dim i As Integer                'カウンタ
    Dim bFlag As Boolean            '入力フラグ
    Dim bFromFlag As Boolean        '入力フラグ(開始日時分)
    Dim bToFlag As Boolean          '入力フラグ(終了日時分)
    Dim iErrorIndex As Integer      'エラーのあるインデックス

    fLogTimeCheck = True
    
    '表示色を元に戻す
    For i = 0 To 5
        txtLogTime(i).ForeColor = MN_COLOR_BLACK
        txtLogTime(i).BackColor = MN_COLOR_WHITE
    Next

    '入力があるかチェックを行う
    bFlag = False                   '無効にする
    bFromFlag = False               '無効にする
    bToFlag = False                 '無効にする
    For i = 0 To 5
        If Not IsNull(txtLogTime(i)) And txtLogTime(i) <> "" Then
            bFlag = True            '有効にする
            If i >= 0 And i <= 2 Then
                bFromFlag = True    '有効にする
            Else
                bToFlag = True      '有効にする
            End If
            Select Case i
            Case 0, 3
                If Int(txtLogTime(i)) < 1 Or Int(txtLogTime(i)) > 31 Then
                    iErrorIndex = i
                    GoTo ErrorHandle
                End If
            Case 1, 4
                If Int(txtLogTime(i)) < 0 Or Int(txtLogTime(i)) > 23 Then
                    iErrorIndex = i
                    GoTo ErrorHandle
                End If
            Case 2, 5
                If Int(txtLogTime(i)) < 0 Or Int(txtLogTime(i)) > 59 Then
                    iErrorIndex = i
                    GoTo ErrorHandle
                End If
            End Select
        End If
    Next
    If bFlag = False Then           '入力がひとつもない
        Exit Function               '処理を終了する
    End If

    '開始日時分のみのチェックを行う
    If bFromFlag = True And bToFlag = False Then
        If IsNull(txtLogTime(0)) Or txtLogTime(0) = "" Then
            iErrorIndex = 0
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(1)) Or txtLogTime(1) = "" Then
            iErrorIndex = 1
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(2)) Or txtLogTime(2) = "" Then
            iErrorIndex = 2
            GoTo ErrorHandle
        End If
    
    '終了日時分のみのチェックを行う
    ElseIf bFromFlag = False And bToFlag = True Then
        If IsNull(txtLogTime(3)) Or txtLogTime(3) = "" Then
            iErrorIndex = 3
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(4)) Or txtLogTime(4) = "" Then
            iErrorIndex = 4
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(5)) Or txtLogTime(5) = "" Then
            iErrorIndex = 5
            GoTo ErrorHandle
        End If
    
    '両方のチェックを行う
    Else
        If IsNull(txtLogTime(0)) Or txtLogTime(0) = "" Then
            iErrorIndex = 0
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(1)) Or txtLogTime(1) = "" Then
            iErrorIndex = 1
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(2)) Or txtLogTime(2) = "" Then
            iErrorIndex = 2
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(4)) Or txtLogTime(4) = "" Then
            iErrorIndex = 4
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(5)) Or txtLogTime(5) = "" Then
            iErrorIndex = 5
            GoTo ErrorHandle
        End If
        
        '終了日がスペースの時は開始日と同じにする
        If IsNull(txtLogTime(3)) Or txtLogTime(3) = "" Then
            txtLogTime(3) = txtLogTime(0)
        End If
        '日時分の比較を行う
        If CInt(txtLogTime(0)) > CInt(txtLogTime(3)) Then
            If CInt(txtLogTime(0)) < 20 _
            Or CInt(txtLogTime(3)) > 10 Then
                iErrorIndex = 0
                GoTo ErrorHandle
            End If
        ElseIf CInt(txtLogTime(0)) = CInt(txtLogTime(3)) Then
            If CInt(txtLogTime(1)) > CInt(txtLogTime(4)) Then
                iErrorIndex = 1
                GoTo ErrorHandle
            ElseIf CInt(txtLogTime(1)) = CInt(txtLogTime(4)) Then
                If CInt(txtLogTime(2)) > CInt(txtLogTime(5)) Then
                    iErrorIndex = 2
                    GoTo ErrorHandle
                End If
            End If
        End If
    End If
    Exit Function
    
ErrorHandle:
    tabLog.Tab = 1
    txtLogTime(iErrorIndex).SetFocus
    txtLogTime(iErrorIndex).BackColor = MN_COLOR_YELLOW
    fLogTimeCheck = False

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fLogSelectCheck
'//  機能名称  : ログファイル取出しチェック処理
'//  機能概要  : 取出しファイル正当性チェックを行う。
'//　　　　　　　表示ファイル指定部：「ログ媒体出力」「ログ圧縮媒体出力」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Integer　　　　　　[OUT]選択中ファイル数
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Function fLogSelectCheck() As Integer
    Dim bRet As Boolean                 '戻り値
    Dim bFlag As Boolean                'フラグ
    Dim lId As Long                     'メールＩＤ
    Dim udtMail As ML_KYOTU_INF         'バッファフラッシュ要求
    Dim lngErrCode As Long              'エラーコード
    
    On Error Resume Next
    
    'リストボックスで選択中のﾌｧｲﾙのﾌﾙﾊﾟｽ文字列をsObjectFilesにセットする。選択中行数を得る。
    fLogSelectCheck = fSelectedFilesGet
    If fLogSelectCheck <= 0 Then
    'ファイル未選択時：「ファイル指定なし」ポップアップを表示
        MsgBox "取出しファイルが選択されていません。" _
               & Chr(vbKeyReturn) & "選択してください。", _
               vbOKOnly + vbExclamation, _
               "操作卓ログ管理"
        Exit Function                   '処理を終了する
    End If

    ' 現在書き込み中のファイル（一番新しいファイル）は対象外とする
'    If lstLogFile.Selected(lstLogFile.ListCount - 1) = True Then       'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    If lstLogFile(tabTakuCorner.Tab).Selected(lstLogFile(tabTakuCorner.Tab).ListCount - 1) = True Then 'EG20 V2.1.0.1 ADD 【フェーズ２対応】
         'バッファフラッシュ要求をログプロセスに送信する
          udtMail.udtlHeader.dwId = ML_ID_LGBUFF_REQ
          udtMail.udtlHeader.dwSize = MlSize.BUFF_FLUSH_CMD
          udtMail.udtlHeader.dwProid = RHOSHU_ID
          udtMail.udtlHeader.dwSubArea = 0
          bRet = DssSendMail(MAIL_SLOT_LOG, MlSize.BUFF_FLUSH_CMD, udtMail.udtlHeader)
          If bRet = False Then
            '「バッファフラッシュ要求送信異常」ログ出力
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
            Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, lngErrCode)
            Exit Function
          Else
            '「バッファフラッシュ要求送信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, 0)
          End If
        
        'バッファフラッシュ終了通知受信
        bFlag = False
        Do Until bFlag = True
            'メール受信処理を行う
            lId = fMailRecieve()
            Select Case lId         'メールＩＤ
                Case ML_ID_PROEND_ORD
                    '「プロセス終了指示」の場合
                    '「プロセス終了指示受信正常」ログ出力
                    Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                    '強制終了処理を行う
                    pfAbortProc
                Case ML_ID_LGBUFF_ANS
                    '「バッファフラッシュ終了」の場合
                    '「バッファフラッシュ終了通知受信正常」ログ出力
                    Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
                    'ループを抜ける
                    Exit Do
                Case Else
            End Select
            Sleep (MN_MAIL_INTERVAL)
        Loop
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sCopyLogFile
'//  機能名称  : ログファイル取出し処理
'//  機能概要  : ログファイルの取出しを行う。
'//　　　　　　　表示ファイル指定部：「ログ媒体出力」
'//
'//              型        名称      意味
'//  引数      : String　　sCopyDir  [IN]書込み先ディレクトリ
'//  　　      : Integer　 iFileNo   [IN]書込み先ファイル数
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yoshimori
'//                 「ログ媒体出力」ポップアップ画面を追加
'//                 「ログ媒体出力」でのエラーメッセージ表示
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub sCopyLogFile(sCopyDir As String, iFileNo As Integer)
    Dim sFileName As String
    Dim sCopyFileName As String
    Dim iResponse As Integer        'MsgBoxボタンコード
    Dim lSts As Long
    Dim iFile As Integer            'ファイル数カウンタ
    Dim iIti As Integer             '選択中ﾌｧｲﾙﾌﾙﾊﾟｽ文字列(sObjectFiles)内の文字位置
    Dim iNext As Integer            '同上、次の文字位置
    Dim lngErrCode As Long
    'V1.8.0.1 ADD START
    Dim slogPath    As String
    Dim sGetLogFile As String
    Dim bRet        As Boolean
    'V1.8.0.1 ADD END
        
On Error GoTo COPY_ERROR
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'リストボックスで選択中の全てのファイルについて、以下を実施する。
    iIti = 1
    For iFile = 0 To iFileNo - 1
        iNext = InStr(iIti, sObjectFiles, " ")  '１行ずつファイルへ書込む。
        If iNext = 0 Then Exit For
        'コピー元ファイル名フルパス（ﾛｸﾞﾄﾚｰｽﾌｧｲﾙ）をセットする。
        sFileName = Mid$(sObjectFiles, iIti, iNext - iIti)
        iIti = iNext + 1
        '書込み先ディレクトリ＋ファイル（コピー元と同じ）名をセットする。
        'sCopyFileName = sCopyDir & "\" & Right$(sFileName, 12) 'V1.8.0.1 DEL
        'V1.8.0.1 ADD START
        'ファイルパスより、ファイル名(最大13バイト)のみを取得する。
        sGetLogFile = Right$(sFileName, 13)
        'L*.dat　or　HOSHU_LOG.datのチェックを行う。
        '判断基準は「\」の有無による。
        If Left$(sGetLogFile, 1) = "\" Then
          '「\」があるのは「L*.dat」のため、「\」を削除する。
           sGetLogFile = Right$(sFileName, 12)
        End If
        sCopyFileName = sCopyDir & "\" & sGetLogFile
        'V1.8.0.1 ADD END
        'ログトレースファイルを指定ファイルに書き出す。
        'FileCopy sFileName, sCopyFileName              'V1.8.0.1 DEL
        'V1.8.0.1 ADD　START
        lSts = CopyFile(sFileName, sCopyFileName, 0)
        If lSts = 0 Then
           GoTo COPY_ERROR
        End If
        'V1.8.0.1 ADD　END
    Next
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'V1.20.0.1 ADD START
    ' ログ媒体出力処理正常時：「ログ媒体出力」ポップアップを表示
    MsgBox "ログ媒体出力処理は正常終了しました。", _
           vbOKOnly + vbInformation, _
           "ログ媒体出力"
    'V1.20.0.1 ADD END
        
    '「操作卓ログ管理画面：ログ媒体出力処理正常」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_OUTPUT_OK, 0)
    
    Exit Sub

COPY_ERROR:
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'Select Case Err.Number        'V1.20.0.1 DEL
    Select Case Err.LastDllError   'V1.20.0.1 ADD
        'Case 61 ' コピー先空き容量不足時：「空き容量無し」ポップアップを表示  'V1.20.0.1 DEL
        Case 112 ' コピー先空き容量不足時：「空き容量無し」ポップアップを表示   'V1.20.0.1 ADD
            iResponse = MsgBox("受け側のドライブのディスクがいっぱいです。" _
               & Chr(vbKeyReturn) & "新しいディスクを挿入してください。", _
               vbOKOnly, _
               "ログ媒体出力")

        'Case 70 ' ライトプロテクト時：「書込み禁止」ポップアップを表示 'V1.20.0.1 DEL
         Case 19 ' ライトプロテクト時：「書込み禁止」ポップアップを表示 'V1.20.0.1 ADD
            lSts = CopyFile(sFileName, sCopyFileName, 0)
            If (lSts = 0) Then
                iResponse = MsgBox("ファイルを作成または置換できません。このディスクはライトプロテクトされてます。" _
                   & Chr(vbKeyReturn) & "ライトプロテクトを解除するか　別のディスクを使ってください。", _
                   vbOKOnly, _
                   "ログ媒体出力")
            End If

        'Case 71 ' ディスクを未挿入時：「媒体未挿入」ポップアップを表示 'V1.20.0.1 DEL
        Case 21, 3    ' ディスクを未挿入時：「媒体未挿入」ポップアップを表示 'V1.20.0.1 ADD
            iResponse = MsgBox("ドライブにディスクが入ってません。" _
               & Chr(vbKeyReturn) & "ディスクを挿入してからやり直してください。", _
               vbOKOnly, _
               "ログ媒体出力")
'V1.20.0.1 DEL START
'        Case 75 ' 権限なし／パス名間違い時：「フォルダ書込み不可」ポップアップを表示
'            iResponse = MsgBox("コピー先の空き容量が不足しています。" _
'               & Chr(vbKeyReturn) & "不要名ファイルを削除するか、ディスクを入れ替えてください ", _
'               vbOKOnly, _
'               "ログ媒体出力")
'V1.20.0.1 DEL END
        Case Else '上記以外時：「ファイル出力異常」ポップアップを表示
            iResponse = MsgBox("予期せぬエラーが発生しました。" _
               & Chr(vbKeyReturn) & "操作をやり直してください。", _
               vbOKOnly, _
               "ログ媒体出力")
    End Select
    
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「操作卓ログ管理画面：ログ媒体出力処理異常」
     Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_OUTPUT_ERROR, lngErrCode)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fWriteLogtxt
'//  機能名称  : ログテキストファイル書込み処理
'//  機能概要  : ログファイルをログテキストファイルに書き込む。
'//　　　　　　　表示項目指定部：「ログ表示(テキスト表示）」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.115修正対応】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Function fWriteLogtxt() As Boolean
    Dim uLogConv As LOGCONV             'ログ検索データ
    Dim bRet As Boolean                 '戻り値
    Dim sFileName As String
    Dim lId As Long                     'メールＩＤ
    Dim bFlag As Boolean                'フラグ
    Dim iResponse As Integer            'MsgBoxボタンコード
    Dim iStatus As Long
    Dim udtMail As ML_KYOTU_INF         'バッファフラッシュ要求
    Dim lngErrCode As Long              'エラーコード
    fWriteLogtxt = False

    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    Dim lngRet As Long                  '戻り値
    Dim iFilePathLen As Integer
    Dim iresult As Integer
    Dim iErrRet As Integer
    Dim sDatFileName As String
    Dim sSourceFileName As String
    Dim fso As New FileSystemObject

    iErrRet = 0
    'EG20 V2.1.0.1 ADD END   【フェーズ２対応】

    On Error Resume Next
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'ログ変換情報を作成する
    sGetSearchData uLogConv
   
' EG20 V3.0.0.2 削除開始（操作卓ログには不要）
'   'バッファフラッシュ要求をログプロセスに送信する
'    udtMail.udtlHeader.dwId = ML_ID_LGBUFF_REQ
'    udtMail.udtlHeader.dwSize = MlSize.BUFF_FLUSH_CMD
'    udtMail.udtlHeader.dwProid = RHOSHU_ID
'    udtMail.udtlHeader.dwSubArea = 0
'    bRet = DssSendMail(MAIL_SLOT_LOG, MlSize.BUFF_FLUSH_CMD, udtMail.udtlHeader)
'    If bRet = False Then
'       '「バッファフラッシュ要求送信異常」ログ出力
'       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, lngErrCode)
'    Else
'       '「バッファフラッシュ要求送信正常」ログ出力
'       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, 0)
'    End If
'
'   If bRet = True Then
'
'       'バッファフラッシュ終了通知受信
'       bFlag = False
'       Do Until bFlag = True
'          'メール受信処理を行う
'          lId = fMailRecieve()
'          Select Case lId         'メールＩＤ
'            Case ML_ID_PROEND_ORD
'              '「プロセス終了指示」の場合
'              '「プロセス終了指示受信正常」ログ出力
'               Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
'              '処理を終了する
'              Exit Function
'            Case ML_ID_LGBUFF_ANS
'              '「バッファフラッシュ終了」の場合
'              '「バッファフラッシュ終了通知受信正常」ログ出力
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
'              'ループを抜ける
'              Exit Do
'            Case Else
'            End Select
'          Sleep (MN_MAIL_INTERVAL)
'         Loop
'    End If
' EG20 V3.0.0.2 削除終了（操作卓ログには不要）

    'ログテキストの作成
'    sFileName = PATH_LOG & sObjectTopFile      'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    If optHoshu(tabTakuCorner.Tab).Value = True Then
        sFileName = PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_SOUSA & sObjectTopFile
        
' EG20 V3.0.0.2 追加開始（操作ログはそのまま表示）
        ' 選択されたファイルをそのままコピー
        If fso.FileExists(sFileName) = True Then
            'ファイルコピー（既に存在した場合は上書きするする）
            fso.CopyFile sFileName, MN_LOG_FILE, True
            fWriteLogtxt = True
        Else
            fWriteLogtxt = False
        End If
        Set fso = Nothing
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        Exit Function
' EG20 V3.0.0.2 追加終了（操作ログはそのまま表示）
    Else
        sFileName = PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_APL & sObjectTopFile
    End If
    'EG20 V2.1.0.1 ADD END
    
        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    iFilePathLen = Len(sFileName)
    
    iresult = 0
    
    iresult = InStr(sFileName, "IDU")
    
    If iFilePathLen = ((iresult - 1) + 3) Then
    
' EG20 V3.6.0.1【03統合TR-No.115修正対応】削除開始
'        'IDUファイル → CABファイル変換
'        bRet = dllCreateDispLogFile2(lngErrCode, sFileName, CAB_LOG_FILE)
'
'        'CABファイル変換正常？
'        If bRet = True Then
'
'            'CABファイル解凍
'             lngRet = pfCabKaito(CAB_LOG_FILE, PATH_WORK)
'
'             If lngRet = 0 Then
'
'                sDatFileName = Replace(sObjectTopFile, "IDU", "DAT")
'                sSourceFileName = PATH_WORK & sDatFileName
'
'                fso.DeleteFile (DAT_LOG_FILE)
'                Name sSourceFileName As DAT_LOG_FILE
'
'                sFileName = DAT_LOG_FILE
'
'             Else
'                fWriteLogtxt = False
'                iErrRet = 1
'             End If
'        Else
'            fWriteLogtxt = False
'            iErrRet = 1
'        End If
' EG20 V3.6.0.1【03統合TR-No.115修正対応】削除終了
' EG20 V3.6.0.1【03統合TR-No.115修正対応】追加開始
        'IDUファイル → DATファイル変換
        bRet = dllCreateDispLogFile2(lngErrCode, sFileName, CAB_LOG_FILE, PATH_WORK)
        ' DATファイル変換正常？
        If bRet <> True Then
            fWriteLogtxt = False
            iErrRet = 1
        End If
        sFileName = DAT_LOG_FILE
' EG20 V3.6.0.1【03統合TR-No.115修正対応】追加終了
    
    End If
    'EG20 V2.1.0.1 ADD END   【フェーズ２対応】
   
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    If iErrRet = 0 Then
    'EG20 V2.1.0.1 ADD END   【フェーズ２対応】
   
        iStatus = dllbLog2Text(sFileName, uLogConv)
        If iStatus = 2 Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
            
            '異常サイズ時：「表示データ量オーバー」ポップアップを表示
            iResponse = MsgBox("データ量が多すぎて、全てを表示できません。" _
                        & Chr(vbKeyReturn) & "一部分のみでも表示しますか？", _
                        vbYesNo + vbExclamation, _
                        "表示データ量オーバー")
            If iResponse = vbYes Then
                fWriteLogtxt = True
            Else
                fWriteLogtxt = False
            End If
            Exit Function          ' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加
        ElseIf iStatus = 1 Then    '正常のとき
            fWriteLogtxt = True
        Else                    'エラーのとき
            fWriteLogtxt = False
        End If
    'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
    End If
    'EG20 V2.1.0.1 ADD END   【フェーズ２対応】
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
'//  機能概要  : 操作卓ログ管理画面より、ログ変換情報を作成する。
'//　　　　　　　表示項目指定部：ログテキストファイル書込み処理
'//
'//              型        名称      意味
'//  引数      : LOGCONV　uLogConv　[OUT]ログ変換情報
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//                 【フェーズ２対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub sGetSearchData(uLogConv As LOGCONV)
    Dim i As Integer                        'カウンタ
    Dim j As Integer                        'コントロール配列数
    Dim sBuff As String                     '文字列バッファ
    Dim byBuff() As Byte                    'バイトバッファ
    Dim iProcessID As Integer               '対象プロセスID
    Dim iChangeCnt As Integer               '変換カウンター(10進→2進(ビット)→10進)
    Dim sChangeProcessId1 As String         '変換後ID[2進]
    Dim lChangeProcessId2 As Long           '変換後ID[10進]
    Dim lSetId As Long                      'エリアセットID
    
    On Error Resume Next
     
    '時刻範囲の作成を行う
    sBuff = ""                              '初期化する
    For i = 0 To 5                          '時刻範囲エリア分繰り返す
        If txtLogTime(i) = "" Then          '入力がない場合
            sBuff = sBuff & "  "            '「空白」をセット
        Else                                '入力がある場合
                                            '２桁数字文字列にする
            sBuff = sBuff & Format(txtLogTime(i), "@@")
        End If
    Next
    byBuff = StrConv(sBuff, vbFromUnicode)  '文字変換する
    For i = 0 To TIMEZONE_LEN - 1           'バイト分繰り返す
        uLogConv.byTimeZone(i) = byBuff(i)  'ログ変換情報に格納する
    Next

    uLogConv.dw1stAssort = ASRT_NOTUSE      '「ログ収集なし」をセット
    uLogConv.dw2stAssort = ASRT_NOTUSE     '「ログ収集なし」をセット
    uLogConv.by2ndAssort = ASRT_NOTUSE      '「ログ収集なし」をセット
    
    '分類の作成を行う
'    If optLogBunrui(0).Value = True Then        'ラジオ釦：「全ての分類を表示」が有効  'EG20 V2.1.0.1 DEL
    'ラジオ釦：「全ての分類を表示」が有効または保守ログ選択
    If optLogBunrui(0).Value = True Or optHoshu(tabTakuCorner.Tab).Value = True Then      'EG20 V2.1.0.1 ADD
       Process_Settei_ALL uLogConv
    Else                                        'ラジオ釦：「指定分類のみ表示」が有効
       Process_Settei uLogConv
    End If

    'ログ種別の作成
    If optLogSyu(0).Value = True Then                 'ラジオ釦：「全ての種別を表示」が有効
        uLogConv.byLogType = LTYP_ALL                 '「全種別」をセット必要
    Else                                              'ラジオ釦：「指定種別のみ表示」が有効
        uLogConv.byLogType = LTYP_NOTUSE              '「無効」をセット
        If chkLogSyu(0).Value = CHECKBOX_ON Then      '「正常」が有効な場合
            uLogConv.byLogType = uLogConv.byLogType + LTYP_NORMAL
        End If
        If chkLogSyu(1).Value = CHECKBOX_ON Then      '「異常」が有効な場合
            uLogConv.byLogType = uLogConv.byLogType + LTYP_ERROR
        End If
        If chkLogSyu(2).Value = CHECKBOX_ON Then      '「警告」が有効な場合
            uLogConv.byLogType = uLogConv.byLogType + LTYP_WARNING
        End If
        If chkLogSyu(4).Value = CHECKBOX_ON Then      '「デバッグ」が有効な場合
            uLogConv.byLogType = uLogConv.byLogType + LTYP_DEBUG
        End If
    End If

    
    '付加情報フラグの作成
    If optLogData(0).Value = True Then          'ラジオ釦：「全行表示」が有効
        uLogConv.byOptFlag = 1                  '「全行表示」をセット
    Else                                        'ラジオ釦：「１行目のみ表示」が有効
        uLogConv.byOptFlag = 0                  '「１行表示」をセット
    End If

    '自改号機情報の作成
    sBuff = ""

'    j = chkLogGouki.UBound         'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    j = UBound(mintStatus)           'EG20 V2.1.0.1 ADD START 【フェーズ２対応】

'EG20 V5.4.0.1 DEL START 【統合No49対応】
'    If optLogGouki(0).Value = True Then         'ラジオ釦：「全号機」が有効        'EG20 V2.1.0.1 DEL 【フェーズ２対応】
'    If optApp(tabTakuCorner.Tab).Value = True Then    'アプリログ選択時               'EG20 V2.1.0.1 ADD 【フェーズ２対応】
'
'        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
''        For i = 0 To j                          '号機分繰り返す
''            sBuff = sBuff & "1"                 '該当号機に「有効」をセット
''        Next
''        For i = j + 1 To GATE_FLAGS_LEN - 1      '号機分繰り返す
''            sBuff = sBuff & "0"                 '該当号機に「無効」をセット
''        Next
'        'EG20 V2.1.0.1 DEL END
'        'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'        For i = 0 To j                          '号機分繰り返す
'            If mintStatus(i) = CHECKBOX_ON Then
'                sBuff = sBuff & "1"                 '該当号機に「有効」をセット
'            Else
'                sBuff = sBuff & "0"                 '該当号機に「無効」をセット
'            End If
'        Next
'        'EG20 V2.1.0.1 ADD END
'
'    Else
'EG20 V5.4.0.1 DEL END
    
        For i = 0 To j                          '号機分繰り返す
'            If chkLogGouki(i).Value = CHECKBOX_ON Then '「？？号機」が有効な場合   'EG20 V2.1.0.1 DEL 【フェーズ２対応】
            If mintStatus(i) = CHECKBOX_ON Then '「？？号機」が有効な場合           'EG20 V2.1.0.1 ADD 【フェーズ２対応】
                sBuff = sBuff & "1"             '該当号機に「有効」をセット
            Else
                sBuff = sBuff & "0"             '該当号機に「無効」をセット
            End If
        Next
        'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'        For i = j + 1 To GATE_FLAGS_LEN - 1     '号機分繰り返す
'            sBuff = sBuff & "0"                 '該当号機に「無効」をセット
'        Next
        'EG20 V2.1.0.1 DEL END
'    End If         'EG20 V5.4.0.1 DEL 【統合No49対応】
    byBuff = StrConv(sBuff, vbFromUnicode)      '文字変換する
    For i = 0 To GATE_FLAGS_LEN - 1             'バイト分繰り返す
        uLogConv.byGateFlag(i) = byBuff(i)      'ログ変換情報に格納する
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtLogTime_DblClick
'//  機能名称  : ログデータ時刻部、ダブルクリック時処理
'//  機能概要  : 擬似テンキー画面を表示
'//　　　　　　　表示項目指定部：ログデータ対象時刻テキストボックス
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]テキストボックスインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub txtLogTime_DblClick(Index As Integer)
    gstrTenKeyData = txtLogTime(Index) ' 現在設定してある情報を渡す
    gstrTenKeySize = 4                 '入力可能文字数を指定する。
    ' 擬似テンキー画面表示
    frmTenKey.Show 1
    ' 設定した情報を更新する
    txtLogTime(Index) = gstrTenKeyData
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtLogTime_KeyPress
'//  機能名称  : ログデータ時刻部、キー入力処理
'//  機能概要  : 入力キーチェックを行う。
'//　　　　　　　表示項目指定部：ログデータ対象時刻テキストボックス
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]テキストボックスインデックス
'//  　　      : Integer　　KeyAscii [IN]入力キー
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub txtLogTime_KeyPress(Index As Integer, KeyAscii As Integer)
    
    '背景色を白色にする
    txtLogTime(Index).BackColor = MN_COLOR_WHITE
    '数字のみ有効とする
    KeyAscii = pfKeyNumeric(KeyAscii)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfKeyNumeric
'//  機能名称  : 数字入力処理
'//  機能概要  : 数字以外の文字を無効にする。。
'//　　　　　　　表示項目指定部：ログデータ対象時刻テキストボックス
'//
'//              型        名称      意味
'//  引数      : Integer　　KeyAscii [IN]入力キー
'//
'//              型        値        意味
'//  戻り値    : Integer　　　　　　 [OUT]キーコード
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Function pfKeyNumeric(iKeyAscii As Integer) As Integer
    
    '初期値として引数のコードを戻り値とする
    pfKeyNumeric = iKeyAscii
    
    'バックスペースキーは有効とする
    If iKeyAscii = vbKeyBack Then
        Exit Function
    End If
    '数字以外は無効とする
    If iKeyAscii < vbKey0 Or iKeyAscii > vbKey9 Then
        pfKeyNumeric = 0
        Beep
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtLogTime_Change
'//  機能名称  : ログデータ対象時刻入力処理
'//  機能概要  : 表示項目指定部：ログデータ対象時刻テキストボックス
'//　　　　　　　　　　　　　　　時刻エリア処理の入力値チェック
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
Private Sub txtLogTime_Change(Index As Integer)
    
    '規定桁数入力
    If Len(txtLogTime(Index)) = 2 Then
        Select Case Index
        Case 0, 3
            '日付(日)の正当性をチェックする
            If pfTextDay(txtLogTime(Index)) <> True Then
                '前面色をエラー色にする
                txtLogTime(Index).ForeColor = MN_COLOR_RED
                Exit Sub
            End If
        Case 1, 4
            '日付(時)の正当性をチェックする
            If pfTextHour(txtLogTime(Index)) <> True Then
                '前面色をエラー色にする
                txtLogTime(Index).ForeColor = MN_COLOR_RED
                Exit Sub
            End If
        Case 2, 5
            '日付(分)の正当性をチェックする
            If pfTextMin(txtLogTime(Index).Text) <> True Then
                '前面色をエラー色にする
                txtLogTime(Index).ForeColor = MN_COLOR_RED
                Exit Sub
            End If
        Case Else
        End Select
        If Index < 5 Then
            'エラーがなければ次の項目へフォーカスを移す
            txtLogTime(Index + 1).SetFocus
        End If
    End If
    '前面色を黒色にする
    txtLogTime(Index).ForeColor = MN_COLOR_BLACK

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfTextDay
'//  機能名称  : 日付正当性チェック処理
'//  機能概要  : 日付の正当性チェックを行う。
'//　　　　　　　表示項目指定部：ログデータ対象時刻テキストボックス
'//
'//              型        名称      意味
'//  引数      : String　　sText　　[IN]入力日値
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Function pfTextDay(sText As String) As Boolean
    
    pfTextDay = False
    '文字数をチェックする
    If LenB(StrConv(sText, vbFromUnicode)) > 2 Then
        Exit Function
    End If
    '数値の正当性チェック
    If IsNumeric(sText) = False Then
        Exit Function
    End If
    '範囲チェックを行う
    If CInt(sText) < 1 Or CInt(sText) > 31 Then
        Exit Function
    End If
    pfTextDay = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfTextHour
'//  機能名称  : 時間正当性チェック処理
'//  機能概要  : 時間の正当性チェックを行う。
'//　　　　　　　表示項目指定部：ログデータ対象時刻テキストボックス
'//
'//              型        名称      意味
'//  引数      : String　　sText　　[IN]入力時間値
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Function pfTextHour(sText As String) As Boolean
    
    pfTextHour = False
    '文字数をチェックする
    If LenB(StrConv(sText, vbFromUnicode)) > 2 Then
        Exit Function
    End If
    '数値の正当性チェック
    If IsNumeric(sText) = False Then
        Exit Function
    End If
    '範囲チェックを行う
    If CInt(sText) < 0 Or CInt(sText) > 23 Then
        Exit Function
    End If
    pfTextHour = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfTextMin
'//  機能名称  : 分数正当性チェック処理
'//  機能概要  : 分数の正当性チェックを行う。
'//　　　　　　　表示項目指定部：ログデータ対象時刻テキストボックス
'//
'//              型        名称      意味
'//  引数      : String　　sText　　[IN]入力分数値
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Function pfTextMin(sText As String) As Boolean
    
    pfTextMin = False
    '文字数をチェックする
    If LenB(StrConv(sText, vbFromUnicode)) > 2 Then
        Exit Function
    End If
    '数値の正当性チェック
    If IsNumeric(sText) = False Then
        Exit Function
    End If
    '範囲チェックを行う
    If CInt(sText) < 0 Or CInt(sText) > 59 Then
        Exit Function
    End If
    pfTextMin = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optLogSyu_Click
'//  機能名称  : 種別ラジオ釦押下時処理
'//  機能概要  : 指定種別のアクティブ・非アクティブの画面更新処理を行う。
'//　　　　　　　表示項目指定部：「種別」部
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下ラジオ釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub optLogSyu_Click(Index As Integer)
    '押下ラジオ釦による画面表示更新処理
    sLogIndexChange
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optlogbunrui_Click
'//  機能名称  : 分類ラジオ釦押下時処理
'//  機能概要  : 指定分類のアクティブ・非アクティブの画面更新処理を行う。
'//　　　　　　　表示項目指定部：「分類」部
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下ラジオ釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub optlogbunrui_Click(Index As Integer)
    '押下ラジオ釦による画面表示更新処理
    sLogIndexChange
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optAll_Click
'//  機能名称  : 「全て選択」「全て非選択」釦押下時処理
'//  機能概要  : 指定分類のチェックON/OFFを行う。
'//　　　　　　　表示項目指定部：「指定分類」部
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub optAll_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To iModCnt
        If Index = 0 Then
        '「全て選択」釦押下時：指定分類を全てチェックする。
            chkMod(i).Value = vbChecked
        Else
        '「全て非選択」釦押下時：指定分類を全てチェックしない。
            chkMod(i).Value = vbUnchecked
        End If
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sLogIndexChange
'//  機能名称  : 項目認識変更処理
'//  機能概要  : 種別、分類の画面表示を更新する。
'//　　　　　　　表示項目指定部：「指定種別」部「指定分類」部
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-30   REVISED BY [TCC] S.Terao
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub sLogIndexChange()
    Dim i As Integer        'カウンタ
    Dim j As Integer        'コントロール配列数

    '***********************
    '* 種別エリアボックス  *
    '***********************
    j = chkLogSyu.UBound
    'ラジオ釦：「全ての種別を表示」が有効
    If optLogSyu(0).Value = True Then
        '全ての種別を非アクティブ表示にする
        For i = 0 To j                      '指定種別数分繰り返す
            chkLogSyu(i).Enabled = False
        Next
     'ラジオ釦：「指定種別のみ表示」が有効
    Else
        '全ての種別をアクティブ表示にする
        For i = 0 To j                      '指定種別数分繰り返す
            chkLogSyu(i).Enabled = True
        Next
    End If

    '***********************
    '* 分類エリアボックス  *
    '***********************
    j = iModCnt
    'ラジオ釦：「全ての分類を表示」が有効
    If optLogBunrui(0).Value = True Then
        '全ての分類を非アクティブ表示にする
        For i = 0 To j                      '指定分類数分繰り返す
             chkMod(i).Enabled = False
             'chkMod(i).Value = CHECKBOX_ON 'V1.7.0.1 DEL
        Next
        optAll(0).Enabled = False  '「全て選択」釦を非アクティブ表示にする。
        optAll(1).Enabled = False  '「全て非選択」釦を非アクティブ表示にする。
    'ラジオ釦：「指定分類のみ表示」が有効
    Else
        '全ての分類をアクティブ表示にする
        For i = 0 To j                     '指定分類数分繰り返す
             chkMod(i).Enabled = True
        Next
        optAll(0).Enabled = True  '「全て選択」釦をアクティブ表示にする。
        optAll(1).Enabled = True  '「全て非選択」釦をアクティブ表示にする。
    End If
End Sub

'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optLogGouki_Click
'//  機能名称  : 項目認識変更処理
'//  機能概要  : 種別、分類の画面表示を更新する。
'//　　　　　　　表示項目指定部：「指定種別」部「指定分類」部
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]ラジオ釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
'Private Sub optLogGouki_Click(Index As Integer)
'    '押下ラジオ釦による画面表示更新処理
'    sOptGoukiChange
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdChkAll_Click
'//  機能名称  : 「全号機選択」釦押下時処理
'//  機能概要  : 全自改号機のチェックをONにする。
'//　　　　　　　表示自改号機指定部：「自改号機」部
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
'Private Sub cmdChkAll_Click()
'
'    Dim i As Integer
'
'    For i = 0 To 17
'       chkLogGouki(i).Value = CHECKBOX_ON
'    Next
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdChkAllKai_Click
'//  機能名称  : 「全号機解除」釦押下時処理
'//  機能概要  : 全自改号機のチェックをOFFにする。
'//　　　　　　　表示自改号機指定部：「自改号機」部
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
'Private Sub cmdChkAllKai_Click()
'
'   Dim i As Integer
'
'    For i = 0 To 17
'       chkLogGouki(i).Value = CHECKBOX_OFF
'    Next
'
'End Sub
'EG20 V2.1.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : chkMod_Click
'//  機能名称  : 指定分類の各チェックボックス押下処理
'//  機能概要  : 指定分類の各チェックボックス状態更新を行う。
'//　　　　　　　表示項目指定部：「指定分類」部
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]各チェックボックスインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub chkMod_Click(Index As Integer)
    Dim iCnt As Integer
    Dim sDai As String
    Dim iChkType As Integer
    
    '付属カウンターが0かどうかチェックする。
    'カウンター0：大分類扱い。カウンター0以外：中分類扱い
    If Int(uModFileData(Index).iFuzokuCnt) = 0 Then
        'インデックス番号が最終の場合、最終以降はないので処理終了
        If Index = iModCnt Then
            Exit Sub
        End If
        
        '中分類扱いのインデックス番号を作成
        iCnt = Index + 1
        '中分類扱い、大分類扱いのプロセスIDを取得する。
        sDai = uModFileData(Index).iProces
        '大分類扱いのチェックボックス状態値を取得する。
        iChkType = chkMod(Index).Value
        Do
            '中分類扱いの付属IDと、大分類扱いのIDとが一致するかどうかチェックする。
            '(※中分類扱いと大分類扱いの繋がり確認)
            If sDai = uModFileData(iCnt).iFuzokuId Then
                '一致した場合、大分類のチェックボックス状態値を中分類扱いにも反映する。
                chkMod(iCnt).Value = iChkType
            Else
                '不一致の場合、処理終了。
                Exit Do
            End If
            '中分類扱いのものがまだいるかチェックする。
            iCnt = iCnt + 1
            If iCnt > iModCnt Then
                'インデックス番号が最終になれば処理終了
                Exit Sub
            End If
        Loop
    End If
End Sub

'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sOptGoukiChange
'//  機能名称  : 表示自改号機指定変更処理
'//  機能概要  : ラジオ釦押下による、画面表示を更新する。
'//　　　　　　　表示自改号機指定部：「自改号機」部
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
'Private Sub sOptGoukiChange()
'    Dim i As Integer            'カウンタ
'    Dim j As Integer        'コントロール配列数
'
'    '指定号機
'    j = chkLogGouki.UBound
'    'ラジオ釦：「全号機」が有効
'    If optLogGouki(0).Value = True Then
'        cmdChkAll.Enabled = False
'        cmdChkAllKai.Enabled = False
'        '全ての号機を非アクティブ表示にする
'        For i = 0 To j                    '号機数分繰り返す
'            chkLogGouki(i).Enabled = False
'        Next
'    'ラジオ釦：「指定号機のみ」が有効
'    Else
'         cmdChkAll.Enabled = True
'         cmdChkAllKai.Enabled = True
'        '全ての号機をアクティブ表示にする
'         For i = 0 To j                   '号機数分繰り返す
'            chkLogGouki(i).Enabled = True
'        Next
'    End If
'
'End Sub
'EG20 V2.1.0.1 DEL END

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

        Case ML_ID_LGBUFF_ANS
             '「バッファフラッシュ終了通知」を受信した場合
             '「バッファフラッシュ終了通知受信正常」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
             '戻り値にメールＩＤをセット
             fMailRecieve = ML_ID_LGBUFF_ANS

        Case ML_ID_HOSHU_ACTIVE_REQ
             '保守画面アクティブ表示の場合
             '「保守画面アクティブ表示要求受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
             AppActivate frmTakuLogKanri.Caption, False
             pfFormActive (frmTakuLogKanri.hwnd)
             fMailRecieve = ML_ID_HOSHU_ACTIVE_REQ

        Case ML_ID_LGCHGREQ_RES
             'ログ切替要求RESの場合
             '「ログ切替要求RES受信正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, LOG_CHANGE_RES_RECV, 0)
             fMailRecieve = ML_ID_LGCHGREQ_RES

        Case Else
        'メールＩＤ不正
          '「メールID不正」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Function

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    Dim lId As Long         'メールＩＤ
    'メールを受信する'
    lId = fMailRecieve()
    If lId = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmTakuLogKanri.Caption, False
        pfFormActive (frmTakuLogKanri.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Process_Settei
'//  機能名称  : 大分類のビット設定処理
'//  機能概要  : 大分類の設定処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   CODED   BY [TCC] C.Terui
'//     REVISIONS :(V30.1.0.1) 2014-05-21   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub Process_Settei(uLogConv As LOGCONV)
    Dim i As Integer                        'カウンタ
    Dim iProcessID As Integer               '対象プロセスID
    Dim iChangeCnt As Integer               '変換カウンター(10進→2進(ビット)→10進)
    Dim sChangeProcessId1 As String         '変換後ID[2進]
    Dim lChangeProcessId2 As Long           '変換後ID[10進]
    Dim lSetId As Long                      'エリアセットID
' V1.3.0.1 ADD START
    Dim bit(0 To 31) As Long
    'ビット宣言
    'EG20 V30.1.0.1 DEL START
'    bit(0) = &H1
'    bit(1) = &H2
'    bit(2) = &H4
'    bit(3) = &H8
'    bit(4) = &H10
'    bit(5) = &H20
'    bit(6) = &H40
'    bit(7) = &H80
'    bit(8) = &H100
'    bit(9) = &H200
'    bit(10) = &H400
'    bit(11) = &H800
'    bit(12) = &H1000
'    bit(13) = &H2000
'    bit(14) = &H4000
'    bit(15) = &H8000
'    bit(16) = &H10000
'    bit(17) = &H20000
'    bit(18) = &H40000
'    bit(19) = &H80000
'    bit(20) = &H100000
'    bit(21) = &H200000
'    bit(22) = &H400000
'    bit(23) = &H800000
'    bit(24) = &H1000000
'    bit(25) = &H2000000
'    bit(26) = &H4000000
'    bit(27) = &H8000000
'    bit(28) = &H10000000
'    bit(29) = &H20000000
'    bit(30) = &H40000000
'    bit(31) = &H80000000
    'EG20 V30.1.0.1 DEL END
    'EG20 V30.1.0.1 ADD START
    '&Hxxxx&と後ろに&をつけないとLONG型として処理されないので修正。&H8000がマイナス値になってしまう。
    'ビット宣言
    bit(0) = &H1&
    bit(1) = &H2&
    bit(2) = &H4&
    bit(3) = &H8&
    bit(4) = &H10&
    bit(5) = &H20&
    bit(6) = &H40&
    bit(7) = &H80&
    bit(8) = &H100&
    bit(9) = &H200&
    bit(10) = &H400&
    bit(11) = &H800&
    bit(12) = &H1000&
    bit(13) = &H2000&
    bit(14) = &H4000&
    bit(15) = &H8000&
    bit(16) = &H10000
    bit(17) = &H20000
    bit(18) = &H40000
    bit(19) = &H80000
    bit(20) = &H100000
    bit(21) = &H200000
    bit(22) = &H400000
    bit(23) = &H800000
    bit(24) = &H1000000
    bit(25) = &H2000000
    bit(26) = &H4000000
    bit(27) = &H8000000
    bit(28) = &H10000000
    bit(29) = &H20000000
    bit(30) = &H40000000
    bit(31) = &H80000000
    'EG20 V30.1.0.1 ADD END
         
    '指定分類分ループする。
      For i = 0 To iModCnt
       '指定分類指定有無チェックを行う。
       If chkMod(i).Value = CHECKBOX_ON Then
          '対象プロセスIDを取得する
          iProcessID = uModFileData(i).iProces
          If (0 < iProcessID) And (iProcessID <= 31) Then
' V1.3.0.1 DEL START
'             'プロセスIDを2進数に変換する。
'             sChangeProcessId1 = 0
'             iChangeCnt = 0
'             For iChangeCnt = 1 To iProcessID
'                If iChangeCnt = 1 Then
'                  'ビットをたたせる。
'                   sChangeProcessId1 = 1
'                Else
'                   sChangeProcessId1 = sChangeProcessId1 & 0
'                End If
'              Next
'
'              lChangeProcessId2 = 0
'              '2進数を10進数に変換する。
'              For iChangeCnt = 0 To Len(sChangeProcessId1) - 1
'                 If Mid(sChangeProcessId1, iChangeCnt + 1, 1) <> 0 Then
'                    lChangeProcessId2 = lChangeProcessId2 + 2 ^ (Len(sChangeProcessId1) - iChangeCnt - 1)
'                 End If
'              Next iChangeCnt
'               uLogConv.dw1stAssort = uLogConv.dw1stAssort + lChangeProcessId2
' V1.3.0.1 DEL END
                uLogConv.dw1stAssort = uLogConv.dw1stAssort + bit(iProcessID)       ' V1.3.0.1 ADD

          ElseIf (31 < iProcessID) And (iProcessID < 63) Then
              iProcessID = iProcessID - 32
' V1.3.0.1 DEL START
'                'プロセスIDを2進数に変換する。
'              iChangeCnt = 0
'              sChangeProcessId1 = 0
'               For iChangeCnt = 1 To iProcessID
'                  If iChangeCnt = 1 Then
'                    'ビットをたたせる｡
'                     sChangeProcessId1 = 1
'                  Else
'                     sChangeProcessId1 = sChangeProcessId1 & 0
'                  End If
'               Next
'
'               lChangeProcessId2 = 0
'               '2進数を10進数に変換する｡
'               For iChangeCnt = 0 To Len(sChangeProcessId1) - 1
'                  If Mid(sChangeProcessId1, iChangeCnt + 1, 1) <> 0 Then
'                     lChangeProcessId2 = lChangeProcessId2 + 2 ^ (Len(sChangeProcessId1) - iChangeCnt - 1)
'                  End If
'               Next iChangeCnt
'                uLogConv.dw2stAssort = uLogConv.dw2stAssort + lChangeProcessId2
' V1.3.0.1 DEL END
                 uLogConv.dw2stAssort = uLogConv.dw2stAssort + bit(iProcessID)       ' V1.3.0.1 ADD
          End If
        End If
      Next
End Sub
'V1.7.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Process_Settei_ALL
'//  機能名称  : 大分類のビット設定処理(無条件全分類)
'//  機能概要  : 大分類の設定処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.7.0.1) 2009-07-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(V30.1.0.1) 2014-05-21   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub Process_Settei_ALL(uLogConv As LOGCONV)
    Dim i As Integer                        'カウンタ
    Dim iProcessID As Integer               '対象プロセスID
    Dim iChangeCnt As Integer               '変換カウンター(10進→2進(ビット)→10進)
    Dim sChangeProcessId1 As String         '変換後ID[2進]
    Dim lChangeProcessId2 As Long           '変換後ID[10進]
    Dim lSetId As Long                      'エリアセットID
    
    Dim bit(0 To 31) As Long
    'ビット宣言
    'EG20 V30.1.0.1 DEL START
'    bit(0) = &H1
'    bit(1) = &H2
'    bit(2) = &H4
'    bit(3) = &H8
'    bit(4) = &H10
'    bit(5) = &H20
'    bit(6) = &H40
'    bit(7) = &H80
'    bit(8) = &H100
'    bit(9) = &H200
'    bit(10) = &H400
'    bit(11) = &H800
'    bit(12) = &H1000
'    bit(13) = &H2000
'    bit(14) = &H4000
'    bit(15) = &H8000
'    bit(16) = &H10000
'    bit(17) = &H20000
'    bit(18) = &H40000
'    bit(19) = &H80000
'    bit(20) = &H100000
'    bit(21) = &H200000
'    bit(22) = &H400000
'    bit(23) = &H800000
'    bit(24) = &H1000000
'    bit(25) = &H2000000
'    bit(26) = &H4000000
'    bit(27) = &H8000000
'    bit(28) = &H10000000
'    bit(29) = &H20000000
'    bit(30) = &H40000000
'    bit(31) = &H80000000
    'EG20 V30.1.0.1 DEL END
    
    'EG20 V30.1.0.1 ADD START
    '&Hxxxx&と後ろに&をつけないとLONG型として処理されないので修正。&H8000がマイナス値になってしまう。
    'ビット宣言
    bit(0) = &H1&
    bit(1) = &H2&
    bit(2) = &H4&
    bit(3) = &H8&
    bit(4) = &H10&
    bit(5) = &H20&
    bit(6) = &H40&
    bit(7) = &H80&
    bit(8) = &H100&
    bit(9) = &H200&
    bit(10) = &H400&
    bit(11) = &H800&
    bit(12) = &H1000&
    bit(13) = &H2000&
    bit(14) = &H4000&
    bit(15) = &H8000&
    bit(16) = &H10000
    bit(17) = &H20000
    bit(18) = &H40000
    bit(19) = &H80000
    bit(20) = &H100000
    bit(21) = &H200000
    bit(22) = &H400000
    bit(23) = &H800000
    bit(24) = &H1000000
    bit(25) = &H2000000
    bit(26) = &H4000000
    bit(27) = &H8000000
    bit(28) = &H10000000
    bit(29) = &H20000000
    bit(30) = &H40000000
    bit(31) = &H80000000
    'EG20 V30.1.0.1 ADD END
         
    '指定分類分ループする。
      For i = 0 To iModCnt
       '指定分類指定有無チェックを行う。
       '対象プロセスIDを取得する
       iProcessID = uModFileData(i).iProces
       If (0 < iProcessID) And (iProcessID <= 31) Then
          uLogConv.dw1stAssort = uLogConv.dw1stAssort + bit(iProcessID)
       ElseIf (31 < iProcessID) And (iProcessID < 63) Then
          iProcessID = iProcessID - 32
          uLogConv.dw2stAssort = uLogConv.dw2stAssort + bit(iProcessID)
       End If
      Next
End Sub
'V1.7.0.1 ADD END
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetEnableFalse
'//  機能名称  : 画面ロック処理
'//  機能概要  : 画面のロックをする。
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
Public Sub SetEnableFalse()

    On Error Resume Next
  
    'タブをFalseにする。
    tabLog.Enabled = False
    
    '「ログ媒体出力」釦をFalseにする。
    cmdLog(1).Enabled = False
    
    '「圧縮媒体出力」釦をFalseにする。
    cmdLzhFileWrite.Enabled = False
       
    '「表示更新」釦をFalseにする。
    cmdUpdateDisplay.Enabled = False
        
    '「メモ帳表示」釦をFalseにする。
    cmdLog(0).Enabled = False

    '「媒体取外」釦をFalseにする。
    cmdInstall.Enabled = False
    
    '「保守画面へ戻る」釦をFalseにする。
    cmdReturn.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetEnableTrue
'//  機能名称  : 画面ロック解除処理
'//  機能概要  : 画面のロックを解除する。
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
Public Sub SetEnableTrue()
  
    On Error Resume Next

    'タブをTrueにする。
    tabLog.Enabled = True
    
    '「ログ媒体出力」釦をTrueにする。
    cmdLog(1).Enabled = True
    
    '「ログ圧縮媒体出力」釦をTrueにする。
    cmdLzhFileWrite.Enabled = True
        
    '「表示更新」釦をTrueにする。
    cmdUpdateDisplay.Enabled = True
    
    '「ログ表示(テキスト表示)」釦をTrueにする。
    cmdLog(0).Enabled = True

    '「媒体取外」釦をTrueにする。
    cmdInstall.Enabled = True
    
    '「保守画面へ戻る」釦をTrueにする。
    cmdReturn.Enabled = True

End Sub

'EG20 V2.1.0.1 ADD START 【統-350対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : tabTakuCorner_Click
'//  機能名称  : コーナ選択タブクリック時処理
'//  機能概要  : 表示号機指定を選択コーナのみにする
'//
'//              型        名称      意味
'//  引数      : Integer　　Index  　[IN]テキストボックスインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-13   CODED   BY [TCC] M.Matsumoto
'//                 【統-350対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub tabTakuCorner_Click(PreviousTab As Integer)

    Dim intIndex As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGoki As Integer
    
    intStIndex = tabTakuCorner.Tab * 16
    intEdIndex = intStIndex + 15
    
    '表示号機指定のコーナタブを選択コーナのみの表示にする
    For intIndex = 0 To tabTakuCorner.Tabs - 1
        If intIndex = tabTakuCorner.Tab Then
            tabCorner.TabVisible(intIndex) = True
            tabCorner.Tab = intIndex
        Else
            tabCorner.TabVisible(intIndex) = False
        End If
    Next
    
    '選択状態（内部変数）は選択コーナの号機のみ有効とする
    For intIndex = 0 To chkLogGouki.UBound
        intGoki = CInt(chkLogGouki(intIndex).Tag) - 1
        If intGoki >= 0 Then
            If intIndex >= intStIndex And intIndex <= intEdIndex Then
                    mintStatus(intGoki) = chkLogGouki(intIndex).Value
            Else
                mintStatus(intGoki) = CHECKBOX_OFF
            End If
        End If
    Next
    
End Sub
'EG20 V2.1.0.1 ADD END

