VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGateVerKanri 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'なし
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   " 媒体 → ワーク　コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      TabIndex        =   87
      Top             =   3240
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMail 
      Left            =   8640
      Top             =   1560
   End
   Begin VB.CommandButton cmdGateVerUpdate 
      Caption         =   "一括更新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   79
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "バージョン情報  媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   0
      Left            =   9360
      TabIndex        =   20
      Top             =   6120
      Width           =   2415
   End
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
      Height          =   550
      Index           =   1
      Left            =   9360
      TabIndex        =   19
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdModoru_Menu 
      Caption         =   "  バージョン管理  画面へ戻る"
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
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   1
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Frame fraDataSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   6255
      Begin VB.OptionButton optData 
         Caption         =   "予備３"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   3960
         TabIndex        =   78
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "予備２"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   3960
         TabIndex        =   77
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "予備１"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   3960
         TabIndex        =   76
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "ｻﾌﾞCPU-Pro3"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2040
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "自改（ＯＳ）"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2040
         TabIndex        =   18
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "ｻﾌﾞCPU-Pro2"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "ｻﾌﾞCPU-Pro1"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "プログラム"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "判定データ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame fraFolderSelect 
      Height          =   1335
      Left            =   6720
      TabIndex        =   8
      Top             =   7080
      Width           =   1935
      Begin VB.CheckBox chkFolder 
         Caption         =   "Ｏ 旧"
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Ｎ 実行"
         BeginProperty Font 
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
         Left            =   240
         TabIndex        =   10
         Top             =   615
         Width           =   1575
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Ｗ ワーク"
         BeginProperty Font 
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
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ワーククリア"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " 圧縮ファイル → ワークコピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyWork_Jikko 
      Caption         =   " ワーク → 実行 コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   旧 → 実行   コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdDLLJikkoGamen 
      Caption         =   " 自改切り離し"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   2
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdKoshin 
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
      Height          =   550
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   21
      Top             =   360
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15266
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   706
      TabMaxWidth     =   3475
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(0)   =   "バージョン管理(自改).frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblKan(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblKan(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblKan(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblKan(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblKan(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblKan(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblKan(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblKan(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblZenVer(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lstKan(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(1)   =   "バージョン管理(自改).frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblKan(8)"
      Tab(1).Control(1)=   "lblKan(9)"
      Tab(1).Control(2)=   "lblKan(10)"
      Tab(1).Control(3)=   "lblKan(11)"
      Tab(1).Control(4)=   "lblKan(12)"
      Tab(1).Control(5)=   "lblKan(13)"
      Tab(1).Control(6)=   "lblKan(14)"
      Tab(1).Control(7)=   "lblKan(15)"
      Tab(1).Control(8)=   "lblZenVer(1)"
      Tab(1).Control(9)=   "lstKan(1)"
      Tab(1).Control(10)=   "Command1(1)"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(2)   =   "バージョン管理(自改).frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblKan(16)"
      Tab(2).Control(1)=   "lblKan(17)"
      Tab(2).Control(2)=   "lblKan(18)"
      Tab(2).Control(3)=   "lblKan(19)"
      Tab(2).Control(4)=   "lblKan(20)"
      Tab(2).Control(5)=   "lblKan(21)"
      Tab(2).Control(6)=   "lblKan(22)"
      Tab(2).Control(7)=   "lblKan(23)"
      Tab(2).Control(8)=   "lblZenVer(2)"
      Tab(2).Control(9)=   "lstKan(2)"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(3)   =   "バージョン管理(自改).frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblKan(24)"
      Tab(3).Control(1)=   "lblKan(25)"
      Tab(3).Control(2)=   "lblKan(26)"
      Tab(3).Control(3)=   "lblKan(27)"
      Tab(3).Control(4)=   "lblKan(28)"
      Tab(3).Control(5)=   "lblKan(29)"
      Tab(3).Control(6)=   "lblKan(30)"
      Tab(3).Control(7)=   "lblKan(31)"
      Tab(3).Control(8)=   "lblZenVer(3)"
      Tab(3).Control(9)=   "lstKan(3)"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(4)   =   "バージョン管理(自改).frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblKan(32)"
      Tab(4).Control(1)=   "lblKan(33)"
      Tab(4).Control(2)=   "lblKan(34)"
      Tab(4).Control(3)=   "lblKan(35)"
      Tab(4).Control(4)=   "lblKan(36)"
      Tab(4).Control(5)=   "lblKan(37)"
      Tab(4).Control(6)=   "lblKan(38)"
      Tab(4).Control(7)=   "lblKan(39)"
      Tab(4).Control(8)=   "lblZenVer(4)"
      Tab(4).Control(9)=   "lstKan(4)"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(5)   =   "バージョン管理(自改).frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstKan(5)"
      Tab(5).Control(1)=   "lblZenVer(5)"
      Tab(5).Control(2)=   "lblKan(47)"
      Tab(5).Control(3)=   "lblKan(46)"
      Tab(5).Control(4)=   "lblKan(45)"
      Tab(5).Control(5)=   "lblKan(44)"
      Tab(5).Control(6)=   "lblKan(43)"
      Tab(5).Control(7)=   "lblKan(42)"
      Tab(5).Control(8)=   "lblKan(41)"
      Tab(5).Control(9)=   "lblKan(40)"
      Tab(5).ControlCount=   10
      Begin VB.CommandButton Command1 
         Caption         =   " 媒体 → ワーク　コピー"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   1
         Left            =   -65640
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   86
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ListBox lstKan 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Index           =   5
         Left            =   -73920
         TabIndex        =   67
         Top             =   2280
         Width           =   7335
      End
      Begin VB.ListBox lstKan 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Index           =   4
         Left            =   -73920
         TabIndex        =   58
         Top             =   2280
         Width           =   7335
      End
      Begin VB.ListBox lstKan 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Index           =   3
         Left            =   -73920
         TabIndex        =   49
         Top             =   2280
         Width           =   7335
      End
      Begin VB.ListBox lstKan 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Index           =   2
         Left            =   -73920
         TabIndex        =   40
         Top             =   2280
         Width           =   7335
      End
      Begin VB.ListBox lstKan 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Index           =   1
         Left            =   -73920
         TabIndex        =   31
         Top             =   2280
         Width           =   7335
      End
      Begin VB.ListBox lstKan 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Index           =   0
         Left            =   1080
         TabIndex        =   22
         Top             =   2280
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '右揃え
         Caption         =   "全体バージョン"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   -73920
         TabIndex        =   85
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '右揃え
         Caption         =   "全体バージョン"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   -73920
         TabIndex        =   84
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '右揃え
         Caption         =   "全体バージョン"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   -73920
         TabIndex        =   83
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '右揃え
         Caption         =   "全体バージョン"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   -73920
         TabIndex        =   82
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '右揃え
         Caption         =   "全体バージョン"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   -73920
         TabIndex        =   81
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '右揃え
         Caption         =   "○○○○○○バージョン（ワーク）：99"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   1080
         TabIndex        =   80
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "コメント"
         Height          =   255
         Index           =   47
         Left            =   -71160
         TabIndex        =   75
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   -67200
         TabIndex        =   74
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "作成日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   -68850
         TabIndex        =   73
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   -70110
         TabIndex        =   72
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "機種名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   -71160
         TabIndex        =   71
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ﾌｫﾙﾀﾞ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   42
         Left            =   -72000
         TabIndex        =   70
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   41
         Left            =   -73680
         TabIndex        =   69
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "タイプ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   40
         Left            =   -73920
         TabIndex        =   68
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "コメント"
         Height          =   255
         Index           =   39
         Left            =   -71160
         TabIndex        =   66
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   -67200
         TabIndex        =   65
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "作成日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   -68850
         TabIndex        =   64
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   -70110
         TabIndex        =   63
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "機種名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   -71160
         TabIndex        =   62
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ﾌｫﾙﾀﾞ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   34
         Left            =   -72000
         TabIndex        =   61
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   33
         Left            =   -73680
         TabIndex        =   60
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "タイプ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   32
         Left            =   -73920
         TabIndex        =   59
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "コメント"
         Height          =   255
         Index           =   31
         Left            =   -71160
         TabIndex        =   57
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   -67200
         TabIndex        =   56
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "作成日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   -68850
         TabIndex        =   55
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   -70110
         TabIndex        =   54
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "機種名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   -71160
         TabIndex        =   53
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ﾌｫﾙﾀﾞ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   26
         Left            =   -72000
         TabIndex        =   52
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   -73680
         TabIndex        =   51
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "タイプ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   -73920
         TabIndex        =   50
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "コメント"
         Height          =   255
         Index           =   23
         Left            =   -71160
         TabIndex        =   48
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   -67200
         TabIndex        =   47
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "作成日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   -68850
         TabIndex        =   46
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   -70110
         TabIndex        =   45
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "機種名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   -71160
         TabIndex        =   44
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ﾌｫﾙﾀﾞ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   -72000
         TabIndex        =   43
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   17
         Left            =   -73680
         TabIndex        =   42
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "タイプ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   -73920
         TabIndex        =   41
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "コメント"
         Height          =   255
         Index           =   15
         Left            =   -71160
         TabIndex        =   39
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   -67200
         TabIndex        =   38
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "作成日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   -68850
         TabIndex        =   37
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   -70110
         TabIndex        =   36
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "機種名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   -71160
         TabIndex        =   35
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ﾌｫﾙﾀﾞ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   -72000
         TabIndex        =   34
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   -73680
         TabIndex        =   33
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "タイプ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   -73920
         TabIndex        =   32
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "タイプ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   1080
         TabIndex        =   30
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   1320
         TabIndex        =   29
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ﾌｫﾙﾀﾞ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   3000
         TabIndex        =   28
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "機種名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   27
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "ファイル"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4890
         TabIndex        =   26
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "作成日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6150
         TabIndex        =   25
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7800
         TabIndex        =   24
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "コメント"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   23
         Top             =   2040
         Width           =   4575
      End
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "自動改札機バージョン管理"
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
      Index           =   0
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   12120
   End
End
Attribute VB_Name = "frmGateVerKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmJGateVerKanri.frm
'//  パッケージ名：バージョン管理(EG20自改)画面
'//
'//  概要：バージョン管理(EG-R自改/NEG自改)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　プロ判正当性チェック処理追加
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応
'//                     ・機種正当性チェック処理追加/「ワーク→実行コピー」時
'//                     ・フェーズ２不具合修正
'//                     ・フェーズ１不具合修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.11.0.1) 2009-10-23  CODED   BY [TCC] D.Yamashita
'//                 ・フェーズ３残件項目対応
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 入力ファイル格納ディレクトリ位置変更
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(1.16.0.1) 2009-12-21   REVISED BY [TCC] S.Terao
'//                 不具合対応
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                ① フォルダ選択画面をOS仕様に変更
'//                ②「メニュー画面へ戻る」釦押下にて、
'//                 　バージョン管理画面のバージョン表示更新を行う
'//                ③表示リソースラジオ釦選択でリストの表示更新
'//                ④ワーク→実行コピーでの機種正当性チェック変更
'//                ⑤ワーク→実行コピーでの正当性チェックiniファイル化
'//                ⑥Dir関数をFileSystemObjectに置き換え
'//                ⑦ファイル選択画面をOS仕様に変更
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(V2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 EG-R(KK)　八丁畷対応　KUK正当性チェック変更
'//                 媒体取外不具合修正
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 ファイル名チェック不具合修正
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-16  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                【運改表示改善対応】
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【種別チェック機能追加】【TOMAS用領域コピー対応】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Dim FolderSyubetu As Integer                 '選択リソース種別

Dim FolderName(0 To 2, 0 To 9) As String     'フォルダ名
Dim TitleBox(0 To 10) As String               'タイトル名
Dim LogBox(0 To 10) As String                 'ログ出力用タイトル名
Dim FileList() As String                     'ファイル名リスト一覧格納エリア
Dim FileListType() As String                 'ファイルリスト一覧格納エリア（次世代自改タイプを含む）
Dim uVersion() As MN_VERSION_JIKAI           'バージョン情報格納エリア
Dim gintUnkaiKind(0 To 8) As Integer         ' 運改種別    ' EG20 V5.11.0.1追加
Dim gintProgramJudgeKind(0 To 8) As Integer  ' プログラム判定種別    ' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD

'選択中リソース種別 =0=MN_RSOC_PRO：プログラム、=1=MN_RSOC_HAN:判定データ
Dim iSelResource As Integer


Private Const MN_MAIL_INTERVAL = 1000       'メールタイマのインターバル値

Private Const MN_FOLD_WRK = 0               '「ワーク」フォルダ
Private Const MN_FOLD_NOW = 1               '「実行」フォルダ
Private Const MN_FOLD_OLD = 2               '「旧」フォルダ

'バージョンデータファイル用の構造体
Private Type MN_VERSION_FILE
    sFileName As String * 12                'ファイル名
    uFooter As MN_FOOT_BYTE                 'フッタ情報
End Type

Private Type MN_VERSION_DAT
    strFolder(0 To 5) As String * 8         'フォルダ名
    intFileNum(0 To 5) As Integer           'ファイル数
End Type
'バージョンデータファイル情報(バージョン2)
Private Type MN_FILE_INFO_V2
    udtInfo As MN_VERSION_DAT               'フォルダ名とファイル数
    uFileInfo() As MN_VERSION_FILE          'ファイル名とフッタ情報
End Type

Dim uVerdataFile As MN_FILE_INFO_V2

Private Const HANKUKA_KUK = "HAN_KUKA.KUK"
Private Const INI_MAX = 5
Dim HAN_KUKA_DATA As HANTEI_DATA
Private Type HANTEI_DATA
    sHederKisyu(0 To 4) As String
    sHederFile(0 To 4) As String
    sFotterKisyu(0 To 4) As String
    sFotterFile(0 To 4) As String
End Type

'V1.4.0.1　ADD　START
Private Const FILE_NAME_MAX_SIZE = 12
Private Const FILE_NAME_SIZE = 19
'【運賃データ正当性チェック異常ステータス定義】
Private sNGSts As String        'NG位置
Private sNGKoumoku As String    'NG項目
'【NG位置】
Private Const ERROR_HEDER = "ヘッダ"  'ヘッダ
Private Const ERROR_FOTTER = "フッタ" 'フッタ
'【NG項目】
Private Const KISHU_NAME_ERROR = "機種名"       '機種名
Private Const FILE_NAME_ERRORE = "ファイル名"   'ファイル名
Private Const CREATE_DATA_ERROR = "作成日付"    '作成日付
Private Const VERSION_ERROR = "バージョン"      'バージョン
Private sJverName As String                     '表示メッセージボックスタイトル
Private Const EG20_JIKAI = "EG20"               'EG20
'V1.4.0.1　ADD　END
'V1.6.0.1 ADD START
Private Const EGR_JIKAI_KISHU = "EG5000"        'EG-R自改機種名
Private Const NEG_JIKAI_KISHU = "EG2000"        'NEG自改機種名
Private Const EG20_JIKAI_KISHU = "EG6000"       'EG20 自改機種名
'V1.20.0.1 DEL START
'EG-R自改
'Private Const EHANTEI_CPU_CHK_FILE = "ko_gateh.vef"
'Private Const EMAIN_CPU_CHK_FILE = "ko_gatep.vef"
'Private Const ESUB_CPU_CHK_FILE = "ko_gatef.vef"
'Private Const EMAIN_OS_CHK_FILE = "ko_gateo.vef"
''NEG自改
'Private Const NHANTEI_CPU_CHK_FILE = "KO_GATEH.VEF"
'Private Const NMAIN_CPU_CHK_FILE = "KO_GATEP.VEF"
'Private Const NSUB_CPU_CHK_FILE = "KO_GATEF.VEF"
'Private Const NMAIN_OS_CHK_FILE = "KO_GATEO.VEF"
'V1.20.0.1 DEL END
'V1.20.0.1 ADD START
'EG-R自改
Private EHANTEI_CPU_CHK_FILE As String
Private EMAIN_CPU_CHK_FILE As String
Private ESUB_CPU_CHK_FILE As String
Private EMAIN_OS_CHK_FILE As String
'NEG自改
Private NHANTEI_CPU_CHK_FILE As String
Private NMAIN_CPU_CHK_FILE As String
Private NSUB_CPU_CHK_FILE As String
Private NMAIN_OS_CHK_FILE As String
'V1.20.0.1 ADD END
'V1.6.0.1 ADD END
'EG20 V30.1.0.1 ADD START
'EG20自改  判定用ファイル名格納エリア
Private EG20_HANTEI_CPU_CHK_FILE As String
Private EG20_MAIN_CPU_CHK_FILE As String
Private EG20_SUB_CPU1_CHK_FILE As String
Private EG20_SUB_CPU2_CHK_FILE As String
Private EG20_SUB_CPU3_CHK_FILE As String
Private EG20_MAIN_OS_CHK_FILE As String
'EG20 V30.1.0.1 ADD END

'データ種別選択
' EG20 V2.0.1.1 ADD START
Public mlngOptDataType          As Long

'フォルダ種別部
Public mlngChkFolderType        As Long

Dim mbVerKanriExecuteFlg                      As Boolean  '出力実行処理中か否か

Private iTab_index As Integer       '　選択中のコーナー番号
' EG20 V2.0.1.1 ADD END

' EG20 V3.0.0.2追加開始
Private Const TITLEDISP_VERNOTHING = "--"       ' 画面上部バージョンなし表示
Private Const TITLEDISP_FIXEDVERNOW = "                      （実行）  ："
Private Const TITLEDISP_FIXEDVEROLD = "                      （旧）    ："

Dim DispTitleBox(0 To 10) As String             ' 画面上部タイトル名（１行目）
Dim DispTitleVersion(0 To 2) As String          ' 画面上部バージョン

' EG20 V3.0.0.2追加終了

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdGateVerUpdate_Click
'//  機能名称  : 一括更新釦押下処理
'//  機能概要  : 改札機一括更新画面を表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//  ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdGateVerUpdate_Click()

    '「自改ﾊﾞｰｼﾞｮﾝ：自改切り離し釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_IKKATSU_BUTTOM, 0)

    '通信接続・切断画面を表示する。
    Load frmGateVerUpdate
    frmGateVerUpdate.Show 1

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン管理(EG20自改)画面(アクティブ時)
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
   On Error Resume Next
    
    'メール受信用のタイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : バージョン管理(EG20自改)画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
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

    'メール受信用のタイマを止める。
    tmrMail.Enabled = False
End Sub

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
Private Sub cmdInstall_Click(Index As Integer)
   On Error Resume Next
   
   If Index = 1 Then                                ' 媒体取外 処理
       '「媒体取外釦押下」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
        '媒体取外処理
        Call pfRemove(Me)
    Else                                            'バージョン情報  媒体出力処理
        '「自改ﾊﾞｰｼﾞｮﾝ：媒体出力釦押下」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OUTPUT_BUTTOM, 0)
 
        '媒体出力処理
        fMakeOutPutFile
    End If
    
ErrorHandle:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Command1_Click
'//  機能名称  : 「媒体→ワークコピー」釦押下時処理
'//  機能概要  : 媒体をワークにコピー
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] T.koyama
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Command2_Click()

   Dim iResponse As Integer         'MsgBoxボタンコード
   Dim lngErrCode As Long           'エラーコード

   On Error Resume Next

   '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ワークコピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_LZH_COPY_WRK_BUTTOM, 0)
    'インストール媒体をワークフォルダ内にコピーする
    sFDInstall "STD"
        
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2005 All Rights Reserved
'/
'/  関数名称     : Form_Load
'/  機能名称     : Form_Load時処理
'/  機能概要     : Form_Load時処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/ ORIGINAL  :(3.1.0.1) 2005-11-29   CODED   BY [TCC] A.Mizuno
'/ REVISIONS :(5.1.0.1) 2006-05-10   CODED   BY [TCC] K.Hayashi
'/ REVISIONS :(5.3.0.1) 2006-06-08   CODED   BY [TCC] K.Hayashi
'/ REVISIONS :(EG20 V2.0.1.1) 2011-11-18   CODED   BY [TCC] T.Koyama
'/ REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/             EG20フェーズ２対応
'/ REVISIONS :(EG20 V3.4.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'/             EG20フェーズ２対応（1コーナ設定で正しく表示が行えない対応）
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'/ REVISIONS :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/             北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

   Dim intCount As Integer
   Dim strCorner1 As String
   Dim strCorner2 As String
   Dim bySelectedFlg     As Byte        'EG20 V30.1.0.1 ADD
   
   On Error Resume Next
 
    sJverName = EG20_JIKAI
    
    '「EG-R自動改札機ﾊﾞｰｼﾞｮﾝ画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EJIKAI_VERASION_KANRI_GAMEN_START, 0)
  
 ' EG20 V2.0.1.1 ADD START【残件№60】
    ' フォルダ選択チェックボックス初期値設定
    For intCount = 0 To chkFolder.UBound
      chkFolder(intCount) = 1
    Next intCount
      
      '号機情報取得
    Call gsGetGateInfo
    Call gsGetCornerName
    Call gsGetCornerType        'EG20 V30.1.0.1
    
   'タブ数を設置コーナ数とする
    SSTab1.Tab = 0
'    SSTab1.Tabs = gintCornerNum            ' EG20 V3.4.0.1 削除
    bySelectedFlg = False       'EG20 V30.1.0.1 ADD
    For intCount = 0 To UBound(gblnCornerSet)
    
        '設定ありのコーナを活性にする
        If gblnCornerSet(intCount) = True Then
            'コーナー名称表示
            strCorner1 = MidB(gstrCornerName(intCount), 1, 12)
            strCorner2 = MidB(gstrCornerName(intCount), 13, 24)
            SSTab1.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
            'EG20 V30.1.0.1 ADD START
'            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
'                '幹線コーナならば押下不可にする
'                SSTab1.TabEnabled(intCount) = False
'            Else
'                '一番始めの在来線コーナーのタブを選択状態にする。
'                If bySelectedFlg = False Then
'                    SSTab1.Tab = intCount
'                    bySelectedFlg = True
'                    '在来線の先頭コーナーならばGATE00にコピーをする必要があるため、先頭インデックスを保存しておく
'                    gintZairaiFirstCornerIdx = intCount
'                End If
'            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
            
            'リストボックスを初期化する
            lstKan(intCount).Clear
        
            '画面タイトル設定
            lbltitle(intCount).Caption = "自動改札機バージョン管理"
   
' EG20 V3.0.0.2削除開始
'            '代表バージョン設定
'            lblZenVer(intCount).Caption = "判定データ　バージョン（ワーク）：  " & vbCrLf & _
'                                          "                      （実行）  ：  " & vbCrLf & _
'                                          "                      （旧）    ：  "
' EG20 V3.0.0.2削除終了
        End If
    Next

    '設定なしのコーナタブを非表示に設定する
    For intCount = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(intCount) = False Then
            SSTab1.TabVisible(intCount) = False
        End If
    Next
 ' EG20 V2.0.1.1 ADD END  【残件№60】

    'データ展開
    sSetFolderName

    '変数の初期化
    FolderSyubetu = 0

    'バージョン情報のリストボックスを作成する
    fMakeListbox

    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

End Sub

  
'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : chkFolder_Click
'/  機能名称     : 「フォルダ選択部」チェック処理
'/  機能概要     : 「フォルダ選択部」チェック処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub chkFolder_Click(Index As Integer)
  
'    Dim ValueCnt                As Integer
'
'    'ログ出力
'    If Index = 0 Then
'        'ワーク
'        Call psPutLog(LOG_frmGateVerKanri_CHKFOLDER0)
'    ElseIf Index = 1 Then
'        '実行
'        Call psPutLog(LOG_frmGateVerKanri_CHKFOLDER1)
'    ElseIf Index = 2 Then
'        '旧
'        Call psPutLog(LOG_frmGateVerKanri_CHKFOLDER2)
'    End If
'
'    '種類によって増減値を変更する
'    ValueCnt = 0
'    'ワーク
'    If Index = 0 Then
'        ValueCnt = 1
'    '実行
'    ElseIf Index = 1 Then
'        ValueCnt = 2
'    '旧
'    ElseIf Index = 2 Then
'        ValueCnt = 4
'    End If
'
'    'チェックがはずされた時
'    If chkFolder(Index).Value = 0 Then
'        mlngChkFolderType = mlngChkFolderType - ValueCnt
'    'チェックされた時
'    ElseIf chkFolder(Index).Value = 1 Then
'        mlngChkFolderType = mlngChkFolderType + ValueCnt
'    End If

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : cmdClear_Click
'/  機能名称     : 「ワーククリア」ボタン押下処理
'/  機能概要     : 「ワーククリア」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()

   Dim iResponse As Integer         'MsgBoxボタンコード
   Dim lngErrCode As Long           'エラーコード

   On Error Resume Next

    '「自改ﾊﾞｰｼﾞｮﾝ管理：ワーククリア釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_CREA_BUTTOM, 0)

    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("「ワーク」フォルダ内の " & TitleBox(FolderSyubetu) & "を、" _
           & Chr(vbKeyReturn) & "全て削除します。    よろしいですか？", _
           vbYesNo + vbExclamation, _
           TitleBox(FolderSyubetu) & "  ワーク クリア")
    If iResponse = vbYes Then
        '[はい] ボタンを選択した場合
        'ワークフォルダ内のファイルを削除する
       If sWrkFolderRemove <> True Then
          '「自改ﾊﾞｰｼﾞｮﾝ管理：ワーククリア処理異常」ログ出力
          lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_CREA_ERROR, lngErrCode)
          Exit Sub
       End If
       '「自改ﾊﾞｰｼﾞｮﾝ管理：ワーククリア処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_CREA_OK, 0)
       
       'リストボックスを初期化する
       lstKan(0).Clear
       lstKan(1).Clear
       lstKan(2).Clear
       lstKan(3).Clear
       lstKan(4).Clear
       lstKan(5).Clear
       
       'バージョン情報リストボックスを作成する
       fMakeListbox
    End If
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : cmdCopyBaitai_Work_Click
'/  機能名称     : 「媒体(圧縮)→ワーク コピー」ボタン押下処理
'/  機能概要     : 「媒体(圧縮)→ワーク コピー」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()

   On Error Resume Next

    '「自改ﾊﾞｰｼﾞｮﾝ：圧縮ﾌｧｲﾙ→ﾜｰｸｺﾋﾟｰ釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_CAB_COPY_WRK_BUTTOM, 0)

    '圧縮ファイルからインストールする。
    sFDInstall "LZH"
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : cmdCopyOld_Jikko_Click
'/  機能名称     : 「旧→実行 コピー」ボタン押下処理
'/  機能概要     : 「旧→実行 コピー」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'/                 【プログレスバー表示機能見直し対応】
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    
   Dim iResponse As Integer         'MsgBoxボタンコード
   Dim lngErrCode As Long           'エラーコード

   On Error Resume Next

   '「自改ﾊﾞｰｼﾞｮﾝ：旧→実行コピー釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OLD_COPY_NOW_BUTTOM, 0)
   '確認ポップアップウィンドウを表示する。
   iResponse = MsgBox("「旧」フォルダの内容を、「実行」フォルダに戻すことにより、" _
             & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "の一世代前のバージョンを、" _
             & Chr(vbKeyReturn) & "実行バージョンとします。  よろしいですか？", _
            vbYesNo + vbExclamation, _
            TitleBox(FolderSyubetu) & "  旧→実行 コピー")
   If iResponse = vbYes Then
   '[はい] ボタンを選択した場合
         
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
         
         '一世代前のバージョンを実行バージョンに戻す
       If fOldVersion <> True Then
          '「自改ﾊﾞｰｼﾞｮﾝ：ワーク→実行コピー処理異常」ログ出力
          lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_ERROR, lngErrCode)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
           'プログレスバーを消去する
           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
          Exit Sub
       End If
       '「自改ﾊﾞｰｼﾞｮﾝ：旧→実行コピー処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_OK, 0)
       
       'リストボックスを初期化する
       lstKan(0).Clear
       lstKan(1).Clear
       lstKan(2).Clear
       lstKan(3).Clear
       lstKan(4).Clear
       lstKan(5).Clear
      
       'バージョン情報リストボックスを作成する
       fMakeListbox
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   End If
       
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : cmdCopyWork_Jikko_Click
'/  機能名称     : 「ワーク→実行 コピー」ボタン押下処理
'/  機能概要     : 「ワーク→実行 コピー」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-27   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (3.10.0.1) 2006-02-02  CODED   BY [TCC] K.Inoue
'/  REVISIONS    : (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'/                 【プログレスバー表示機能見直し対応】
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
   
   Dim iResponse As Integer         'MsgBoxボタンコード
   Dim lngErrCode As Long           'エラーコード

   On Error Resume Next

   '「ワーク→実行コピー」ボタンの場合。
   '「自改ﾊﾞｰｼﾞｮﾝ：ワーク→実行コピー釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_COPY_NOW_BUTTOM, 0)
    
   '確認ポップアップウィンドウを表示する。
   iResponse = MsgBox("「ワーク」フォルダの内容を、「実行」フォルダに登録することにより、" _
            & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & " の最新のバージョンを、実行バージョンとします。" _
            & Chr(vbKeyReturn) & "よろしいですか？", _
           vbYesNo + vbExclamation, _
           TitleBox(FolderSyubetu) & "  ワーク→実行 コピー")
   If iResponse = vbYes Then
   '[はい] ボタンを選択した場合
            
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
            '最新バージョンを実行バージョンとして登録する
        If fNewVersion <> True Then
           '「自改ﾊﾞｰｼﾞｮﾝ：ワーク→実行コピー処理異常」ログ出力
           lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_ERROR, lngErrCode)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
           'プログレスバーを消去する
           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           Exit Sub
        End If
        '「自改ﾊﾞｰｼﾞｮﾝ：ワーク→実行コピー処理正常」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_OK, 0)
        
        'リストボックスを初期化する
        lstKan(0).Clear
        lstKan(1).Clear
        lstKan(2).Clear
        lstKan(3).Clear
        lstKan(4).Clear
        lstKan(5).Clear
        
        'バージョン情報リストボックスを作成する
        fMakeListbox
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   End If
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : cmdDLLJikkoGamen_Click
'/  機能名称     : 「DLL実行画面へ」ボタン押下処理
'/  機能概要     : 「DLL実行画面へ」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdDLLJikkoGamen_Click()

    Dim i As Integer                'カウンタ
    Dim bFlag As Boolean            'フラグ
    Dim lRetVal As Long             '戻り値
    Dim sCommand As String          'コマンド文字列
    Dim sWriteDir As String
    
    On Error GoTo ErrorHandle

    '「自改ﾊﾞｰｼﾞｮﾝ：自改切り離し釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_KIRIHANASI_BUTTOM, 0)

    '通信接続・切断画面を表示する。
    Load frmConectSts
    frmConectSts.Show 1

ErrorHandle:
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : cmdKoshin_Click
'/  機能名称     : 「表示更新」ボタン押下処理
'/  機能概要     : 「表示更新」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdKoshin_Click()
    
    Dim i As Integer                'カウンタ
    Dim bFlag As Boolean            'フラグ
    Dim lRetVal As Long             '戻り値
    Dim sCommand As String          'コマンド文字列
    Dim sWriteDir As String
    
    On Error GoTo ErrorHandle

    '「自改ﾊﾞｰｼﾞｮﾝ：表示更新釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)

    'フォルダ選択部に指定有無チェック
    bFlag = False                                 'フラグを「偽」にする
    For i = 0 To 2                                'フォルダ数分繰り返す
        If chkFolder(i).Value = CHECKBOX_ON Then   '「？？」フォルダが指定されている
            bFlag = True                            'フラグを「真」にする
            Exit For                                'ループを抜ける
        End If
    Next
              
    If bFlag = False Then                       'フォルダ指定無し
        '「表示フォルダ指定なし」ポップアップ表示
        MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                          vbOKOnly + vbExclamation, _
                          "自動改札機 バージョン管理"
        '処理を抜ける
        Exit Sub
    End If
    
    'リストボックスを初期化する
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
    
    'バージョン情報リストボックスを作成する
    fMakeListbox
              
ErrorHandle:
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : cmdModoru_Menu_Click
'/  機能名称     : メニュー画面に戻るボタン押下処理
'/  機能概要     : メニュー画面に戻るボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Menu_Click()
    
'    'ログ出力
'    Call psPutLog(LOG_frmGateVerKanri_CMDMODORU_MENU)
'
'    'メニュー画面表示
'    frmProgramHanteiData.Show

    '画面のUnload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : sOptDataDisp
'/  機能名称     : データ種別選択部表示処理
'/  機能概要     : データ種別選択部を選択されたタブ別に表示処理を行う
'/
'/                 型          名称                   意味
'/  引数         : Long        判定IC-Mメーカー選択部 クリックしたタブインデックス(1～6)
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-26   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sOptDataDisp(Index As Long)
'
'    'データ種別選択部表示
'    Dim intCnt                  As Long
'
'    'データ種別選択部を再表示する
'    For intCnt = 5 To 0 Step -1
'        If gudtMaker(Index).strType(intCnt) = "" Then
'            Me.optData(intCnt).Caption = ""
'            Me.optData(intCnt).Visible = False
'        Else
'            Me.optData(intCnt).Caption = gudtMaker(Index).strType(intCnt)
'            Me.optData(intCnt).Visible = True
'
'            'Ver1.0.0.6 ADD Start
'            mlngOptDataType = intCnt + 1
'            'Ver1.0.0.6 ADD End
'        End If
'    Next
'
'    'Ver1.0.0.6 UPD Start
'    '選択状態にする
'    Me.optData(mlngOptDataType - 1).Value = True
'    'Ver1.0.0.6 UPD End


End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : sCmdBtnEnabled
'/  機能名称     : コマンドボタン押下可・不可処理
'/  機能概要     : コマンドボタンを引数に基いて押下可・不可処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-26   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (1.0.0.5) 2005-04-06   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (3.1.0.1) 2005-12-09   CODED   BY [TCC] A.Mizuno
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)
'
'    'すべての釦押下可能
'    Me.cmdClear.Enabled = blnFlg
'    Me.cmdCopyBaitai_Work.Enabled = blnFlg
'    If blnFlg = vbTrue Then
'      Call CopyBtm_Enabled
'    Else
'      Me.cmdCopyOld_Jikko.Enabled = blnFlg
'      Me.cmdCopyWork_Jikko.Enabled = blnFlg
'    End If
'    Me.cmdDLLJikkoGamen.Enabled = blnFlg
'    Me.cmdKoshin.Enabled = blnFlg
'    Me.cmdModoru_Menu.Enabled = blnFlg
''V3.1.0.1 Add Start
'    'DLL許可画面のボタン制御追加
'    Me.cmdDLLKyokaGamen.Enabled = blnFlg
''V3.1.0.1 Add End
    
End Sub

Private Sub Form_Paint()
'    glnghwndTabCnt = gudtVerTbsInfo.lnghwndTabCnt
'    glnghwndOwnDrwTab1 = gudtVerTbsInfo.lnghwndOwnDrwTab
'    glngPrevWndProc = gudtVerTbsInfo.lngPrevWndProc
'    tbsICMVersion.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    ' サブクラス化開始
'    UnSubClass Me, gudtVerTbsInfo.lngPrevWndProc
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  関数名称     : optData_Click
'/  機能名称     : 「データ種別選択部」押下処理
'/  機能概要     : 「データ種別選択部」押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-26   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ２対応
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub optData_Click(Index As Integer)
  
    Dim i As Integer                'カウンタ
    Dim bFlag As Boolean            'フラグ

    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    'リソース種別を変更する。'
    FolderSyubetu = Index
    
' EG20 V3.0.0.2削除開始
'    ' EG20 V2.0.1.1 ADD START【残件№60】
'    Select Case FolderSyubetu           'リソース種別
'        Case 0                              '判定データ
'           lblZenVer(iTab_index).Caption = "判定データ  バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 1                              'プログラム
'           lblZenVer(iTab_index).Caption = "プログラム  バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 2                              'ｻﾌﾞCPU-Pro1
'           lblZenVer(iTab_index).Caption = "ｻﾌﾞCPU-Pro1 バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 3                              'ｻﾌﾞCPU-Pro2
'           lblZenVer(iTab_index).Caption = "ｻﾌﾞCPU-Pro2 バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 4                              'ｻﾌﾞCPU-Pro3
'           lblZenVer(iTab_index).Caption = "ｻﾌﾞCPU-Pro3 バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 5                              '自改（ＯＳ）
'           lblZenVer(iTab_index).Caption = "自改（ＯＳ）バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 6                              '予備１
'           lblZenVer(iTab_index).Caption = "予備１      バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 7                              '予備２
'           lblZenVer(iTab_index).Caption = "予備２      バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'        Case 8                              '予備１
'           lblZenVer(iTab_index).Caption = "予備３      バージョン（ワーク）：" & vbCrLf & _
'                                           "                      （実行）  ：" & vbCrLf & _
'                                           "                      （旧）    ："
'    End Select
'    ' EG20 V2.0.1.1 ADD START【残件№60】
' EG20 V3.0.0.2削除終了

    
    
    
    'フォルダ選択部に指定有無チェック
    bFlag = False                                 'フラグを「偽」にする
    For i = 0 To 2                                'フォルダ数分繰り返す
        If chkFolder(i).Value = CHECKBOX_ON Then   '「？？」フォルダが指定されている
            bFlag = True                            'フラグを「真」にする
            Exit For                                'ループを抜ける
        End If
    Next
    
    If bFlag = False Then                       'フォルダ指定無し
        '「表示フォルダ指定なし」ポップアップ表示
        MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                    vbOKOnly + vbExclamation, _
                    "自動改札機 バージョン管理"
        '処理を抜ける
        Exit Sub
    End If
    
    'リストボックスを初期化する
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
    
    'バージョン情報リストボックスを作成する
    fMakeListbox
    'V1.20.0.1 ADD END

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sVersionDisp
'//  機能名称  : バージョン情報リストボックス追加
'//  機能概要  : バージョン情報をファイル名単位でリストボックスに追加する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.4.0.1) 2012-06-17 REVISED BY [TCC] H.Sugimoto
'//                【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sVersionDisp(uVerData() As MN_VERSION_JIKAI)
    Dim sFileName As String         'ファイル名文字列（次世代自改タイプを含む）
    Dim sFileSize As String         'ファイルサイズ文字列
    Dim sFileInfo(2) As String      'バージョン情報文字列
    Dim sComment1(2) As String      'コメント文字列
    Dim sComment2(2) As String      'コメント文字列

   On Error Resume Next
    
    If uVerData(0).sFileName <> "" Then     '「ワーク」フォルダにファイルがある
        'ファイル名格納
        sFileName = StrConv(MidB(StrConv(uVerData(0).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    ElseIf uVerData(1).sFileName <> "" Then '「実行」フォルダにファイルがある
        'ファイル名格納
        sFileName = StrConv(MidB(StrConv(uVerData(1).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    Else                                    '「旧」フォルダにファイルがある
        'ファイル名格納
        sFileName = StrConv(MidB(StrConv(uVerData(2).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    End If
    sFileName = sFileName & " "

    If uVerData(0).sFileName <> "" Then     '「ワーク」フォルダにファイルがある
        'バージョン情報格納
        sFileInfo(0) = " " & StrConv(MidB(StrConv(uVerData(0).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
        sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
        sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
        sFileInfo(0) = sFileInfo(0) & uVerData(0).sVersion
        sComment1(0) = " " & StrConv(MidB(StrConv(uVerData(0).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(0) = " " & StrConv(MidB(StrConv(uVerData(0).sComment, vbFromUnicode), 33, 32), vbUnicode)
    End If
    If uVerData(1).sFileName <> "" Then     '「実行フォルダにファイルがある
        'バージョン情報格納
        sFileInfo(1) = " " & StrConv(MidB(StrConv(uVerData(1).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
        sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
        sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
        sFileInfo(1) = sFileInfo(1) & uVerData(1).sVersion
        sComment1(1) = " " & StrConv(MidB(StrConv(uVerData(1).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(1) = " " & StrConv(MidB(StrConv(uVerData(1).sComment, vbFromUnicode), 33, 32), vbUnicode)
    End If
    If uVerData(2).sFileName <> "" Then     '「旧」フォルダにファイルがある
        'バージョン情報格納
        sFileInfo(2) = " " & StrConv(MidB(StrConv(uVerData(2).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
        sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
        sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
        sFileInfo(2) = sFileInfo(2) & uVerData(2).sVersion
        sComment1(2) = " " & StrConv(MidB(StrConv(uVerData(2).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(2) = " " & StrConv(MidB(StrConv(uVerData(2).sComment, vbFromUnicode), 33, 32), vbUnicode)
    End If


    If chkFolder(0).Value = CHECKBOX_ON Then               '｢ワーク｣フォルダ表示
        If uVerData(0).sFileName <> "" Then         '｢ワーク｣フォルダにファイルはある
            If chkFolder(1).Value = CHECKBOX_ON Then       '｢実行｣フォルダ表示
                If uVerData(1).sFileName <> "" Then '｢実行｣フォルダにファイルはある
                    '｢ワーク｣フォルダと｢実行｣フォルダを比較する
                    If sFileInfo(0) = sFileInfo(1) Then
                        If chkFolder(2).Value = CHECKBOX_ON Then   '｢旧｣フォルダ表示
                            If uVerData(2).sFileName <> "" Then
                                '｢実行｣フォルダと｢旧｣フォルダを比較する
                                If sFileInfo(1) = sFileInfo(2) Then
'                                    lstKan(0).AddItem sFileName & "W N O" & sFileInfo(0)
                                    lstKan(iTab_index).AddItem sFileName & "W N O" & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
                                Else
'                                    lstKan(0).AddItem sFileName & "W N  " & sFileInfo(0)
                                    lstKan(iTab_index).AddItem sFileName & "W N  " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
'                                    lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                    lstKan(iTab_index).AddItem Space(17) & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                    End If
                                    If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '｢旧｣フォルダにファイルがない
'                                lstKan(0).AddItem sFileName & "W N  " & sFileInfo(1)
                                lstKan(iTab_index).AddItem sFileName & "W N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                     lstKan(0).AddItem Space(22) & sComment2(1)
                                     lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                                lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                            End If
                        Else                                '｢旧｣フォルダ非アクティブ表示
'                            lstKan(0).AddItem sFileName & "W N  " & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W N  " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
                        End If
                    Else                            '｢ワーク｣フォルダと｢実行｣フォルダのバージョンが違う
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】削除開始
''                        lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
'                        lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
'                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
''                            lstKan(0).AddItem Space(22) & sComment1(0)
'                            lstKan(iTab_index).AddItem Space(22) & sComment1(0)
'                        End If
'                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
''                            lstKan(0).AddItem Space(22) & sComment2(0)
'                            lstKan(iTab_index).AddItem Space(22) & sComment2(0)
'                        End If
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】削除終了
                        If chkFolder(2).Value = CHECKBOX_ON Then   '｢旧｣フォルダ表示
                            If uVerData(2).sFileName <> "" Then
                                '｢実行｣フォルダと｢旧｣フォルダを比較する
                                If sFileInfo(1) = sFileInfo(2) Then
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加開始
                                    lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加終了
'                                    lstKan(0).AddItem Space(17) & "  N O" & sFileInfo(1)
                                    lstKan(iTab_index).AddItem Space(17) & "  N O" & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                    End If
                                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                    End If
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加開始
                                ElseIf sFileInfo(0) = sFileInfo(2) Then
                                    ' 「ワーク」＝「旧」の場合
                                    lstKan(iTab_index).AddItem sFileName & "W   O" & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
                                    lstKan(iTab_index).AddItem Space(17) & "  N  " & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                    End If
                                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                    End If
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加終了
                                Else
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加開始
                                    ' 「ワーク」≠ 「実行」 ≠「旧」の場合
                                    lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加終了
'                                    lstKan(0).AddItem Space(17) & "  N  " & sFileInfo(1)
                                    lstKan(iTab_index).AddItem Space(17) & "  N  " & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                    End If
                                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                    End If
'                                    lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                    lstKan(iTab_index).AddItem Space(17) & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                    End If
                                    If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '｢旧｣フォルダにファイルがない
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加開始
                                ' 「ワーク」≠ 「実行」 ≠「旧」の場合
                                lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                End If
                                If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                End If
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加終了
'                                lstKan(0).AddItem Space(17) & "  N  " & sFileInfo(1)
                                lstKan(iTab_index).AddItem Space(17) & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                                lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                            End If
                        Else
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加開始
                            ' 「ワーク」≠ 「実行」 ≠「旧」の場合
                            lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
' EG20 V6.4.0.1【総点検修正対応：ワーク≠実行、ワーク＝旧の場合の表示不正】追加終了
'                            lstKan(0).AddItem Space(17) & "  N  " & sFileInfo(1)
                            lstKan(iTab_index).AddItem Space(17) & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
                        End If
                    End If
                Else                                    '｢実行｣フォルダにファイルがない
                    If chkFolder(2).Value = CHECKBOX_ON Then   '｢旧｣フォルダ表示
                        If uVerData(2).sFileName <> "" Then
                            If sFileInfo(0) = sFileInfo(2) Then
'                                lstKan(0).AddItem sFileName & "W   O" & sFileInfo(0)
                                lstKan(iTab_index).AddItem sFileName & "W   O" & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                End If
                                If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                End If
'                                lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                                lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                            Else
'                                lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                                lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                End If
                                If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                lstKan(iTab_index).AddItem Space(17) & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                End If
                                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                End If
'                                lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                                lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                            End If
                        Else                            '｢旧｣フォルダにファイルがない
'                            lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
'                            lstKan(0).AddItem Space(17) & "  N O" & " -------- --------  -------- ----"
                            lstKan(iTab_index).AddItem Space(17) & "  N O" & " -------- --------  -------- ----"
                        End If
                    Else                                '｢旧｣フォルダ非アクティブ表示
'                        lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                        lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                        End If
'                        lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                    End If
                End If
            Else                                        '｢実行｣フォルダ非アクティブ表示
                If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                    If uVerData(2).sFileName <> "" Then
                        If sFileInfo(0) = sFileInfo(2) Then
'                            lstKan(0).AddItem sFileName & "W   O" & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W   O" & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
                        Else
'                            lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
'                            lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                            lstKan(iTab_index).AddItem Space(17) & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                    '｢旧｣フォルダにファイルがない
'                        lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                        lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                        End If
'                        lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
'                    lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                    lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment1(0)
                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                    End If
                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment2(0)
                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                    End If
                End If
            End If
        Else                                '｢ワーク｣フォルダにファイルがない
            If chkFolder(1).Value = CHECKBOX_ON Then               '｢実行｣フォルダ表示
                If uVerData(1).sFileName <> "" Then         '｢実行｣フォルダにファイルはある
                    If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                        If uVerData(2).sFileName <> "" Then '｢旧｣フォルダにファイルはある
                            '｢実行｣フォルダと｢旧｣フォルダを比較する
                            If sFileInfo(1) = sFileInfo(2) Then
'                                lstKan(0).AddItem sFileName & "  N O" & sFileInfo(1)
                                lstKan(iTab_index).AddItem sFileName & "  N O" & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                                lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                            Else                            '｢実行｣フォルダと｢旧｣フォルダのバージョンが違う
'                                lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                                lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                lstKan(iTab_index).AddItem Space(17) & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                End If
                                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                End If
'                                lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                                lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                            End If
                        Else                                '｢旧｣フォルダにファイルはない
'                            lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                            lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
'                            lstKan(0).AddItem Space(17) & "W   O" & " -------- --------  -------- ----"
                            lstKan(iTab_index).AddItem Space(17) & "W   O" & " -------- --------  -------- ----"
                        End If
                    Else                                    '｢旧｣フォルダ非アクティブ表示
'                        lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                        lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                        End If
                        If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                        End If
'                        lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                    End If
                Else                                        '｢実行｣フォルダにファイルがない
                    If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                        If uVerData(2).sFileName <> "" Then
'                            lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                            lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                            End If
'                            lstKan(0).AddItem Space(17) & "W N  " & " -------- --------  -------- ----"
                            lstKan(iTab_index).AddItem Space(17) & "W N  " & " -------- --------  -------- ----"
                        Else                                '｢旧｣フォルダにファイルがない
'                            lstKan(0).AddItem sFileName & "W N O" & " -------- --------  -------- ----"
                            lstKan(iTab_index).AddItem sFileName & "W N O" & " -------- --------  -------- ----"
                        End If
                    Else                                    '｢旧｣フォルダ非アクティブ表示
'                        lstKan(0).AddItem sFileName & "W N  " & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem sFileName & "W N  " & " -------- --------  -------- ----"
                    End If
                End If
            Else                                        '｢実行｣フォルダ非アクティブ表示
                If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                    If uVerData(2).sFileName <> "" Then
'                        lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                        lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                        End If
                        If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                        End If
'                        lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                    Else                                '｢旧｣フォルダにファイルがない
'                        lstKan(0).AddItem sFileName & "W   O" & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem sFileName & "W   O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
'                    lstKan(0).AddItem sFileName & "W    " & " -------- --------  -------- ----"
                    lstKan(iTab_index).AddItem sFileName & "W    " & " -------- --------  -------- ----"
                End If
            End If
        End If
    Else                                                '｢ワーク｣フォルダ非アクティブ表示
        If chkFolder(1).Value = CHECKBOX_ON Then               '｢実行｣フォルダ表示
            If uVerData(1).sFileName <> "" Then         '｢実行｣フォルダにファイルはある
                If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                    If uVerData(2).sFileName <> "" Then '｢旧｣フォルダにファイルはある
                        '｢実行｣フォルダと｢旧｣フォルダを比較する
                        If sFileInfo(1) = sFileInfo(2) Then
'                            lstKan(0).AddItem sFileName & "  N O" & sFileInfo(1)
                            lstKan(iTab_index).AddItem sFileName & "  N O" & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
                        Else
'                            lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                            lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
'                            lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                            lstKan(iTab_index).AddItem Space(17) & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                '｢旧｣フォルダにファイルはない
'                        lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                        lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                        End If
                        If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                        End If
'                        lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
'                    lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                    lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment1(1)
                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                    End If
                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment2(1)
                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                    End If
                End If
            Else                                        '｢実行｣フォルダにファイルがない
                If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                    If uVerData(2).sFileName <> "" Then
'                        lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                        lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                        End If
                        If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                        End If
'                        lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                    Else                                '｢旧｣フォルダにファイルがない
'                        lstKan(0).AddItem sFileName & "  N O" & " -------- --------  -------- ----"
                        lstKan(iTab_index).AddItem sFileName & "  N O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
'                    lstKan(0).AddItem sFileName & "  N  " & " -------- --------  -------- ----"
                    lstKan(iTab_index).AddItem sFileName & "  N  " & " -------- --------  -------- ----"
                End If
            End If
        Else                                    '｢実行｣フォルダ非アクティブ表示
            If uVerData(2).sFileName <> "" Then '｢旧｣フォルダにファイルはある
'                lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                    lstKan(0).AddItem Space(22) & sComment1(2)
                    lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                End If
                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
'                    lstKan(0).AddItem Space(22) & sComment2(2)
                    lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                End If
            Else                                '｢旧｣フォルダにファイルがない
'                lstKan(0).AddItem sFileName & "    O" & " -------- --------  -------- ----"
                lstKan(iTab_index).AddItem sFileName & "    O" & " -------- --------  -------- ----"
            End If
        End If
    End If
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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
   On Error Resume Next
    
    '汎用メール受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
'        AppActivate frmJVer.Caption, False                 ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
        AppActivate frmGateVerKanri.Caption, False          ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
        pfFormActive (frmGateVerKanri.hwnd)                 ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSetFolderName
'//  機能名称  : データ展開
'//  機能概要  : フォルダ名などのデータをグローバルエリアに展開する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                ワーク→実行コピーでの正当性チェックINI読込み
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【種別チェック機能追加】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSetFolderName()

        TitleBox(0) = "判定データ"
        TitleBox(1) = "プログラム"
        TitleBox(2) = "ｻﾌﾞCPU-Pro1"
        TitleBox(3) = "ｻﾌﾞCPU-Pro2"
        TitleBox(4) = "ｻﾌﾞCPU-Pro3"
        TitleBox(5) = "自改（ＯＳ）"
        TitleBox(6) = "予備１"
        TitleBox(7) = "予備２"
        TitleBox(8) = "予備３"
    
        LogBox(0) = "判定"
        LogBox(1) = "プログメイン"
        LogBox(2) = "サブ1"
        LogBox(3) = "サブ2"
        LogBox(4) = "サブ3"
        LogBox(5) = "OS"
        LogBox(5) = "予備1"
        LogBox(5) = "予備2"
        LogBox(5) = "予備3"
        
        'フォルダ名に設定を行う
        FolderName(0, 0) = EG20_NHAN1WRK
        FolderName(1, 0) = EG20_NHAN1NOW
        FolderName(2, 0) = EG20_NHAN1OLD
        FolderName(0, 1) = EG20_NPRO1WRK
        FolderName(1, 1) = EG20_NPRO1NOW
        FolderName(2, 1) = EG20_NPRO1OLD
        FolderName(0, 2) = EG20_NSCP1WRK
        FolderName(1, 2) = EG20_NSCP1NOW
        FolderName(2, 2) = EG20_NSCP1OLD
        FolderName(0, 3) = EG20_NSCP2WRK
        FolderName(1, 3) = EG20_NSCP2NOW
        FolderName(2, 3) = EG20_NSCP2OLD
        FolderName(0, 4) = EG20_NSCP3WRK
        FolderName(1, 4) = EG20_NSCP3NOW
        FolderName(2, 4) = EG20_NSCP3OLD
        FolderName(0, 5) = EG20_NOSWRK
        FolderName(1, 5) = EG20_NOSNOW
        FolderName(2, 5) = EG20_NOSOLD
        FolderName(0, 6) = EG20_NYOBI1WRK
        FolderName(1, 6) = EG20_NYOBI1NOW
        FolderName(2, 6) = EG20_NYOBI1OLD
        FolderName(0, 7) = EG20_NYOBI2WRK
        FolderName(1, 7) = EG20_NYOBI2NOW
        FolderName(2, 7) = EG20_NYOBI2OLD
' EG20 V5.11.0.1追加開始
        FolderName(0, 8) = EG20_NYOBI3WRK
        FolderName(1, 8) = EG20_NYOBI3NOW
        FolderName(2, 8) = EG20_NYOBI3OLD
' EG20 V5.11.0.1追加終了
' EG20 V5.11.0.1削除開始
'        FolderName(0, 8) = EG20_NYOBI2WRK
'        FolderName(1, 8) = EG20_NYOBI2NOW
'        FolderName(2, 8) = EG20_NYOBI2OLD
' EG20 V5.11.0.1削除終了

' EG20 V3.0.0.2追加開始
        DispTitleBox(0) = "判定データ  バージョン（ワーク）："
        DispTitleBox(1) = "プログラム  バージョン（ワーク）："
        DispTitleBox(2) = "ｻﾌﾞCPU-Pro1 バージョン（ワーク）："
        DispTitleBox(3) = "ｻﾌﾞCPU-Pro2 バージョン（ワーク）："
        DispTitleBox(4) = "ｻﾌﾞCPU-Pro3 バージョン（ワーク）："
        DispTitleBox(5) = "自改（ＯＳ）バージョン（ワーク）："
        DispTitleBox(6) = "予備１      バージョン（ワーク）："
        DispTitleBox(7) = "予備２      バージョン（ワーク）："
        DispTitleBox(8) = "予備３      バージョン（ワーク）："
' EG20 V3.0.0.2追加終了


'EG20 V30.1.0.1 DEL START
''V1.20.0.1 ADD START
''-------EG-R自改-------
'    ' キー名:判定CPU-PRO代表
'    EHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
'
'    ' キー名:メインCPU-PRO代表
'    EMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_PRO, PATH_GATEVER_FILE)
'
'    ' キー名：サブCPU-PRO代表
'    ESUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_SUB_PRO, PATH_GATEVER_FILE)
'
'    ' キー名:メインCPU-OS代表
'    EMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_OS, PATH_GATEVER_FILE)
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
    ' キー名:判定CPU-PRO代表
    EG20_HANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    
    ' キー名:メインCPU-PRO代表
    EG20_MAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_MAIN_PRO, PATH_GATEVER_FILE)
        
    ' キー名：サブCPU-PRO代表
    EG20_SUB_CPU1_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO1, PATH_GATEVER_FILE)
    EG20_SUB_CPU2_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO2, PATH_GATEVER_FILE)
    EG20_SUB_CPU3_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO3, PATH_GATEVER_FILE)
    
    ' キー名:メインCPU-OS代表
    EG20_MAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_MAIN_OS, PATH_GATEVER_FILE)
'EG20 V30.1.0.1 ADD END
    
''-------NEG自改-------
'    ' キー名:判定CPU-PRO代表
'    NHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
'
'    ' キー名:メインCPU-PRO代表
'    NMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_PRO, PATH_GATEVER_FILE)
'
'    ' キー名：サブCPU-PRO代表
'    NSUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_SUB_PRO, PATH_GATEVER_FILE)
'
'    ' キー名:メインCPU-OS代表
'    NMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_OS, PATH_GATEVER_FILE)
''V1.20.0.1 ADD END

' EG20 V5.11.0.1【運改表示改善対応】追加開始
    gintUnkaiKind(0) = BootInfoGateType.TYPE_NHAN
    gintUnkaiKind(1) = BootInfoGateType.TYPE_NPRO
    gintUnkaiKind(2) = BootInfoGateType.TYPE_NSCP1
    gintUnkaiKind(3) = BootInfoGateType.TYPE_NSCP2
    gintUnkaiKind(4) = BootInfoGateType.TYPE_NSCP3
    gintUnkaiKind(5) = BootInfoGateType.TYPE_NOS
    gintUnkaiKind(6) = BootInfoGateType.TYPE_NYOBI1
    gintUnkaiKind(7) = BootInfoGateType.TYPE_NYOBI2
    gintUnkaiKind(8) = BootInfoGateType.TYPE_NYOBI3
' EG20 V5.11.0.1【運改表示改善対応】追加終了

' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD START
    gintProgramJudgeKind(0) = ProgramJudgeKind.JUDGE_NHAN       ' 判定データ
    gintProgramJudgeKind(1) = ProgramJudgeKind.JUDGE_NPRO       ' プログラム
    gintProgramJudgeKind(2) = ProgramJudgeKind.JUDGE_NSCP1      ' サブCPU-Pro1
    gintProgramJudgeKind(3) = ProgramJudgeKind.JUDGE_NSCP2      ' サブCPU-Pro2
    gintProgramJudgeKind(4) = ProgramJudgeKind.JUDGE_NSCP3      ' サブCPU-Pro3
    gintProgramJudgeKind(5) = ProgramJudgeKind.JUDGE_NOS        ' 自改（OS）
    gintProgramJudgeKind(6) = ProgramJudgeKind.JUDGE_NOCHECK    ' 予備1
    gintProgramJudgeKind(7) = ProgramJudgeKind.JUDGE_NOCHECK    ' 予備2
    gintProgramJudgeKind(8) = ProgramJudgeKind.JUDGE_NOCHECK    ' 予備3
' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMakeListbox
'//  機能名称  : バージョン情報リストボックス作成
'//  機能概要  : 各フォルダからバージョン取得を行い、リストボックス作成
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeListbox() As Boolean
    
    Dim bRet As Boolean                        '戻り値
    
    Dim sCorner As String                      'コーナー番号
    Dim sGatePath As String                    'コーナー番号付ファイルパス
    Dim sFilePath As String                    'ファイルファイルパス
    Dim i As Integer                           'ループカウンタ
    Dim sWorkVer As String                      ' ワークバージョン
    Dim sNowVer As String                       ' 現行バージョン
    Dim sOldVer As String                       ' 旧バージョン

    On Error Resume Next

    sWorkVer = TITLEDISP_VERNOTHING
    sNowVer = TITLEDISP_VERNOTHING
    sOldVer = TITLEDISP_VERNOTHING
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab

    sCorner = Format(iTab_index + 1, "00")

    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

    '***********************************************
    '* 次世代自改フォルダから全てのバージョン情報を取得する *
    '***********************************************

    ReDim uVersion(0)

    '｢ワーク｣フォルダからファイルリストを取得する
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    If bRet = True Then
        'ファイルリストからバージョン情報を取得する
        sWorkVer = sVersionInfo(sFilePath, MN_FLDWRK)
    End If

    '｢実行｣フォルダからファイルリストを取得する
    sFilePath = sGatePath & FolderName(1, FolderSyubetu)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    If bRet = True Then
        'ファイルリストからバージョン情報を取得する
        sNowVer = sVersionInfo(sFilePath, MN_FLDNOW)
    End If

    '｢旧｣フォルダからファイルリストを取得する
    sFilePath = sGatePath & FolderName(2, FolderSyubetu)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    If bRet = True Then
        'ファイルリストからバージョン情報を取得する
        sOldVer = sVersionInfo(sFilePath, MN_FLDOLD)
    End If

    'バージョン情報をファイル名順にソートする
    sListboxSort

    'バージョン情報をリストボックスにセットする
    Call sVerListDisp(sWorkVer, sNowVer, sOldVer)

End Function

' EG20 V3.0.0.2 削除開始
'Private Function fMakeListbox() As Boolean
'
'    Dim bRet As Boolean                        '戻り値
'
'    Dim sCorner As String                      'コーナー番号
'    Dim sGatePath As String                    'コーナー番号付ファイルパス
'    Dim sFilePath As String                    'ファイルファイルパス
'    Dim i As Integer                           'ループカウンタ
'
'    On Error Resume Next
'
''    ' 選択中のコーナー番号取得
''    iTab_index = SSTab1.Tab
''
''    sCorner = Format(iTab_index + 1, "00")
''
''    ' コーナー番号付ファイルパス作成
''    sGatePath = PATH_N_GATE & sCorner
'
'    '***********************************************
'    '* 次世代自改フォルダから全てのバージョン情報を取得する *
'    '***********************************************
'    For i = 0 To 5
'
'        iTab_index = i
'
'        sCorner = Format(iTab_index + 1, "00")
'
'        ' コーナー番号付ファイルパス作成
'        sGatePath = PATH_N_GATE & sCorner
'
'        ReDim uVersion(0)
'
'        '｢ワーク｣フォルダからファイルリストを取得する
'        sFilePath = sGatePath & FolderName(0, FolderSyubetu)
'
''       bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
'        bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
'        If bRet = True Then
'            'ファイルリストからバージョン情報を取得する
''            sVersionInfo FolderName(0, FolderSyubetu), MN_FLDWRK
'            sVersionInfo sFilePath, MN_FLDWRK
'        End If
'
'        '｢実行｣フォルダからファイルリストを取得する
'        sFilePath = sGatePath & FolderName(1, FolderSyubetu)
'
''       bRet = fReadFileList(FolderName(1, FolderSyubetu) & "\" & MN_FILELIST)
'        bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
'        If bRet = True Then
'            'ファイルリストからバージョン情報を取得する
''           sVersionInfo FolderName(1, FolderSyubetu), MN_FLDNOW
'            sVersionInfo sFilePath, MN_FLDNOW
'        End If
'
'        '｢旧｣フォルダからファイルリストを取得する
'        sFilePath = sGatePath & FolderName(2, FolderSyubetu)
'
''       bRet = fReadFileList(FolderName(2, FolderSyubetu) & "\" & MN_FILELIST)
'        bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
'        If bRet = True Then
'            'ファイルリストからバージョン情報を取得する
''           sVersionInfo FolderName(2, FolderSyubetu), MN_FLDOLD
'            sVersionInfo sFilePath, MN_FLDOLD
'        End If
'
'        'バージョン情報をファイル名順にソートする
'        sListboxSort
'
'        'バージョン情報をリストボックスにセットする
'        sVerListDisp
'
'    Next i
'End Function
' EG20 V3.0.0.2 削除終了
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sVerListDisp
'//  機能名称  : バージョン情報リストボックス設定
'//  機能概要  : 取得したバージョン情報を、リストボックスに設定
'//
'//              型        名称             意味
'//  引数      : String    szWorkVersion    ワークバージョン
'//  引数      : String    szNowVersion     実行バージョン
'//  引数      : String    szOldVersion     旧バージョン
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'/  REVISIONS    : (EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sVerListDisp()                 ' EG20 V3.0.0.2削除
' EG20 V3.0.0.2追加開始
Private Sub sVerListDisp(szWorkVersion As String, _
                            szNowVersion As String, _
                            szOldVersion As String)
' EG20 V3.0.0.2追加終了

    Dim i As Integer                        'カウンタ
    Dim uVerData(2) As MN_VERSION_JIKAI     'バージョン情報（各フォルダ）
    Dim lDataNum As Long                    'バージョン情報数
    Dim szWorkBuffer As String              ' ワークバッファ        ' EG20 V3.0.0.2追加
    Dim szTitleBuffer As String             ' ワークバッファ        ' EG20 V3.0.0.2追加

    On Error Resume Next

'    'リストボックスを初期化する
'    lstKan(0).Clear
'    lstKan(1).Clear
'    lstKan(2).Clear
'    lstKan(3).Clear
'    lstKan(4).Clear
'    lstKan(5).Clear

    lDataNum = UBound(uVersion)             'バージョン情報数取得
    For i = 1 To lDataNum

        uVerData(0).sFileName = ""          'ファイル名をクリアする
        uVerData(1).sFileName = ""          'ファイル名をクリアする
        uVerData(2).sFileName = ""          'ファイル名をクリアする

        Select Case uVersion(i).iFolder     'フォルダ名を対象とする
        Case MN_FLDWRK                      '「ワーク」フォルダの場合
            uVerData(0) = uVersion(i)       '「ワーク」フォルダ内に格納する
            If i + 1 <= lDataNum Then       '次のデータがある?
                If uVersion(i).sFileName = uVersion(i + 1).sFileName Then
                                                        'ファイル名が同じ?
                    Select Case uVersion(i + 1).iFolder 'フォルダ名を対象とする
                    Case MN_FLDNOW                      '「実行」フォルダの場合
                        uVerData(1) = uVersion(i + 1)   '「実行」フォルダ内に格納する
                        If i + 2 <= lDataNum Then       '次のデータがある?
                            If uVersion(i + 1).sFileName = uVersion(i + 2).sFileName Then
                                                        'ファイル名が同じ?
                                uVerData(2) = uVersion(i + 2)
                                                        '「旧」フォルダ内に格納する
                                i = i + 2               'カウンタを次々にする
                            Else
                                i = i + 1               'カウンタを次にする
                            End If
                        Else
                            i = i + 1                   'カウンタを次にする
                        End If
                    Case MN_FLDOLD                      '「旧」フォルダの場合
                        uVerData(2) = uVersion(i + 1)   '「旧」フォルダ内に格納する
                        i = i + 1                       'カウンタを次にする
                    End Select
                End If
            End If
        Case MN_FLDNOW                      '「実行」フォルダの場合
            uVerData(1) = uVersion(i)       '「実行」フォルダ内に格納する
            If i + 1 <= lDataNum Then       '次のデータがある
                If uVersion(i).sFileName = uVersion(i + 1).sFileName Then
                                                    'ファイル名が同じ?
                    uVerData(2) = uVersion(i + 1)   '「旧」フォルダ内に格納する
                    i = i + 1                       'カウンタを次にする
                End If
            End If
        Case MN_FLDOLD                      '「旧」フォルダの場合
            uVerData(2) = uVersion(i)       '「旧」フォルダ内に格納する
        End Select
        'ファイル名をまとめてリストボックスに設定
        sVersionDisp uVerData()
    Next

' EG20 V3.0.0.2追加開始
    ' ワーク行編集
    szWorkBuffer = DispTitleBox(FolderSyubetu) & szWorkVersion & vbCrLf
    szTitleBuffer = szWorkBuffer
    ' 実行行編集
    szWorkBuffer = TITLEDISP_FIXEDVERNOW & szNowVersion & vbCrLf
    szTitleBuffer = szTitleBuffer & szWorkBuffer
    ' 旧行編集
    szWorkBuffer = TITLEDISP_FIXEDVEROLD & szOldVersion
    szTitleBuffer = szTitleBuffer & szWorkBuffer

    lblZenVer(iTab_index).Caption = szTitleBuffer

    DispTitleVersion(MN_FOLD_WRK) = szWorkVersion
    DispTitleVersion(MN_FOLD_NOW) = szNowVersion
    DispTitleVersion(MN_FOLD_OLD) = szOldVersion
' EG20 V3.0.0.2追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSetChkFile
'//  機能名称  : ワーク→実行コピーで使用する正当性チェックINI読込み
'//  機能概要  : INIファイルにの内容をエリアに展開する。
'//
'//              型        名称      意味
'//  引数      : String    セクション名
'//              String    キー名
'//              String    ファイル名
'//
'//              型        値        意味
'//  戻り値    : String    正当性チェックINIの内容（異常時はブランク）
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sSetChkFile(sSec As String, sKey As String, sFilePath As String) As String

    Dim iRet As Integer             '関数の戻り値
    Dim sIni_Data As String * 128   'INIファイルより1行分取得
    
    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード

    
    'エラールーチンを宣言
    On Error Resume Next

    'iniファイル取得
    sIni_Data = ""
    iRet = GetPrivateProfileString(sSec, sKey, DEFAILT, sIni_Data, Len(sIni_Data), sFilePath)
    
    '異常処理
    If iRet = 0 Then
        
        'ログ出力「INIファイル読込異常」
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_INI_READ, lngErrCode)
        'ログ出力　┗ファイル名
        Call psFileNameGet(sFilePath, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
        'ログ出力　┗キー名
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗Key:" & sKey, lngErrCode)
        
    End If
    
    sSetChkFile = Left$(sIni_Data, iRet)
    
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fReadFileList
'//  機能名称  : ファイルリストの取得
'//  機能概要  : ファイルリストより、ファイル名を取得する。
'//
'//              型        名称      意味
'//  引数      : String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fReadFileList(sFileList As String) As Boolean
    Dim iFileNumber As Integer      'ファイル番号
    Dim sFileName As String         'ファイル名
    Dim iListCnt As Integer         'ファイル格納数

    On Error GoTo ErrorHandler      'エラーハンドル設定

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '未使用のファイル番号を取得する

    Open sFileList For Input Access Read As #iFileNumber    'ファイルリストのオープン
    Do While Not EOF(iFileNumber)                           'ファイルの終端までループを繰り返します。
        Line Input #iFileNumber, sFileName                  'データを読み込みます。
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                'ファイル名が存在する
            iListCnt = iListCnt + 1                         'ファイル数のカウンタをアップする
            ReDim Preserve FileList(iListCnt)               'ファイル名格納エリアを拡張する
            ReDim Preserve FileListType(iListCnt)           'ファイル名格納エリアを拡張する
            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
                                                            'ファイル名をファイル名格納エリアにセット
        End If
    Loop
    Close #iFileNumber      'ファイルを閉じます。

    fReadFileList = True    '戻り値を正常とする

    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    'V1.21.0.1 ADD  START
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    'V1.21.0.1 ADD  END
    fReadFileList = False   '戻り値をエラーとする
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sVersionInfo
'//  機能名称  : バージョン情報の取得
'//  機能概要  : ファイルリスト一覧からバージョン情報を取得する。
'//
'//              型        名称      意味
'//  引数      : String　　sPath
'//  　　　    : Integer　 iFolder
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sVersionInfo(sPath As String, iFolder As Integer)                  ' EG20 V3.0.0.2削除
Private Function sVersionInfo(sPath As String, iFolder As Integer) As String    ' EG20 V3.0.0.2追加
    Dim i As Integer                    'カウンタ
    Dim j As Integer                    'カウンタ
    Dim sMyName As String               'ファイル名
    Dim iFileNumber As Integer          'ファイル番号
    Dim lLen As Long                    'ファイルサイズ
    Dim uFooter As MN_FOOT              'フッタ情報格納エリア
    Dim lPos As Long                    'バージョン情報格納位置
    Dim sDateTime As String
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD

    Dim szResultVersion As String        ' 出力バージョン               ' EG20 V3.0.0.2追加

    szResultVersion = TITLEDISP_VERNOTHING                              ' EG20 V3.0.0.2追加
   On Error Resume Next

    For i = 0 To UBound(FileList) - 1   'ファイルリスト数

        sMyName = sPath & "\" & FileList(i)     'ファイルフルパス名の作成

        'If Dir(sMyName) <> "" Then              'ファイルが存在する?    'V1.20.0.1 DEL
        If objFso.FileExists(sMyName) = True Then  'ファイルが存在する?    'V1.20.0.1 ADD
            lLen = FileLen(sMyName)             'ファイルサイズの取得

            iFileNumber = FreeFile              '未使用のファイル番号を取得する

            Open sMyName For Binary Access Read As #iFileNumber
                                                'ファイルのオープン
            Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
                                                'フッタ情報の取得
            ReDim Preserve uVersion(UBound(uVersion) + 1)
                                                'バージョン情報格納エリアの拡張
            lPos = UBound(uVersion)             'バージョン情報格納位置の取得
            uVersion(lPos).sFileName = UCase(FileListType(i))       'ファイル名を大文字にしてセット
            uVersion(lPos).iFolder = iFolder                    'フォルダ名セット
            uVersion(lPos).sMachineName = uFooter.sKisyu        '機種名セット
            uVersion(lPos).sFooterFile = uFooter.sFileName      'ファイル名セット

            sDateTime = ""
            For j = 0 To 3
                sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
            Next
            sDateTime = sDateTime & " "
            For j = 4 To 5
                sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
            Next
            uVersion(lPos).sFileDate = sDateTime
            uVersion(lPos).sVersion = uFooter.sVersion          'バージョン情報セット
            uVersion(lPos).sComment = uFooter.sHyoji            '表示文字列セット

' EG20 V3.0.0.2追加開始
            ' ファイルリストの先頭で、かつ最初に見つかったファイルのバージョンを設定
            If szResultVersion = TITLEDISP_VERNOTHING Then
                szResultVersion = uFooter.sVersion
            End If
' EG20 V3.0.0.2追加終了

            Close #iFileNumber                  'ファイルを閉じます
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD

    sVersionInfo = szResultVersion              ' EG20 V3.0.0.2追加

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sListboxSort
'//  機能名称  : バージョン情報のソート
'//  機能概要  : バージョン情報をファイル名順にソートする。
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
Private Sub sListboxSort()
    Dim i As Integer                'カウンタ
    Dim j As Integer                'カウンタ
    Dim uBuff As MN_VERSION_JIKAI   'バージョン情報格納バッファ

    On Error Resume Next
   
    For i = 1 To UBound(uVersion) - 1
        For j = i + 1 To UBound(uVersion)
            'ファイル名の比較を行う
            If uVersion(j).sFileName < uVersion(i).sFileName Then
                'ファイル名が小さければ移し替える
                uBuff = uVersion(i)
                uVersion(i) = uVersion(j)
                uVersion(j) = uBuff
            ElseIf uVersion(j).sFileName = uVersion(i).sFileName Then
                'フォルダの比較を行う
                If uVersion(j).iFolder = MN_FLDWRK And uVersion(i).iFolder = MN_FLDNOW Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                ElseIf uVersion(j).iFolder = MN_FLDNOW And uVersion(i).iFolder = MN_FLDOLD Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                ElseIf uVersion(j).iFolder = MN_FLDWRK And uVersion(i).iFolder = MN_FLDOLD Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                End If
            End If
        Next
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psVersionDisp
'//  機能名称  : バージョン情報表示処理
'//  機能概要  : バージョン情報表示部の表示処理を行う。
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
'Public Sub psVersionDisp()
'
'    Dim strFilePath     As String   'バージョンファイルパス
'    Dim bRet            As Boolean  '戻り値
'    Dim intFileNo       As Integer  'ファイル番号
'    Dim strWork         As String   '作業エリア
'    Dim strVerData      As String   '全体バージョン
'    Dim intCnt          As Integer  'カウンター
'    Dim lngErrCode      As Long     'エラーコード
'
''*******************************
''VBエラー処理
'On Error GoTo Error_psVersionDisp
''*******************************
'
'    '媒体出力釦押下不可
'    cmdOutput.Enabled = False
'
'    'リスト初期化
'    LstFile.Clear
'
'    '全体バージョン初期化
'    lblZenVer.Caption = "全体バージョン（ワーク）:--.--.--.--" & vbCrLf & _
'                        "　　　　　　　（実行）　:--.--.--.--" & vbCrLf & _
'                        "　　　　　　　（旧）    :--.--.--.--"
'
'    '作業エリア初期化
'    strWork = ""
'
'    '全体バージョン初期化
'    strVerData = ""
'
'    'LDユーティリティ画面表示用バージョンファイルパス作成
'    strFilePath = PATH_LDU_APP & PATH_LDU_WORK & LDU_VER_FILE
'
'    bRet = True
'    '///////////////////////////////////////////////////////////////////////////////////////////
'    '/ 共通DA:LDユーティリティ画面表示用バージョンファイル作成
'    '///////////////////////////////////////////////////////////////////////////////////////////
'    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, strFilePath, PATH_LDU_APP)
'
'    'LDユーティリティ画面表示用バージョンファイル作成成功
'    If bRet Then
'       '「LDユーティリティバージョン管理画面：バージョン情報ファイル作成正常」ログ出力
'       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
'    'LDユーティリティ画面表示用バージョンファイル作成失敗
'    Else
'       '「LDユーティリティバージョン管理画面：バージョン情報ファイル作成異常」ログ出力
'       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'       Exit Sub
'    End If
'
'    'LDユーティリティ画面表示用バージョンファイルの有無確認
'    If Len(Trim(Dir(strFilePath))) = 0 Then
'        Exit Sub
'    End If
'
'    'LDユーティリティ画面表示用バージョンファイルのファイル番号を取得する。
'    intFileNo = FreeFile
'
'    'LDユーティリティ画面表示用バージョンファイルオープン
'    Open strFilePath For Input As #intFileNo
'
'
'        'ワーク
'        Line Input #intFileNo, strWork
'
'        If (Trim(strWork) = "") Then
'            strVerData = "全体バージョン（ワーク）：--.--.--.--" & vbCrLf
'        Else
'            '全体バージョン文字列作成
'            strVerData = strVerData & strWork & vbCrLf
'        End If
'
'        '実行
'        Line Input #intFileNo, strWork
'        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "　　　　　　　（実行）　：--.--.--.--" & vbCrLf
'        Else
'            '全体バージョン文字列作成
'            strVerData = strVerData & strWork & vbCrLf
'        End If
'
'        '旧
'        Line Input #intFileNo, strWork
'        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "　　　　　　　（旧）    ：--.--.--.--" & vbCrLf
'        Else
'            '全体バージョン文字列作成
'            strVerData = strVerData & strWork & vbCrLf
'        End If
'
'        '全体バージョン出力
'        lblZenVer.Caption = strVerData
'
'        strWork = ""
'
'        'リスト表示分読み込み（ファイル終端までループを繰り返す）
'        Do While Not EOF(1)
'
'            Line Input #intFileNo, strWork
'
'            '改行コードのみは読みとばす
'            If Trim(strWork) <> "" Then
'
'                'リストに出力
'                LstFile.AddItem (strWork)
'
'            End If
'        Loop
'
'    'ファイルクローズ
'    Close #intFileNo
'
'    '媒体出力釦押下可
'    cmdOutput.Enabled = True
'
'    Exit Sub
'
''*******************************
''VBエラー処理
'Error_psVersionDisp:
'   '「LDユーティリティバージョン管理画面：バージョン情報ファイル作成異常」ログ出力
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
''    ファイルクローズ
'    Close #intFileNo
''*******************************
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfInstallSeitouseiChck
'//  機能名称  : 外部入力プログラム判定データ正当性チェック処理
'//  機能概要  : 外部入力プログラム判定データ正当性チェック処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応不具合修正
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 ファイル名チェック不具合修正
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応【種別チェック機能追加】
'//     REVISIONS :(EG20 V6.11.0.1) 2013-03-27 REVISED BY  [TCC] H.Kondoh
'//                 媒体投入機能変更対応
'//                   種別０の場合も異常とするように変更
'//     REVISIONS :(X.X.X.X)----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfInstallSeitouseiChck(sInputPass As String) As Boolean
    Dim lngFileListCnt As Long               'ファイルリスト数
    Dim strWork     As String                '作業エリア
    Dim iFileNumber As Integer               '未使用ファイル番号
    Dim myLen As Long                        '文字列の長さ
    Dim SysCodeTxt As String                 'バイト変換後(全角→半角)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           'ファイルリスト内記載ファイル名
    Dim bRet As Boolean
    Dim lngCnt As Long
    Dim iRet   As Integer                    'バージョンチェックDLL戻り値
    Dim iGouki As Integer                    '号機番号
    Dim sVersionInfoPath As String           'バージョン情報ファイル(号機別)
    Dim sSrcFileName As String               'ファイルリスト名
    Dim lngErrCode   As Long
    Dim intCheckKind As Integer              ' チェック種別     ' EG20 V6.9.0.1ADD
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD
    
    On Error Resume Next
    
    pfInstallSeitouseiChck = True
    
    '********************************
    '*プロ判正当性チェック
    '********************************
    '外部媒体フォルダ内ファイル名を作成
    sSrcFileName = sInputPass & MN_FILELIST
    '外部媒体の検索をする
    'If Dir(sSrcFileName) <> "" Then     'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
      
      'ファイルが存在しない
      MsgBox "媒体内に、ファイルリストが存在しません。", _
             vbOKOnly + vbExclamation, _
             "→ワーク コピー"
     '「自改ﾊﾞｰｼﾞｮﾝ：ファイルリスト無し」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      pfInstallSeitouseiChck = False
      Set objFso = Nothing    'V1.20.0.1 ADD
      Exit Function
    End If

   '｢ワーク｣フォルダからファイルリストを取得する
    bRet = fReadFileList(sInputPass & MN_FILELIST)

    'サム値チェック
    For lngCnt = 0 To UBound(FileList) - 1
        If pfFileSumChk(sInputPass & FileList(lngCnt), lngSumRet) <> True Then

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           
            'サム値異常
            If lngSumRet = SUM_CHK.SumErr Then
               MsgBox "サム値が異常です。" _
                      & Chr(vbKeyReturn) & "データを確認してください。", _
                      vbOKOnly + vbExclamation, _
                      "自動改札機 バージョン管理"
            'サム値異常以外異常
            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
                   '「ワークコピー異常終了」ポップアップ画面表示
               MsgBox "コピーエラーが発生しました。" _
                     & Chr(vbKeyReturn) & "エラーコード＝" _
                     & str$(Err.Number), _
                     vbOKOnly + vbExclamation, _
                     "自動改札機 バージョン管理"
            End If
            pfInstallSeitouseiChck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    'ファイル数最大チェック
    If UBound(FileList) > FILECNT_MAX Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       
       MsgBox "ファイル数が上限を超えています。" _
              & Chr(vbKeyReturn) & "データを確認してください。", _
              vbOKOnly + vbExclamation, _
              "自動改札機 バージョン管理"
      pfInstallSeitouseiChck = False

      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)

      Exit Function
    End If
'V2.6.0.1 DEL START
'    'ファイル名サイズチェック
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '未使用のファイル番号を取得する
'
'    bRet = True
'
'    'ファイルリストをオープン。
'    Open sInputPass & MN_FILELIST For Input As #iFileNumber
'    For i = 0 To lngFileListCnt
'       If i = lngFileListCnt Then
'          Exit For
'       End If
'       'ファイル名を取得する。
'       Input #iFileNumber, strWork
'       'ファイル名定義なし
'       If strWork = "" Then
'          'ループ抜け
'          MsgBox "ファイル名が異常です。" _
'                  & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "自動改札機 バージョン管理"
'          bRet = False
'          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'          Exit For
'       'フォーマット異常
'       ElseIf " " <> Mid(strWork, 2, 1) And Left$(strWork, 1) <> "/" Then
'          'ループ抜け
'          MsgBox "ファイル名が異常です。" _
'                  & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "自動改札機 バージョン管理"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       'フォーマット異常
'       ElseIf (InStr(strWork, ".") - 1) = -1 And Left$(strWork, 1) <> "/" Then
'           MsgBox "ファイル名が異常です。" _
'                  & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "自動改札機 バージョン管理"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       '「/*--」等のコメント部は除く
'       ElseIf Left$(strWork, 1) = "/" Then
'               '何もしない。
'       Else
'          'ファイル名のみを抽出
'          sGetFileListName = Mid(strWork, 3, 16)
'          '取得ファイル名のサイズを取得
'          myLen = LenB(StrConv(Trim(sGetFileListName), vbFromUnicode))                                              '半角換算のバイト数を取得
'          If FILE_NAME_MAX_SIZE < myLen Then
'            '13バイト以上の場合
'            MsgBox "ファイル名が異常です。" _
'                   & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                   vbOKOnly + vbExclamation, _
'                   sJverName & "自動改札機 バージョン管理"
'            bRet = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'            Exit For
'           End If
'        End If
'     Next
'    'ファイルリストをクローズ。
'    Close #iFileNumber
'V2.6.0.1 DEL END
'V2.6.0.1 ADD START
    For i = 0 To UBound(FileList) - 1
       '取得ファイル名のサイズを取得
       myLen = LenB(StrConv(Trim(FileList(i)), vbFromUnicode))                                              '半角換算のバイト数を取得
       If FILE_NAME_MAX_SIZE < myLen Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
          'プログレスバーを消去する
          Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
          
          '13バイト以上の場合
          MsgBox "ファイル名が異常です。" _
                 & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
                  vbOKOnly + vbExclamation, _
                  "自動改札機 バージョン管理"
           bRet = False
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
           Exit For
       End If
    Next
'V2.6.0.1 ADD END

' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD START
    If bRet = False Then
        pfInstallSeitouseiChck = bRet
        Exit Function
    End If

    For i = 0 To UBound(FileList) - 1
        ' ファイルリスト内の種別を抽出
        intCheckKind = CInt(Left$(FileListType(i), 1))
'EG20 V6.11.0.1 DEL Start
'        If ((gintProgramJudgeKind(FolderSyubetu) = intCheckKind) Or _
'            (intCheckKind = ProgramJudgeKind.JUDGE_NOCHECK)) Then
'            ' データ種別選択部の選択内容とファイルリスト内の種別の比較結果が「一致」、もしくは
'            ' ファイルリスト内の種別が「チェックなし」
'            ' →チェック結果正常
'EG20 V6.11.0.1 DEL End
'EG20 V6.11.0.1 ADD Start
        If (gintProgramJudgeKind(FolderSyubetu) = intCheckKind) Then
            ' データ種別選択部の選択内容とファイルリスト内の種別の比較結果が「一致」
            ' →チェック結果正常
'EG20 V6.11.0.1 ADD End
            bRet = True
        Else
            ' 上記以外
            ' →チェック結果異常
            bRet = False
            ' プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' メッセージ表示
            MsgBox "選択したデータ種別とインストール部材が" & Chr(vbKeyReturn) _
                     & "一致しません", _
                   vbOKOnly + vbExclamation, _
                   "自動改札機 バージョン管理"
            ' エラーログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_PRGKIND_ERROR, 0)
            Exit For
        End If
    Next
' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD END

    pfInstallSeitouseiChck = bRet
Exit Function

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pfInstallSeitouseiChck = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fSelectFile
'//  機能名称  : バージョンチェックファイル名
'//  機能概要  : 対象バージョンチェックファイル名を取得する
'//
'//              型        名称      意味
'//  引数      : String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSelectFile() As String
 
 'EG20 V30.1.0.1 DEL START
' If gStrCurrentForm = sFormName_EJVer Then
'    'バージョンチェックファイル名を設定する。
'    Select Case FolderSyubetu
'       Case 0 '判定CPU-Pro
'            fSelectFile = EHANTEI_CPU_CHK_FILE
'
'       Case 1 'メインCPU-Pro
'            fSelectFile = EMAIN_CPU_CHK_FILE
'
'       Case 2 'サブCPU-Pro
'            fSelectFile = ESUB_CPU_CHK_FILE
'
'       Case 3 'メインCPU-OS
'            fSelectFile = EMAIN_OS_CHK_FILE
'
'     End Select
'  Else
'    'バージョンチェックファイル名を設定する。
'    Select Case FolderSyubetu
'       Case 0 '判定CPU-Pro
'             fSelectFile = NHANTEI_CPU_CHK_FILE
'
'       Case 1 'メインCPU-Pro
'            fSelectFile = NMAIN_CPU_CHK_FILE
'
'       Case 2 'サブCPU-Pro
'            fSelectFile = NSUB_CPU_CHK_FILE
'
'       Case 3 'メインCPU-OS
'            fSelectFile = NMAIN_OS_CHK_FILE
'
'    End Select
'   End If
  'EG20 V30.1.0.1 DEL END
    
    'EG20 V30.1.0.1 ADD START
    'バージョンチェックファイル名を設定する。
    Select Case FolderSyubetu
       Case 0 '判定CPU-Pro
            fSelectFile = EG20_HANTEI_CPU_CHK_FILE
       
       Case 1 'メインCPU-Pro
            fSelectFile = EG20_MAIN_CPU_CHK_FILE
       
       Case 2 'サブCPU1-Pro
            fSelectFile = EG20_SUB_CPU1_CHK_FILE
       
       Case 3 'サブCPU2-Pro
            fSelectFile = EG20_SUB_CPU2_CHK_FILE
       
       Case 4 'サブCPU3-Pro
            fSelectFile = EG20_SUB_CPU3_CHK_FILE
       
       Case 5 'メインCPU-OS
            fSelectFile = EG20_MAIN_OS_CHK_FILE
     
    End Select
    'EG20 V30.1.0.1 ADD END

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fNewVersion
'//  機能名称  : 最新バージョン処理
'//  機能概要  : 最新(ワーク)バージョンを、実行(実行)バージョンに登録
'//
'//              型        名称      意味
'//  引数      : String　　sPath
'//  　　　    : Integer　 iFolder
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　プロ判正当性チェック処理追加
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ１対応不具合修正
'//                 フェーズ３対応　機種正当性チェック処理追加
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                【運改表示改善対応】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fNewVersion() As Boolean
    Dim bRet As Boolean                      '戻り値
    Dim lngCnt                  As Long      'カウンター
    Dim sSrcFileName            As String    'ワークフォルダ内ファイルリスト
    Dim sFileName As String
    Dim lngErrCode As Long                   'エラーコード
    'V1.4.0.1 ADD START
    Dim lngFileListCnt As Long               'ファイルリスト数
    Dim strWork     As String                '作業エリア
    Dim iFileNumber As Integer               '未使用ファイル番号
    Dim myLen As Long                        '文字列の長さ
    Dim SysCodeTxt As String                 'バイト変換後(全角→半角)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           'ファイルリスト内記載ファイル名
    'V1.4.0.1 ADD END
    Dim iKansiAplChk As Integer              'アプリ起動チェック戻り値　'V1.6.0.1 ADD
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD
    
    Dim sCorner As String                    'コーナー番号
    Dim sGatePath As String                  'コーナー番号付ファイルパス
    Dim sFilePath As String                  'ファイルファイルパス
    
    On Error Resume Next
    
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

    sFilePath = sGatePath & FolderName(0, FolderSyubetu)

    '｢ワーク｣フォルダのファイルリストを検索する
    'ワークフォルダ内ファイル名を作成
'    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sFilePath & "\" & MN_FILELIST
    'ファイルの検索をする
    'If Dir(sSrcFileName) <> "" Then     'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
      Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else
      
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
      'プログレスバーを消去する
      Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
      'ファイルが存在しない
      MsgBox "「ワーク」フォルダ内の " & TitleBox(FolderSyubetu) & "に、" _
             & Chr(vbKeyReturn) & "ファイルリストが存在しません。", _
             vbOKOnly + vbExclamation, _
             TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
     '「自改ﾊﾞｰｼﾞｮﾝ：ファイルリスト無し」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      fNewVersion = False
      Set objFso = Nothing    'V1.20.0.1 ADD
      Exit Function
    End If
  
    '｢ワーク｣フォルダからファイルリストを取得する
    'bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)　'V1.8.0.1 DEL
    
    bRet = pfSeitouseiChck    'V1.4.0.1　ADD
    '自改プログラム判定データ正当性チェックを行う(対象ファイル：HAN_KUKA.KUK)
'    bRet = fDataFileCheck(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST) 'V1.4.0.1　DEL
'V1.8.0.1 ADD START
    '｢ワーク｣フォルダからファイルリストより、登録ファイル数をカウントする
    If bRet = True Then
'       bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
       bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    End If
'V1.8.0.1 ADD END

  If bRet = True Then
    '｢旧｣フォルダ内のファイルを全て削除する
     If sOldFolderRemove <> True Then
'        Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加    EG20 V3.6.0.1削除
        Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1追加
         fNewVersion = False
         Exit Function
     End If

    '｢旧｣フォルダ内のファイルを｢実行｣フォルダの内容に置換える
    If sCopyNOWtoOLD <> True Then
'        Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加    EG20 V3.6.0.1削除
        Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1追加
        fNewVersion = False
        Exit Function
    End If

    '｢実行｣フォルダ内のファイルを｢ワーク｣フォルダの内容に置換える
    If sCopyWRKtoNOW <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加
        fNewVersion = False
        Exit Function
    End If
    
' EG20 V3.0.0.2 追加開始
    ' 改札機共通エリア更新処理
    Call pubfuncCommonAreaUpdate
' EG20 V3.0.0.2 追加終了
 
    '自改バージョン情報更新要求メールを管理プロセスへ送信する。
    'V1.6.0.1　ADD　START
    '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '監視盤起動時：管理プロセスに自改バージョン情報更新要求メールを送信する。
    'V1.6.0.1 ADD END
      If gStrCurrentForm = sFormName_EJVer Then
         psVersionUpdateReqest (ML_REQUEST_EGATE)
      Else
         psVersionUpdateReqest (ML_REQUEST_NGATE)
      End If
    'V1.6.0.1 ADD START
    Else
        '監視盤未起動時：改札機バージョン更新処理結果に正常を設定する。
        gintGateVerInfUpdRes = MailSts.stsNormal
    End If
    'V1.6.0.1 ADD END

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '改札機バージョン更新処理結果
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
' EG20 V5.8.0.1削除開始
'        ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'        Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
        ' 運改状態更新
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1追加終了
'        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1)   ' EG20 V5.6.0.1追加           ' EG20 V5.11.0.1削除
        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1追加
        '正常
        MsgBox "「ワーク」フォルダの内容を,「実行」フォルダに登録して、" _
                & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & " の最新のバージョンとしました。", _
                vbOKOnly + vbExclamation, _
                TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
        fNewVersion = True
    Else
        '異常
        If gStrCurrentForm = sFormName_EJVer Then
           MsgBox "改札機のバージョン作成で異常が発生しました。", _
                  vbOKOnly + vbExclamation, _
                  "自動改札機 バージョン管理"
        Else
         MsgBox "改札機のバージョン作成で異常が発生しました。", _
                 vbOKOnly + vbExclamation, _
                 "自動改札機 バージョン管理"
        End If
        
        fNewVersion = False
    End If
  
    fNewVersion = True
  Else
    fNewVersion = False
  End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSeitouseiChck
'//  機能名称  : プログラム判定データ正当性チェック処理
'//  機能概要  : プログラム判定データ正当性チェック処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　プロ判正当性チェック処理
'//     REVISIONS :(1.6.0.1) 2009-06-16  REVISED BY [TCC] S.Terao
'//                 フェーズ２対応不具合修正
'//                 フェーズ３対応　機種正当性チェック追加
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSeitouseiChck() As Boolean
    Dim lngFileListCnt As Long               'ファイルリスト数
    Dim strWork     As String                '作業エリア
    Dim iFileNumber As Integer               '未使用ファイル番号
    Dim myLen As Long                        '文字列の長さ
    Dim SysCodeTxt As String                 'バイト変換後(全角→半角)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           'ファイルリスト内記載ファイル名
    Dim bRet As Boolean
    Dim lngCnt As Long
    Dim iRet   As Integer                    'バージョンチェックDLL戻り値
    Dim iGouki As Integer                    '号機番号
    Dim sVersionInfoPath As String           'バージョン情報ファイル(号機別)
    Dim iCnt             As Integer          '号機カウンター　V1.6.0.1　ADD
    
    Dim sCorner As String                    'コーナー番号
    Dim sGatePath As String                  'コーナー番号付ファイルパス
    Dim sFilePath As String                  'ファイルファイルパス
    
    On Error Resume Next
    
    pfSeitouseiChck = True
   
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

    '********************************
    '*プロ判正当性チェック
    '********************************
    '自改プログラム判定データ正当性チェックを行う(対象ファイル：HAN_KUKA.KUK)
'    bRet = fDataFileCheck(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    
    bRet = fDataFileCheck(sFilePath & "\" & MN_FILELIST)
    If bRet = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
       'プログレスバーを消去する
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       If sNGSts <> "" And sNGKoumoku <> "" Then
          MsgBox "運賃データ正当性チェック異常(" & sNGSts & "：" & sNGKoumoku & "）", _
                 vbOKOnly + vbExclamation, _
                 "自動改札機 バージョン管理"
       Else
          MsgBox "異常終了しました。", _
                 vbOKOnly + vbExclamation, _
                 TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
       End If
'       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加
       Call pubfuncErrorOccur(MN_FOLD_WRK)          ' EG20 V3.6.0.1追加
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2追加開始
    ' 改札機共通判定処理
    bRet = pubfuncCommonGateCheck(MN_FOLD_WRK)
    If bRet = False Then
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2追加終了

'V1.6.0.1 DEL START
'    'サム値チェック
'    For lngCnt = 0 To UBound(FileList) - 1
'        If pfFileSumChk(FolderName(0, FolderSyubetu) & "\" & FileList(lngCnt), lngSumRet) <> True Then
'            'サム値異常
'            If lngSumRet = SUM_CHK.SumErr Then
'               MsgBox "サム値が異常です。" _
'                      & Chr(vbKeyReturn) & "データを確認してください。", _
'                      vbOKOnly + vbExclamation, _
'                      sJverName & "自動改札機 バージョン管理"
'            'サム値異常以外異常
'            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
'               MsgBox "異常終了しました。", _
'                     vbOKOnly + vbExclamation, _
'                     TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
'            End If
'            pfSeitouseiChck = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
'            Exit Function
'        End If
'    Next
'
'    'ファイル数最大チェック
'    If UBound(FileList) > FILECNT_MAX Then
'       MsgBox "ファイル数が上限を超えています。" _
'              & Chr(vbKeyReturn) & "データを確認してください。", _
'              vbOKOnly + vbExclamation, _
'              sJverName & "自動改札機 バージョン管理"
'      pfSeitouseiChck = False
'
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
'
'      Exit Function
'    End If
'
'    'ファイル名サイズチェック
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '未使用のファイル番号を取得する
'    'ファイルリストをオープン。
'    Open FolderName(0, FolderSyubetu) & "\" & MN_FILELIST For Input As #iFileNumber
'    For i = 0 To lngFileListCnt
'       If i = lngFileListCnt Then
'          Exit For
'       End If
'       'ファイル名を取得する。
'       Input #iFileNumber, strWork
'       'ファイル名定義なし
'       If strWork = "" Then
'          'ループ抜け
'          MsgBox "ファイル名が異常です。" _
'                  & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "自動改札機 バージョン管理"
'          bRet = False
'          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'          Exit For
'       'フォーマット異常
'       ElseIf " " <> Mid(strWork, 2, 1) Then
'          'ループ抜け
'          MsgBox "ファイル名が異常です。" _
'                  & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "自動改札機 バージョン管理"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       ElseIf (InStr(strWork, ".") - 1) = -1 Then
'           MsgBox "ファイル名が異常です。" _
'                  & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "自動改札機 バージョン管理"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       Else
'          'ファイル名のみを抽出
'          sGetFileListName = Mid(strWork, 3, 16)
'          '取得ファイル名のサイズを取得
'          myLen = LenB(StrConv(sGetFileListName, vbFromUnicode))                                              '半角換算のバイト数を取得
'          If FILE_NAME_MAX_SIZE < myLen Then
'            '13バイト以上の場合
'            MsgBox "ファイル名が異常です。" _
'                   & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                   vbOKOnly + vbExclamation, _
'                   sJverName & "自動改札機 バージョン管理"
'            bRet = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'            Exit For
'           End If
'        End If
'     Next
'    'ファイルリストをクローズ。
'    Close #iFileNumber
'V1.6.0.1 DEL END
'V1.11.0.1 DEL START
'    If gStrCurrentForm = sFormName_EJVer Then
''V1.6.0.1 ADD　START
'   For iCnt = 1 To MAX_GATE_NO
'      'EG-R自改のみ：自改バージョンチェックDLL処理
'      iGouki = pfGetGoukiNo(iCnt)
'      If iGouki <> 0 Then
''V1.6.0.1 ADD　END
'       'iGouki = pfGetGoukiNo 'V1.6.0.1 DEL
'       sVersionInfoPath = Replace(GATE_VERSION_INFO_FILE, "##", Format(iGouki, "0#"))
'
'       'iRet = dllVerChk(E_EPRO1WRK & "\\" & GATE_VERSION_KANRI_FILE, PATH_GATE & sVersionInfoPath, PATH_HOSHU_LOG & GATE_VERSION_NGLIST_FILE)　　　　　　　　　'V1.6.0.1　DEL
'       iRet = dllVerChk(FolderName(0, FolderSyubetu) & "\" & GATE_VERSION_KANRI_FILE, PATH_GATE & sVersionInfoPath, PATH_HOSHU_LOG & GATE_VERSION_NGLIST_FILE)  'V1.6.0.1　ADD
'       If iRet = 1 Then
'          bRet = True
'       Else
'          bRet = False
'          MsgBox "異常終了しました。", _
'                 vbOKOnly + vbExclamation, _
'                 TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
'          'V1.6.0.1 ADD START
'           pfSeitouseiChck = False
'           Exit Function
'          'V1.6.0.1 ADD END
'       End If
'       End If 'V1.6.0.1 ADD
'      Next 'V1.6.0.1 ADD
'    End If
''V1.6.0.1 ADD START
'V1.11.0.1 DEL END
    '機種正当性チェック(対象ファイル：XX_GATEY.VEF　XX:ユーザー名　Y：データ種別)
'    bRet = fKishuCheck(FolderName(0, FolderSyubetu) & "\")
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    
    bRet = fKishuCheck(sFilePath & "\")
    
    If bRet = False Then
       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
       'プログレスバーを消去する
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       MsgBox "異常終了しました。", _
                  vbOKOnly + vbExclamation, _
                  TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
       pfSeitouseiChck = False
       Exit Function
    End If
'V1.6.0.1 ADD END

    pfSeitouseiChck = bRet
Exit Function

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加
    pfSeitouseiChck = False
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sOldFolderRemove
'//  機能名称  : 旧フォルダ内ファイル削除処理
'//  機能概要  : 旧フォルダ内のファイルを削除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sOldFolderRemove() As Boolean
    Dim MyName As String                'ファイル名
    Dim lngErrCode As Long              'エラーコード
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    'V1.20.0.1 ADD END
    
    Dim sCorner As String                      'コーナー番号
    Dim sGatePath As String                    'コーナー番号付ファイルパス
    Dim sFilePath As String                    'ファイルファイルパス
    
   '戻り値初期化
    sOldFolderRemove = True
    
    On Error GoTo ErrorHandler          'エラーハンドルの登録
    
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner
 
    '「実行」フォルダ内のディレクトリの名前を表示します。
'    gstrMyPath = FolderName(2, FolderSyubetu) & "\"
    gstrMyPath = sGatePath & FolderName(2, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' 最初のディレクトリ名を返します。
'    Do While MyName <> ""                   ' ループを開始します。
'        ' 現在のディレクトリと親ディレクトリは無視します。
'        If MyName <> "." And MyName <> ".." Then
'            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'                'ファイルを削除する
'                Kill gstrMyPath & MyName
'            End If
'        End If
'        MyName = Dir        ' 次のディレクトリ名を返します。
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  'ループを開始
        If objFso.FileExists(objFi.Path) = True Then      'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
                'ファイルを削除する
                Kill gstrMyPath & MyName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    Exit Function           '処理を終了する

ErrorHandler:   ' エラー処理ルーチン。
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    '「ワーク→実行コピー異常終了」ポップアップ画面表示
     MsgBox "異常終了しました。", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
    '「自改ﾊﾞｰｼﾞｮﾝ：旧フォルダﾌｧｲﾙ削除異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLDFILE_DELETE_ERROR, lngErrCode)

    sOldFolderRemove = False
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sNowFolderRemove
'//  機能名称  : 実行フォルダ内のファイル削除処理
'//  機能概要  : 実行フォルダ内のファイルを削除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sNowFolderRemove() As Boolean
    Dim MyName As String                'ファイル名
    Dim lngErrCode As Long              'エラーコード
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    'V1.20.0.1 ADD END

    Dim sCorner As String                 'コーナー番号
    Dim sGatePath As String               'コーナー番号付ファイルパス
    Dim sFilePath As String
    On Error GoTo ErrorHandler          'エラーハンドルの登録

    '初期値設定
    sNowFolderRemove = True
    
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

    sFilePath = sGatePath & FolderName(1, FolderSyubetu)
    
    '「実行」フォルダ内のディレクトリの名前を表示します。
'    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
    gstrMyPath = sFilePath & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' 最初のディレクトリ名を返します。
'    Do While MyName <> ""                   ' ループを開始します。
'        ' 現在のディレクトリと親ディレクトリは無視します。
'        If MyName <> "." And MyName <> ".." Then
'            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'
'                Kill gstrMyPath & MyName        'ファイルを削除する
'
'            End If
'        End If
'        MyName = Dir        ' 次のディレクトリ名を返します。
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  'ループを開始
        If objFso.FileExists(objFi.Path) = True Then      'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                Kill gstrMyPath & MyName        'ファイルを削除する

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END

    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    '「旧→実行コピー異常終了」ポップアップ画面表示
    MsgBox "異常終了しました。", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  旧→実行 コピー"

    '「自改ﾊﾞｰｼﾞｮﾝ：実行フォルダﾌｧｲﾙ削除異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOWFILE_DELETE_ERROR, lngErrCode)

    sNowFolderRemove = False
    
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sWrkFolderRemove
'//  機能名称  : ワークフォルダ内ファイル削除処理
'//  機能概要  : ワークフォルダ内のファイルを削除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-17  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                【運改表示改善対応】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim MyName As String                'ファイル名
    Dim lngErrCode As Long              'エラーコード
    Dim lngPgmHanteiStsWork As Long     'プログラム判定状態（ワーク）   ' EG20 V3.6.0.1追加
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    'V1.20.0.1 ADD END
    
    Dim sCorner As String               'コーナー番号
    Dim sGatePath As String             'コーナー番号付ファイルパス
    
    On Error GoTo ErrorHandler          'エラーハンドルの登録

    '初期値設定
    sWrkFolderRemove = True
   
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner
  
    'ワークフォルダ内のディレクトリの名前を表示します。
'    gstrMyPath = FolderName(0, FolderSyubetu) & "\"
    gstrMyPath = sGatePath & FolderName(0, FolderSyubetu) & "\"
    
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' 最初のディレクトリ名を返します。
'    Do While MyName <> ""                   ' ループを開始します。
'        ' 現在のディレクトリと親ディレクトリは無視します。
'        If MyName <> "." And MyName <> ".." Then
'            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'                'ファイルを削除する
'                Kill gstrMyPath & MyName
'            End If
'        End If
'        MyName = Dir        ' 次のディレクトリ名を返します。
'    Loop
    'V1.20.0.1 DEL END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  'ループを開始
        If objFso.FileExists(objFi.Path) = True Then      'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
                'ファイルを削除する
                Kill gstrMyPath & MyName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END

' EG20 V3.6.0.1追加開始
    '監視設定エリア「プログラム判定異常状態（ワーク）」の状態を取得する
    lngPgmHanteiStsWork = pfGetKansiSet(IdKansiSet.PG_HANTEI_ERR_STS_WORK)

    '「プログラム判定異常状態（ワーク）」（正常）
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_ERR_STS_WORK, ErrCode.Normal)
    
    '変化があった場合、「状態変化通知」を送信する
    If lngPgmHanteiStsWork <> ErrCode.Normal Then
        Call sSendMailStsChgInf(MailSts.stsNormal, lngPgmHanteiStsWork)
    End If
' EG20 V3.6.0.1追加終了
    
' EG20 V5.11.0.1削除開始
'' EG20 V5.8.0.1削除開始
''    ' 運改状態更新                                              ' EG20 V5.5.0.1追加
''    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1追加
'' EG20 V5.8.0.1削除終了
'' EG20 V5.8.0.1追加開始
'    ' 運改状態更新
'    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_NASHI)
'' EG20 V5.8.0.1追加終了
'    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, iTab_index + 1)   ' EG20 V5.6.0.1追加
' EG20 V5.11.0.1削除終了
' EG20 V5.11.0.1追加開始
    ' 運改状態更新
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_CLEAR)
    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1追加
' EG20 V5.11.0.1追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    '「ワーククリア正常終了」ポップアップ画面表示
    MsgBox "「ワーク」フォルダ内の " & TitleBox(FolderSyubetu) & "を、" _
               & Chr(vbKeyReturn) & "全て削除しました。", _
               vbOKOnly + vbExclamation, _
               TitleBox(FolderSyubetu) & "  ワーク クリア"

    Exit Function '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    Call pubfuncErrorOccur(MN_FOLD_WRK)             ' EG20 V3.6.0.1追加
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '「ワーククリア異常終了」ポップアップ画面表示
     MsgBox "異常終了しました。", _
           vbOKOnly + vbCritical, _
           "ワーク クリア"
           
   '「自改ﾊﾞｰｼﾞｮﾝ：ﾜｰｸﾌｫﾙﾀﾞﾌｧｲﾙ削除異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sCopyNOWtoOLD
'//  機能名称  : 実行バージョン保存処理
'//  機能概要  : 実行フォルダ内のファイルを、旧フォルダにコピーする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sCopyNOWtoOLD() As Boolean
    Dim MyName As String                'ファイル名
    Dim sSrcFileName As String          'コピー元ファイルのフルパス名
    Dim sDstFileName As String          'コピー先ファイルのフルパス名
    Dim iResponse As Integer            'MsgBoxボタンコード
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    'V1.20.0.1 ADD END
    
    Dim sCorner As String                      'コーナー番号
    Dim sGatePath As String                    'コーナー番号付ファイルパス
    
    On Error GoTo ErrorHandler              'エラーハンドル設定
  
    '戻り値初期化
    sCopyNOWtoOLD = True
   
       ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

    '実行フォルダ内のディレクトリの名前を表示します。
'    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
    gstrMyPath = sGatePath & FolderName(1, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' 最初のディレクトリ名を返します。
'    Do While MyName <> ""                   ' ループを開始します。
'        ' 現在のディレクトリと親ディレクトリは無視します。
'        If MyName <> "." And MyName <> ".." Then
'            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'
'                '実行フォルダ内ファイル名を作成する
'                sSrcFileName = gstrMyPath & MyName
'
'                '旧フォルダ内ファイル名を作成する
'                sDstFileName = FolderName(2, FolderSyubetu) & "\" & MyName
'
'                'ワークフォルダ内のファイルを実行フォルダにコピーする
'                FileCopy sSrcFileName, sDstFileName
'
'            End If
'        End If
'        MyName = Dir        ' 次のディレクトリ名を返します。
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  'ループを開始
        If objFso.FileExists(objFi.Path) = True Then      'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます｡
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                '実行フォルダ内ファイル名を作成する
                sSrcFileName = gstrMyPath & MyName

                '旧フォルダ内ファイル名を作成する
'                sDstFileName = FolderName(2, FolderSyubetu) & "\" & MyName
                sDstFileName = sGatePath & FolderName(2, FolderSyubetu) & "\" & MyName

                'ワークフォルダ内のファイルを実行フォルダにコピーする
                FileCopy sSrcFileName, sDstFileName

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
           
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           ' 「ワーク→実行コピー異常終了」ポップアップ画面表示
            MsgBox "異常終了しました。", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
    
    sCopyNOWtoOLD = False
    
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sCopyWRKtoNOW
'//  機能名称  : 最新バージョンコピー
'//  機能概要  : ワークフォルダ内のファイルを、実行フォルダにコピー
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応（PASSINFコピー対応）
'//     REVISIONS :(EG20 V3.5.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sCopyWRKtoNOW() As Boolean
    
    Dim i As Integer                'カウンタ
    Dim sSrcFileName As String      'コピー元ファイル名
    Dim sDstFileName As String      'コピー先ファイル名
    Dim bError As Boolean           'フラグ
    Dim bRet As Boolean             '戻り値
    
    Dim sCorner As String                'コーナー番号
    Dim sGatePath As String              'コーナー番号付ファイルパス
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler      'エラーハンドルの登録
  
    '戻り値初期化
    sCopyWRKtoNOW = True
    
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

    '****************************
    '* ファイルリストをコピーする *
    '****************************
      
'    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sGatePath & FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
                                    'ワークフォルダ内ファイル名を作成する
'    sDstFileName = FolderName(1, FolderSyubetu) & "\" & MN_FILELIST
    sDstFileName = sGatePath & FolderName(1, FolderSyubetu) & "\" & MN_FILELIST
                                    '実行フォルダ内ファイル名を作成する
    'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする   'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then     'ファイルの検索をする   'V1.20.0.1 ADD
        'ファイルリストを「ワーク」フォルダから「実行」フォルダにコピーする
        FileCopy sSrcFileName, sDstFileName
    Else                                'ファイルが存在しない
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
     '「ワークフォルダファイルリストなし」ポップアップ画面表示
     MsgBox "「ワーク」フォルダ内の " & TitleBox(FolderSyubetu) & "に、" _
             & Chr(vbKeyReturn) & "ファイルリストが存在しません。", _
             vbOKOnly + vbExclamation, _
             TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
     sCopyWRKtoNOW = False
     Set objFso = Nothing    'V1.20.0.1 ADD
     Exit Function                   '処理を終了する
    End If

    bError = False                  'エラーフラグを「偽」にする
    For i = 0 To UBound(FileList) - 1
                                    'ファイルリスト一覧数分繰り返す
'        sSrcFileName = FolderName(0, FolderSyubetu) & "\" & FileList(i)
        sSrcFileName = sGatePath & FolderName(0, FolderSyubetu) & "\" & FileList(i)
                                    'ワークフォルダ内ファイル名を作成する
'        sDstFileName = FolderName(1, FolderSyubetu) & "\" & FileList(i)
        sDstFileName = sGatePath & FolderName(1, FolderSyubetu) & "\" & FileList(i)
                                    '実行フォルダ内ファイル名を作成する

        'ワークフォルダ内のファイルを実行フォルダにコピーする
        'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする   'V1.20.0.1 DEL
        If objFso.FileExists(sSrcFileName) = True Then   'ファイルの検索をする   'V1.20.0.1 ADD
            'ファイルを「ワーク」フォルダから「実行」フォルダにコピーする
            FileCopy sSrcFileName, sDstFileName
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2追加開始
    If pfuncCopyPASSINF(iTab_index, MN_FOLD_WRK) = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
' EG20 V3.5.0.1追加開始
        MsgBox "異常終了しました。", _
                vbOKOnly + vbExclamation, _
                TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
' EG20 V3.5.0.1追加終了
        sCopyWRKtoNOW = False
    End If
' EG20 V3.0.0.2追加終了
    
    Exit Function                           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    Select Case Err.Number
        Case 53 '「ワーク→実行コピー異常終了」ポップアップ画面表示
            MsgBox "異常終了しました。", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
            
            sCopyWRKtoNOW = False
            Set objFso = Nothing    'V1.20.0.1 ADD
            Exit Function
        Case Else
                ' 他のエラー処理をここに記述します。
    End Select
    sCopyWRKtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fDataFileCheck
'//  機能名称  : 自改プログラム判定データ正当性チェック処理
'//  機能概要  : 対象となるHAN_KUKA.KUK有無チェックを行う。
'//
'//              型        名称      意味
'//  引数      : String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-23  CODED   BY [TCC] D.Yamashita
'//                 ・フェーズ３残件項目対応　異常時クローズ処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fDataFileCheck(sFileList As String) As Boolean
    Dim iFileNumber As Integer      'ファイル番号
    Dim sFileName As String         'ファイル名
    Dim iListCnt As Integer         'ファイル格納数
    Dim sFolderPath As String       'HAN_KUKA.KUKフォルダパス用
    Dim sHANKUKAPath As String      'HAN_KUKA.KUKフルパス用
     
    On Error GoTo ErrorHandler      'エラーハンドル設定

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '未使用のファイル番号を取得する

    Open sFileList For Input Access Read As #iFileNumber    'ファイルリストのオープン
    Do While Not EOF(iFileNumber)                           'ファイルの終端までループを繰り返します。
        Line Input #iFileNumber, sFileName                  'データを読み込みます。
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                'ファイル名が存在する
            iListCnt = iListCnt + 1                         'ファイル数のカウンタをアップする
            ReDim Preserve FileList(iListCnt)               'ファイル名格納エリアを拡張する
            ReDim Preserve FileListType(iListCnt)           'ファイル名格納エリアを拡張する
            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
                                                            'ファイル名をファイル名格納エリアにセット
            If HANKUKA_KUK = FileList(iListCnt - 1) Then
               'HAN_KUKA.KUKファイルが有った場合、データ正当性チェックを行う。
               psFolderPathGet sFileList, sFolderPath
               sHANKUKAPath = sFolderPath & HANKUKA_KUK
               If fHankukaChck(sHANKUKAPath) = False Then
                 'データ正当性チェック異常時は、戻り値にFalseを設定する。
                  fDataFileCheck = False
                  Close #iFileNumber        'ファイルを閉じます。   'V1.11.0.1 ADD
                  Exit Function
               End If
            End If
        End If
  Loop
  
  Close #iFileNumber        'ファイルを閉じます。

  fDataFileCheck = True     '戻り値を正常とする

    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:               ' エラー処理ルーチン。
    fDataFileCheck = False  '戻り値をエラーとする
    Close #iFileNumber      'ファイルを閉じます。
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fKishuCheck
'//  機能名称  : 自改プログラム判定データ正当性チェック処理
'//  機能概要  : 対象となるデータの機種正当性チェックを行う。
'//
'//              型        名称      意味
'//  引数      : String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                ワーク→実行コピーでの機種正当性チェック変更
'//                Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fKishuCheck(sFileList As String) As Boolean
    Dim sKisyu       As String * 8     '取得機種名
    Dim sMyName      As String         '機種正当性チェックリストファイル名
    Dim sFileName    As String         'ファイルリスト記載ファイル名
    Dim sChkFileName As String         '機種正当性チェックファイルパス
    Dim sVerChkFile  As String         'バージョンチェックファイル名
    
    Dim lLen         As Long           'ファイルサイズ
    Dim lPos         As Long           'バージョン情報格納位置
           
    Dim i            As Integer        'カウンター
    Dim iCnt         As Integer        '登録レコード数
    Dim iListCnt     As Integer        'ファイル格納数
    Dim iFileNumber  As Integer        'ファイル番号

    Dim bRet         As Boolean        '機種正当性チェック結果

    Dim uHeder       As MN_HEDER       'ヘッダ情報格納エリア
    Dim uFotter      As MN_FOOT        'フッタ情報格納エリア
    
    Dim sChkData As String             '比較文字抽出    'V1.20.0.1 ADD
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD
    
    On Error GoTo ErrorHandler      'エラーハンドル設定
     
    '初期化
    iCnt = 0
    iListCnt = 0
    iFileNumber = 0
    fKishuCheck = False
        
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)
    
    'バージョンデータ(機種正当性チェックリストファイルパス)作成
    sVerChkFile = fSelectFile
    
    'ファイル名取得不可=機種正当性チェックファイルなし
    If sVerChkFile = "" Then
       '正当性チェックを行う必要ないため、正常を返す。
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    
    sMyName = sFileList & sVerChkFile
    
    'If Dir(sMyName) <> "" Then              'ファイルが存在する?     'V1.20.0.1 DEL
    If objFso.FileExists(sMyName) = True Then    'ファイルが存在する?  'V1.20.0.1 ADD
       
       iFileNumber = FreeFile               '未使用のファイル番号を取得する
       
       Open sMyName For Input Access Read As #iFileNumber     'バージョンデータのオープン
       
       'データ読み込み
       Line Input #iFileNumber, sFileName
          
       '読み込みデータより、ヘッダ部を除く。
       sFileName = Mid(sFileName, Len(uHeder) - 3)
       
       'ファイルの終端までループを繰り返します。
       Do While Not EOF(iFileNumber)
          
          '読み込み。
          Line Input #iFileNumber, sFileName
           
           '取得情報が「/」以降のコメントなら対象外。
           'データ部本文以外なら対象外
           'データ部本文のみの場合のみ、ファイル名取得を行う。
           If sFileName <> "" And Left$(sFileName, 1) <> "/" _
                              And " " = Mid(sFileName, 2, 1) Then   'ファイル名が存在する
              iListCnt = iListCnt + 1                         'ファイル数のカウンタをアップする
              ReDim Preserve FileList(iListCnt)               'ファイル名格納エリアを拡張する
              ReDim Preserve FileListType(iListCnt)           'ファイル名格納エリアを拡張する
              'ファイル名をファイル名格納エリアにセット
              FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
              FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 12)
              '登録レコード数をカウント
              iCnt = iCnt + 1
            End If
       Loop
       
       Close #iFileNumber                                     'ファイルを閉じます。
       iFileNumber = 0
    Else
       'ファイルが存在しない場合：正当性チェックを行わない。
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    
    'V1.20.0.1 ADD  START
    If iCnt = 0 Then
       'ファイルリストコードが存在しない場合：正当性チェックを行わない。
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    'V1.20.0.1 ADD  END
    
    'ファイル機種正当性チェックを行う。
    For i = 0 To iCnt - 1
         'チェック対象ファイルパス作成
        sChkFileName = sFileList & FileList(i)
    
        'If Dir(sChkFileName) <> "" Then              'ファイルが存在する?  'V1.20.0.1 DEL
        If objFso.FileExists(sChkFileName) = True Then  'ファイルが存在する?   'V1.20.0.1 ADD
            
            lLen = FileLen(sChkFileName)             'ファイルサイズの取得

            iFileNumber = FreeFile                   '未使用のファイル番号を取得する
            'ファイルのオープンを行う。
            Open sChkFileName For Binary Access Read As #iFileNumber
            'フッタ情報の取得
            Get #iFileNumber, lLen - Len(uFotter) + 1, uFotter
            
            Close #iFileNumber                       'ファイルを閉じます
            iFileNumber = 0
            
            '機種名セット
            sKisyu = uFotter.sKisyu
            
            sChkData = "" '初期化　'V1.20.0.1 ADD
            
' EG20 V3.0.0.2 追加開始
            '文字抽出
            sChkData = Left(sKisyu, Len(EG20_JIKAI_KISHU))
            If EG20_JIKAI_KISHU = sChkData Then
                bRet = True  '機種正当性：正常
            Else
                bRet = False '機種正当性：異常
                fKishuCheck = bRet
                Set objFso = Nothing    'V1.20.0.1 ADD
                Exit Function
            End If
' EG20 V3.0.0.2 追加終了
            
' EG20 V3.0.0.2 削除開始
'            '自改チェック
'            If gStrCurrentForm = sFormName_EJVer Then
'               'EG-R自改時
'               'If EGR_JIKAI_KISHU = Trim(sKisyu) Then  'V1.20.0.1 DEL
'               'V1.20.0.1 ADD START
'               '文字抽出
'               sChkData = Left(sKisyu, Len(EGR_JIKAI_KISHU))
'               If EGR_JIKAI_KISHU = sChkData Then
'               'V1.20.0.1 ADD END
'                   bRet = True  '機種正当性：正常
'               Else
'                   bRet = False '機種正当性：異常
'                   fKishuCheck = bRet
'                   Set objFso = Nothing    'V1.20.0.1 ADD
'                   Exit Function
'               End If
'            Else
'               'NEG自改時
'               'If NEG_JIKAI_KISHU = Trim(sKisyu) Then    'V1.20.0.1 DEL
'               'V1.20.0.1 ADD START
'               '文字抽出
'               sChkData = Left(sKisyu, Len(NEG_JIKAI_KISHU))
'               If NEG_JIKAI_KISHU = sChkData Then
'               'V1.20.0.1 ADD END
'                   bRet = True  '機種正当性：正常
'               Else
'                   bRet = False '機種正当性：異常
'                   fKishuCheck = bRet
'                   Set objFso = Nothing    'V1.20.0.1 ADD
'                   Exit Function
'               End If
'            End If
' EG20 V3.0.0.2 削除終了

        End If
    Next

  fKishuCheck = bRet
  
  Set objFso = Nothing    'V1.20.0.1 ADD
  
 Exit Function

ErrorHandler:
   If iFileNumber <> 0 Then
       Close #iFileNumber                                     'ファイルを閉じます。
   End If
    
   '戻り値を異常とする
   fKishuCheck = False
       
   Set objFso = Nothing    'V1.20.0.1 ADD

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fHankukaChck
'//  機能名称  : HAN_KUKA.KUK正当性チェック処理
'//  機能概要  : 対象となるHAN_KUKA.KUKの内容をチェックする。
'//
'//              型        名称      意味
'//  引数      : String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-06   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応
'//     REVISIONS :(1.16.0.1) 2009-12-21   REVISED BY [TCC] S.Terao
'//                 不具合対応
'//     REVISIONS :(V2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 EG-R(KK)　八丁畷対応　KUK正当性チェック変更
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fHankukaChck(sFilePath As String) As Boolean
    Dim iFileNumber As Integer           'ファイル番号
    Dim i As Integer
    Dim lSts As Long
    Dim sKeyName As String
    Dim lPos As Long                     'バージョン情報格納位置
    Dim lLen As Long                     'ファイルサイズ
    Dim uFooter As MN_FOOT          'フッタ情報格納エリア
'    Dim uHeder As MN_FOOT           'ヘッダ情報格納エリア     'V1.4.0.1 DEL
    Dim sDateTime As String
    Dim j As Integer
    Dim lngErrCode As Long          'エラーコード
    'V1.4.0.1 ADD START
    Dim uHeder As HAN_KUKA_KUK_HEADER       'ヘッダ情報格納エリア
    Dim sGetInfo As String * MAX_PATH_SIZE  'INIファイル取得用
    Dim sChkFileData As String
    Dim iMojisu As Integer
    
    'V1.16.0.1 ADD Start
    Dim bChkSts As Boolean              'チェック結果フラグ
    Dim sChkData As String              '比較文字抽出
    'V1.16.0.1 ADD End
    
   '初期化：正常(ブランク）
    sNGSts = ""
    sNGKoumoku = ""
    'V1.4.0.1 ADD END
    Dim oFs As New FileSystemObject 'V2.5.0.1 ADD
    
    fHankukaChck = False
    
'V2.5.0.1 ADD START
 'ファイル有無チェックを行う。
 If oFs.FileExists(sFilePath) = False Then
    'ファイルが無ければ正当性チェックを行わない。
    fHankukaChck = True
    Set oFs = Nothing
    Exit Function
 End If
'V2.5.0.1 ADD END

 'V1.4.0.1 DEL START
'   For i = 0 To INI_MAX
'      'ヘッダ：期待値機種名取得
'      sKeyName = Format(HEDER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sHederKisyu(i), _
'                                     Len(HAN_KUKA_DATA.sHederKisyu(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      'ヘッダ：期待値ファイル名取得
'      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sHederFile(i), _
'                                     Len(HAN_KUKA_DATA.sHederFile(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      'フッタ：期待値機種名取得
'      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sFotterKisyu(i), _
'                                     Len(HAN_KUKA_DATA.sFotterKisyu(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      'フッタ：期待値ファイル名取得
'      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sFotterFile(i), _
'                                     Len(HAN_KUKA_DATA.sFotterFile(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'    Next i
    'V1.4.0.1 DEL END
    'V1.4.0.1 ADD START
    '初期化
    For i = 0 To INI_MAX - 1
        HAN_KUKA_DATA.sHederKisyu(i) = ""
        HAN_KUKA_DATA.sHederFile(i) = ""
        HAN_KUKA_DATA.sFotterKisyu(i) = ""
        HAN_KUKA_DATA.sFotterFile(i) = ""
    Next
    For i = 0 To INI_MAX - 1
      'ヘッダ：期待値機種名取得
      sKeyName = Format(HEDER_KISHU_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
       
      Else
        HAN_KUKA_DATA.sHederKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      'ヘッダ：期待値ファイル名取得
      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
        
      Else
         HAN_KUKA_DATA.sHederFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      'フッタ：期待値機種名取得
      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
        
      Else
        HAN_KUKA_DATA.sFotterKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      'フッタ：期待値ファイル名取得
      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
        
      Else
        HAN_KUKA_DATA.sFotterFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
    Next i
    'V1.4.0.1 ADD END

    On Error GoTo ErrorHandler      'エラーハンドル設定
    
    'HAN_KUKA.KUKファイルサイズ取得
    lLen = FileLen(sFilePath)
    
    '未使用のファイル番号を取得する
    iFileNumber = FreeFile
    
    'V1.4.0.1 DEL START
'    'HAN_KUKA.KUKファイルをオープンする。
'    Open sFilePath For Input Access Read As #iFileNumber
'
'    'HAN_KUKA.KUKファイルのヘッダ情報を取得する。
''    Get #iFileNumber, 1, uHeder
    'V1.4.0.1 DEL END

    'V1.4.0.1 ADD START
    'HAN_KUKA.KUKファイルをオープンする。
    Open sFilePath For Binary Access Read As #iFileNumber
            
    'HAN_KUKA.KUKファイルのヘッダ情報を取得する。
    Get #iFileNumber, 1, uHeder
    'V1.4.0.1 ADD END

   'HAN_KUKA.KUKファイルのフッタ情報を取得する。
    Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter

    'HAN_KUKA.KUKファイルをクローズする。
    Close #iFileNumber
    
    iFileNumber = 0                          'V1.4.0.1 ADD
'V1.4.0.1 DEL START
    '機種名/ファイル名チェック
'    For i = 0 To 5
'       'ヘッダ情報：機種名チェック
'       If uHeder.sKisyu <> HAN_KUKA_DATA.sHederKisyu(i) Then
'          Exit Function
'       End If
'       'ヘッダ情報：ファイル名チェック
'       If uHeder.sFileName <> HAN_KUKA_DATA.sHederFile(i) Then
'          Exit Function
'       End If
'       'フッタ情報：機種名チェック
'       If uFooter.sKisyu <> HAN_KUKA_DATA.sFotterKisyu(i) Then
'          Exit Function
'       End If
'       'フッタ情報：ファイル名チェック
'       If uFooter.sFileName <> HAN_KUKA_DATA.sFotterFile(i) Then
'          Exit Function
'       End If
'     Next
'V1.4.0.1 DEL END
   'V1.4.0.1 ADD START
   'ヘッダ情報：機種名チェック
   iMojisu = InStr(uHeder.sKisyuName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sKisyuName, 1)
   Else
     sChkFileData = Mid(uHeder.sKisyuName, 1, iMojisu)
   End If
'V1.16.0.1 DEL START
'    For i = 0 To INI_MAX - 1
'      If sChkFileData <> HAN_KUKA_DATA.sHederKisyu(i) Then
'         If i = INI_MAX - 1 Then
'            '機種名期待値全不一致：
'            sNGSts = ERROR_HEDER
'            sNGKoumoku = KISHU_NAME_ERROR
'            GoTo ErrorHandler
'         End If
'      Else
'        Exit For
'      End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
    bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sHederKisyu(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sHederKisyu(i)))
          If sChkData = HAN_KUKA_DATA.sHederKisyu(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    'チェック結果フラグ判定
    If bChkSts = False Then
       '機種名期待値全不一致：
        sNGSts = ERROR_HEDER
        sNGKoumoku = KISHU_NAME_ERROR
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END

   'ヘッダ情報：ファイル名チェック
   iMojisu = InStr(uHeder.sProgrumName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sProgrumName, 1)
   Else
     sChkFileData = Mid(uHeder.sProgrumName, 1, iMojisu)
   End If

'V1.16.0.1 DEL START
'   For i = 0 To INI_MAX - 1
'       If sChkFileData <> HAN_KUKA_DATA.sHederFile(i) Then
'         If i = INI_MAX - 1 Then
'            'ファイル名期待値全不一致：
'            sNGSts = ERROR_HEDER
'            sNGKoumoku = FILE_NAME_ERRORE
'            GoTo ErrorHandler
'         End If
'      Else
'         Exit For
'      End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
    bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sHederFile(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sHederFile(i)))
          If sChkData = HAN_KUKA_DATA.sHederFile(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    'チェック結果フラグ判定
    If bChkSts = False Then
       '機種名期待値全不一致：
        sNGSts = ERROR_HEDER
        sNGKoumoku = FILE_NAME_ERRORE
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END
    
   '作成日付チェック
   'ヘッダ情報：作成日付が数値かどうか
    sDateTime = ""
    For j = 0 To 3
        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
    Next
    For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
    Next
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_HEDER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    
    'バージョン数値チェック
    If IsNumeric(uHeder.sVersion) = False Then
       sNGSts = ERROR_HEDER
       sNGKoumoku = VERSION_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    
   'フッタ情報：機種名チェック
   iMojisu = InStr(uFooter.sKisyu, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uFooter.sKisyu, 1)
   Else
     sChkFileData = Mid(uFooter.sKisyu, 1, iMojisu)
   End If
'V1.16.0.1 DEL START
'    For i = 0 To INI_MAX - 1
'      If sChkFileData <> HAN_KUKA_DATA.sFotterKisyu(i) Then
'         If i = INI_MAX - 1 Then
'             '機種名期待値全不一致：
'             sNGSts = ERROR_FOTTER
'             sNGKoumoku = KISHU_NAME_ERROR
'             GoTo ErrorHandler
'          End If
'       Else
'         Exit For
'       End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
   bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sFotterKisyu(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterKisyu(i)))
          If sChkData = HAN_KUKA_DATA.sFotterKisyu(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    'チェック結果フラグ判定
    If bChkSts = False Then
       '機種名期待値全不一致：
        sNGSts = ERROR_FOTTER
        sNGKoumoku = KISHU_NAME_ERROR
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END

   'フッタ情報：ファイル名チェック
   iMojisu = InStr(uFooter.sFileName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uFooter.sFileName, 1)
   Else
     sChkFileData = Mid(uFooter.sFileName, 1, iMojisu)
   End If
'V1.16.0.1 DEL START
'    For i = 0 To INI_MAX - 1
'       If sChkFileData <> HAN_KUKA_DATA.sFotterFile(i) Then
'          If i = INI_MAX - 1 Then
'             '機種名期待値全不一致：
'             sNGSts = ERROR_FOTTER
'             sNGKoumoku = FILE_NAME_ERRORE
'             GoTo ErrorHandler
'          End If
'       Else
'         Exit For
'       End If
'    Next
'   'V1.4.0.1 ADD END
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
   bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sFotterFile(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterFile(i)))
          If sChkData = HAN_KUKA_DATA.sFotterFile(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    'チェック結果フラグ判定
    If bChkSts = False Then
       '機種名期待値全不一致：
        sNGSts = ERROR_FOTTER
        sNGKoumoku = FILE_NAME_ERRORE
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END

'V1.4.0.1 DEL START
'   '作成日付チェック
'   'ヘッダ情報：作成日付が数値かどうか
'    sDateTime = ""
'    For j = 0 To 3
'        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
'    Next
'    sDateTime = sDateTime & " "
'    For j = 4 To 5
'        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
'    Next
'    If (Chr(sDateTime) >= "A" And Chr(sDateTime) <= "Z") And _
'        (Chr(sDateTime) >= "a" And Chr(sDateTime) <= "z") Then
'         Exit Function
'    End If
'V1.4.0.1 DEL END
      
    'フッタ情報：作成日付が数値かどうか
     sDateTime = ""
     For j = 0 To 3
         sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
     Next
    'sDateTime = sDateTime & " " 'V1.4.0.1 DEL
     For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    'V1.4.0.1 DEL START
'    If (Chr(sDateTime) >= "A" And Chr(sDateTime) <= "Z") And _
'       (Chr(sDateTime) >= "a" And Chr(sDateTime) <= "z") Then
'        Exit Function
'    End If
    'V1.4.0.1 DEL END
    
    'V1.4.0.1 ADD START
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_FOTTER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    'V1.4.0.1 ADD END
    'V1.4.0.1 DEL START
'      'バージョン値チェック
'    'ヘッダ情報：バージョン値が数値かどうか
'    If (Chr(uHeder.sVersion) >= "A" And Chr(uHeder.sVersion) <= "Z") And _
'        (Chr(uHeder.sVersion) >= "a" And Chr(uHeder.sVersion) <= "z") Then
'        Exit Function
'    End If
'
'    'フッタ情報：バージョン値が数値かどうか
'    If (Chr(uFooter.sVersion) >= "A" And Chr(uFooter.sVersion) <= "Z") And _
'       (Chr(uFooter.sVersion) >= "a" And Chr(uFooter.sVersion) <= "z") Then
'        Exit Function
'    End If
    'V1.4.0.1 DEL END
    
    'V1.4.0.1 ADD START
    'バージョン値チェック
    'フッタ情報：バージョン値が数値かどうか
    If IsNumeric(uFooter.sVersion) = False Then
       sNGSts = ERROR_FOTTER
       sNGKoumoku = VERSION_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    'V1.4.0.1 ADD END
    
    '「自改ﾊﾞｰｼﾞｮﾝ：正当チェック正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_OK, 0)
    
    'すべてOKの場合、TRUEでかえる。
    fHankukaChck = True

Exit Function 'V1.4.0.1 ADD
'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    'V1.4.0.1 ADD START
    If iFileNumber > 0 Then
       'HAN_KUKA.KUKファイルをクローズする。
       Close #iFileNumber
    End If
    iFileNumber = 0
    'V1.4.0.1 ADD END
    
    '「自改ﾊﾞｰｼﾞｮﾝ：正当チェック異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   ' Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_OK, 0) 'V1.4.0.1 DEL
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_ERROR, lngErrCode)  'V1.4.0.1 ADD
    fHankukaChck = False   '戻り値をエラーとする
    'HAN_KUKA.KUKファイルをクローズする。
    'Close #iFileNumber                        'V1.4.0.1 DEL
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fOldVersion
'//  機能名称  : 旧バージョン処理
'//  機能概要  : 一世代前のバージョンを実行(実行)バージョンに返す。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-29   REVISED BY [TCC] S.Terao
'//                フェーズ３対応　管理へのメール送信処理を「ワーク→実行コピー」時にあわせた
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                【運改表示改善対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fOldVersion() As Boolean
    Dim bRet As Boolean                     '戻り値
    Dim lngCnt                  As Long     'カウンター
    Dim sSrcFileName            As String   '旧フォルダ内ファイルリスト
    Dim lngSumRet               As Long
    Dim lngErrCode              As Long     'エラーコード
    Dim iKansiAplChk As Integer              'アプリ起動チェック戻り値　'V1.6.0.1 ADD

    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD
    
    Dim sCorner As String                      'コーナー番号
    Dim sGatePath As String                    'コーナー番号付ファイルパス
    Dim sFilePath As String                    'ファイルファイルパス
    
    On Error Resume Next
 
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

   '旧フォルダ内のファイルリストを検索する。
'    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '「旧」フォルダ内ファイル名を作成する
    'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする  'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else                                'ファイルが存在しない
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        MsgBox "「旧」フォルダ内の " & TitleBox(FolderSyubetu) & "に、" _
                   & Chr(vbKeyReturn) & "ファイルリストが存在しません。", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  旧→実行 コピー"
        '「自改ﾊﾞｰｼﾞｮﾝ：ファイルリスト無し」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)
 
        fOldVersion = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '処理を終了する
    End If
    
    '｢旧｣フォルダからファイルリストを取得する
    sFilePath = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu)

'    bRet = fReadFileList(FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
  
' EG20 V3.6.0.1 【統合TR-No.260】追加開始
    bRet = fDataFileCheck(sFilePath & "\" & MN_FILELIST)
    If bRet = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       
       If sNGSts <> "" And sNGKoumoku <> "" Then
          MsgBox "運賃データ正当性チェック異常(" & sNGSts & "：" & sNGKoumoku & "）", _
                 vbOKOnly + vbExclamation, _
                 "自動改札機 バージョン管理"
       Else
          MsgBox "異常終了しました。", _
                 vbOKOnly + vbExclamation, _
                 TitleBox(FolderSyubetu) & "  旧→実行 コピー"
       End If
       Call pubfuncErrorOccur(MN_FOLD_OLD)
       fOldVersion = False
       Exit Function
    End If
' EG20 V3.6.0.1 【統合TR-No.260】追加終了
  
' EG20 V3.0.0.2 追加開始
    If pubfuncCommonGateCheck(MN_FOLD_OLD) = False Then
        fOldVersion = False
       Exit Function
    End If
' EG20 V3.0.0.2 追加終了
  
    '｢実行｣フォルダ内のファイルを全て削除する
    If sNowFolderRemove <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.0.0.2追加
        fOldVersion = False
        Exit Function
    End If
    
    '｢旧｣フォルダ内のファイルを｢実行｣フォルダの内容に置換える
    If sCopyOLDtoNOW <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.0.0.2追加
        fOldVersion = False
        Exit Function
    End If

    Call pubfuncCommonAreaUpdate                ' EG20 V3.0.0.2 追加

'V1.6.0.1 DEL START
'   '自改バージョン情報更新要求メールを管理プロセスへ送信する。
'     If gStrCurrentForm = sFormName_EJVer Then
'        psVersionUpdateReqest (ML_REQUEST_EGATE)
'     Else
'        psVersionUpdateReqest (ML_REQUEST_NGATE)
'     End If
'V1.6.0.1 DEL END
'V1.6.0.1 ADD START
    '自改バージョン情報更新要求メールを管理プロセスへ送信する。
    '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '監視盤起動時：管理プロセスに自改バージョン情報更新要求メールを送信する。
      If gStrCurrentForm = sFormName_EJVer Then
         psVersionUpdateReqest (ML_REQUEST_EGATE)
      Else
         psVersionUpdateReqest (ML_REQUEST_NGATE)
      End If
    Else
        '監視盤未起動時：改札機バージョン更新処理結果に正常を設定する。
        gintGateVerInfUpdRes = MailSts.stsNormal
    End If
'V1.6.0.1 ADD END
     
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
     
     '改札機バージョン更新処理異常
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
' EG20 V5.8.0.1追加開始
        ' 運改状態更新
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1追加終了
'        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1)   ' EG20 V5.6.0.1追加           ' EG20 V5.11.0.1削除
        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1追加
        '正常
        MsgBox "「旧」フォルダの内容を、「実行」フォルダに戻して、" _
                    & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "の一世代前のバージョンを、" _
                    & Chr(vbKeyReturn) & "実行バージョンとしました。", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  旧→実行 コピー"
        fOldVersion = True
    Else
        '異常
        If gStrCurrentForm = sFormName_EJVer Then
          MsgBox "改札機のバージョン作成で異常が発生しました。", _
                  vbOKOnly + vbExclamation, _
                  "自動改札機 バージョン管理"
        Else
           MsgBox "改札機のバージョン作成で異常が発生しました。", _
                   vbOKOnly + vbExclamation, _
                   "自動改札機 バージョン管理"
        End If
        fOldVersion = False
    End If

    fOldVersion = True
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sCopyOLDtoNOW
'//  機能名称  : 旧バージョンに戻す処理
'//  機能概要  : 旧フォルダ内のファイルを、実行フォルダにコピーする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.5.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sCopyOLDtoNOW() As Boolean
    Dim i As Integer                'カウンタ
    Dim sSrcFileName As String      'コピー元ファイル名
    Dim sDstFileName As String      'コピー先ファイル名
    Dim bError As Boolean           'エラーフラグ
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD

    Dim sCorner As String                      'コーナー番号
    Dim sGatePath As String                    'コーナー番号付ファイルパス
    
    On Error GoTo ErrorHandler
    
    '初期値設定
    sCopyOLDtoNOW = True

    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

    '****************************
    '* ファイルリストをコピーする *
    '****************************
'    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '「旧」フォルダ内ファイル名を作成する
'    sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
    sDstFileName = sGatePath & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '「実行」フォルダ内ファイル名を作成する
    'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする  'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then 'ファイルの検索をする   'V1.20.0.1 ADD
        'ファイルリストを「旧」フォルダから「実行」フォルダにコピーする
        FileCopy sSrcFileName, sDstFileName
    Else
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       
       '「旧フォルダファイルリストなし」ポップアップ画面表示
        MsgBox "「旧」フォルダ内の " & TitleBox(FolderSyubetu) & "に、" _
                   & Chr(vbKeyReturn) & "ファイルリストが存在しません。", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  旧→実行 コピー"
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '処理を終了する
    End If

    bError = False                  'エラーフラグを「偽」にする
    For i = 0 To UBound(FileList) - 1
                                    'ファイルリスト数分繰り返す
        '旧フォルダ内ファイル名を作成する
'        sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)
        sSrcFileName = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)

        '実行フォルダ内ファイル名を作成する
'        sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)
        sDstFileName = sGatePath & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)

        '旧フォルダ内のファイルを実行フォルダにコピーする
        'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする  'V1.20.0.1 DEL
        If objFso.FileExists(sSrcFileName) = True Then 'ファイルの検索をする   'V1.20.0.1 ADD
            'ファイルを「旧」フォルダから「実行」フォルダにコピーする
            FileCopy sSrcFileName, sDstFileName
        Else                                'ファイルが存在しない
            bError = True                   'エラーフラグを「真」にする
        End If
    Next
    If bError = True Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '「旧フォルダファイルリスト登録なし」ポップアップ画面表示
        MsgBox "「旧」フォルダ内の " & TitleBox(FolderSyubetu) & "に、" _
                   & Chr(vbKeyReturn) & "ファイルリストに登録されていて、存在しないファイルがありました。", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  旧→実行 コピー"
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function
    End If

    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2追加開始
    If pfuncCopyPASSINF(iTab_index, MN_FOLD_OLD) = False Then
' EG20 V3.5.0.1追加開始
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        MsgBox "異常終了しました。", _
               vbOKOnly + vbExclamation, _
               TitleBox(FolderSyubetu) & "  旧→実行 コピー"
' EG20 V3.5.0.1追加終了
        sCopyOLDtoNOW = False
    End If
' EG20 V3.0.0.2追加終了
    
    Exit Function       '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    '「旧→実行コピー異常終了」ポップアップ画面表示
    MsgBox "異常終了しました。", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  旧→実行 コピー"
        
    sCopyOLDtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMakeOutPutFile
'//  機能名称  : 媒体出力処理を行う。
'//  機能概要  : 媒体出力ファイル作成と出力を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-17  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-16  REVISED BY [TCC] M.Matsumoto
'//                 【統-273対応】
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-14 REVISED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeOutPutFile()
   Dim sOutFileName As String '媒体出力ファイル名[種別別]
   Dim iFileNumber As Integer 'ファイル番号
   Dim i As Integer           'カウンター
   Dim bFlag As Boolean       'フラグ
   Dim iResponse As Integer   'MsgBox戻り値
   Dim lngErrCode As Long     'エラーコード
   Dim fso         As New FileSystemObject   'ファイルシステムオブジェクト
   Dim strWriteDir As String               '出力先フォルダ
   Dim strStationName As String
' EG20 V2.0.1.1 ADD START【残件№60】
   Dim iTab_index  As Integer
   Dim strSyubetu As String     ' 種別名
' EG20 V2.0.1.1 ADD END【残件№60】
    
   On Error Resume Next 'V1.21.0.1 ADD

' EG20 V2.0.1.1 ADD START【残件№60】
    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
' EG20 V2.0.1.1 ADD END  【残件№60】

  'フォルダ選択部に指定有無チェック
  bFlag = False                                 'フラグを「偽」にする
  For i = 0 To 2                                'フォルダ数分繰り返す
     If chkFolder(i).Value = CHECKBOX_ON Then   '「？？」フォルダが指定されている
        bFlag = True                            'フラグを「真」にする
        Exit For                                'ループを抜ける
     End If
  Next
              
  If bFlag = False Then                       'フォルダ指定無し
     If gStrCurrentForm = sFormName_EJVer Then
       '「表示フォルダ指定なし」ポップアップ表示
         MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                 vbOKOnly + vbExclamation, _
                 "自動改札機 バージョン管理"
     Else
       '「表示フォルダ指定なし」ポップアップ表示
         MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                vbOKOnly + vbExclamation, _
                "自動改札機 バージョン管理"
     End If
         '処理を抜ける
     Exit Function
   End If
  
  
    'EG20 V2.1.0.1 ADD START 【統-273対応】
    If lstKan(iTab_index).ListCount = 0 Then
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Exit Function
    End If
    'EG20 V2.1.0.1 ADD END
  
  'フォルダ選択ポップアップ画面表示
'  strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")                         'V1.12.0.1 DEL
  strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.12.0.1 ADD

  '指定フォルダなし
  If Len(strWriteDir) = 0 Then
       Exit Function
  End If

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
  'プログレスバーを表示する
  Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

  'コピー先フォルダの有無確認
  If fso.FolderExists(strWriteDir) = False Then
     'コピー先フォルダ作成
     fso.CreateFolder (strWriteDir)
  End If
   
  '駅名取得
   strStationName = gsGetStationEkiName
   
   strSyubetu = ""
   '処理中フォームにより、媒体出力するファイル名作成
'   If gStrCurrentForm = sFormName_EJVer Then
       'リソース選択部分岐
       Select Case FolderSyubetu
        Case 0      '判定CPU-Pro
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJHANTEIPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJHANTEIPRO      'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJHANTEIPRO 'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJHANTEIPRO  'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJHANTEIPRO    'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJHANTEIPRO 'V1.8.0.1 ADD     'EG20　V30.1.0.1 DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJHANTEIPRO 'V30.1.0.1 ADD
          strSyubetu = "判定データ"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 1      'メインCPU-Pro
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINPRO   'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJMAINPRO    'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJMAINPRO   'EG20 EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJMAINPRO   'V1.8.0.1 ADD 'EG20 V30.1.0.1 DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJMAINPRO     'EG20 V30.1.0.1 ADD
          strSyubetu = "プログラム"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 2      'サブCPU-Pro1
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJSUBPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJSUBPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJSUBPRO   'V1.8.0.1 ADD
         ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJSUBPRO1    'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJSUBPRO1  'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJSUBPRO1   'V1.8.0.1 ADD     'EG20 V30.1.0.1 DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJSUBPRO1   'EG20 V30.1.0.1 ADD
          strSyubetu = "サブCPU-Pro1"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 3      'サブCPU-Pro2
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINOS        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINOS   'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJSUBPRO2    'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJSUBPRO2  'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJSUBPRO2   'V1.8.0.1 ADD     'EG20 V30.1.0.1 DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJSUBPRO2   'EG20 V30.1.0.1 ADD
          strSyubetu = "サブCPU-Pro2"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 4      'サブCPU-Pro3
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI1         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJSUBPRO3    'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJSUBPRO3  'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJSUBPRO3    'V1.8.0.1 ADD    'EG20 V30.1.0.1 DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJSUBPRO3     'EG20 V30.1.0.1 ADD
          strSyubetu = "サブCPU-Pro3"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 5      'メインCPU-OS
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINOS        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINOS   'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJMAINOS     'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJMAINOS   'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJMAINOS   'V1.8.0.1 ADD      'EG20 V30.1.0.1　DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJMAINOS    'EG20 V30.1.0.1 ADD
          strSyubetu = "自改（ＯＳ）"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 6      '予備1
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI1         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJYOBI1  'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJYOBI1    'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJYOBI1      'EG20 V30.1.0.1 ADD
          strSyubetu = "予備１"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 7      '予備2
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI2
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI2         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI2    'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJYOBI2  'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJYOBI2     'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJYOBI2    'V1.8.0.1 ADD  'EG20 V30.1.0.1 DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJYOBI2       'EG20 V30.1.0.1 ADD
          strSyubetu = "予備２"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        Case 8      '予備3
'        ' EG20 V2.0.1.1 DEL START【残件№60】
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI3
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI2         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI3    'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   【残件№60】
        ' EG20 V2.0.1.1 ADD START【残件№60】
          'sOutFileName = PATH_WORK & strStationName & VER_TXT_EJYOBI3  'EG20 V30.1.0.1 DEL
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJYOBI3    'EG20 V30.1.0.1 ADD
          'strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJYOBI3    'V1.8.0.1 ADD      'EG20 V30.1.0.1 DEL
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_EJYOBI3    'V1.8.0.1 ADD
          strSyubetu = "予備３"
        ' EG20 V2.0.1.1 ADD END  【残件№60】
        End Select
'  Else
'       'リソース選択部分岐
'       Select Case FolderSyubetu
'        Case 0      '判定CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJHANTEIPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJHANTEIPRO      'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJHANTEIPRO 'V1.8.0.1 ADD
'        Case 1      'メインCPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJMAINPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJMAINPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJMAINPRO   'V1.8.0.1 ADD
'        Case 2      'サブCPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJSUBPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJSUBPRO         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJSUBPRO    'V1.8.0.1 ADD
'        Case 3      'メインCPU-OS
'          sOutFileName = PATH_WORK & VER_TXT_NJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_NJMAINOS         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJMAINOS    'V1.8.0.1 ADD
'        Case 4      '予備1
'          sOutFileName = PATH_WORK & VER_TXT_NJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_NJYOBI1          'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJYOBI1     'V1.8.0.1 ADD
'        Case 5      '予備2
'          sOutFileName = PATH_WORK & VER_TXT_NJYOBI2
'          'strWriteDir = strWriteDir & VER_TXT_NJYOBI2          'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJYOBI2     'V1.8.0.1 ADD
'        End Select
'  End If

  iFileNumber = FreeFile              '未使用のファイル番号を取得する
 
  '対象ファイルをオープンする。
  Open sOutFileName For Output Access Write As #iFileNumber
  
  ' 設置駅名書き込み
   Print #iFileNumber, "設置駅：" & strStationName
   Print #iFileNumber, ""
     
  ' データ種別（ワーク）書き込み
   Print #iFileNumber, "データ種別：" & strSyubetu
   Print #iFileNumber, ""

  ' 全体バージョン書き込み
   Print #iFileNumber, "全体バージョン（ワーク）：" & DispTitleVersion(MN_FOLD_WRK)
   Print #iFileNumber, "　　　　　　　（実行）　：" & DispTitleVersion(MN_FOLD_NOW)
   Print #iFileNumber, "　　　　　　　（旧）　　：" & DispTitleVersion(MN_FOLD_OLD)
   Print #iFileNumber, ""

'  For i = 0 To lstKan(0).ListCount - 1
  For i = 0 To lstKan(iTab_index).ListCount - 1
  'リストボックスに表示されている分だけ、書き込む。
'       Print #iFileNumber, lstKan(0).List(i) & Chr(vbKeyReturn)
'       Print #iFileNumber, lstKan(iTab_index).List(i) & Chr(vbKeyReturn)   ' EG20 V3.0.0.2削除
       Print #iFileNumber, lstKan(iTab_index).List(i)                       ' EG20 V3.0.0.2追加
  Next
 
  '対象ファイルをクローズする。
  Close #iFileNumber

  'ファイルの有無確認
  If fso.FileExists(sOutFileName) = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
     'プログレスバーを消去する
     Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
     'ファイル無し異常ポップアップ画面表示
     MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
     Exit Function
  End If
    
  On Error GoTo COPY_ERROR
  'ファイルコピー
  fso.CopyFile sOutFileName, strWriteDir
  '「媒体出力正常終了」ポップアップ画面表示
  'V1.8.0.1 DEL START
  'iResponse = MsgBox("正常終了しました。", vbOKOnly, _
  '                   "出力結果")
  'V1.8.0.1 DEL END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
  'プログレスバーを消去する
  Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
  
  MsgBox "正常終了しました。", vbInformation, "出力結果"   'V1.8.0.1 ADD
                   
  '「自改ﾊﾞｰｼﾞｮﾝ：媒体出力処理正常」ログ出力
  Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_OK, 0)
  
  Set fso = Nothing

  Exit Function
    
'*******************************
'VBエラー処理
COPY_ERROR:
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        '処理異常の場合、出力結果ポップアップ(異常)表示
        MsgBox "異常終了しました。", vbCritical, "出力結果"
        '「自改ﾊﾞｰｼﾞｮﾝ：媒体出力処理異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_ERROR, lngErrCode)
        Set fso = Nothing
'*******************************
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sFDInstall
'//  機能名称  : 媒体インストール処理
'//  機能概要  : インストール媒体ファイルを、ワークフォルダにコピーする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//                 フェーズ２不具合修正
'//                 フェーズ３対応
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 入力ファイル格納ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//                 ファイル選択画面をOS仕様に変更
'//                 Dir関数をFileSystemObjectに置き換え
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                  EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                【運改表示改善対応】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【TOMAS用領域コピー対応】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall(sFlag As String)
    Dim MyName As String            'ファイルフルパス名
    Dim sSrcFileName As String      'コピー元ファイル名
    Dim sDstFileName As String      'コピー先ファイル名
    Dim iResponse As Integer        'MsgBoxボタンコード
    Dim sInputPass As String        'インストール元ディレクトリ名(STD)orファイル名(LZH)
    Dim sInputFolder As String      'インストール元フォルダ名。LZHの時、解凍先フォルダ。
    Dim lngErrCode As Long          'エラーコード
    'V1.6.0.1 ADD START
    Dim bRet As Boolean             '正当性チェック戻り値
    Dim sChkName As String          'チェックファイル
    'V1.6.0.1 ADD END
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    Dim objFi As File                    'ファイルオブジェクト
    'V1.20.0.1 ADD END
    
    Dim sCorner As String            'コーナー番号
    Dim sGatePath As String          'コーナー番号付ファイルパス
    Dim sFilePath As String          'ファイルファイルパス
    Dim lngPgmHanteiStsWork As Long     'プログラム判定状態（ワーク）   ' EG20 V3.0.0.2追加
    Dim szTargetFolder As String     ' 属性変更先フォルダ名             ' EG20 V5.8.0.1追加
    
    Dim sTomasPath As String         ' TOMAS用領域ファイルパス
    
    On Error GoTo ErrorHandler      'エラーハンドルの登録

    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner

' EG20 V5.8.0.1追加開始
    szTargetFolder = sGatePath & FolderName(0, FolderSyubetu)
' EG20 V5.8.0.1追加終了

    If sFlag = "STD" Then
    '標準（非圧縮）ファイル指定の時:
    'ディレクトリ選択画面を表示させ、入力ファイル格納ディレクトリ名を得る。
'       sInputPass = pfDirSelection("a:", "インストール媒体のディレクトリ選択")     'V1.12.0.1 DEL
        'sInputPass = pfDirSelection("H:", "インストール媒体のディレクトリ選択")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
        sInputPass = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
        If sInputPass = "" Then
        'ディレクトリが指定なし時は処理終了
            'V1.20.0.1 ADD START
            Set objFso = Nothing
            Set objFi = Nothing
            'V1.20.0.1 ADD END
            Exit Sub
        End If
        sInputFolder = sInputPass
    Else
    '圧縮ファイル指定の時:
    '圧縮ファイル選択画面を表示させ、LZHファイルフルパス名を得る（デフォルトはＦＤを表示。）。
'       sInputPass = pfCabFileSelection("a:")     'V1.12.0.1 DEL
        'V1.20.0.1 DEL START
       'sInputPass = pfCabFileSelection("H:")      'V1.12.0.1 ADD
        'If sInputPass = "" Then Exit Sub 'ファイルが選択されなければ戻る。
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
        '取得ファイル名を初期化
        CommonDialog1.FileName = ""
        '初期ディレクトリを設定
        If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'フォルダ選択画面デフォルトパス１が存在するか
            '存在するため、デフォルトパス１（H:）を設定
            CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
        Else
            '存在しないため、デフォルトパス２（C:）を設定
            CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
        End If
        '拡張子を設定
        CommonDialog1.Filter = "圧縮ファイル（*.cab）|*.cab|"
        'ファイル選択画面を開く
        CommonDialog1.ShowOpen
        '選択したファイル名を取得
        sInputPass = CommonDialog1.FileName
        If sInputPass = "" Then 'ファイル未選択
            Set objFso = Nothing
            Set objFi = Nothing
            Exit Sub    'ファイルが選択されなければ処理中断
        End If
        
        Call ChDrive("D")  'V2.5.0.1 ADD
        
        'V1.20.0.1 ADD END
       '解凍用一時フォルダを作成する。
       psMakeFolder MELTED_FOLDER_FULLPASS
       '圧縮ファイルを、解凍用一時フォルダに解凍・格納させる。
        Call psCabReqest(CABREQEST.CAB_THAW, sInputPass, MELTED_FOLDER_FULLPASS)
        If glngCabErrCd <> 0 Then
            psDeleteFolder MELTED_FOLDER_FULLPASS
            'V1.20.0.1 ADD START
            Set objFso = Nothing
            Set objFi = Nothing
            'V1.20.0.1 ADD END
            Exit Sub
        End If
        sInputFolder = MELTED_FOLDER_FULLPASS
    End If
    
    '「ワークコピー確認」ポップアップ画面表示
    iResponse = MsgBox(sInputPass & " の全てのファイルを、" _
                       & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                       & TitleBox(FolderSyubetu) & "の「ワーク」フォルダにコピーします。 " _
                       & "よろしいですか？", _
                       vbYesNo + vbExclamation, _
                       TitleBox(FolderSyubetu) & "  媒体→ワーク コピー")
    If iResponse = vbNo Then
    '[いいえ] ボタンを選択:何もしない。
    '但し、圧縮ファイル指定の時は、解凍用一時フォルダを削除する。
       If sFlag = "LZH" Then
           psDeleteFolder MELTED_FOLDER_FULLPASS
       End If
        'V1.20.0.1 ADD START
        Set objFso = Nothing
        Set objFi = Nothing
        'V1.20.0.1 ADD END
        Exit Sub
    End If
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'V1.6.0.1 ADD START
    '外部入力プロ判正当性チェック
    If sFlag = "STD" Then
       '媒体→ワーク コピー時
       bRet = pfInstallSeitouseiChck(sInputPass)
    Else
       '圧縮ファイル→ワーク コピー時
       bRet = pfInstallSeitouseiChck(MELTED_FOLDER_FULLPASS & "\")
    End If
    If bRet = False Then
        Call pubfuncErrorOccur(MN_FOLD_WRK)         ' EG20 V3.0.0.2追加
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
       If sFlag = "LZH" Then
           psDeleteFolder MELTED_FOLDER_FULLPASS
       End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
       'V1.20.0.1 ADD START
       Set objFso = Nothing
       Set objFi = Nothing
       'V1.20.0.1 ADD END
       Exit Sub
    End If
    
    'バージョンチェックファイル有無チェックを行う。
    sChkName = fSelectFile
    'V1.20.0.1 DEL START
'    sChkName = Dir(FolderName(0, FolderSyubetu) & "\" & sChkName)
'    If sChkName <> "" Then
'      Kill FolderName(0, FolderSyubetu) & "\" & sChkName
'    End If
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    
'    If objFso.FileExists(FolderName(0, FolderSyubetu) & "\" & sChkName) = True Then
    If objFso.FileExists(sFilePath & "\" & sChkName) = True Then
        '指定ファイルが存在する
'        sChkName = objFso.GetFileName(FolderName(0, FolderSyubetu) & "\" & sChkName)
        sChkName = objFso.GetFileName(sFilePath & "\" & sChkName)
'        Kill FolderName(0, FolderSyubetu) & "\" & sChkName
        Kill sFilePath & "\" & sChkName
    Else
        sChkName = ""
    End If
    'V1.20.0.1 ADD END
    'V1.6.0.1 ADD START
    
    '指定フォルダ内のファイルを、全て「ワーク」フォルダにコピーする。
    'V1.20.0.1 DEL START
'    MyName = Dir(sInputFolder & "\*.*", vbNormal)  ' 最初のディレクトリ名を返します。
'    Do While MyName <> ""                   ' ループを開始します。
'        ' 現在のディレクトリと親ディレクトリは無視します。
'        If MyName <> "." And MyName <> ".." Then
'            '媒体内ファイル名を作成する
'            sSrcFileName = sInputFolder & "\" & MyName
'            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
'            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
'                'ワークフォルダ内ファイル名を作成する
'                sDstFileName = FolderName(0, FolderSyubetu) & "\" & MyName
'                '媒体内のファイルをワークフォルダにコピーする
'                FileCopy sSrcFileName, sDstFileName
'            End If
'        End If
'        MyName = Dir                    ' 次のディレクトリ名を返します。
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(sInputFolder).files   'ループを開始
        If objFso.FileExists(objFi.Path) = True Then  'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            '媒体内ファイル名を作成
            sSrcFileName = sInputFolder & "\" & MyName
            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
                'ワークフォルダ内ファイル名を作成する
'                sDstFileName = FolderName(0, FolderSyubetu) & "\" & MyName
                sDstFileName = sGatePath & FolderName(0, FolderSyubetu) & "\" & MyName

                '媒体内のファイルをワークフォルダにコピーする
                FileCopy sSrcFileName, sDstFileName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    '圧縮ファイル指定の時は、解凍用一時フォルダを削除する。(使用済みのため)
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
    
' EG20 V5.8.0.1削除開始
'    ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
    '読み取り外しの関数を実行
    dllChangeAttributeContents (szTargetFolder)

' EG20 V6.9.0.1 【量産対応：TOMAS用領域コピー対応】ADD START
    ' 処理すべき対象がコーナ1の場合
    ' TOMAS領域（N_GATE00）もN_GATE01の内容でコピー
    'If iTab_index = 0 Then     'EG20  V30.1.0.1 DEL
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
    'ワークコピーしようとするたびにそのコーナから00へコピーするため、先頭コーナの判定を削除
    'If iTab_index = gintZairaiFirstCornerIdx Then  'EG20 V30.1.0.1 ADD
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
        ' 削除先のフォルダ（TOMAS領域）を指定
        sTomasPath = PATH_N_GATE & "00" & FolderName(0, FolderSyubetu) & "\"
        sInputFolder = sGatePath & FolderName(0, FolderSyubetu) & "\"
        
        ' TOMAS領域を削除
        If funcRemoveFile(sTomasPath) = False Then
            
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
            MsgBox "ＴＯＭＡＳ用領域コピー異常終了", _
                    vbOKOnly + vbExclamation, _
                    "自動改札機　バージョン管理"
            
            '「自改ﾊﾞｰｼﾞｮﾝ：TOMASﾌｫﾙﾀﾞﾌｧｲﾙ削除異常」ログ出力
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_DELETE_ERROR, lngErrCode)
        
            GoTo TomasErrorHandler
        End If
        
        ' TOMAS領域へコピー
        If funcCopyFile(sInputFolder, sTomasPath, lngErrCode) = False Then
            
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
            MsgBox "ＴＯＭＡＳ用領域コピー異常終了", _
                    vbOKOnly + vbExclamation, _
                    "自動改札機　バージョン管理"
            
            '「自改ﾊﾞｰｼﾞｮﾝ：TOMAS領域ｺﾋﾟｰ処理異常」ログ出力
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_COPY_ERROR, lngErrCode)
        
            GoTo TomasErrorHandler
        End If
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
    'End If
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
' EG20 V6.9.0.1 【量産対応：TOMAS用領域コピー対応】ADD END

    ' 運改状態更新
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1追加終了
'    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iTab_index + 1)   ' EG20 V5.6.0.1追加           ' EG20 V5.11.0.1削除
    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1追加
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '「ワークコピー正常終了」ポップアップ画面表示
    MsgBox "インストール媒体の全てのファイルを、" _
            & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "の「ワーク」フォルダに" _
            & Chr(vbKeyReturn) & "コピーしました。", _
            vbOKOnly + vbExclamation, _
            TitleBox(FolderSyubetu) & "  媒体→ワーク コピー"
    
    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_OK, 0)
    
    'リストボックスを初期化する
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
  
    'バージョン情報リストボックスを作成する
    fMakeListbox
    
    '監視設定エリア「プログラム判定異常状態（ワーク）」の状態を取得する
    lngPgmHanteiStsWork = pfGetKansiSet(IdKansiSet.PG_HANTEI_ERR_STS_WORK)

    '「プログラム判定異常状態（ワーク）」（正常）
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_ERR_STS_WORK, ErrCode.Normal)
    
    '変化があった場合、「状態変化通知」を送信する
    If lngPgmHanteiStsWork <> ErrCode.Normal Then
        Call sSendMailStsChgInf(MailSts.stsNormal, lngPgmHanteiStsWork)
    End If
    
    Exit Sub    '処理を終了する

ErrorHandler:   ' エラー処理。
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    'V1.20.0.1 ADD END
    Select Case Err.Number
        Case 53 ' 「指定種別ファイルなし」ポップアップ画面表示
            MsgBox "インストール媒体に " & TitleBox(FolderSyubetu) & "は、" _
                   & Chr(vbKeyReturn) & "ひとつも存在しません。", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  →ワーク コピー"
            Exit Sub
        Case 71 '「媒体なし」ポップアップ画面表示
            iResponse = MsgBox("媒体が準備されていません。", _
                    vbRetryCancel + vbExclamation, _
                    TitleBox(FolderSyubetu) & "  →ワーク コピー")
            If iResponse = vbRetry Then    '「やり直し」ボタンを選択した場合
                Resume      ' エラーが発生した行から処理再開
            Else                            '「キャンセル」ボタンを選択した場合
                Exit Sub    '処理を終了する
            End If
        Case Else  '「ワークコピー異常終了」ポップアップ画面表示
           MsgBox "インストール媒体からのコピーエラーが発生しました。" _
                   & Chr(vbKeyReturn) & "エラーコード＝" _
                   & str$(Err.Number), _
                   vbOKOnly + vbExclamation, _
                   "→ワーク コピー"
    End Select
    
    Call pubfuncErrorOccur(MN_FOLD_WRK)         ' EG20 V3.0.0.2追加

' EG20 V5.8.0.1追加開始
    dllChangeAttributeContents (szTargetFolder)
' EG20 V5.8.0.1追加終了
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    '圧縮ファイル指定の時は、解凍用一時フォルダを削除する。
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End

    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)

' EG20 V6.9.0.1 【量産対応：TOMAS用領域コピー対応】ADD START
    Exit Sub    '処理を終了する

TomasErrorHandler:   ' TOMAS処理用エラー処理。
' EG20 V6.9.0.1 【量産対応：TOMAS用領域コピー対応】ADD END
    
    Call pubfuncErrorOccur(MN_FOLD_WRK)
    
    'リストボックスを初期化する
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
  
    'バージョン情報リストボックスを作成する
    fMakeListbox

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : SSTab1_Click
'//  機能名称  : コーナタブ選択処理
'//  機能概要  : コーナ表示を切り替える
'//
'//              型        名称             意味
'//  引数      : Integer   PreviousTab      選択タブ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'/  REVISIONS    : (EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/                  EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub SSTab1_Click(PreviousTab As Integer)
    
    On Error GoTo ErrorHandle
    
    'リストボックスを初期化する
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
    
    'バージョン情報リストボックスを作成する
    fMakeListbox
ErrorHandle:

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pubfuncCommonGateCheck
'//  機能名称  : 改札機共通判定処理
'//  機能概要  : サム値チェック、ファイル数最大チェックの実行
'//
'//              型         名称            意味
'//  引数      : Integer    nKind           MN_FOLD_WRK(0):ワーク
'//                                         MN_FOLD_NOW(1):実行
'//                                         MN_FOLD_OLD(2):旧
'//
'//              型        値        意味
'//  戻り値    : BOOL      TRUE      正常
'//                        FALSE     異常
'//
'//  ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                  EG20フェーズ２対応
'//  REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pubfuncCommonGateCheck(nKind As Integer) As Boolean

    Dim lngSumRet As Long
    Dim lngCnt As Long
    Dim lngFileListCnt As Long               'ファイルリスト数
    Dim i As Integer
    Dim strWork     As String                '作業エリア
    Dim iFileNumber As Integer               '未使用ファイル番号
    Dim bRet As Boolean
    Dim sGetFileListName As String           'ファイルリスト内記載ファイル名
    Dim myLen As Long                        '文字列の長さ
    Dim sCorner As String                    'コーナー番号
    Dim sGatePath As String                  'コーナー番号付ファイルパス
    Dim sFilePath As String                  'ファイルファイルパス
    Dim lTotalCount As Long                  ' 結果件数

    Dim lngPgmHanteiRcvErrSts   As Long     'プログラム判定受信異常状態
    Dim lngPgmHanteiSndErrSts   As Long     'プログラム判定配信異常状態
    Dim lngPgmHanteiErrSts      As Long     'プログラム判定異常状態（実行）
    Dim lngPgmHanteiErrStsOld   As Long     'プログラム判定異常状態（旧）
    Dim lngPgmHanteiElseErrSts  As Long     'プログラム判定その他異常状態

    
    On Error Resume Next

    ' 選択中のコーナー番号取得
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' コーナー番号付ファイルパス作成
    sGatePath = PATH_N_GATE & sCorner


    ' /////////////////////////////////////////////////////
    ' // サム値チェック
    For lngCnt = 0 To UBound(FileList) - 1
        sFilePath = sGatePath & FolderName(nKind, FolderSyubetu)
        If pfFileSumChk(sFilePath & "\" & FileList(lngCnt), lngSumRet) <> True Then
            
            '「プログラム判定受信異常状態」取得
            lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
        
            '監視設定エリア「プログラム判定受信異常状態」を更新
            Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_SumChk)
                    
            '監マプロセスに「状態変化通知」を送信
            If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_SumChk Then
                Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_SumChk)
            End If
            
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
            'サム値異常
            If lngSumRet = SUM_CHK.SumErr Then
               MsgBox "サム値が異常です。" _
                      & Chr(vbKeyReturn) & "データを確認してください。", _
                      vbOKOnly + vbExclamation, _
                      "自動改札機 バージョン管理"
            
            'サム値異常以外異常
            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
               MsgBox "異常終了しました。", _
                     vbOKOnly + vbExclamation, _
                      "自動改札機 バージョン管理"
            End If
            pubfuncCommonGateCheck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    ' /////////////////////////////////////////////////////
    ' // ファイル数最大チェック
    If UBound(FileList) > FILECNT_MAX Then

        '「プログラム判定受信異常状態」
        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)

        '監視設定エリア「プログラム判定受信異常状態」を更新
        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
                
        '監マプロセスに「状態変化通知」を送信
        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
        End If
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

        MsgBox "ファイル数が上限を超えています。" _
                & Chr(vbKeyReturn) & "データを確認してください。", _
                vbOKOnly + vbExclamation, _
                "自動改札機 バージョン管理"
        pubfuncCommonGateCheck = False

        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
        Exit Function
    End If

    ' /////////////////////////////////////////////////////
    ' // 全ファイル数最大チェック（実行＋追加分）
    bRet = True
    lTotalCount = pfuncTotalListCount()
    lTotalCount = lTotalCount + UBound(FileList)
    If lTotalCount > TOTALFILECNT_MAX Then
        bRet = False
    End If
    If bRet = False Then
        '「プログラム判定受信異常状態」
        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)

        '監視設定エリア「プログラム判定受信異常状態」を更新
        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
                
        '監マプロセスに「状態変化通知」を送信
        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
        End If
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        MsgBox "ファイル数が上限を超えています。" _
                & Chr(vbKeyReturn) & "データを確認してください。", _
                vbOKOnly + vbExclamation, _
                "自動改札機 バージョン管理"
        pubfuncCommonGateCheck = False

        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
        Exit Function
    End If

    pubfuncCommonGateCheck = True
    Exit Function

' 未実施
'    ' /////////////////////////////////////////////////////
'    ' // ファイル名サイズチェック
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '未使用のファイル番号を取得する
'
'    sFilePath = sGatePath & FolderName(nKind, FolderSyubetu)
'    'ファイルリストをオープン。
'    Open sFilePath & "\" & MN_FILELIST For Input As #iFileNumber
'
'    bRet = True
'    For i = 0 To lngFileListCnt
'        If i = lngFileListCnt Then
'            Exit For
'        End If
'
'        'ファイル名を取得する。
'        Input #iFileNumber, strWork
'        If strWork <> "" And Left$(strWork, 1) <> "/" Then  'ファイル名が存在する
'            'ファイル名定義なし
'            If strWork = "" Then
'                'ループ抜け
'                MsgBox "ファイル名が異常です。" _
'                        & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                        vbOKOnly + vbExclamation, _
'                        "自動改札機 バージョン管理"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            'フォーマット異常
'            ElseIf " " <> Mid(strWork, 2, 1) Then
'              'ループ抜け
'                MsgBox "ファイル名が異常です。" _
'                        & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                        vbOKOnly + vbExclamation, _
'                        "自動改札機 バージョン管理"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            ElseIf (InStr(strWork, ".") - 1) = -1 Then
'                MsgBox "ファイル名が異常です。" _
'                        & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                        vbOKOnly + vbExclamation, _
'                        "自動改札機 バージョン管理"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            Else
'                'ファイル名のみを抽出
'                sGetFileListName = Mid(strWork, 3, 16)
'                '取得ファイル名のサイズを取得
'                myLen = LenB(StrConv(sGetFileListName, vbFromUnicode))      '半角換算のバイト数を取得
'                If FILE_NAME_MAX_SIZE < myLen Then
'                    '13バイト以上の場合
'                    MsgBox "ファイル名が異常です。" _
'                            & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
'                            vbOKOnly + vbExclamation, _
'                            "自動改札機 バージョン管理"
'                    bRet = False
'                    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                    Exit For
'                End If
'            End If
'        End If
'    Next
'
'    If bRet = False Then
'        '「プログラム判定受信異常状態」
'        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
'
'        '監視設定エリア「プログラム判定受信異常状態」を更新
'        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
'
'        '監マプロセスに「状態変化通知」を送信
'        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
'            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
'        End If
'    End If
'    'ファイルリストをクローズ。
'    Close #iFileNumber
'    pubfuncCommonGateCheck = bRet

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pubfuncCommonGateCheck = False
    
    '「プログラム判定受信異常状態」
    lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)

    '監視設定エリア「プログラム判定受信異常状態」を更新
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
            
    '監マプロセスに「状態変化通知」を送信
    If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
        Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfuncTotalListCount
'//  機能名称  : 総リスト数の取得
'//  機能概要  : 指定種別以外の総ファイル数を算出する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値               意味
'//  戻り値    : LONG      lResultCount     件数
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：frmJVer.fReadFileList流用
'///////////////////////////////////////////////////////////////////
Private Function pfuncTotalListCount() As Long
    Dim lResultCount As Long                ' 結果件数
    Dim iLoop As Integer                    ' ループ
    
    Dim iFileNumber As Integer              'ファイル番号
    Dim sFileName As String                 'ファイル名
    Dim sSrcFileName As String              'ファイル名
    Dim iListCnt As Integer                 'ファイル格納数
    Dim sCorner As String                   'コーナー番号
    Dim sGatePath As String                 'コーナー番号付ファイルパス
    Dim objFso As New FileSystemObject      ' ファイルシステムオブジェクト

    On Error GoTo ErrorHandler      'エラーハンドル設定
    
    
    ' コーナー番号付ファイルパス作成
    sCorner = Format(iTab_index + 1, "00")
    sGatePath = PATH_N_GATE & sCorner
    
    lResultCount = 0
    iFileNumber = FreeFile   '未使用のファイル番号を取得する
    For iLoop = 0 To 8
        
        iFileNumber = FreeFile   '未使用のファイル番号を取得する
        sSrcFileName = sGatePath & FolderName(1, iLoop) & "\" & MN_FILELIST
   
        If objFso.FileExists(sSrcFileName) = True Then
   
            Open sSrcFileName For Input Access Read As #iFileNumber     'ファイルリストのオープン
            iListCnt = 0
            Do While Not EOF(iFileNumber)                               'ファイルの終端までループを繰り返します。
                Line Input #iFileNumber, sFileName                      'データを読み込みます。
                If sFileName <> "" And Left$(sFileName, 1) <> "/" Then  'ファイル名が存在する
                    iListCnt = iListCnt + 1                             'ファイル数のカウンタをアップする
                End If
            Loop
            Close #iFileNumber      'ファイルを閉じます。
            iFileNumber = 0
            If iLoop <> FolderSyubetu Then
                lResultCount = lResultCount + iListCnt
            End If
        End If
    Next

    pfuncTotalListCount = lResultCount    '戻り値を設定する
    Set objFso = Nothing

    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    pfuncTotalListCount = 0    '戻り値を設定する
    Set objFso = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pfuncCopyPASSINF
'//  機能名称  : 実行フォルダへのPASSINFコピー
'//  機能概要  : 指定種別以外の総ファイル数を算出する。
'//
'//              型        名称      意味
'//  引数      : Integer   nCorner   コーナ番号（0～5）
'//  引数      : Integer    nKind           MN_FOLD_WRK(0):ワーク
'//                                         MN_FOLD_NOW(1):実行
'//                                         MN_FOLD_OLD(2):旧
'//
'//              型        値               意味
'//  戻り値    : BOOL      TRUE             正常
'//            : BOOL      FALSE            異常
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：frmJVer.fReadFileList流用
'///////////////////////////////////////////////////////////////////
Private Function pfuncCopyPASSINF(nCorner As Integer, nKind As Integer) As Boolean
    
    Dim objFso As New FileSystemObject      ' ファイルシステムオブジェクト
    Dim szSrcFile As String                 ' コピー元ファイル
    Dim szDstFile As String                 ' コピー先ファイル
    Dim sCorner As String           'コーナー番号
    Dim sGatePath As String         'コーナー番号付ファイルパス

    On Error GoTo ErrorHandler              ' エラーハンドルの登録

    ' 対象が判定データの場合のみ処理を行う
    ' 上記に該当しない場合は正常終了
    If FolderSyubetu <> 0 Then
        pfuncCopyPASSINF = True
        Set objFso = Nothing
        Exit Function
    End If

    ' コーナー番号付ファイルパス作成
    sCorner = Format(nCorner + 1, "00")
    sGatePath = PATH_N_GATE & sCorner
    ' コピー元ファイル
    szSrcFile = sGatePath & FolderName(nKind, 0) & "\" & "PASSINF"
    szDstFile = sGatePath & FolderName(MN_FOLD_NOW, 0) & "\" & "PASSINF"

    If objFso.FileExists(szSrcFile) = True Then
        'ファイルコピー（既に存在した場合は上書きするする）
        objFso.CopyFile szSrcFile, szDstFile, True
        pfuncCopyPASSINF = True
    Else
        pfuncCopyPASSINF = False
    End If

    Set objFso = Nothing
    Exit Function

ErrorHandler:
    pfuncCopyPASSINF = False
    Set objFso = Nothing
End Function

