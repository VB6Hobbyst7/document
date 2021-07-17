VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJVer 
   BorderStyle     =   0  'なし
   Caption         =   "自動改札機バージョン管理"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   -210
   ClientWidth     =   12000
   ClipControls    =   0   'False
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "   メニュー     画面へ戻る"
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
      Left            =   9360
      TabIndex        =   30
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   9720
      TabIndex        =   29
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "テキスト表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   9720
      TabIndex        =   28
      Top             =   3240
      Width           =   2055
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
      Height          =   735
      Left            =   9720
      TabIndex        =   27
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdLzhFileCopy 
      Caption         =   " 圧縮ファイル      →ワーク コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   26
      Top             =   7200
      Width           =   2295
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
      Height          =   5580
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Frame fraResource 
      Caption         =   "表示リソース指定"
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
      Height          =   1575
      Left            =   7560
      TabIndex        =   17
      Top             =   720
      Width           =   4215
      Begin VB.OptionButton optSyubetu 
         Caption         =   "判定CPU-Pro"
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
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "メインCPU-Pro"
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
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "サブCPU-Pro"
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
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "メインCPU-OS"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "予備1"
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
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "予備2"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "旧→実行 コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   5400
      TabIndex        =   14
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "ワーク→実行 コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   5400
      TabIndex        =   13
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "媒体→ワーク コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2760
      TabIndex        =   12
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "ワーク クリア"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton cmdVer 
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
      Height          =   735
      Index           =   0
      Left            =   9720
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "自改切り離し"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   9720
      TabIndex        =   10
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Frame fraVersion 
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
      Height          =   1815
      Left            =   7560
      TabIndex        =   16
      Top             =   2520
      Width           =   2055
      Begin VB.CheckBox chkFolder 
         Caption         =   "W ワーク"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   1  'ﾁｪｯｸ
         Width           =   1815
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "O 旧"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Value           =   1  'ﾁｪｯｸ
         Width           =   1815
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "N 実行"
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
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Value           =   1  'ﾁｪｯｸ
         Width           =   1815
      End
   End
   Begin VB.Timer tmrMail 
      Left            =   8160
      Top             =   8040
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "NEG自動改札機バージョン管理"
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
      TabIndex        =   31
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "コメント"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   25
      Top             =   1200
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
      Index           =   5
      Left            =   6840
      TabIndex        =   24
      Top             =   840
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
      Index           =   4
      Left            =   5190
      TabIndex        =   23
      Top             =   840
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
      Index           =   3
      Left            =   3930
      TabIndex        =   22
      Top             =   840
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
      Index           =   2
      Left            =   2880
      TabIndex        =   21
      Top             =   840
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
      Index           =   1
      Left            =   2040
      TabIndex        =   20
      Top             =   840
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
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   840
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
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "frmJVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmJVer.frm
'//  パッケージ名：バージョン管理(EG-R自改/NEG自改)画面
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
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Dim FolderSyubetu As Integer                 '選択リソース種別

Dim FolderName(0 To 2, 0 To 7) As String     'フォルダ名
Dim TitleBox(0 To 7) As String               'タイトル名
Dim LogBox(0 To 7) As String                 'ログ出力用タイトル名
Dim FileList() As String                     'ファイル名リスト一覧格納エリア
Dim FileListType() As String                 'ファイルリスト一覧格納エリア（次世代自改タイプを含む）
Dim uVersion() As MN_VERSION_JIKAI           'バージョン情報格納エリア

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
Private Const EGR_JIKAI = "EG-R"                'EG-R
Private Const NEG_JIKAI = "NEG"                 'NEG
'V1.4.0.1　ADD　END
'V1.6.0.1 ADD START
Private Const EGR_JIKAI_KISHU = "EG5000"        'EG-R自改機種名
Private Const NEG_JIKAI_KISHU = "EG2000"        'NEG自改機種名
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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン管理(EG-R/NEG自改)画面(アクティブ時)
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
'//  機能名称  : バージョン管理(EG-R/NEG自改)画面(ディアクティブ時)
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
'//  関数名称  : Form_Load
'//  機能名称  : バージョン管理(EG-R/NEG自改)画面(ロード時)
'//  機能概要  : メール受信用のタイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-10   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
   On Error Resume Next
 
    If gStrCurrentForm = sFormName_EJVer Then
       sJverName = EGR_JIKAI                        'V1.4.0.1 ADD
       Label1.Caption = "EG-R自動改札機バージョン管理"
       '「EG-R自動改札機ﾊﾞｰｼﾞｮﾝ画面：表示」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EJIKAI_VERASION_KANRI_GAMEN_START, 0)
    Else
       sJverName = NEG_JIKAI                        'V1.4.0.1 ADD
       '「NEG自動改札機ﾊﾞｰｼﾞｮﾝ画面：表示」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, NJIKAI_VERASION_KANRI_GAMEN_START, 0)
    End If
  
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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdUpdate_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 各釦押下による処理を行う。
'//             「ワーククリア」「媒体→ワークコピー」「ワーク→実行コピー」
'//             「旧→実行コピー」
'//
'//              型        名称      意味
'//  引数      : Integer   Index     [IN]押下釦インデックス値
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                  ・フェーズ１不具合対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdUpdate_Click(Index As Integer)
   Dim iResponse As Integer         'MsgBoxボタンコード
   Dim lngErrCode As Long           'エラーコード

   On Error Resume Next

' 押されたボタンを判定する。
Select Case Index
   Case 0
        '「ワーククリア」ボタンの場合。
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
           'バージョン情報リストボックスを作成する
           fMakeListbox
        End If
        
   Case 1
        '「媒体→ワークコピー」ボタンの場合。
        '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ワークコピー釦押下」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_LZH_COPY_WRK_BUTTOM, 0)
        'インストール媒体をワークフォルダ内にコピーする
        sFDInstall "STD"
        
   Case 2
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
            'V1.6.0.1　DEL START
            'If CheckAppStart(PROC_KANRI) <> 0 Then
            '   '異常
            '   If gStrCurrentForm = sFormName_EJVer Then
            '      MsgBox "改札機のバージョン作成で異常が発生しました。", _
            '             vbOKOnly + vbExclamation, _
            '             "EG-R自動改札機 バージョン管理"
            '      Exit Sub
            '   Else
            '      MsgBox "改札機のバージョン作成で異常が発生しました。", _
            '             vbOKOnly + vbExclamation, _
            '             "NEG自動改札機 バージョン管理"
            '      Exit Sub
            '   End If
            'End If
            'V1.6.0.1　DEL END
            '最新バージョンを実行バージョンとして登録する
            If fNewVersion <> True Then
               '「自改ﾊﾞｰｼﾞｮﾝ：ワーク→実行コピー処理異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
               Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_ERROR, lngErrCode)
               Exit Sub
            End If
            '「自改ﾊﾞｰｼﾞｮﾝ：ワーク→実行コピー処理正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_OK, 0)
            'バージョン情報リストボックスを作成する
            fMakeListbox
        End If
        
   Case Else
        '「旧→実行コピー」ボタンの場合。
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
            '一世代前のバージョンを実行バージョンに戻す
           If fOldVersion <> True Then
              '「自改ﾊﾞｰｼﾞｮﾝ：ワーク→実行コピー処理異常」ログ出力
              lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
              Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_ERROR, lngErrCode)
              Exit Sub
            End If
            '「自改ﾊﾞｰｼﾞｮﾝ：旧→実行コピー処理正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_OK, 0)
            'バージョン情報リストボックスを作成する
            fMakeListbox
       End If
       
  End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdLzhFileCopy_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 各釦押下による処理を行う。
'//             「圧縮ファイル→ワークコピー」
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
Private Sub cmdLzhFileCopy_Click()
   
   On Error Resume Next
    
    '「自改ﾊﾞｰｼﾞｮﾝ：圧縮ﾌｧｲﾙ→ﾜｰｸｺﾋﾟｰ釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_CAB_COPY_WRK_BUTTOM, 0)
 
    '圧縮ファイルからインストールする。
    sFDInstall "LZH"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdVer_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 各釦押下による処理を行う。
'//             「表示更新」「テキスト表示」「媒体出力」
'//             「自改切り離し」
'//
'//              型        名称      意味
'//  引数      : Integer   Index     [IN]押下釦インデックス値
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdVer_Click(Index As Integer)
    Dim i As Integer                'カウンタ
    Dim bFlag As Boolean            'フラグ
    Dim lRetVal As Long             '戻り値
    Dim sCommand As String          'コマンド文字列
    Dim sWriteDir As String
    
    On Error GoTo ErrorHandle

    Select Case Index
        Case 0  '「表示更新」釦
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
                If gStrCurrentForm = sFormName_EJVer Then
                  '「表示フォルダ指定なし」ポップアップ表示
                   MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                          vbOKOnly + vbExclamation, _
                          "EG-R自動改札機 バージョン管理"
                Else
                  '「表示フォルダ指定なし」ポップアップ表示
                   MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                          vbOKOnly + vbExclamation, _
                          "NEG自動改札機 バージョン管理"
                End If
                '処理を抜ける
                Exit Sub
              End If
              'バージョン情報リストボックスを作成する
              fMakeListbox
              
        Case 1 '「テキスト表示」釦
            '「自改ﾊﾞｰｼﾞｮﾝ：テキスト表示釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_TXT_BUTTOM, 0)

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
                           "EG-R自動改札機 バージョン管理"
               Else
                 '「表示フォルダ指定なし」ポップアップ表示
                   MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                          vbOKOnly + vbExclamation, _
                          "NEG自動改札機 バージョン管理"
               End If
                   '処理を抜ける
               Exit Sub
             End If


            'リストボックスの内容をファイルに書き込み
            sWriteListbox
            sCommand = MN_EXE_MEMO & MN_VERSI_FILE 'メモ帳実行コマンドを作成する
            'メモ帳を起動する｡
            lRetVal = Shell(sCommand, vbMaximizedFocus)
            'メモ帳をアクティブ（前面表示）にする
            AppActivate lRetVal, True
            SendKeys "{LEFT}", True
           
        Case 2 '「媒体出力」釦
            '「自改ﾊﾞｰｼﾞｮﾝ：媒体出力釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OUTPUT_BUTTOM, 0)
 
            '媒体出力処理
             fMakeOutPutFile
           
        Case 3  '「自改切り離し」釦
            '「自改ﾊﾞｰｼﾞｮﾝ：自改切り離し釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_KIRIHANASI_BUTTOM, 0)

            '通信接続・切断画面を表示する。
            Load frmConectSts
            frmConectSts.Show 1
        Case Else
   End Select
ErrorHandle:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : optSyubetu_Click
'//  機能名称  : リソース釦押下時処理
'//  機能概要  : 対象リソース種別を変更する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                表示リソースラジオ釦選択でリストの表示更新
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub optSyubetu_Click(Index As Integer)
    'V1.20.0.1 ADD START
    Dim i As Integer                'カウンタ
    Dim bFlag As Boolean            'フラグ
    'V1.20.0.1 ADD END

    'リソース種別を変更する。'
    FolderSyubetu = Index
    
    'V1.20.0.1 ADD START
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
                    "EG-R自動改札機 バージョン管理"
        Else
            '「表示フォルダ指定なし」ポップアップ表示
            MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
                    vbOKOnly + vbExclamation, _
                    "NEG自動改札機 バージョン管理"
        End If
        '処理を抜ける
        Exit Sub
    End If
    
    'バージョン情報リストボックスを作成する
    fMakeListbox
    'V1.20.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                ②「メニュー画面へ戻る」釦押下にて、
'//                 　バージョン管理画面のバージョン表示更新を行う
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
   
   On Error Resume Next
    
   If gStrCurrentForm = sFormName_EJVer Then
      '「EG-R自動改札機ﾊﾞｰｼﾞｮﾝ画面：消去」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EJIKAI_VERASION_KANRI_GAMEN_END, 0)
   Else
      '「NEG自動改札機ﾊﾞｰｼﾞｮﾝ画面：消去」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, NJIKAI_VERASION_KANRI_GAMEN_END, 0)
   End If
   
   'V1.20.0.1 ADD START
   'バージョン管理画面のバージョン表示更新処理を行う。
   frmVersion.psGetVersion
   'V1.20.0.1 ADD END
   
   'NEG/EG-R自動改札機バージョン管理画面を閉じる
   Unload Me
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
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeListbox() As Boolean
    Dim bRet As Boolean                        '戻り値

    On Error Resume Next

    '***********************************************
    '* 次世代自改フォルダから全てのバージョン情報を取得する *
    '***********************************************
    ReDim uVersion(0)

    '｢ワーク｣フォルダからファイルリストを取得する
    bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = True Then
        'ファイルリストからバージョン情報を取得する
        sVersionInfo FolderName(0, FolderSyubetu), MN_FLDWRK
    End If

    '｢実行｣フォルダからファイルリストを取得する
    bRet = fReadFileList(FolderName(1, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = True Then
        'ファイルリストからバージョン情報を取得する
        sVersionInfo FolderName(1, FolderSyubetu), MN_FLDNOW
    End If

    '｢旧｣フォルダからファイルリストを取得する
    bRet = fReadFileList(FolderName(2, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = True Then
        'ファイルリストからバージョン情報を取得する
        sVersionInfo FolderName(2, FolderSyubetu), MN_FLDOLD
    End If

    'バージョン情報をファイル名順にソートする
    sListboxSort

    'バージョン情報をリストボックスにセットする
    sVerListDisp
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sVerListDisp
'//  機能名称  : バージョン情報リストボックス設定
'//  機能概要  : 取得したバージョン情報を、リストボックスに設定
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
Private Sub sVerListDisp()
    Dim i As Integer                        'カウンタ
    Dim uVerData(2) As MN_VERSION_JIKAI     'バージョン情報（各フォルダ）
    Dim lDataNum As Long                    'バージョン情報数

    On Error Resume Next

    'リストボックスを初期化する
    lstKan.Clear

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
End Sub

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
                                    lstKan.AddItem sFileName & "W N O" & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(0)
                                    End If
                                Else
                                    lstKan.AddItem sFileName & "W N  " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(0)
                                    End If
                                    lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(2)
                                    End If
                                    If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '｢旧｣フォルダにファイルがない
                                lstKan.AddItem sFileName & "W N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                     lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                            End If
                        Else                                '｢旧｣フォルダ非アクティブ表示
                            lstKan.AddItem sFileName & "W N  " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                        End If
                    Else                            '｢ワーク｣フォルダと｢実行｣フォルダのバージョンが違う
                        lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(0)
                        End If
                        If chkFolder(2).Value = CHECKBOX_ON Then   '｢旧｣フォルダ表示
                            If uVerData(2).sFileName <> "" Then
                                '｢実行｣フォルダと｢旧｣フォルダを比較する
                                If sFileInfo(1) = sFileInfo(2) Then
                                    lstKan.AddItem Space(17) & "  N O" & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(1)
                                    End If
                                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(1)
                                    End If
                                Else
                                    lstKan.AddItem Space(17) & "  N  " & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(1)
                                    End If
                                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(1)
                                    End If
                                    lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(2)
                                    End If
                                    If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '｢旧｣フォルダにファイルがない
                                lstKan.AddItem Space(17) & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                            End If
                        Else
                            lstKan.AddItem Space(17) & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                        End If
                    End If
                Else                                    '｢実行｣フォルダにファイルがない
                    If chkFolder(2).Value = CHECKBOX_ON Then   '｢旧｣フォルダ表示
                        If uVerData(2).sFileName <> "" Then
                            If sFileInfo(0) = sFileInfo(2) Then
                                lstKan.AddItem sFileName & "W   O" & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(0)
                                End If
                                If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(0)
                                End If
                                lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                            Else
                                lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(0)
                                End If
                                If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(0)
                                End If
                                lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(2)
                                End If
                                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(2)
                                End If
                                lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                            End If
                        Else                            '｢旧｣フォルダにファイルがない
                            lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                            lstKan.AddItem Space(17) & "  N O" & " -------- --------  -------- ----"
                        End If
                    Else                                '｢旧｣フォルダ非アクティブ表示
                        lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(0)
                        End If
                        lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                    End If
                End If
            Else                                        '｢実行｣フォルダ非アクティブ表示
                If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                    If uVerData(2).sFileName <> "" Then
                        If sFileInfo(0) = sFileInfo(2) Then
                            lstKan.AddItem sFileName & "W   O" & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                        Else
                            lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                            lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                    '｢旧｣フォルダにファイルがない
                        lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(0)
                        End If
                        lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
                    lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                        lstKan.AddItem Space(22) & sComment1(0)
                    End If
                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                        lstKan.AddItem Space(22) & sComment2(0)
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
                                lstKan.AddItem sFileName & "  N O" & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                            Else                            '｢実行｣フォルダと｢旧｣フォルダのバージョンが違う
                                lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(2)
                                End If
                                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(2)
                                End If
                                lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                            End If
                        Else                                '｢旧｣フォルダにファイルはない
                            lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                            lstKan.AddItem Space(17) & "W   O" & " -------- --------  -------- ----"
                        End If
                    Else                                    '｢旧｣フォルダ非アクティブ表示
                        lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(1)
                        End If
                        If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(1)
                        End If
                        lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                    End If
                Else                                        '｢実行｣フォルダにファイルがない
                    If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                        If uVerData(2).sFileName <> "" Then
                            lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(2)
                            End If
                            lstKan.AddItem Space(17) & "W N  " & " -------- --------  -------- ----"
                        Else                                '｢旧｣フォルダにファイルがない
                            lstKan.AddItem sFileName & "W N O" & " -------- --------  -------- ----"
                        End If
                    Else                                    '｢旧｣フォルダ非アクティブ表示
                        lstKan.AddItem sFileName & "W N  " & " -------- --------  -------- ----"
                    End If
                End If
            Else                                        '｢実行｣フォルダ非アクティブ表示
                If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                    If uVerData(2).sFileName <> "" Then
                        lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(2)
                        End If
                        If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(2)
                        End If
                        lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                    Else                                '｢旧｣フォルダにファイルがない
                        lstKan.AddItem sFileName & "W   O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
                    lstKan.AddItem sFileName & "W    " & " -------- --------  -------- ----"
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
                            lstKan.AddItem sFileName & "  N O" & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                        Else
                            lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                            lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                '｢旧｣フォルダにファイルはない
                        lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(1)
                        End If
                        If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(1)
                        End If
                        lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
                    lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                        lstKan.AddItem Space(22) & sComment1(1)
                    End If
                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                        lstKan.AddItem Space(22) & sComment2(1)
                    End If
                End If
            Else                                        '｢実行｣フォルダにファイルがない
                If chkFolder(2).Value = CHECKBOX_ON Then       '｢旧｣フォルダ表示
                    If uVerData(2).sFileName <> "" Then
                        lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(2)
                        End If
                        If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(2)
                        End If
                        lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                    Else                                '｢旧｣フォルダにファイルがない
                        lstKan.AddItem sFileName & "  N O" & " -------- --------  -------- ----"
                    End If
                Else                                    '｢旧｣フォルダ非アクティブ表示
                    lstKan.AddItem sFileName & "  N  " & " -------- --------  -------- ----"
                End If
            End If
        Else                                    '｢実行｣フォルダ非アクティブ表示
            If uVerData(2).sFileName <> "" Then '｢旧｣フォルダにファイルはある
                lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                    lstKan.AddItem Space(22) & sComment1(2)
                End If
                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                    lstKan.AddItem Space(22) & sComment2(2)
                End If
            Else                                '｢旧｣フォルダにファイルがない
                lstKan.AddItem sFileName & "    O" & " -------- --------  -------- ----"
            End If
        End If
    End If
End Sub

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sVersionInfo(sPath As String, iFolder As Integer)
    Dim i As Integer                    'カウンタ
    Dim j As Integer                    'カウンタ
    Dim sMyName As String               'ファイル名
    Dim iFileNumber As Integer          'ファイル番号
    Dim lLen As Long                    'ファイルサイズ
    Dim uFooter As MN_FOOT              'フッタ情報格納エリア
    Dim lPos As Long                    'バージョン情報格納位置
    Dim sDateTime As String
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD

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

            Close #iFileNumber                  'ファイルを閉じます
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    
End Sub

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
    
    On Error Resume Next
    
    '｢ワーク｣フォルダのファイルリストを検索する
    'ワークフォルダ内ファイル名を作成
    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
    'ファイルの検索をする
    'If Dir(sSrcFileName) <> "" Then     'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
      Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else
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
       bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    End If
'V1.8.0.1 ADD END

  If bRet = True Then
    '｢旧｣フォルダ内のファイルを全て削除する
     If sOldFolderRemove <> True Then
         fNewVersion = False
         Exit Function
     End If

    '｢旧｣フォルダ内のファイルを｢実行｣フォルダの内容に置換える
    If sCopyNOWtoOLD <> True Then
        fNewVersion = False
        Exit Function
    End If

    '｢実行｣フォルダ内のファイルを｢ワーク｣フォルダの内容に置換える
    If sCopyWRKtoNOW <> True Then
        fNewVersion = False
        Exit Function
    End If
    
 
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
    
    '改札機バージョン更新処理結果
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
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
                  "EG-R自動改札機 バージョン管理"
        Else
         MsgBox "改札機のバージョン作成で異常が発生しました。", _
                 vbOKOnly + vbExclamation, _
                 "NEG自動改札機 バージョン管理"
        End If
        
        fNewVersion = False
    End If
  
    fNewVersion = True
  Else
    fNewVersion = False
  End If
End Function
'V1.4.0.1 ADD START
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
    On Error Resume Next
    
    pfSeitouseiChck = True
    
    '********************************
    '*プロ判正当性チェック
    '********************************
    '自改プログラム判定データ正当性チェックを行う(対象ファイル：HAN_KUKA.KUK)
    bRet = fDataFileCheck(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = False Then
       If sNGSts <> "" And sNGKoumoku <> "" Then
          MsgBox "運賃データ正当性チェック異常(" & sNGSts & "：" & sNGKoumoku & "）", _
                 vbOKOnly + vbExclamation, _
                 sJverName & "自動改札機 バージョン管理"
       Else
          MsgBox "異常終了しました。", _
                 vbOKOnly + vbExclamation, _
                 TitleBox(FolderSyubetu) & "  ワーク→実行 コピー"
       End If
       pfSeitouseiChck = False
       Exit Function
    End If

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
    bRet = fKishuCheck(FolderName(0, FolderSyubetu) & "\")
    If bRet = False Then
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
    pfSeitouseiChck = False
End Function
'V1.4.0.1 ADD END
'V1.6.0.1 ADD START
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
            
            '自改チェック
            If gStrCurrentForm = sFormName_EJVer Then
               'EG-R自改時
               'If EGR_JIKAI_KISHU = Trim(sKisyu) Then  'V1.20.0.1 DEL
               'V1.20.0.1 ADD START
               '文字抽出
               sChkData = Left(sKisyu, Len(EGR_JIKAI_KISHU))
               If EGR_JIKAI_KISHU = sChkData Then
               'V1.20.0.1 ADD END
                   bRet = True  '機種正当性：正常
               Else
                   bRet = False '機種正当性：異常
                   fKishuCheck = bRet
                   Set objFso = Nothing    'V1.20.0.1 ADD
                   Exit Function
               End If
            Else
               'NEG自改時
               'If NEG_JIKAI_KISHU = Trim(sKisyu) Then    'V1.20.0.1 DEL
               'V1.20.0.1 ADD START
               '文字抽出
               sChkData = Left(sKisyu, Len(NEG_JIKAI_KISHU))
               If NEG_JIKAI_KISHU = sChkData Then
               'V1.20.0.1 ADD END
                   bRet = True  '機種正当性：正常
               Else
                   bRet = False '機種正当性：異常
                   fKishuCheck = bRet
                   Set objFso = Nothing    'V1.20.0.1 ADD
                   Exit Function
               End If
            End If

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSelectFile() As String
 
 If gStrCurrentForm = sFormName_EJVer Then
    'バージョンチェックファイル名を設定する。
    Select Case FolderSyubetu
       Case 0 '判定CPU-Pro
            fSelectFile = EHANTEI_CPU_CHK_FILE
       
       Case 1 'メインCPU-Pro
            fSelectFile = EMAIN_CPU_CHK_FILE
       
       Case 2 'サブCPU-Pro
            fSelectFile = ESUB_CPU_CHK_FILE
       
       Case 3 'メインCPU-OS
            fSelectFile = EMAIN_OS_CHK_FILE
     
     End Select
  Else
    'バージョンチェックファイル名を設定する。
    Select Case FolderSyubetu
       Case 0 '判定CPU-Pro
             fSelectFile = NHANTEI_CPU_CHK_FILE
      
       Case 1 'メインCPU-Pro
            fSelectFile = NMAIN_CPU_CHK_FILE
       
       Case 2 'サブCPU-Pro
            fSelectFile = NSUB_CPU_CHK_FILE
       
       Case 3 'メインCPU-OS
            fSelectFile = NMAIN_OS_CHK_FILE
     
    End Select
   End If

End Function
'V.1.6.0.1 ADD END

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
    
    On Error Resume Next
 
   '旧フォルダ内のファイルリストを検索する。
    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '「旧」フォルダ内ファイル名を作成する
    'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする  'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else                                'ファイルが存在しない
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
    bRet = fReadFileList(FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST)
  
    '｢実行｣フォルダ内のファイルを全て削除する
    If sNowFolderRemove <> True Then
        fOldVersion = False
        Exit Function
    End If
    
    '｢旧｣フォルダ内のファイルを｢実行｣フォルダの内容に置換える
    If sCopyOLDtoNOW <> True Then
        fOldVersion = False
        Exit Function
    End If
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
     
     '改札機バージョン更新処理異常
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
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
                  "EG-R自動改札機 バージョン管理"
        Else
           MsgBox "改札機のバージョン作成で異常が発生しました。", _
                   vbOKOnly + vbExclamation, _
                   "NEG自動改札機 バージョン管理"
        End If
        fOldVersion = False
    End If

    fOldVersion = True
End Function

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSetFolderName()

        TitleBox(0) = "判定CPU-Pro"
        TitleBox(1) = "メインCPU-Pro"
        TitleBox(2) = "サブCPU-Pro"
        TitleBox(3) = "メインCPU-OS"
        TitleBox(4) = "予備1"
        TitleBox(5) = "予備2"
    
        LogBox(0) = "判定"
        LogBox(1) = "メイン"
        LogBox(2) = "サブ"
        LogBox(3) = "OS"
        LogBox(4) = "予備1"
        LogBox(5) = "予備2"
        
  If gStrCurrentForm = sFormName_EJVer Then
        'フォルダ名に設定を行う
        FolderName(0, 0) = E_EHAN1WRK
        FolderName(1, 0) = E_EHAN1NOW
        FolderName(2, 0) = E_EHAN1OLD
        FolderName(0, 1) = E_EPRO1WRK
        FolderName(1, 1) = E_EPRO1NOW
        FolderName(2, 1) = E_EPRO1OLD
        FolderName(0, 2) = E_ESCPUWRK
        FolderName(1, 2) = E_ESCPUNOW
        FolderName(2, 2) = E_ESCPUOLD
        FolderName(0, 3) = E_EOSWRK
        FolderName(1, 3) = E_EOSNOW
        FolderName(2, 3) = E_EOSOLD
        FolderName(0, 4) = E_EYOBI1WRK
        FolderName(1, 4) = E_EYOBI1NOW
        FolderName(2, 4) = E_EYOBI1OLD
        FolderName(0, 5) = E_EYOBI2WRK
        FolderName(1, 5) = E_EYOBI2NOW
        FolderName(2, 5) = E_EYOBI2OLD
    Else
        'フォルダ名に設定を行う
        FolderName(0, 0) = N_NHAN1WRK
        FolderName(1, 0) = N_NHAN1NOW
        FolderName(2, 0) = N_NHAN1OLD
        FolderName(0, 1) = N_NPRO1WRK
        FolderName(1, 1) = N_NPRO1NOW
        FolderName(2, 1) = N_NPRO1OLD
        FolderName(0, 2) = N_NSCPUWRK
        FolderName(1, 2) = N_NSCPUNOW
        FolderName(2, 2) = N_NSCPUOLD
        FolderName(0, 3) = N_NOSWRK
        FolderName(1, 3) = N_NOSNOW
        FolderName(2, 3) = N_NOSOLD
        FolderName(0, 4) = N_NYOBI1WRK
        FolderName(1, 4) = N_NYOBI1NOW
        FolderName(2, 4) = N_NYOBI1OLD
        FolderName(0, 5) = N_NYOBI2WRK
        FolderName(1, 5) = N_NYOBI2NOW
        FolderName(2, 5) = N_NYOBI2OLD
    End If

'V1.20.0.1 ADD START
'-------EG-R自改-------
    ' キー名:判定CPU-PRO代表
    EHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    
    ' キー名:メインCPU-PRO代表
    EMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_PRO, PATH_GATEVER_FILE)
        
    ' キー名：サブCPU-PRO代表
    ESUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_SUB_PRO, PATH_GATEVER_FILE)
    
    ' キー名:メインCPU-OS代表
    EMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_OS, PATH_GATEVER_FILE)
    
'-------NEG自改-------
    ' キー名:判定CPU-PRO代表
    NHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    
    ' キー名:メインCPU-PRO代表
    NMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_PRO, PATH_GATEVER_FILE)
    
    ' キー名：サブCPU-PRO代表
    NSUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_SUB_PRO, PATH_GATEVER_FILE)
    
    ' キー名:メインCPU-OS代表
    NMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_OS, PATH_GATEVER_FILE)
'V1.20.0.1 ADD END

End Sub

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
'//  関数名称  : sWriteListbox
'//  機能名称  : バージョンテキストファイル書込み。
'//  機能概要  : リストボックスの内容を、バージョンテキストファイルに書き込む。
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
Private Sub sWriteListbox()
    Dim iFileNumber As Integer
    Dim i As Integer

    On Error Resume Next
    
    iFileNumber = FreeFile              '未使用のファイル番号を取得する

    Open MN_VERSI_FILE For Output Access Write As #iFileNumber
                                        'ファイル名を作成します。

    For i = 0 To lstKan.ListCount - 1

        Print #iFileNumber, lstKan.List(i) & Chr(vbKeyReturn)
                                        'データを書き込む
    Next

    Close #iFileNumber                  'ファイルを閉じます。
End Sub

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
    
    On Error GoTo ErrorHandler      'エラーハンドルの登録

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
    If objFso.FileExists(FolderName(0, FolderSyubetu) & "\" & sChkName) = True Then
        '指定ファイルが存在する
        sChkName = objFso.GetFileName(FolderName(0, FolderSyubetu) & "\" & sChkName)
        Kill FolderName(0, FolderSyubetu) & "\" & sChkName
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
                sDstFileName = FolderName(0, FolderSyubetu) & "\" & MyName
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
    
    '「ワークコピー正常終了」ポップアップ画面表示
    MsgBox "インストール媒体の全てのファイルを、" _
            & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "の「ワーク」フォルダに" _
            & Chr(vbKeyReturn) & "コピーしました。", _
            vbOKOnly + vbExclamation, _
            TitleBox(FolderSyubetu) & "  媒体→ワーク コピー"
    
    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_OK, 0)
    
    'バージョン情報リストボックスを作成する
    fMakeListbox
    Exit Sub    '処理を終了する

ErrorHandler:   ' エラー処理。
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
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
    
    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)
End Sub

'V1.6.0.1 ADD START
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
            'サム値異常
            If lngSumRet = SUM_CHK.SumErr Then
               MsgBox "サム値が異常です。" _
                      & Chr(vbKeyReturn) & "データを確認してください。", _
                      vbOKOnly + vbExclamation, _
                      sJverName & "自動改札機 バージョン管理"
            'サム値異常以外異常
            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
                   '「ワークコピー異常終了」ポップアップ画面表示
               MsgBox "インストール媒体からのコピーエラーが発生しました。" _
                     & Chr(vbKeyReturn) & "エラーコード＝" _
                     & str$(Err.Number), _
                     vbOKOnly + vbExclamation, _
                    "→ワーク コピー"
            End If
            pfInstallSeitouseiChck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    'ファイル数最大チェック
    If UBound(FileList) > FILECNT_MAX Then
       MsgBox "ファイル数が上限を超えています。" _
              & Chr(vbKeyReturn) & "データを確認してください。", _
              vbOKOnly + vbExclamation, _
              sJverName & "自動改札機 バージョン管理"
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
          '13バイト以上の場合
          MsgBox "ファイル名が異常です。" _
                 & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
                  vbOKOnly + vbExclamation, _
                  sJverName & "自動改札機 バージョン管理"
           bRet = False
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
           Exit For
       End If
    Next
'V2.6.0.1 ADD END

    pfInstallSeitouseiChck = bRet
Exit Function

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pfInstallSeitouseiChck = False
End Function
'V1.6.0.1 ADD END

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sCopyWRKtoNOW() As Boolean
    
    Dim i As Integer                'カウンタ
    Dim sSrcFileName As String      'コピー元ファイル名
    Dim sDstFileName As String      'コピー先ファイル名
    Dim bError As Boolean           'フラグ
    Dim bRet As Boolean             '戻り値
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler      'エラーハンドルの登録
  
    '戻り値初期化
    sCopyWRKtoNOW = True
    
    '****************************
    '* ファイルリストをコピーする *
    '****************************
    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
                                    'ワークフォルダ内ファイル名を作成する
    sDstFileName = FolderName(1, FolderSyubetu) & "\" & MN_FILELIST
                                    '実行フォルダ内ファイル名を作成する
    'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする   'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then     'ファイルの検索をする   'V1.20.0.1 ADD
        'ファイルリストを「ワーク」フォルダから「実行」フォルダにコピーする
        FileCopy sSrcFileName, sDstFileName
    Else                                'ファイルが存在しない
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
        sSrcFileName = FolderName(0, FolderSyubetu) & "\" & FileList(i)
                                    'ワークフォルダ内ファイル名を作成する
        sDstFileName = FolderName(1, FolderSyubetu) & "\" & FileList(i)
                                    '実行フォルダ内ファイル名を作成する

        'ワークフォルダ内のファイルを実行フォルダにコピーする
        'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする   'V1.20.0.1 DEL
        If objFso.FileExists(sSrcFileName) = True Then   'ファイルの検索をする   'V1.20.0.1 ADD
            'ファイルを「ワーク」フォルダから「実行」フォルダにコピーする
            FileCopy sSrcFileName, sDstFileName
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    Exit Function                           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
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
    
    On Error GoTo ErrorHandler              'エラーハンドル設定
  
    '戻り値初期化
    sCopyNOWtoOLD = True
   
    '実行フォルダ内のディレクトリの名前を表示します。
    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
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
                sDstFileName = FolderName(2, FolderSyubetu) & "\" & MyName

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sCopyOLDtoNOW() As Boolean
    Dim i As Integer                'カウンタ
    Dim sSrcFileName As String      'コピー元ファイル名
    Dim sDstFileName As String      'コピー先ファイル名
    Dim bError As Boolean           'エラーフラグ
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler
    
    '初期値設定
    sCopyOLDtoNOW = True

    '****************************
    '* ファイルリストをコピーする *
    '****************************
    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '「旧」フォルダ内ファイル名を作成する
    sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '「実行」フォルダ内ファイル名を作成する
    'If Dir(sSrcFileName) <> "" Then     'ファイルの検索をする  'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then 'ファイルの検索をする   'V1.20.0.1 ADD
        'ファイルリストを「旧」フォルダから「実行」フォルダにコピーする
        FileCopy sSrcFileName, sDstFileName
    Else
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
        sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)

        '実行フォルダ内ファイル名を作成する
        sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)

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
    Exit Function       '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
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
    
    On Error GoTo ErrorHandler          'エラーハンドルの登録
   
   '戻り値初期化
    sOldFolderRemove = True
 
    '「実行」フォルダ内のディレクトリの名前を表示します。
    gstrMyPath = FolderName(2, FolderSyubetu) & "\"
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

    On Error GoTo ErrorHandler          'エラーハンドルの登録

    '初期値設定
    sNowFolderRemove = True
    
    '「実行」フォルダ内のディレクトリの名前を表示します。
    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim MyName As String                'ファイル名
    Dim lngErrCode As Long              'エラーコード
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    'V1.20.0.1 ADD END
    
    On Error GoTo ErrorHandler          'エラーハンドルの登録

    '初期値設定
    sWrkFolderRemove = True
   
    'ワークフォルダ内のディレクトリの名前を表示します。
    gstrMyPath = FolderName(0, FolderSyubetu) & "\"
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeOutPutFile()
'   Dim sOutFileName As String '媒体出力ファイル名[種別別]
'   Dim iFileNumber As Integer 'ファイル番号
'   Dim i As Integer           'カウンター
'   Dim bFlag As Boolean       'フラグ
'   Dim iResponse As Integer   'MsgBox戻り値
'   Dim lngErrCode As Long     'エラーコード
'   Dim fso         As New FileSystemObject   'ファイルシステムオブジェクト
'   Dim strWriteDir As String               '出力先フォルダ
'
'   On Error Resume Next 'V1.21.0.1 ADD
'
'  'フォルダ選択部に指定有無チェック
'  bFlag = False                                 'フラグを「偽」にする
'  For i = 0 To 2                                'フォルダ数分繰り返す
'     If chkFolder(i).Value = CHECKBOX_ON Then   '「？？」フォルダが指定されている
'        bFlag = True                            'フラグを「真」にする
'        Exit For                                'ループを抜ける
'     End If
'  Next
'
'  If bFlag = False Then                       'フォルダ指定無し
'     If gStrCurrentForm = sFormName_EJVer Then
'       '「表示フォルダ指定なし」ポップアップ表示
'         MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
'                 vbOKOnly + vbExclamation, _
'                 "EG-R自動改札機 バージョン管理"
'     Else
'       '「表示フォルダ指定なし」ポップアップ表示
'         MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
'                vbOKOnly + vbExclamation, _
'                "NEG自動改札機 バージョン管理"
'     End If
'         '処理を抜ける
'     Exit Function
'   End If
'
'  'フォルダ選択ポップアップ画面表示
''  strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")                         'V1.12.0.1 DEL
'  strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.12.0.1 ADD
'
'  '指定フォルダなし
'  If Len(strWriteDir) = 0 Then
'       Exit Function
'  End If
'
'  'コピー先フォルダの有無確認
'  If fso.FolderExists(strWriteDir) = False Then
'     'コピー先フォルダ作成
'     fso.CreateFolder (strWriteDir)
'  End If
'
'   '処理中フォームにより、媒体出力するファイル名作成
'   If gStrCurrentForm = sFormName_EJVer Then
'       'リソース選択部分岐
'       Select Case FolderSyubetu
'        Case 0      '判定CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_EJHANTEIPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJHANTEIPRO      'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJHANTEIPRO 'V1.8.0.1 ADD
'        Case 1      'メインCPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINPRO   'V1.8.0.1 ADD
'        Case 2      'サブCPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_EJSUBPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJSUBPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJSUBPRO   'V1.8.0.1 ADD
'        Case 3      'メインCPU-OS
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINOS        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINOS   'V1.8.0.1 ADD
'        Case 4      '予備1
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI1         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
'        Case 5      '予備2
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI2
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI2         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI2    'V1.8.0.1 ADD
'        End Select
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
'
'  iFileNumber = FreeFile              '未使用のファイル番号を取得する
'
'  '対象ファイルをオープンする。
'  Open sOutFileName For Output Access Write As #iFileNumber
'
'  For i = 0 To lstKan.ListCount - 1
'  'リストボックスに表示されている分だけ、書き込む。
'       Print #iFileNumber, lstKan.List(i) & Chr(vbKeyReturn)
'  Next
'
'  '対象ファイルをクローズする。
'  Close #iFileNumber
'
'  'ファイルの有無確認
'  If fso.FileExists(sOutFileName) = False Then
'     'ファイル無し異常ポップアップ画面表示
'     MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
'     Exit Function
'  End If
'
'  On Error GoTo COPY_ERROR
'  'ファイルコピー
'  fso.CopyFile sOutFileName, strWriteDir
'  '「媒体出力正常終了」ポップアップ画面表示
'  'V1.8.0.1 DEL START
'  'iResponse = MsgBox("正常終了しました。", vbOKOnly, _
'  '                   "出力結果")
'  'V1.8.0.1 DEL END
'  MsgBox "正常終了しました。", vbInformation, "出力結果"   'V1.8.0.1 ADD
'
'  '「自改ﾊﾞｰｼﾞｮﾝ：媒体出力処理正常」ログ出力
'  Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_OK, 0)
'
'  Set fso = Nothing
'
'  Exit Function
'
''*******************************
''VBエラー処理
'COPY_ERROR:
'        '処理異常の場合、出力結果ポップアップ(異常)表示
'        MsgBox "異常終了しました。", vbCritical, "出力結果"
'        '「自改ﾊﾞｰｼﾞｮﾝ：媒体出力処理異常」ログ出力
'        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_ERROR, lngErrCode)
'        Set fso = Nothing
''*******************************
End Function

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
   On Error Resume Next
    
    '汎用メール受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmJVer.Caption, False
    End If
End Sub
'V1.4.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetGoukiNo
'//  機能名称  : 論理号機番号を取得する。
'//  機能概要  : GATE.INIより論理号機番号を取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Integer            [OUT]先頭号機番号
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-18   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Function pfGetGoukiNo() As Integer                        'V1.6.0.1 DEL
Private Function pfGetGoukiNo(iGoukiCunter As Integer) As Integer  'V1.6.0.1 ADD

    Dim lngRet As Long          '関数の返り値
    Dim iGate As Integer        '自改INDEX
    Dim j As Integer            'ワークINDEX
    Dim sGoukiNo As String      'GLTファイルレコードデータ(号機番号表示文字)
    Dim cWork As Byte           'ワークエリア
    Dim lngErrCode As Long      'エラーコード
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim intGLTFileNo As Integer     'ﾌｧｲﾙ番号
'   Dim iGoukiCunter As Integer　　 'V1.6.0.1 DEL
    

    On Error Resume Next

 '   For iGoukiCunter = 1 To MAX_GATE_NO   'V1.6.0.1 DEL
         '自動改札機情報取得
         sKeyName = "gate" & Format(iGoukiCunter, "00")
         iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                        sKeyName, _
                                        DEFAILT, sGateData, Len(sGateData), _
                                        PATH_GATE_FILE)
         If iRet = 0 Then
            '「EG-R自動改札機バージョン管理画面：自動改札機INIファイル読込異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
            pfGetGoukiNo = 0
            Exit Function
         End If
             
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
        End If
'        If Trim(sFData(4)) = EGR Then                          'V1.0.6.1 DEL
        If Trim(sFData(4)) = EGR Or Trim(sFData(4)) = NEG Then  'V1.0.6.1 ADD
           pfGetGoukiNo = iGoukiCunter
           Exit Function
        End If
 'Next 'V1.6.0.1 DEL
End Function
'V1.4.0.1 ADD END

'V1.20.0.1 ADD START
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
'V1.20.0.1 ADD END

