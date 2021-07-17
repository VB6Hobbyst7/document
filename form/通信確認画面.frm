VERSION 5.00
Begin VB.Form frmPing 
   BorderStyle     =   0  'なし
   Caption         =   "通信確認"
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
      Left            =   10800
      Top             =   6360
   End
   Begin VB.Frame frmKekka 
      Caption         =   "ＰＩＮＧ結果"
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   8535
      Begin VB.CommandButton cmdZikko 
         Caption         =   "ＰＩＮＧ実行"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1905
      End
      Begin VB.ListBox LstStatus 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   6015
      End
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
   Begin VB.Frame frmkiki 
      Caption         =   "機器選択"
      Enabled         =   0   'False
      Height          =   4455
      Left            =   6120
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
      Begin VB.ListBox LstKiki 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         ItemData        =   "通信確認画面.frx":0000
         Left            =   240
         List            =   "通信確認画面.frx":0002
         TabIndex        =   37
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   32
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblIP20 
         Caption         =   "．"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   36
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP21 
         Caption         =   "．"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   35
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP22 
         Caption         =   "．"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   34
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame frmTe 
      Caption         =   "ＩＰアドレス手入力"
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
      Begin VB.CommandButton cmdC 
         Caption         =   "クリア"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   360
         TabIndex        =   26
         Top             =   960
         Width           =   1875
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   9
         Left            =   4080
         TabIndex        =   25
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   8
         Left            =   3240
         TabIndex        =   24
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   7
         Left            =   2400
         TabIndex        =   23
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   4
         Left            =   2400
         TabIndex        =   22
         Top             =   1800
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   5
         Left            =   3240
         TabIndex        =   21
         Top             =   1800
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   6
         Left            =   4080
         TabIndex        =   20
         Top             =   1800
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   1
         Left            =   2400
         TabIndex        =   19
         Top             =   2640
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   2
         Left            =   3240
         TabIndex        =   18
         Top             =   2640
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   3
         Left            =   4080
         TabIndex        =   17
         Top             =   2640
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   0
         Left            =   2400
         TabIndex        =   16
         Top             =   3480
         Width           =   800
      End
      Begin VB.CommandButton cmdOct 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3240
         TabIndex        =   15
         Top             =   3480
         Width           =   800
      End
      Begin VB.CommandButton cmdBs 
         Caption         =   "BS"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   4080
         TabIndex        =   14
         Top             =   3480
         Width           =   800
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   3
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   2
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   360
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblIP10 
         Caption         =   "．"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP11 
         Caption         =   "．"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP12 
         Caption         =   "．"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame frmSentaku 
      Caption         =   "ＩＰアドレス入力方法選択"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   11535
      Begin VB.OptionButton OptTe 
         Caption         =   "手入力"
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptKiki 
         Caption         =   "機器選択"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   $"通信確認画面.frx":0004
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
      TabIndex        =   3
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "通信確認"
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
      TabIndex        =   38
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmPing.frm
'//  パッケージ名：通信確認画面
'//
'//  概要：通信確認画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//               EG10より、通信確認(frmPing.frm)画面流用
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//     REVISIONS :(EG30 31.2.0.1) 2015-07-17   REVISED BY [TCC] T.Nakajima
'//                 pingのステータスを表示できるよう修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private bIP0 As Boolean
Private bIP1 As Boolean
Private bIP2 As Boolean

Private Const MAXIPKIKIINFO = 96            '機器構成情報最大

'Private sKikiIP(45) As String              'IPアドレス格納エリア   ' EG20 V3.4.0.1削除
Private sKikiIP(MAXIPKIKIINFO) As String    'IPアドレス格納エリア   ' EG20 V3.4.0.1追加
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値 '1.3.0.1 ADD

' EG20 V3.4.0.1【接続機器見直し対応】追加開始
' 上位機器設定構成
Private Type TRANSKIKI_INFO
    bStatus As Boolean              ' 設定有無（TRUE:有り,FALSE:無し）
    sGetInf As String               ' 画面表示用名称
    iAreaID As Integer              ' 対象外部機器上位機器通信状態エリアID
    nIniListNo As Integer           ' 外部機器リスト番号
    nCorner As Integer              ' コーナ番号
End Type
Private gTransKikiInfo(1 To CONECT_KIKI_INI_MAX) As TRANSKIKI_INFO

' EG20 V3.4.0.1【接続機器見直し対応】追加終了

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 通信確認画面(アクティブ時)
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
'//  機能名称  : 通信確認画面(ディアクティブ時)
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
'//  機能名称  : 通信確認画面(ロード時)
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
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    Dim sKeyName As String
    Dim sGateData As String * 128    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim sIP As String * 15
    Dim sCPUReg As String           'LDU_APLROOTレジストリ取得用
    Dim sCPUData As String * 128    '１行分ファイル内容取得用
    
    '配置設定
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '初期化
    LstStatus.Clear
    LstKiki.Clear
    
    txtIP1(0).Text = ""
    txtIP1(1).Text = ""
    txtIP1(2).Text = ""
    txtIP1(3).Text = ""
    txtIP2(0).Text = ""
    txtIP2(1).Text = ""
    txtIP2(2).Text = ""
    txtIP2(3).Text = ""

    OptTe.Value = True
    frmTe.Enabled = True
    txtIP1(0).Enabled = True
    txtIP1(1).Enabled = True
    txtIP1(2).Enabled = True
    txtIP1(3).Enabled = True
    lblIP10.Enabled = True
    lblIP11.Enabled = True
    lblIP12.Enabled = True
    cmdNum(0).Enabled = True
    cmdNum(1).Enabled = True
    cmdNum(2).Enabled = True
    cmdNum(3).Enabled = True
    cmdNum(4).Enabled = True
    cmdNum(5).Enabled = True
    cmdNum(6).Enabled = True
    cmdNum(7).Enabled = True
    cmdNum(8).Enabled = True
    cmdNum(9).Enabled = True
    cmdOct.Enabled = True
    cmdBs.Enabled = True
    cmdC.Enabled = True
    
    frmkiki.Enabled = False
    txtIP2(0).Enabled = False
    txtIP2(1).Enabled = False
    txtIP2(2).Enabled = False
    txtIP2(3).Enabled = False
    lblIP20.Enabled = False
    lblIP21.Enabled = False
    lblIP22.Enabled = False
    LstKiki.Enabled = False
    
    bIP0 = False
    bIP1 = False
    bIP2 = False
    
' EG20 V3.4.0.1追加開始
    '号機情報取得
    Call gsGetGateInfo
    ' コーナ名称設定処理
    Call gsGetCornerName
' EG20 V3.4.0.1追加終了
    
    'V1.8.0.1 ADD START
'    For i = 0 To 45                    ' EG20 V3.4.0.1削除
    For i = 0 To MAXIPKIKIINFO          ' EG20 V3.4.0.1追加
     sKikiIP(i) = ""
    Next
    'V1.8.0.1 ADD END
    
    'V1.3.0.1 ADD START
    'メイル受信用のメイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
    
    On Error GoTo FileError
   
    '外部機器情報(OUTKIKI_LIST.ini)取得表示
    OverKikiPing
    '自動改札機(Gate.ini)情報より自改情報取得表示
    GatePing
    '自動改札機(Gate.ini)情報より判定ICM情報取得得表示
    ICMPing
    
    ' 操作卓情報表示
    OperatePing
    
    '「通信確認画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_START, 0)
    
FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : OverKikiPing
'//  機能名称  : 通信確認画面(ロード時)
'//  機能概要  : 外部機器情報を取得表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub OverKikiPing()
  Dim sIniFilPath As String              '対象外部機器専用INIファイル
  Dim iCnt As Integer                    'カウンター
  Dim sKey As String                     'キー名
  Dim sGetInf As String * PING_SIZE      '取得情報(表示名称)
  Dim sFilePath As String * PING_SIZE    '取得情報(対象外部機器INIパス)
  Dim sSectionName As String * PING_SIZE '取得情報(セクション名)
  Dim sKeyName As String * PING_SIZE     '取得情報(キー名)
  Dim lSts As Long                       'INI取得処理戻り値
'  Dim sIP As String * PING_IP_SIZE       '取得IPアドレス           ' EG20 V6.1.0.1削除
  Dim sIP As String                      ' 取得IPアドレス           ' EG20 V6.1.0.1追加
  Dim iType As Integer                   '取得情報(機器タイプ)
  Dim sTargetPath As String              '対象外部機器INIファイルパス
  Dim iAreaID As Integer                 '取得情報(エリアID)        ' EG20 V3.4.0.1追加
    
  Dim sGetString As String * 128         ' INI取得文字列
  Dim nNullIndex As Integer              ' 文字数ワーク
    
  On Error Resume Next
 
  '初期化
  sIP = ""
  
'  For iCnt = 1 To 10                               ' EG20 V3.4.0.1削除

' EG20 V3.4.0.1【接続機器見直し対応】追加開始
  For iCnt = 1 To CONECT_KIKI_INI_MAX
    gTransKikiInfo(iCnt).bStatus = False               ' 設定有無（TRUE:有り,FALSE:無し）
    gTransKikiInfo(iCnt).sGetInf = ""                  ' 画面表示用名称
    gTransKikiInfo(iCnt).iAreaID = 0                   ' 対象外部機器上位機器通信状態エリアID
    gTransKikiInfo(iCnt).nIniListNo = 0                ' 外部機器リスト番号
    gTransKikiInfo(iCnt).nCorner = 0                   ' コーナ番号

    ' OUTKIKI_LIST.iniから上位通信エリアIDを取得する。
    sKey = ""
    sKey = PROFILE_KEY_KIKIAREA_NAME & Format(iCnt, "00")
    iAreaID = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                sKey, _
                                DEFAILT_Int, _
                                OUTKIKI_LIST_FILE)

' EG20 V3.4.0.1【接続機器見直し対応】追加終了

    ' OUTKIKI_LIST.iniから表示用外部機器名称を取得する。
    sGetInf = ""
    sKey = ""
    sKey = PROFILE_KEY_KIKINAME & Format(iCnt, "00")
    lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                  sKey, _
                                  DEFAILT, _
                                  sGetInf, _
                                  Len(sGetInf), _
                                  OUTKIKI_LIST_FILE)
     If lSts = False Then
        '何もしない
     Else
'       LstKiki.AddItem sGetInf                             ' EG20 V3.4.0.1削除
        Call psAddKikiCornerName(sGetInf, iAreaID, iCnt)    ' EG20 V3.4.0.1追加
     End If

    If gTransKikiInfo(iCnt).bStatus = True Then             ' EG20 V3.4.0.1追加

        sKey = ""
        sFilePath = ""
        ' OUTKIKI_LIST.iniから表示対象外部機器INIファイルパスを取得する。
        sKey = PROFILE_KEY_KIKIPATH & Format(iCnt, "00")
        lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                       sKey, _
                                       DEFAILT, _
                                       sFilePath, _
                                       Len(sFilePath), _
                                       OUTKIKI_LIST_FILE)
                                   
        sKey = ""
        ' OUTKIKI_LIST.iniから機器タイプ(監視盤orIDUorLDU)を取得する。
        sKey = PROFILE_KEY_TYPE & Format(iCnt, "00")
        iType = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                     sKey, _
                                     DEFAILT_Int, _
                                     OUTKIKI_LIST_FILE)
        
        sTargetPath = ""
        If iType = 1 Then  '機器タイプが監視盤の場合
           sTargetPath = PATH_KANSI & sFilePath
        End If
        If iType = 2 Then  '機器タイプがIDUの場合
           sTargetPath = PATH_IDU_APP & "\\" & sFilePath
        End If
        If iType = 3 Then  '機器タイプがLDUの場合
           sTargetPath = PATH_LDU_APP & "\\" & sFilePath
        End If
        sKey = ""

        ' OUTKIKI_LIST.iniから対象外部機器INIファイルのセクション名を取得する。
        sKey = PROFILE_KEY_SECTION_NAME & Format(iCnt, "00")
        lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                       sKey, _
                                       DEFAILT, _
                                       sSectionName, _
                                       Len(sSectionName), _
                                       OUTKIKI_LIST_FILE)

         sKey = ""

        ' OUTKIKI_LIST.iniから対象外部機器INIファイルのキー名を取得する。
        sKey = PROFILE_KEY_KEY_NAME & Format(iCnt, "00")
        lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                       sKey, _
                                       DEFAILT, _
                                       sKeyName, _
                                       Len(sKeyName), _
                                       OUTKIKI_LIST_FILE)
        sKey = ""
        sIP = "" 'V1.8.0.1 ADD
        ' 対象外部機器INIファイルからIPアドレスを取得する。
        lSts = GetPrivateProfileString(sSectionName, _
                                       sKeyName, _
                                       DEFAILT, _
                                       sGetString, _
                                       Len(sGetString), _
                                       sTargetPath)
        If lSts > 0 Then                             ' V1.3.0.1 ADD
            LstKiki.AddItem gTransKikiInfo(iCnt).sGetInf    ' EG20 V3.4.0.1追加
            
            nNullIndex = InStr(sGetString, Chr(0))
            If nNullIndex <> 0 Then
                sIP = Left(sGetString, nNullIndex - 1)
            Else
                sIP = sGetString
            End If
            sKikiIP(LstKiki.ListCount - 1) = Trim(sIP)
        End If                                       ' V1.3.0.1 ADD
    End If                                                  ' EG20 V3.4.0.1追加
  Next iCnt
  
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : GatePing
'//  機能名称  : 通信確認画面(ロード時)
'//  機能概要  : 自動改札機情報を取得表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-12  CODED BY  [TCC] H.Sugimoto
'//                 【表示号機改善対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub GatePing()
    Dim sKeyName As String
    Dim sGateData As String * PING_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim sIP As String * PING_IP_SIZE
    Dim nCorner As Integer                      ' コーナ番号    ' EG20 V6.1.0.1追加

    On Error Resume Next

   '自動改札機情報取得
    For i = 1 To MAX_GATE_NO
        sKeyName = "gate" & Format(i, "00")
        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                       sKeyName, _
                                       DEFAILT, sGateData, Len(sGateData), _
                                       PATH_GATE_FILE)
        If iRet = 0 Then
            '「通信確認画面：自動改札機INIファイル読込異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
            Exit Sub
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
            
' EG20 V3.4.0.1【接続機器見直し対応】削除開始
'            '機種タイプによって表示を行う。
'            'E：EG-R自動改札機＊：表示しない。
'            If Trim(sFData(4)) = EGR Then
'                LstKiki.AddItem "EG-R自動改札機" & "#" & i
'                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(5))
'            End If
'            If Trim(sFData(4)) = MISETI Then
'               '処理を行わない。
'            End If
' EG20 V3.4.0.1【接続機器見直し対応】削除終了
' EG20 V3.4.0.1【接続機器見直し対応】追加開始
            ' EG20改札機であれば表示
            If Trim(sFData(GATE_IDX.IDX_KISHU)) = EG20 Then
' EG20 V6.1.0.1【表示号機改善対応】削除開始
'                LstKiki.AddItem "自動改札機" & "#" & i
' EG20 V6.1.0.1【表示号機改善対応】削除終了
' EG20 V6.1.0.1【表示号機改善対応】追加開始
                nCorner = CInt(Trim(sFData(GATE_IDX.IDX_RONRI_CORNER)))
                LstKiki.AddItem "自動改札機" & "#" & Trim(sFData(GATE_IDX.IDX_DISP_GOKI)) & _
                                    "(" & Format(nCorner, "00") & ")"
' EG20 V6.1.0.1【表示号機改善対応】追加終了
                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(GATE_IDX.IDX_ADDRESS))
            End If
' EG20 V3.4.0.1【接続機器見直し対応】追加終了
       End If
    Next
    
' EG20 V3.4.0.1【接続機器見直し対応】削除開始
'      '自動改札機情報取得(NEG)
'    For i = 1 To MAX_GATE_NO
'        sKeyName = "gate" & Format(i, "00")
'        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
'                                       sKeyName, _
'                                       DEFAILT, sGateData, Len(sGateData), _
'                                       PATH_GATE_FILE)
'        If iRet = 0 Then
'            '「通信確認画面：自動改札機INIファイル読込異常」ログ出力
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
'            Exit Sub
'        End If
'
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
'
'            '機種タイプによって表示を行う。
'            'N：NEG自動改札機。＊：表示しない。
'            If Trim(sFData(4)) = NEG Then
'                LstKiki.AddItem "NEG自動改札機" & "#" & i
'                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(5))
'            End If
'            If Trim(sFData(4)) = MISETI Then
'               '処理を行わない。
'            End If
'       End If
'    Next
' EG20 V3.4.0.1【接続機器見直し対応】削除終了

FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : GatePing
'//  機能名称  : 通信確認画面(ロード時)
'//  機能概要  : 自動改札機情報を取得表示する。
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
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-12  CODED BY  [TCC] H.Sugimoto
'//                 【表示号機改善対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub ICMPing()
    Dim sKeyName As String
    Dim sGateData As String * PING_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim sIP As String * PING_IP_SIZE
    Dim szIniFilePath As String     ' INIファイルパス   ' EG20 V3.4.0.1【接続機器見直し対応】追加
    Dim nCorner As Integer                      ' コーナ番号    ' EG20 V6.1.0.1追加

    On Error Resume Next

   '自動改札機情報取得
    For i = 1 To MAX_GATE_NO
' EG20 V3.4.0.1【接続機器見直し対応】削除開始
'        sKeyName = "gate" & Format(i, "00")
'        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
'                                       sKeyName, _
'                                       DEFAILT, sGateData, Len(sGateData), _
'                                       PATH_GATE_FILE)
' EG20 V3.4.0.1【接続機器見直し対応】削除終了
' EG20 V3.4.0.1【接続機器見直し対応】追加開始
        ' IDUのICM.INIから改札機情報を取得
        szIniFilePath = PATH_IDU_APP & IDU_ICM_FILE
        sKeyName = "icm" & Format(i, "00")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ICM, _
                                    sKeyName, _
                                    DEFAILT, sGateData, Len(sGateData), _
                                    szIniFilePath)
' EG20 V3.4.0.1【接続機器見直し対応】追加終了
        If iRet = 0 Then
            '「通信確認画面：自動改札機INIファイル読込異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
            Exit Sub
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
                       
' EG20 V3.4.0.1【接続機器見直し対応】削除開始
'            If Trim(sFData(4)) <> MISETI Then   'V1.8.0.1 ADD
'            '判定IC-Mのアドレスチェックを行う。
'             LstKiki.AddItem "判定IC-Mモジュール" & "#" & i
'             sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(14))
'             End If  'V1.8.0.1 ADD
' EG20 V3.4.0.1【接続機器見直し対応】削除終了
' EG20 V3.4.0.1【接続機器見直し対応】追加開始
            If Trim(sFData(5)) <> MISETI Then   'V1.8.0.1 ADD
                '判定IC-Mのアドレスチェックを行う。
' EG20 V6.1.0.1【表示号機改善対応】削除開始
'                LstKiki.AddItem "ＩＣＭ" & "#" & i
' EG20 V6.1.0.1【表示号機改善対応】削除終了
' EG20 V6.1.0.1【表示号機改善対応】追加開始
                nCorner = CInt(Trim(sFData(3)))
                LstKiki.AddItem "ＩＣＭ" & "#" & Trim(sFData(1)) & _
                                    "(" & Format(nCorner, "00") & ")"
' EG20 V6.1.0.1【表示号機改善対応】追加終了
                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(7))
             End If  'V1.8.0.1 ADD
' EG20 V3.4.0.1【接続機器見直し対応】追加終了
       End If
    Next
        
FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : OptTe_Click
'//  機能名称  : ラジオ釦：手入力選択時処理
'//  機能概要  : 画面を更新する。
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
Private Sub OptTe_Click()
    
    On Error Resume Next
   
    '「通信確認画面：手入力選択」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_HAND_PING, 0)
 
    frmTe.Enabled = True
    txtIP1(0).Enabled = True
    txtIP1(1).Enabled = True
    txtIP1(2).Enabled = True
    txtIP1(3).Enabled = True
    lblIP10.Enabled = True
    lblIP11.Enabled = True
    lblIP12.Enabled = True
    cmdNum(0).Enabled = True
    cmdNum(1).Enabled = True
    cmdNum(2).Enabled = True
    cmdNum(3).Enabled = True
    cmdNum(4).Enabled = True
    cmdNum(5).Enabled = True
    cmdNum(6).Enabled = True
    cmdNum(7).Enabled = True
    cmdNum(8).Enabled = True
    cmdNum(9).Enabled = True
    cmdOct.Enabled = True
    cmdBs.Enabled = True
    cmdC.Enabled = True
    
    frmkiki.Enabled = False
    txtIP2(0).Enabled = False
    txtIP2(1).Enabled = False
    txtIP2(2).Enabled = False
    txtIP2(3).Enabled = False
    lblIP20.Enabled = False
    lblIP21.Enabled = False
    lblIP22.Enabled = False
    LstKiki.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : OptKiki_Click
'//  機能名称  : ラジオ釦：機器選択選択時処理
'//  機能概要  : 画面を更新する。
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
Private Sub OptKiki_Click()
    
    On Error Resume Next
    
    '「通信確認画面：機器選択」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_KIKI_PING, 0)
    
    frmTe.Enabled = False
    txtIP1(0).Enabled = False
    txtIP1(1).Enabled = False
    txtIP1(2).Enabled = False
    txtIP1(3).Enabled = False
    lblIP10.Enabled = False
    lblIP11.Enabled = False
    lblIP12.Enabled = False
    cmdNum(0).Enabled = False
    cmdNum(1).Enabled = False
    cmdNum(2).Enabled = False
    cmdNum(3).Enabled = False
    cmdNum(4).Enabled = False
    cmdNum(5).Enabled = False
    cmdNum(6).Enabled = False
    cmdNum(7).Enabled = False
    cmdNum(8).Enabled = False
    cmdNum(9).Enabled = False
    cmdOct.Enabled = False
    cmdBs.Enabled = False
    cmdC.Enabled = False
    
    frmkiki.Enabled = True
    txtIP2(0).Enabled = True
    txtIP2(1).Enabled = True
    txtIP2(2).Enabled = True
    txtIP2(3).Enabled = True
    lblIP20.Enabled = True
    lblIP21.Enabled = True
    lblIP22.Enabled = True
    LstKiki.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdNum_Click
'//  機能名称  : 各数字釦押下時処理
'//  機能概要  : テキストボックスにIP表示
'//
'//              型        名称      意味
'//  引数      :Integer　　Index　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdNum_Click(Index As Integer)

    If Len(txtIP1(0).Text) <> 3 And bIP0 = False Then
        txtIP1(0).Text = txtIP1(0).Text & Trim(str(Index))
        Exit Sub
    End If
    If Len(txtIP1(1).Text) <> 3 And bIP1 = False Then
        txtIP1(1).Text = txtIP1(1).Text & Trim(str(Index))
        Exit Sub
    End If
    If Len(txtIP1(2).Text) <> 3 And bIP2 = False Then
        txtIP1(2).Text = txtIP1(2).Text & Trim(str(Index))
        Exit Sub
    End If
    If Len(txtIP1(3).Text) <> 3 Then
        txtIP1(3).Text = txtIP1(3).Text & Trim(str(Index))
        Exit Sub
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdOct_Click
'//  機能名称  : オクテッド(「.」)釦押下時処理
'//  機能概要  :
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
Private Sub cmdOct_Click()

    If Len(txtIP1(0).Text) <> 3 And bIP0 = False And Len(txtIP1(0)) <> 0 Then
        bIP0 = True
        Exit Sub
    End If
    If Len(txtIP1(1).Text) <> 3 And bIP1 = False And Len(txtIP1(1)) <> 0 Then
        bIP1 = True
        Exit Sub
    End If
    If Len(txtIP1(2).Text) <> 3 And bIP2 = False And Len(txtIP1(2)) <> 0 Then
        bIP2 = True
        Exit Sub
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdBs_Click
'//  機能名称  : 「BS」釦押下時処理
'//  機能概要  :
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
Private Sub cmdBs_Click()

    If Len(txtIP1(3).Text) <> 0 Then
        txtIP1(3).Text = Left(txtIP1(3).Text, Len(txtIP1(3).Text) - 1)
        Exit Sub
    End If
    
    If bIP2 = True Then
        bIP2 = False
    End If

    If Len(txtIP1(2).Text) <> 0 Then
        txtIP1(2).Text = Left(txtIP1(2).Text, Len(txtIP1(2).Text) - 1)
        Exit Sub
    End If
    
    If bIP1 = True Then
        bIP1 = False
    End If

    If Len(txtIP1(1).Text) <> 0 Then
        txtIP1(1).Text = Left(txtIP1(1).Text, Len(txtIP1(1).Text) - 1)
        Exit Sub
    End If
    
    If bIP0 = True Then
        bIP0 = False
    End If

    If Len(txtIP1(0).Text) <> 0 Then
        txtIP1(0).Text = Left(txtIP1(0).Text, Len(txtIP1(0).Text) - 1)
        Exit Sub
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdC_Click
'//  機能名称  : 「クリア」釦押下時処理
'//  機能概要  : IPテキストボックスをクリアする。
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
Private Sub cmdC_Click()

    txtIP1(0).Text = ""
    txtIP1(1).Text = ""
    txtIP1(2).Text = ""
    txtIP1(3).Text = ""
    
    bIP0 = False
    bIP1 = False
    bIP2 = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtIP1_KeyPress
'//  機能名称  : テキストボックス手入力処理
'//  機能概要  :
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
Private Sub txtIP1_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 46 Then
        If Index <> 3 Then
            txtIP1(Index + 1).SetFocus
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtIP1_Change
'//  機能名称  : テキストボックス手入力処理
'//  機能概要  :
'//
'//              型        名称      意味
'//  引数      : Integer　Index
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtIP1_Change(Index As Integer)
    If InStr(txtIP1(Index).Text, ".") <> 0 Then
        txtIP1(Index).Text = Replace(txtIP1(Index).Text, ".", "")
        Select Case Index
            Case 0:
                bIP0 = True
            Case 1:
                bIP1 = True
            Case 2:
                bIP2 = True
        End Select
    End If
    If Len(txtIP1(Index).Text) = 3 Then
        If Index <> 3 Then
            txtIP1(Index + 1).SetFocus
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : LstKiki_Click
'//  機能名称  : 機器リストボックス押下時処理
'//  機能概要  : 機器入力テキストボックスにIPアドレス表示を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　Index
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub LstKiki_Click()
    Dim sText As String
    Dim i As Integer
    Dim iTop As Integer
    Dim iText As Integer
    
    For i = 0 To 3
      txtIP2(i).Text = " "
    Next
    
    sText = sKikiIP(LstKiki.ListIndex)
    iTop = 1
    iText = 0
    
    'sTextがない場合、テキストにブランクをセットし、処理終了
    If Len(sText) = 0 Then
        For i = 0 To 3
            txtIP2(i).Text = ""
        Next
        Exit Sub
    End If
        
    'IPアドレスをテキストにセット
    For i = 1 To Len(sText)
        If Mid(sText, i, 1) = "." Then
            txtIP2(iText).Text = Mid(sText, iTop, i - iTop)
            iTop = i + 1
            iText = iText + 1
        End If
    Next
       
    txtIP2(iText).Text = Right(sText, Len(sText) - iTop + 1)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdZikko_Click
'//  機能名称  : テキストボックス手入力処理
'//  機能概要  :
'//
'//              型        名称      意味
'//  引数      : Integer　Index
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG30 31.2.0.1) 2015-07-17   REVISED BY [TCC] T.Nakajima
'//                 pingのステータスを表示できるよう修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZikko_Click()
  Dim host_path As String
  Dim wsaDD As wsaDATA
  Dim i As Integer
  Dim VerReq As Integer
  Dim rc As Long
  Dim HostAddress As Long
  Dim IcmpHandle As Long
  Dim RepryBuffer As ICMP_REPRY_BUFFER  'ICMP応答受信バッファ

    On Error GoTo ERR_SPACE
    
    '「通信確認画面：PING実行釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_PING_BUTTOM, 0)
    
    LstStatus.Clear
    DoEvents

    Me.Enabled = False
    cmdZikko.Enabled = False
    If OptTe.Value = True Then
        host_path = Trim(txtIP1(0).Text) & "." & _
                    Trim(txtIP1(1).Text) & "." & _
                    Trim(txtIP1(2).Text) & "." & _
                    Trim(txtIP1(3).Text)
        If Len(Trim(txtIP1(0).Text)) = 0 Or _
            Len(Trim(txtIP1(1).Text)) = 0 Or _
            Len(Trim(txtIP1(2).Text)) = 0 Or _
            Len(Trim(txtIP1(3).Text)) = 0 Then
            
            LstStatus.AddItem "Unknown host " & host_path
            Me.Enabled = True
            cmdZikko.Enabled = True
            Exit Sub
        End If
    Else
        host_path = txtIP2(0).Text & "." & _
                    txtIP2(1).Text & "." & _
                    txtIP2(2).Text & "." & _
                    txtIP2(3).Text
        If Len(txtIP2(0).Text) = 0 Or _
            Len(txtIP2(1).Text) = 0 Or _
            Len(txtIP2(2).Text) = 0 Or _
            Len(txtIP2(3).Text) = 0 Then

            LstStatus.AddItem "Unknown host " & host_path
            Me.Enabled = True
            cmdZikko.Enabled = True
            Exit Sub
        End If
    End If
    DoEvents
    
    '①WinSockAPIの初期化
    VerReq = MakeInteger(1, 1)                      'WinSock1.1を要求
    rc = WSAStartup(VerReq, wsaDD)
    If rc <> 0 Then
        'Winsockリソースの確保に失敗
        Me.Enabled = True
        cmdZikko.Enabled = True
        Exit Sub
    End If
    
    '②送信先のIPアドレスの取得
    HostAddress = inet_addr(host_path)              'IPアドレスに変換(数値の場合 ex:127.0.0.1)
    
    Call WSACleanup                                 'WinSockのクローズ

    '③ICMPを使ってエコーを送る
    If HostAddress <> INADDR_NONE Then
        'ICMP操作ハンドル取得
        IcmpHandle = IcmpCreateFile()
        
        LstStatus.AddItem "Pinging " & host_path & " with 32 bytes of data:"
        DoEvents
        
        'エコーを４回送る
        For i = 1 To 4
            'EG30 V31.2.0.1 DEL START
            'rc = IcmpSendEcho(IcmpHandle, HostAddress, 8, 0, 0, RepryBuffer, Len(RepryBuffer), 300)
            'If rc = 0 Then
            '    LstStatus.AddItem "Request timed out."
            'Else
            '    LstStatus.AddItem "Reply From " & host_path & _
            '                     ": bytes=32 " & _
            '                     "time=" & RepryBuffer.EchoRepry.RoundTripTime & "ms " & _
            '                     "TTL=" & CByte("128")
            'End If
            'EG30 V31.2.0.1 DEL END
            'EG30 V31.2.0.1 ADD START
            call IcmpSendEcho(IcmpHandle, HostAddress, 8, 0, 0, RepryBuffer, Len(RepryBuffer), 300)
            If RepryBuffer.EchoRepry.Status = ICMP_SUCCESS Then
                LstStatus.AddItem "Reply From " & host_path & _
                                 ": bytes=32 " & _
                                 "time=" & RepryBuffer.EchoRepry.RoundTripTime & "ms " & _
                                 "TTL=" & CByte("128")
            Else
                LstStatus.AddItem EvaluatePingResponse(RepryBuffer.EchoRepry.Status)
            End If
            'EG30 V31.2.0.1 ADD END
            DoEvents
        Next
        
        'ICMP操作ハンドルクローズ
        rc = IcmpCloseHandle(IcmpHandle)
    Else
        LstStatus.AddItem "Unknown host " & host_path
    End If
    Me.Enabled = True
    cmdZikko.Enabled = True
ERR_SPACE:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdCancel_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下時処理
'//  機能概要  :　自画面を消去する
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
    
    '「通信確認画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_END, 0)
    Unload Me
End Sub

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmPing.Caption, False
        pfFormActive (frmPing.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : psAddKikiCornerName
'//  機能名称  : 上位機器コーナ名称追加処理
'//  機能概要  : 上位機器名称に対してコーナ名称を付加する必要があれば追加する。
'//
'//              型        名称      意味
'//  引数      : String 　 sName     [IN]上位機器名称
'//  引数      : Integer　 iAreaID   [IN]上位機器通信状態エリアID
'//  引数      : Integer　 nIndex    [IN]上位機器設定構成
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V3.4.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psAddKikiCornerName(sName As String, iAreaID As Integer, nIndex As Integer)

    Dim nCorner As Integer                  ' コーナインデックス
    Dim szCornerName As String              ' コーナ名称
    Dim szResultName As String              ' 出力名称

    szResultName = ""
    nCorner = 0                                         ' コーナ設定不要
    gTransKikiInfo(nIndex).nIniListNo = nIndex          ' 外部機器リスト番号
    gTransKikiInfo(nIndex).iAreaID = iAreaID            ' 参照エリア
    gTransKikiInfo(nIndex).nCorner = nCorner
    gTransKikiInfo(nIndex).bStatus = True
    ' 1.対象外部機器上位機器通信状態エリアIDをチェックして
    '   接続対象を選別する。
    Select Case iAreaID
    Case IdKikiComSts.ID_DESYU_COM                                       ' 1:デ集通信状態
        nCorner = 1
    Case IdKikiComSts.ID_DESYU2_COM                                      ' 9:デ集2通信状態
        nCorner = 2
    Case IdKikiComSts.ID_DESYU3_COM                                      ' 10:デ集3通信状態
        nCorner = 3
    Case IdKikiComSts.ID_DESYU4_COM                                      ' 11:デ集4通信状態
        nCorner = 4
    Case IdKikiComSts.ID_DESYU5_COM                                      ' 12:デ集5通信状態
        nCorner = 5
    Case IdKikiComSts.ID_DESYU6_COM                                      ' 13:デ集6通信状態
        nCorner = 6
    Case IdKikiComSts.ID_ENKAKU_COM                                      ' 2:遠隔通信状態
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 1
    Case IdKikiComSts.ID_ENKAKU2_COM                                     ' 21:遠隔2通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 2
    Case IdKikiComSts.ID_ENKAKU3_COM                                     ' 22:遠隔3通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 3
    Case IdKikiComSts.ID_ENKAKU4_COM                                     ' 23:遠隔4通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 4
    Case IdKikiComSts.ID_ENKAKU5_COM                                     ' 24:遠隔5通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 5
    Case IdKikiComSts.ID_ENKAKU6_COM                                     ' 25:遠隔6通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 6
    Case Else
    End Select

    If nCorner <> 0 Then
        If gblnCornerSet(nCorner - 1) <> True Then
            gTransKikiInfo(nIndex).bStatus = False
        End If
        szCornerName = "(" & Format(nCorner, "00") & ")"
    End If
    szResultName = Left(sName, InStr(sName, Chr(0)) - 1)
    szResultName = szResultName + szCornerName
    gTransKikiInfo(nIndex).sGetInf = szResultName

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : OperatePing
'//  機能名称  : 操作卓設定作成
'//  機能概要  : 操作卓情報を取得表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V6.1.0.1) 2012-06-08  CODED BY  [TCC] H.Sugimoto
'//                 【操作卓ＰＩＮＧ対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub OperatePing()
  
    Dim iLoop As Integer                ' ループ
    Dim lSts As Long                    ' INI取得処理戻り値
    Dim ikousei As Integer              ' 設置構成
    Dim szSection As String             ' セクション
    Dim sIP As String                   ' 取得IPアドレス
    Dim sGetString As String * 128      ' INI取得文字列
    Dim szDispName As String            ' 表示名
    Dim nNullIndex As Integer           ' 文字数ワーク
  
    On Error Resume Next
  
  
    ' 操作卓コーナ数分ループ
    For iLoop = 1 To 6
        szSection = "KOUSEI" & Format(iLoop, "0") & "_INFO"
        szDispName = "操作卓" & "(" & Format(iLoop, "00") & ")"
        
        ' コーナ設定有り
        ' OPERATE.INIファイルから「接続有無(0:設置なし、1:設置あり）」を取得する。
        ikousei = GetPrivateProfileInt(szSection, "kousei", _
                                           0, OPERATEINI_FILE)
        
        If ikousei = 1 Then
            ' OPERATE.INIファイルから「ＩＰアドレス」を取得する。
            lSts = GetPrivateProfileString(szSection, _
                                           "ip_address", _
                                           "0.0.0.0", _
                                           sGetString, _
                                           Len(sGetString), _
                                           OPERATEINI_FILE)
        
            LstKiki.AddItem szDispName
            
            nNullIndex = InStr(sGetString, Chr(0))
            If nNullIndex <> 0 Then
                sIP = Left(sGetString, nNullIndex - 1)
            Else
                sIP = sGetString
            End If
            sKikiIP(LstKiki.ListCount - 1) = Trim(sIP)
        End If
    Next iLoop
  
End Sub


