VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKansenGateVerUpdate 
   BorderStyle     =   0  'なし
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   8160
      Top             =   3600
   End
   Begin VB.ListBox LstStatus 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   360
      TabIndex        =   42
      Top             =   5880
      Width           =   7935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGateComConf 
      Caption         =   " 自改切り離し"
      Height          =   525
      Left            =   9240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   23
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   旧 → 実行   コピー"
      Height          =   525
      Left            =   9240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   22
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyWork_Jikko 
      Caption         =   " ワーク → 実行 コピー"
      Height          =   525
      Left            =   9240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   21
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " 圧縮ファイル → ワークコピー"
      Height          =   525
      Left            =   9240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   20
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ワーククリア"
      Height          =   525
      Left            =   9240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   19
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdUSBRemove 
      Caption         =   "媒体取外"
      Height          =   525
      Left            =   9240
      TabIndex        =   18
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work2 
      Caption         =   " 媒体 → ワーク　コピー"
      Height          =   525
      Left            =   9240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Frame fraDataSelect 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   13
      Top             =   4440
      Width           =   7935
      Begin VB.CheckBox optData 
         Caption         =   "予備３"
         Height          =   240
         Index           =   8
         Left            =   5520
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "予備２"
         Height          =   240
         Index           =   7
         Left            =   5520
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "予備１"
         Height          =   240
         Index           =   6
         Left            =   5520
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "予備２"
         Height          =   240
         Index           =   5
         Left            =   2760
         TabIndex        =   26
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "予備１"
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   25
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "ＯＳ"
         Height          =   240
         Index           =   3
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "サブＣＰＵ"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "メインＣＰＵ"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "判定ＣＰＵ"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdModoru 
      Caption         =   " バージョン管理 画面へ戻る"
      Height          =   855
      Left            =   9240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   0
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdSelectNone 
      Caption         =   "全コーナ非選択"
      Height          =   525
      Left            =   3000
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "全コーナ選択"
      Height          =   525
      Left            =   360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Frame fraICMDLL 
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   11775
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "未選択"
         Height          =   855
         Index           =   5
         Left            =   9840
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   35
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "未選択"
         Height          =   855
         Index           =   4
         Left            =   7920
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   34
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "未選択"
         Height          =   855
         Index           =   3
         Left            =   6000
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   33
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "未選択"
         Height          =   855
         Index           =   2
         Left            =   4080
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   32
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "未選択"
         Height          =   855
         Index           =   1
         Left            =   2160
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   31
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FF80&
         Caption         =   "選択"
         Height          =   855
         Index           =   0
         Left            =   240
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   30
         Top             =   1080
         Value           =   1  'ﾁｪｯｸ
         Width           =   1515
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '中央揃え
         Caption         =   "コーナ６"
         Height          =   255
         Index           =   5
         Left            =   9720
         TabIndex        =   41
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '中央揃え
         Caption         =   "コーナ５"
         Height          =   255
         Index           =   4
         Left            =   7800
         TabIndex        =   40
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '中央揃え
         Caption         =   "コーナ４"
         Height          =   255
         Index           =   3
         Left            =   5895
         TabIndex        =   39
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '中央揃え
         Caption         =   "コーナ３"
         Height          =   255
         Index           =   2
         Left            =   3945
         TabIndex        =   38
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '中央揃え
         Caption         =   "コーナ２"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   37
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '中央揃え
         Caption         =   "コーナ１"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         Height          =   855
         Index           =   1
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         Height          =   855
         Index           =   2
         Left            =   3945
         TabIndex        =   10
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         Height          =   855
         Index           =   3
         Left            =   5895
         TabIndex        =   9
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         Height          =   855
         Index           =   4
         Left            =   7800
         TabIndex        =   8
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         Height          =   855
         Index           =   5
         Left            =   9720
         TabIndex        =   7
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.Frame fraMakerName 
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   10575
      Begin VB.Label lblMakerName 
         BackStyle       =   0  '透明
         Caption         =   "コーナ選択"
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   10335
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "新幹線自動改札機バージョン一括更新"
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
      Width           =   12120
   End
End
Attribute VB_Name = "frmKansenGateVerUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmGateVerUpdate.frm
'//  パッケージ名：自改バージョン一括更新画面
'//  概要        ：自改バージョン一括更新画面
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-17  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.79】
'//  REVISIONS   ： (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 【運改表示改善対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【種別チェック機能追加】【TOMAS用領域コピー対応】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////

Option Explicit

'コーナ選択釦 使用定数
Private Const SELECTSW_ON_MESSAGE = "選択"    ' 釦メッセージ：選択
Private Const SELECTSW_OFF_MESSAGE = "未選択" ' 釦メッセージ：未選択
Private Const SELECTSW_ON_COLOR = &H80FF80    ' 釦色：選択
Private Const SELECTSW_OFF_COLOR = &H80FFFF   ' 釦色：未選択
Private Const SELECTSW_ON_VALUE = 1           ' 釦状態：選択
Private Const SELECTSW_OFF_VALUE = 0          ' 釦状態：未選択

Private Const MN_FOLD_WRK = 0                   '「ワーク」フォルダ
Private Const MN_FOLD_NOW = 1                   '「実行」フォルダ
Private Const MN_FOLD_OLD = 2                   '「旧」フォルダ

Private Const MN_MAIL_INTERVAL = 1000           'メールタイマのインターバル値

Private Const FILE_NAME_MAX_SIZE = 12

Private Const EG20_JIKAI_KISHU = "EG6000"       'EG20 自改機種名
Private Const EG30_JIKAI_KISHU = "EG7000"       'EG30 自改機種名
Private Const HANKUKA_KUK = "HAN_KUKA.KUK"
Private Const INI_MAX = 5

Private Const DATA_KIND_MAX = 6                 'データ種別数       'EG20 V30.1.0.1 ADD

'【NG位置】
Private Const ERROR_HEDER = "ヘッダ"  'ヘッダ
Private Const ERROR_FOTTER = "フッタ" 'フッタ
'【NG項目】
Private Const KISHU_NAME_ERROR = "機種名"       '機種名
Private Const FILE_NAME_ERRORE = "ファイル名"   'ファイル名
Private Const CREATE_DATA_ERROR = "作成日付"    '作成日付
Private Const VERSION_ERROR = "バージョン"      'バージョン

Dim FolderSyubetu As Integer                    ' 選択リソース種別

Dim FolderName(0 To 2, 0 To 8) As String        ' フォルダ名
Dim TitleBox(0 To 8) As String                  ' タイトル名
Dim LogBox(0 To 8) As String                    ' ログ出力用タイトル名
Dim FileList() As String                        'ファイル名リスト一覧格納エリア
Dim FileListType() As String                 'ファイルリスト一覧格納エリア（次世代自改タイプを含む）
Dim gintUnkaiKind(0 To 8) As Integer            ' 運改種別    ' EG20 V5.11.0.1追加
Dim gintProgramJudgeKind(0 To 8) As Integer     ' プログラム判定種別    ' EG20 V6.9.0.1【量産対応】ADD

Private sNGSts As String        'NG位置
Private sNGKoumoku As String    'NG項目
Dim HAN_KUKA_DATA As HANTEI_DATA
Private Type HANTEI_DATA
    sHederKisyu(0 To 4) As String
    sHederFile(0 To 4) As String
    sFotterKisyu(0 To 4) As String
    sFotterFile(0 To 4) As String
End Type

' 判定用ファイル名格納エリア
'EG20 V30.1.0.1 DEL START
'Private EG20_HANTEI_CPU_CHK_FILE As String
'Private EG20_MAIN_CPU_CHK_FILE As String
'Private EG20_SUB_CPU1_CHK_FILE As String
'Private EG20_SUB_CPU2_CHK_FILE As String
'Private EG20_SUB_CPU3_CHK_FILE As String
'Private EG20_MAIN_OS_CHK_FILE As String
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
'新幹線自改
Private EG30_HANTEI_CPU_CHK_FILE As String
Private EG30_MAIN_CPU_CHK_FILE As String
Private EG30_SUB_CPU_CHK_FILE As String
Private EG30_MAIN_OS_CHK_FILE As String
'EG20 V30.1.0.1 ADD END

Private Const NGATE_00 = -1         'TOMAS領域フォルダ  EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD


'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    : Form_Load
'//  機能名称    : Form_Load時処理
'//  機能概要    : Form_Load時処理を行う
'//
'//                   型          名称            意味
'//  引数        :
'//  戻り値      :
'//
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                EG20フェーズ２対応
'//                EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   : (EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   :(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    Dim intLoop         As Integer          ' ループカウンタ
    
    On Error Resume Next
    
    '「自改バージョン一括更新画面：表示」ログ出力
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_GAMEN_START, 0)      'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_GAMEN_START, 0)      'EG20 V30.1.0.1 ADD
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
        
  
    ' /////////////////////////////////////////////////////////////////////////
    ' // コーナ設定
    ' /////////////////////////////////////////////////////////////////////////
    ' コーナ名称設定処理
    Call gsGetCornerName
    
    For intLoop = 0 To CONECT_CORNER_MAXINDEX
    
        '設定ありのコーナを活性にする
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
            ' /////////////////////////////////////////////////
            ' // ラベル（コーナー名称表示）
            lblCornerNo(intLoop).Visible = True
            lblGokiBetsuNumber(intLoop).Caption = gstrCornerName(intLoop)
            lblGokiBetsuNumber(intLoop).Visible = True
            
            ' /////////////////////////////////////////////////
            ' // 釦
            chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
            chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
            chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
            
            chkUpdate(intLoop).Visible = True
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
            'EG20 V30.1.0.1 ADD START
            '幹線コーナーのみ押下可能とする。
'            If gintCornerType(intLoop) = CORNER_TYPE_KANSEN Then
'                chkUpdate(intLoop).Enabled = True
'            Else
'                chkUpdate(intLoop).Enabled = False
'            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 【HKRK_Kansi06_005_01】 DEL END

        Else
            lblCornerNo(intLoop).Visible = False
            lblGokiBetsuNumber(intLoop).Caption = ""
            lblGokiBetsuNumber(intLoop).Visible = False
        
            chkUpdate(intLoop).Visible = False
        End If
    
    Next intLoop
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // その他コントロール設定
    ' /////////////////////////////////////////////////////////////////////////
    LstStatus.Clear
   
    'For intLoop = 0 To 8       'EG20 V30.1.0.1 DEL
    For intLoop = 0 To DATA_KIND_MAX - 1    'EG20 V30.1.0.1 ADD
        optData(intLoop).Value = SELECTSW_ON_VALUE      ' 未チェック
    Next intLoop
    
    ' 自改ＤＬＬフォルダ設定
    sSetFolderName
    
    ' 変数の初期化
    FolderSyubetu = 0
    
    ' コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(True)
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：Form_Activate
'//  機能名称    ：自改バージョン一括更新画面(アクティブ時)
'//  機能概要    ：画面再表示処理を行う。
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    ' 現在フォームを更新
    'gStrCurrentForm = sFormName_GateVerUpdate      'EG20 V30.1.0.1 DEL
    gStrCurrentForm = sFormName_KGateVerUpdate      'EG20 V30.1.0.1 ADD
    
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：Form_Deactivate
'//  機能名称    ：自改バージョン一括更新画面(ディアクティブ時)
'//  機能概要    ：メール受信用のタイマ停止
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
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
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdModoru_Click
'//  機能名称    ：「バージョン管理画面へ戻る」ボタン押下処理
'//  機能概要    ：「バージョン管理画面へ戻る」ボタン押下処理を行う
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Click()
    On Error Resume Next
    
    '「自改バージョン一括更新画面：消去」ログ出力
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_GAMEN_END, 0)        'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_GAMEN_END, 0)        'EG20 V30.1.0.1 ADD

    '画面のUnload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称     : cmdSelectAll_Click
'//  機能名称     : 「全コーナ選択」ボタン押下処理
'//  機能概要     : 「全コーナ選択」ボタン押下処理を行う
'//
'//                   型          名称            意味
'//  引数         :
'//  戻り値       :
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   : (EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdSelectAll_Click()

    Dim intLoop     As Integer          ' ループカウンタ

    On Error Resume Next
    
    '「自改バージョン一括更新画面：全コーナ選択釦押下」ログ出力
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_SELECTALL_BUTTON, 0)     'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_SELECTALL_BUTTON, 0)     'EG20 V30.1.0.1 ADD

    For intLoop = 0 To CONECT_CORNER_MAXINDEX

        If chkUpdate(intLoop).Visible = True Then
            'EG20 V30.1.0.1 DEL START
'            chkUpdate(intLoop).Caption = SELECTSW_ON_MESSAGE
'            chkUpdate(intLoop).BackColor = SELECTSW_ON_COLOR
'            chkUpdate(intLoop).Value = SELECTSW_ON_VALUE
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
            'EG20 V30.1.0.1 ADD START
            '新幹線コーナーに対してのみ切り替えを行う。
'            If gintCornerType(intLoop) = CORNER_TYPE_KANSEN Then
'                chkUpdate(intLoop).Caption = SELECTSW_ON_MESSAGE
'                chkUpdate(intLoop).BackColor = SELECTSW_ON_COLOR
'                chkUpdate(intLoop).Value = SELECTSW_ON_VALUE
'            End If
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】DEL END
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】ADD START
            chkUpdate(intLoop).Caption = SELECTSW_ON_MESSAGE
            chkUpdate(intLoop).BackColor = SELECTSW_ON_COLOR
            chkUpdate(intLoop).Value = SELECTSW_ON_VALUE
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】ADD END
        End If
    Next

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称     : cmdSelectAll_Click
'//  機能名称     : 「全コーナ非選択」ボタン押下処理
'//  機能概要     : 「全コーナ非選択」ボタン押下処理を行う
'//
'//                   型          名称            意味
'//  引数         :
'//  戻り値       :
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdSelectNone_Click()

    Dim intLoop     As Integer          ' ループカウンタ

    On Error Resume Next
    
    '「自改バージョン一括更新画面：全コーナ非選択釦押下」ログ出力
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_SELECTALLOFF_BUTTON, 0)      'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_SELECTALLOFF_BUTTON, 0)      'EG20 V30.1.0.1 ADD

    For intLoop = 0 To CONECT_CORNER_MAXINDEX

        If chkUpdate(intLoop).Visible = True Then
            'EG20 V30.1.0.1 DEL START
'            chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
'            chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
'            chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_01】 DEL START
            'EG20 V30.1.0.1 ADD START
            '幹線コーナーに対してのみ切替を行う。
'            If gintCornerType(intLoop) = CORNER_TYPE_KANSEN Then
'                chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
'                chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
'                chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
'            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
            chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
            chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
            chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
            'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
        End If
    Next

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称     : chkUpdate_Click
'//  機能名称     : 「コーナ別選択」ボタン押下処理
'//  機能概要     : 「コーナ別選択」ボタン押下処理を行う
'//
'//                   型          名称            意味
'//  引数         :
'//  戻り値       :
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub chkUpdate_Click(Index As Integer)

    '「自改バージョン一括更新画面：コーナ選択釦押下」ログ出力
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_SELECT_BUTTON, 0)        'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_SELECT_BUTTON, 0)        'EG20 V30.1.0.1 ADD

    If chkUpdate(Index).Value = SELECTSW_ON_VALUE Then
        chkUpdate(Index).Caption = SELECTSW_ON_MESSAGE
        chkUpdate(Index).BackColor = SELECTSW_ON_COLOR
'        chkUpdate(Index).Value = SELECTSW_ON_VALUE
    Else
        chkUpdate(Index).Caption = SELECTSW_OFF_MESSAGE
        chkUpdate(Index).BackColor = SELECTSW_OFF_COLOR
'        chkUpdate(Index).Value = SELECTSW_OFF_VALUE
    End If

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： cmdClear_Click
'//  機能名称    ：「ワーククリア」ボタン押下処理
'//  機能概要    ：「ワーククリア」ボタン押下処理を行う
'//
'//                 型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ： (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ： (EG20 V30.3.0.1) 2014-11-13 REVISED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()
    
    Dim iResponse As Integer        ' MsgBoxボタンコード
    Dim iCornerLoop As Integer      ' ループ
    Dim iSelctLoop As Integer       ' ループ
    Dim bStatus As Boolean          ' 処理結果
    Dim iTomasFlg   As Integer      ' TOMAS処理済フラグ（コーナ一括処理中に一つのコーナだけTOMAS領域のコピーを行う）    'EG20 V30.3.0.1 ADD
    
    
    On Error Resume Next
    
    '「自改バージョン一括更新画面：ワーククリア釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_CREA_BUTTOM, 0)
    
    ' コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(False)

    ' 確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("選択されたコーナ・種別の「ワーク」フォルダ内のファイルを全て削除します。" _
                         & Chr(vbKeyReturn) & "よろしいですか？", _
                        vbYesNo + vbExclamation, _
                        "ワーク クリア")
    If iResponse = vbYes Then

        ' コーナ選択・種別選択チェック
        If sSelectChk = False Then
            'コマンドボタン押下可・不可処理
             Call sCmdBtnEnabled(True)
            Exit Sub
        End If

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

        LstStatus.Clear

        ' /////////////////////////////////////////////////
        ' // コーナ単位での処理
        iTomasFlg = 0       ' EG20 V30.3.0.1 ADD
        For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
            If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
'                For iSelctLoop = 0 To 8        'EG20 V30.1.0.1 DEL
                For iSelctLoop = 0 To DATA_KIND_MAX - 1
                    If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                        FolderSyubetu = iSelctLoop
                        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
                        If iTomasFlg = 0 Then
                            bStatus = sWrkFolderRemove(NGATE_00)
                            'Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
                        End If
                        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
                        bStatus = sWrkFolderRemove(iCornerLoop)
                        Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
                    End If
                Next
                'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
                '最初のコーナ1回のみTOMAS領域のフォルダにコピーすればOKなので、以降やらないようにフラグをON
                iTomasFlg = 1
                'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
            End If
        Next
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    End If

    'コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(True)
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： cmdCopyBaitai_Work_Click
'//  機能名称    ：「圧縮ファイル→ワークコピー」釦押下時処理
'//  機能概要    ：「圧縮ファイル→ワークコピー」釦押下時処理を行う
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()
    
    On Error Resume Next
    
    'コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(False)
    
    '「自改ﾊﾞｰｼﾞｮﾝ：圧縮ﾌｧｲﾙ→ﾜｰｸｺﾋﾟｰ釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_CAB_COPY_WRK_BUTTOM, 0)
 
    ' コーナ選択・種別選択チェック
    If sSelectChk = False Then
        'コマンドボタン押下可・不可処理
        Call sCmdBtnEnabled(True)
        Exit Sub
    End If
    
    LstStatus.Clear
    
    '圧縮ファイルからインストールする。
    sFDInstall "LZH"

   'コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： cmdCopyBaitai_Work2_Click
'//  機能名称    ：「媒体 →ワークコピー」釦押下時処理
'//  機能概要    ：「媒体 →ワークコピー」釦押下時処理を行う
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work2_Click()

    On Error Resume Next
    
    'コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(False)
    
    '「自改ﾊﾞｰｼﾞｮﾝ：圧縮ﾌｧｲﾙ→ﾜｰｸｺﾋﾟｰ釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_LZH_COPY_WRK_BUTTOM, 0)
 
    ' コーナ選択・種別選択チェック
    If sSelectChk = False Then
        'コマンドボタン押下可・不可処理
        Call sCmdBtnEnabled(True)
        Exit Sub
    End If
    
    LstStatus.Clear
    
    'インストール媒体をワークフォルダ内にコピーする。
    sFDInstall "STD"

   'コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称     : cmdCopyWork_Jikko_Click
'//  機能名称     :「ワーク → 実行コピー」釦押下時処理
'//  機能概要     :「ワーク → 実行コピー」釦押下時処理を行う
'//
'//                   型          名称            意味
'//  引数         :
'//  戻り値       :
'//
'//  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS    :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS    :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS    :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS    :(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ： (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 【運改表示改善対応】
'//  REVISIONS   ： (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(EG20 V30.3.0.1) 2014-11-11  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
    Dim iResponse               As Integer      'MsgBoxボタンコード
    Dim iCornerLoop As Integer      ' ループ
    Dim iSelctLoop As Integer       ' ループ
    Dim bStatus As Boolean          ' 処理結果
    Dim iTomasFlg   As Integer      ' TOMAS処理済フラグ（コーナ一括処理中に一つのコーナだけTOMAS領域のコピーを行う）'EG20 V30.3.0.1 ADD

    
    On Error Resume Next
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_COPY_NOW_BUTTOM, 0)

    Call sCmdBtnEnabled(False)

    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("選択されたコーナ・種別の実行フォルダをクリアして" _
                        & Chr(vbKeyReturn) & "ワークフォルダのファイルをコピーしますが" _
                        & Chr(vbKeyReturn) & "よろしいですか？", _
                        vbYesNo + vbExclamation, _
                        "ワーク→実行コピー")
    If iResponse = vbYes Then

        ' コーナ選択・種別選択チェック
        If sSelectChk = False Then
            'コマンドボタン押下可・不可処理
             Call sCmdBtnEnabled(True)
            Exit Sub
        End If

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        LstStatus.Clear
        ' /////////////////////////////////////////////////
        ' // コーナ単位での処理
        iTomasFlg = 0   'EG20 V30.3.0.1 ADD
        For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
            If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
                'For iSelctLoop = 0 To 8        'EG20 V30.1.0.1 DEL
                For iSelctLoop = 0 To DATA_KIND_MAX - 1     'EG20 V30.1.0.1 ADD
                    If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                        FolderSyubetu = iSelctLoop
                        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
                        If iTomasFlg = 0 Then
                            bStatus = fNewVersion(NGATE_00)
                            'Call AddMessageLstStatus(NGATE_00, FolderSyubetu, bStatus)
                        End If
                        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
                        bStatus = fNewVersion(iCornerLoop)
                        Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
' EG20 V3.0.0.2追加開始
                        If bStatus = True Then
                            '改札機共通エリア更新処理（正常）
                            Call pubfuncCommonAreaUpdate
' EG20 V5.8.0.1削除開始
'                            ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'                            Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
                            ' 運改状態更新
                            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1追加終了
'                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1)   ' EG20 V5.6.0.1追加           ' EG20 V5.11.0.1削除
                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1追加
                        End If
' EG20 V3.0.0.2追加終了
                    End If
                Next
                'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
                '最初のコーナ1回のみTOMAS領域のフォルダにコピーすればOKなので、以降やらないようにフラグをON
                iTomasFlg = 1
                'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
            End If
        Next

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    End If

   'コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称     : cmdCopyOld_Jikko_Click
'//  機能名称     :「旧 → 実行コピー」釦押下時処理
'//  機能概要     :「旧 → 実行コピー」釦押下時処理を行う
'//
'//                   型          名称            意味
'//  引数         :
'//  戻り値       :
'//
'//  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS    :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS    :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS    : (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 【運改表示改善対応】
'//  REVISIONS    : (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(EG20 V30.3.0.1) 2014-11-11  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    Dim iResponse               As Integer      'MsgBoxボタンコード
    Dim iCornerLoop As Integer      ' ループ
    Dim iSelctLoop As Integer       ' ループ
    Dim bStatus As Boolean          ' 処理結果
    
    Dim iTomasFlg   As Integer      ' TOMAS処理済フラグ（コーナ一括処理中に一つのコーナだけTOMAS領域のコピーを行う）'EG20 V30.3.0.1 ADD


    On Error Resume Next
    '「自改ﾊﾞｰｼﾞｮﾝ：旧→実行コピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OLD_COPY_NOW_BUTTOM, 0)

    Call sCmdBtnEnabled(False)
    
    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("選択されたコーナ・種別の実行フォルダをクリアして" _
                        & Chr(vbKeyReturn) & "旧フォルダのファイルをコピーしますが" _
                        & Chr(vbKeyReturn) & "よろしいですか？", _
                        vbYesNo + vbExclamation, _
                        "旧→実行コピー")
    If iResponse = vbYes Then

        ' コーナ選択・種別選択チェック
        If sSelectChk = False Then
            'コマンドボタン押下可・不可処理
             Call sCmdBtnEnabled(True)
            Exit Sub
        End If

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

        LstStatus.Clear

        ' /////////////////////////////////////////////////
        ' // コーナ単位での処理
        iTomasFlg = 0       ' EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD
        For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
            If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
                'For iSelctLoop = 0 To 8                    'EG20 V30.1.0.1 DEL
                For iSelctLoop = 0 To DATA_KIND_MAX - 1     'EG20 V30.1.0.1 ADD
                    If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                        FolderSyubetu = iSelctLoop
                        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
                        If iTomasFlg = 0 Then
                            bStatus = fOldVersion(NGATE_00)
                            'Call AddMessageLstStatus(NGATE_00, FolderSyubetu, bStatus)
                        End If
                        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
                        bStatus = fOldVersion(iCornerLoop)
                        Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
' EG20 V3.0.0.2追加開始
                        If bStatus = True Then
                            '改札機共通エリア更新処理（正常）
                            Call pubfuncCommonAreaUpdate
' EG20 V5.8.0.1追加開始
                            ' 運改状態更新
                            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1追加終了
'                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1)   ' EG20 V5.6.0.1追加          ' EG20 V5.11.0.1削除
                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1, gintUnkaiKind(FolderSyubetu))   ' EG20 V5.11.0.1追加
                        Else
                            Call pubfuncErrorOccur(MN_FOLD_NOW)
                        End If
' EG20 V3.0.0.2追加終了
                    End If
                Next
                'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
                '最初のコーナ1回のみTOMAS領域のフォルダにコピーすればOKなので、以降やらないようにフラグをON
                iTomasFlg = 1       ' EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD
                'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
            End If
        Next

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    End If

   'コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： cmdGateComConf_Click
'//  機能名称    ：「自改切り離し」釦押下時処理
'//  機能概要    ：「自改切り離し」釦押下時処理を行う
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdGateComConf_Click()
    '「自改ﾊﾞｰｼﾞｮﾝ：自改切り離し釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_KIRIHANASI_BUTTOM, 0)

    '通信接続・切断画面を表示する。
    Load frmConectSts
    frmConectSts.Show 1
End Sub


'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称     : cmdUSBRemove_Click
'//  機能名称     :「媒体取り外し」釦押下時処理
'//  機能概要     :「媒体取り外し」釦押下時処理を行う
'//
'//                   型          名称            意味
'//  引数         :
'//  戻り値       :
'//
'//  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdUSBRemove_Click()
    On Error Resume Next
   
    '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
    ' コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(False)
 
    ' 媒体取外処理
    Call pfRemove(Me)

    ' コマンドボタン押下可・不可処理
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'/
'/  関数名称     : sSelectChk
'/  機能名称     : コーナ選択・種別選択チェック
'/  機能概要     : コーナ選択・種別選択チェック処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.79】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Function sSelectChk() As Boolean
    Dim iCnt            As Integer
    Dim bRet            As Boolean

    '初期値設定
    sSelectChk = True

    'コーナ選択チェック
    bRet = False
    For iCnt = 0 To CONECT_CORNER_MAXINDEX
        
        If chkUpdate(iCnt).Value = SELECTSW_ON_VALUE Then
            bRet = True
        End If
    Next
    
    If bRet = False Then
        sSelectChk = False
        MsgBox "コーナが選択されていません。", _
                vbOKOnly + vbExclamation, _
                "コーナ選択"
        Exit Function
    End If

    '種別選択チェック
    bRet = False
    'For iCnt = 0 To 8                      'EG20 V30.1.0.1 DEL
    For iCnt = 0 To DATA_KIND_MAX - 1       'EG20 V30.1.0.1 ADD
        
        If optData(iCnt).Value = SELECTSW_ON_VALUE Then
            bRet = True
        End If
    Next
    
    If bRet = False Then
        sSelectChk = False
' EG20 V3.6.0.1【03統合TR-No.79】削除開始
'        MsgBox "種別がされていません。", _
'                vbOKOnly + vbExclamation, _
'                "種別選択"
' EG20 V3.6.0.1【03統合TR-No.79】削除終了
' EG20 V3.6.0.1【03統合TR-No.79】追加開始
        MsgBox "種別が選択されていません。", _
                vbOKOnly + vbExclamation, _
                "種別選択"
' EG20 V3.6.0.1【03統合TR-No.79】追加終了
        Exit Function
    End If

End Function

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'/
'/  関数名称     : sCmdBtnEnabled
'/  機能名称     : コマンドボタン押下可・不可処理
'/  機能概要     : コマンドボタンを引数に基いて押下可・不可処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                EG20フェーズ２対応
'//                EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)
    Dim iLoopCnt    As Integer

    'コーナ選択釦
    For iLoopCnt = 0 To CONECT_CORNER_MAXINDEX
        'chkUpdate(iLoopCnt).Enabled = blnFlg       'EG20 V30.1.0.1 DEL
        chkUpdate(iLoopCnt).Enabled = blnFlg       'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】ADD
        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
        'EG20 V30.1.0.1 ADD START
        '幹線コーナーに対してのみ設定可能とする。
'        If gintCornerType(iLoopCnt) = CORNER_TYPE_KANSEN Then
'            chkUpdate(iLoopCnt).Enabled = blnFlg
'        End If
        'EG20 V30.1.0.1 ADD END
        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
    Next
    
    '全コーナ選択・非選択釦
    cmdSelectAll.Enabled = blnFlg
    cmdSelectNone.Enabled = blnFlg
    
    '種別選択
    'For iLoopCnt = 0 To 8                  'EG20 V30.1.0.1 DEL
    For iLoopCnt = 0 To DATA_KIND_MAX - 1   'EG20 V30.1.0.1 ADD
        optData(iLoopCnt).Enabled = blnFlg
    Next
    
    '各コマンド釦
    cmdClear.Enabled = blnFlg             ' 「ワーククリア」
    cmdCopyBaitai_Work.Enabled = blnFlg   ' 「圧縮ファイル→ワークコピー」
    cmdCopyBaitai_Work2.Enabled = blnFlg  ' 「媒体→ワークコピー」
    cmdCopyWork_Jikko.Enabled = blnFlg    ' 「ワーク→実行コピー」
    cmdCopyOld_Jikko.Enabled = blnFlg     ' 「旧→実行コピー」
    cmdGateComConf.Enabled = blnFlg       ' 「自改切り離し」
    cmdUSBRemove.Enabled = blnFlg         ' 「媒体取外」
    cmdModoru.Enabled = blnFlg            ' 「戻る」

End Sub
'EG20 V30.1.0.1 DEL START
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
''//
''//  関数名称    ： sSetFolderName
''//  機能名称    ： データ展開
''//  機能概要    ： フォルダ名などのデータをグローバルエリアに展開する。
''//
''//                 型        名称      意味
''//  引数        ： なし
''//
''//                 型        値        意味
''//  戻り値      ： なし
''//
''//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
''//                 EG20フェーズ２対応
''//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
''//  REVISIONS   ： (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
''//                 【運改表示改善対応】
''//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
''//                 量産対応【種別チェック機能追加】
''//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
''//  備考        ：
''///////////////////////////////////////////////////////////////////
'Private Sub sSetFolderName()
'
'    TitleBox(0) = "判定データ  "
'    TitleBox(1) = "プログラム  "
'    TitleBox(2) = "サブCPU-Pro1"
'    TitleBox(3) = "サブCPU-Pro2"
'    TitleBox(4) = "サブCPU-Pro3"
'    TitleBox(5) = "自改（ＯＳ）"
'    TitleBox(6) = "予備1       "
'    TitleBox(7) = "予備2       "
'    TitleBox(8) = "予備3       "
'
'    LogBox(0) = "判定"
'    LogBox(1) = "プロ"
'    LogBox(2) = "サブ1"
'    LogBox(3) = "サブ2"
'    LogBox(4) = "サブ3"
'    LogBox(5) = "OS"
'    LogBox(6) = "予備1"
'    LogBox(7) = "予備2"
'    LogBox(8) = "予備3"
'
'    'フォルダ名に設定を行う
'    FolderName(MN_FOLD_WRK, 0) = EG20_NHAN1WRK       ' 判定データCPU-PRO（ワーク）
'    FolderName(MN_FOLD_NOW, 0) = EG20_NHAN1NOW       ' 判定データCPU-PRO（実行）
'    FolderName(MN_FOLD_OLD, 0) = EG20_NHAN1OLD       ' 判定データCPU-PRO（旧）
'
'    FolderName(MN_FOLD_WRK, 1) = EG20_NPRO1WRK       ' 改札機プログラム情報（ワーク）
'    FolderName(MN_FOLD_NOW, 1) = EG20_NPRO1NOW       ' 改札機プログラム情報（実行）
'    FolderName(MN_FOLD_OLD, 1) = EG20_NPRO1OLD       ' 改札機プログラム情報（旧）
'
'    FolderName(MN_FOLD_WRK, 2) = EG20_NSCP1WRK       ' サブCPU1-PRO（ワーク）
'    FolderName(MN_FOLD_NOW, 2) = EG20_NSCP1NOW       ' サブCPU1-PRO（実行）
'    FolderName(MN_FOLD_OLD, 2) = EG20_NSCP1OLD       ' サブCPU1-PRO（旧）
'
'    FolderName(MN_FOLD_WRK, 3) = EG20_NSCP2WRK       ' サブCPU2-PRO（ワーク）
'    FolderName(MN_FOLD_NOW, 3) = EG20_NSCP2NOW       ' サブCPU2-PRO（実行）
'    FolderName(MN_FOLD_OLD, 3) = EG20_NSCP2OLD       ' サブCPU2-PRO（旧）
'
'    FolderName(MN_FOLD_WRK, 4) = EG20_NSCP3WRK       ' サブCPU3-PRO（ワーク）
'    FolderName(MN_FOLD_NOW, 4) = EG20_NSCP3NOW       ' サブCPU3-PRO（実行）
'    FolderName(MN_FOLD_OLD, 4) = EG20_NSCP3OLD       ' サブCPU3-PRO（旧）
'
'    FolderName(MN_FOLD_WRK, 5) = EG20_NOSWRK         ' 改札機（OS）情報（ワーク）
'    FolderName(MN_FOLD_NOW, 5) = EG20_NOSNOW         ' 改札機（OS）情報（実行）
'    FolderName(MN_FOLD_OLD, 5) = EG20_NOSOLD         ' 改札機（OS）情報（旧）
'
'    FolderName(MN_FOLD_WRK, 6) = EG20_NYOBI1WRK      ' 予備1（ワーク）
'    FolderName(MN_FOLD_NOW, 6) = EG20_NYOBI1NOW      ' 予備1（実行）
'    FolderName(MN_FOLD_OLD, 6) = EG20_NYOBI1OLD      ' 予備1（旧）
'
'    FolderName(MN_FOLD_WRK, 7) = EG20_NYOBI2WRK      ' 予備2（ワーク）
'    FolderName(MN_FOLD_NOW, 7) = EG20_NYOBI2NOW      ' 予備2（実行）
'    FolderName(MN_FOLD_OLD, 7) = EG20_NYOBI2OLD      ' 予備2（旧）
'
'    FolderName(MN_FOLD_WRK, 8) = EG20_NYOBI3WRK      ' 予備3（ワーク）
'    FolderName(MN_FOLD_NOW, 8) = EG20_NYOBI3NOW      ' 予備3（実行）
'    FolderName(MN_FOLD_OLD, 8) = EG20_NYOBI3OLD      ' 予備3（旧）
'
'    ' /////////////////////////////////////////////////////
'    ' // EG20自改
'    ' キー名:判定CPU-PRO代表
'    EG20_HANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
'
'    ' キー名:メインCPU-PRO代表
'    EG20_MAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_MAIN_PRO, PATH_GATEVER_FILE)
'
'    ' キー名：サブCPU-PRO代表
'    EG20_SUB_CPU1_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO1, PATH_GATEVER_FILE)
'    EG20_SUB_CPU2_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO2, PATH_GATEVER_FILE)
'    EG20_SUB_CPU3_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO3, PATH_GATEVER_FILE)
'
'    ' キー名:メインCPU-OS代表
'    EG20_MAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_MAIN_OS, PATH_GATEVER_FILE)
'
'' EG20 V5.11.0.1【運改表示改善対応】追加開始
'    gintUnkaiKind(0) = BootInfoGateType.TYPE_NHAN
'    gintUnkaiKind(1) = BootInfoGateType.TYPE_NPRO
'    gintUnkaiKind(2) = BootInfoGateType.TYPE_NSCP1
'    gintUnkaiKind(3) = BootInfoGateType.TYPE_NSCP2
'    gintUnkaiKind(4) = BootInfoGateType.TYPE_NSCP3
'    gintUnkaiKind(5) = BootInfoGateType.TYPE_NOS
'    gintUnkaiKind(6) = BootInfoGateType.TYPE_NYOBI1
'    gintUnkaiKind(7) = BootInfoGateType.TYPE_NYOBI2
'    gintUnkaiKind(8) = BootInfoGateType.TYPE_NYOBI3
'' EG20 V5.11.0.1【運改表示改善対応】追加終了
'
'' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD START
'    gintProgramJudgeKind(0) = ProgramJudgeKind.JUDGE_NHAN       ' 判定データ
'    gintProgramJudgeKind(1) = ProgramJudgeKind.JUDGE_NPRO       ' プログラム
'    gintProgramJudgeKind(2) = ProgramJudgeKind.JUDGE_NSCP1      ' サブCPU-Pro1
'    gintProgramJudgeKind(3) = ProgramJudgeKind.JUDGE_NSCP2      ' サブCPU-Pro2
'    gintProgramJudgeKind(4) = ProgramJudgeKind.JUDGE_NSCP3      ' サブCPU-Pro3
'    gintProgramJudgeKind(5) = ProgramJudgeKind.JUDGE_NOS        ' 自改（OS）
'    gintProgramJudgeKind(6) = ProgramJudgeKind.JUDGE_NOCHECK    ' 予備1
'    gintProgramJudgeKind(7) = ProgramJudgeKind.JUDGE_NOCHECK    ' 予備2
'    gintProgramJudgeKind(8) = ProgramJudgeKind.JUDGE_NOCHECK    ' 予備3
'' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD END
'
'End Sub
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-18  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSetFolderName()

        TitleBox(0) = "判定ＣＰＵ"
        TitleBox(1) = "メインＣＰＵ"
        TitleBox(2) = "サブＣＰＵ"
        TitleBox(3) = "ＯＳ"
        TitleBox(4) = "予備１"
        TitleBox(5) = "予備２"
    
        LogBox(0) = "判定"
        LogBox(1) = "プログメイン"
        LogBox(2) = "サブ"
        LogBox(3) = "ＯＳ"
        LogBox(4) = "予備１"
        LogBox(5) = "予備２"
        
        'フォルダ名に設定を行う
        FolderName(0, 0) = EG30_JHANWRK
        FolderName(1, 0) = EG30_JHANNOW
        FolderName(2, 0) = EG30_JHANOLD
        FolderName(0, 1) = EG30_JPROWRK
        FolderName(1, 1) = EG30_JPRONOW
        FolderName(2, 1) = EG30_JPROOLD
        FolderName(0, 2) = EG30_JSCPUWRK
        FolderName(1, 2) = EG30_JSCPUNOW
        FolderName(2, 2) = EG30_JSCPUOLD
        FolderName(0, 3) = EG30_JOSWRK
        FolderName(1, 3) = EG30_JOSNOW
        FolderName(2, 3) = EG30_JOSOLD
        FolderName(0, 4) = EG30_JYOBIWK1
        FolderName(1, 4) = EG30_JYOBINW1
        FolderName(2, 4) = EG30_JYOBIOD1
        FolderName(0, 5) = EG30_JYOBIWRK
        FolderName(1, 5) = EG30_JYOBINOW
        FolderName(2, 5) = EG30_JYOBIOLD

'-------新幹線自改-------
    ' キー名:判定CPU-PRO代表
    EG30_HANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    ' キー名:メインCPU-PRO代表
    EG30_MAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_MAIN_PRO, PATH_GATEVER_FILE)
        
    ' キー名：サブCPU-PRO代表
    EG30_SUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_SUB_PRO1, PATH_GATEVER_FILE)
    
    ' キー名:メインCPU-OS代表
    EG30_MAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_MAIN_OS, PATH_GATEVER_FILE)

    gintUnkaiKind(0) = BootInfoGateType.TYPE_JHAN
    gintUnkaiKind(1) = BootInfoGateType.TYPE_JPRO
    gintUnkaiKind(2) = BootInfoGateType.TYPE_JSCPU
    gintUnkaiKind(3) = BootInfoGateType.TYPE_JOS

    gintProgramJudgeKind(0) = ProgramJudgeKind.JUDGE_JHAN       'a:判定CPU用プログラム（幹線）
    gintProgramJudgeKind(1) = ProgramJudgeKind.JUDGE_JPRO       'b:メインCPU用プログラム（幹線）
    gintProgramJudgeKind(2) = ProgramJudgeKind.JUDGE_JSCPU     'c:サブCPUプログラム（幹線）
    gintProgramJudgeKind(3) = ProgramJudgeKind.JUDGE_JOS        ' d:OSプログラム（幹線）
    gintProgramJudgeKind(4) = ProgramJudgeKind.JUDGE_YOBI1      'e:予備１（幹線） チェック無し
    gintProgramJudgeKind(5) = ProgramJudgeKind.JUDGE_YOBI       'f:予備（幹線） チェック無し
    gintProgramJudgeKind(6) = ProgramJudgeKind.JUDGE_NOCHECK
    gintProgramJudgeKind(7) = ProgramJudgeKind.JUDGE_NOCHECK
    gintProgramJudgeKind(8) = ProgramJudgeKind.JUDGE_NOCHECK

End Sub
'EG20 V30.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sWrkFolderRemove
'//  機能名称    ： ワークフォルダ内ファイル削除処理
'//  機能概要    ： ワークフォルダ内のファイルを削除する。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-17  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ： (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 【運改表示改善対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove(nCorner As Integer) As Boolean
    Dim MyName As String                    ' ファイル名
    Dim lngErrCode As Long                  ' エラーコード
    Dim lngPgmHanteiStsWork As Long         ' プログラム判定状態（ワーク）   ' EG20 V3.6.0.1追加
    
    Dim objFso As New FileSystemObject      ' ファイルシステムオブジェクト
    Dim objFi As File                       ' ファイルオブジェクト
    
    On Error GoTo ErrorHandler              ' エラーハンドルの登録

    '初期値設定
    sWrkFolderRemove = True
   
    'ワークフォルダ内のディレクトリの名前を表示します。（コーナ単位）
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
    
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
'    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, nCorner + 1)   ' EG20 V5.6.0.1追加
' EG20 V5.11.0.1削除終了
' EG20 V5.11.0.1追加開始
    ' 運改状態更新
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_CLEAR)
    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, nCorner + 1, gintUnkaiKind(FolderSyubetu))
' EG20 V5.11.0.1追加終了

    Exit Function '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    Call pubfuncErrorOccur(MN_FOLD_WRK)             ' EG20 V3.6.0.1追加
           
   '「自改ﾊﾞｰｼﾞｮﾝ：ﾜｰｸﾌｫﾙﾀﾞﾌｧｲﾙ削除異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    Set objFso = Nothing
    Set objFi = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sFDInstall
'//  機能名称    ： 媒体インストール処理
'//  機能概要    ： インストール媒体ファイルを、ワークフォルダにコピーする。
'//
'//                 型        名称      意味
'//  引数        ： String    sFlag     処理種別
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ： (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 【運改表示改善対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応【種別チェック機能追加】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall(sFlag As String)
    Dim iResponse As Integer        'MsgBoxボタンコード
    Dim sInputPass As String        'インストール元ディレクトリ名(STD)orファイル名(LZH)
    Dim sInputFolder As String      'インストール元フォルダ名。LZHの時、解凍先フォルダ。
    Dim lngErrCode As Long          'エラーコード
    Dim bRet As Boolean             '正当性チェック戻り値
    Dim bStatus As Boolean          ' 処理結果
    Dim iCornerLoop As Integer      ' ループ
    Dim iSelctLoop As Integer       ' ループ
   
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    Dim objFi As File                    'ファイルオブジェクト

    Dim lngPgmHanteiStsWork As Long     'プログラム判定状態（ワーク）   ' EG20 V3.0.0.2追加

    On Error GoTo ErrorHandler      'エラーハンドルの登録

    If sFlag = "STD" Then
    '標準（非圧縮）ファイル指定の時:
    'ディレクトリ選択画面を表示させ、入力ファイル格納ディレクトリ名を得る。
        sInputPass = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
        If sInputPass = "" Then
        'ディレクトリが指定なし時は処理終了
            Set objFso = Nothing
            Set objFi = Nothing
            Exit Sub
        End If
        sInputFolder = sInputPass
    Else
    '圧縮ファイル指定の時:
    '圧縮ファイル選択画面を表示させ、LZHファイルフルパス名を得る（デフォルトはＦＤを表示。）。
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
        
        Call ChDrive("D")
        
       '解凍用一時フォルダを作成する。
       psMakeFolder MELTED_FOLDER_FULLPASS
       '圧縮ファイルを、解凍用一時フォルダに解凍・格納させる。
        Call psCabReqest(CABREQEST.CAB_THAW, sInputPass, MELTED_FOLDER_FULLPASS)
        If glngCabErrCd <> 0 Then
            psDeleteFolder MELTED_FOLDER_FULLPASS
            Set objFso = Nothing
            Set objFi = Nothing
            Exit Sub
        End If
        sInputFolder = MELTED_FOLDER_FULLPASS
    End If
    
    Set objFso = Nothing
    Set objFi = Nothing
    
    '「ワークコピー確認」ポップアップ画面表示
    iResponse = MsgBox(sInputPass & " の全てのファイルを、" _
                       & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                       & "「ワーク」フォルダにコピーします。 " _
                       & "よろしいですか？", _
                       vbYesNo + vbExclamation, _
                       "媒体→ワーク コピー")
    If iResponse = vbNo Then
    '[いいえ] ボタンを選択:何もしない。
    '但し、圧縮ファイル指定の時は、解凍用一時フォルダを削除する。
        If sFlag = "LZH" Then
            psDeleteFolder MELTED_FOLDER_FULLPASS
        End If
        Exit Sub
    End If
    
' EG20 V6.9.0.1 【量産対応：種別チェック機能追加】DEL START
'    '外部入力プロ判正当性チェック
'    If sFlag = "STD" Then
'       '媒体→ワーク コピー時
'       bRet = pfInstallSeitouseiChck(sInputPass)
'    Else
'       '圧縮ファイル→ワーク コピー時
'       bRet = pfInstallSeitouseiChck(MELTED_FOLDER_FULLPASS & "\")
'    End If
'    If bRet = False Then
'       Exit Sub
'    End If
' EG20 V6.9.0.1 【量産対応：種別チェック機能追加】DEL END
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    ' /////////////////////////////////////////////////
    ' // コーナ単位での処理
    For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
        If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
            'For iSelctLoop = 0 To 8                    'EG20 V30.1.0.1 DEL
            For iSelctLoop = 0 To DATA_KIND_MAX - 1     'EG20 V30.1.0.1 ADD
                If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                    FolderSyubetu = iSelctLoop
                    bStatus = sFDInstall2(iCornerLoop, sFlag, sInputFolder)
                    Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
' EG20 V3.0.0.2追加開始
                    If bStatus = True Then
                        '監視設定エリア「プログラム判定異常状態（ワーク）」の状態を取得する
                        lngPgmHanteiStsWork = pfGetKansiSet(IdKansiSet.PG_HANTEI_ERR_STS_WORK)

                        '「プログラム判定異常状態（ワーク）」（正常）
                        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_ERR_STS_WORK, ErrCode.Normal)
                        
                        '変化があった場合、「状態変化通知」を送信する
                        If lngPgmHanteiStsWork <> ErrCode.Normal Then
                            Call sSendMailStsChgInf(MailSts.stsNormal, lngPgmHanteiStsWork)
                        End If
                    
' EG20 V5.8.0.1削除開始
'                        ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'                        Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
                        ' 運改状態更新
                        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1追加終了
'                        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iCornerLoop + 1)   ' EG20 V5.6.0.1追加           ' EG20 V5.11.0.1削除
                        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iCornerLoop + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1追加
                    Else
                        Call pubfuncErrorOccur(MN_FOLD_WRK)
                    End If
' EG20 V3.0.0.2追加終了
                End If
            Next
        End If
    Next
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '圧縮ファイル指定の時は、解凍用一時フォルダを削除する。(使用済みのため)
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
    
    Exit Sub    '処理を終了する

ErrorHandler:   ' エラー処理。
    Set objFso = Nothing
    Set objFi = Nothing
    Call pubfuncErrorOccur(MN_FOLD_WRK)             ' EG20 V3.0.0.2追加

' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    '圧縮ファイル指定の時は、解凍用一時フォルダを削除する。
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sFDInstall2
'//  機能名称    ： 媒体インストール処理
'//  機能概要    ： インストール媒体ファイルを、ワークフォルダにコピーする。
'//
'//                 型        名称      意味
'//  引数        ： String    sFlag     処理種別
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応【種別チェック機能追加】【TOMAS用領域コピー対応】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function sFDInstall2(nCorner As Integer, sFlag As String, sInputFolder As String) As Boolean

    Dim MyName As String            'ファイルフルパス名
    Dim sSrcFileName As String      'コピー元ファイル名
    Dim sDstFileName As String      'コピー先ファイル名
    Dim lngErrCode As Long          'エラーコード
    Dim sChkName As String          'チェックファイル
    Dim szTargetFolder As String    ' 属性変更先フォルダ名            ' EG20 V5.8.0.1追加
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    Dim objFi As File                    'ファイルオブジェクト
    Dim bRet As Boolean             '正当性チェック戻り値             ' EG20 V6.9.0.1ADD
    
    Dim sTomasPath As String        ' TOMAS用領域ファイルパス
    
    On Error GoTo ErrorHandler      'エラーハンドルの登録

    
    sFDInstall2 = True

' EG20 V6.9.0.1 【量産対応：種別チェック機能追加】ADD START
    bRet = pfInstallSeitouseiChck(sInputFolder & "\")
    If bRet = False Then
        Set objFso = Nothing
        Set objFi = Nothing
        sFDInstall2 = False
        Exit Function           '処理を終了する
    End If
' EG20 V6.9.0.1 【量産対応：種別チェック機能追加】ADD END

' EG20 V5.8.0.1追加開始
    szTargetFolder = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu)
' EG20 V5.8.0.1追加終了
    'バージョンチェックファイル有無チェックを行う。
    sChkName = fSelectFile
    
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
   
    If objFso.FileExists(gstrMyPath & sChkName) = True Then
        '指定ファイルが存在する
        sChkName = objFso.GetFileName(gstrMyPath & sChkName)
        Kill gstrMyPath & sChkName
    Else
        sChkName = ""
    End If
    
    '指定フォルダ内のファイルを、全て「ワーク」フォルダにコピーする。
    For Each objFi In objFso.GetFolder(sInputFolder).files   'ループを開始
        If objFso.FileExists(objFi.Path) = True Then  'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            '媒体内ファイル名を作成
            sSrcFileName = sInputFolder & "\" & MyName
            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
                'ワークフォルダ内ファイル名を作成する
                sDstFileName = gstrMyPath & MyName
                '媒体内のファイルをワークフォルダにコピーする
                FileCopy sSrcFileName, sDstFileName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    
'    '圧縮ファイル指定の時は、解凍用一時フォルダを削除する。(使用済みのため)
'    If sFlag = "LZH" Then
'        psDeleteFolder MELTED_FOLDER_FULLPASS
'    End If

' EG20 V5.8.0.1追加開始
    ' 属性変更処理
    dllChangeAttributeContents (szTargetFolder)
' EG20 V5.8.0.1追加終了
    
' EG20 V6.9.0.1 【量産対応：TOMAS用領域コピー対応】ADD START
    ' 処理すべき対象がコーナ1の場合
    ' TOMAS領域（N_GATE00）もN_GATE01の内容でコピー
    'If nCorner = 0 Then                            'EG20 V30.1.0.1 DEL
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
    'ワークコピーしようとするたびにそのコーナから00へコピーするため、先頭コーナの判定を削除
    'If nCorner = gintKansenFirstCornerIdx Then      'EG20 V30.1.0.1 ADD
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
        ' 削除先のフォルダ（TOMAS領域）を指定
        sTomasPath = PATH_GATE_EG20 & "00" & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
        
        ' TOMAS領域を削除
        If funcRemoveFile(sTomasPath) = False Then
           
            sFDInstall2 = False
            '「自改ﾊﾞｰｼﾞｮﾝ：TOMASﾌｫﾙﾀﾞﾌｧｲﾙ削除異常」ログ出力
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_DELETE_ERROR, lngErrCode)
        
            Exit Function '処理を終了する
        End If
        
        ' TOMAS領域へコピー
        If funcCopyFile(gstrMyPath, sTomasPath, lngErrCode) = False Then
            
            sFDInstall2 = False
            '「自改ﾊﾞｰｼﾞｮﾝ：TOMAS領域ｺﾋﾟｰ処理異常」ログ出力
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_COPY_ERROR, lngErrCode)
        
            Exit Function '処理を終了する
        End If
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
    'End If
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
' EG20 V6.9.0.1 【量産対応：TOMAS用領域コピー対応】ADD END
    
    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_OK, 0)
    
    Exit Function '処理を終了する

ErrorHandler:   ' エラー処理。
    Set objFso = Nothing
    Set objFi = Nothing
    
    sFDInstall2 = False

' EG20 V5.8.0.1追加開始
    ' 属性変更処理
    dllChangeAttributeContents (szTargetFolder)
' EG20 V5.8.0.1追加終了

    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： AddLstStatus
'//  機能名称    ： 処理結果リストへのメッセージ出力処理
'//  機能概要    ： 処理結果リストへのメッセージ出力処理
'//
'//                 型        名称      意味
'//  引数        ： Integer   nProc     処理番号（押下釦）
'//                 Integer   nCorner   コーナ番号（0～5）
'//                 Integer   nDataKind 処理種別（0～）
'//                 Boolean   bResult   処理結果（TRUE:正常、FALSE:異常）
'//
'//
'//                 型        値        意味
'//  戻り値      ： Boolean   bResult   処理結果（TRUE:正常、FALSE:異常）
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Sub AddMessageLstStatus(nCorner As Integer, _
                                      nDataKind As Integer, _
                                      bResult As Boolean)
    Dim szOutMessae As String       ' 出力メッセージ
    Dim szWorkMsg As String         ' 出力メッセージ（ワーク）
    
    On Error Resume Next

    szOutMessae = "コーナ" & Format(nCorner + 1, "00") & "：" & _
                    TitleBox(nDataKind) & "："
    
    If bResult = True Then
        szWorkMsg = "正常終了"
    Else
        szWorkMsg = "異常終了"
    End If
    
    szOutMessae = szOutMessae & szWorkMsg

    LstStatus.AddItem (szOutMessae)
    LstStatus.Selected(LstStatus.ListCount - 1) = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： pfInstallSeitouseiChck
'//  機能名称    ： 外部入力プログラム判定データ正当性チェック処理
'//  機能概要    ： 外部入力プログラム判定データ正当性チェック処理を行う。
'//
'//                 型        名称      意味
'//  引数        ： なし
'//
'//                 型        値        意味
'//  戻り値      ： Boolean   bResult   処理結果（TRUE:正常、FALSE:異常）
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応【種別チェック機能追加】
'//     REVISIONS :(EG20 V6.11.0.1) 2013-03-27 REVISED BY  [TCC] H.Kondoh
'//                 媒体投入機能変更対応
'//                   種別０の場合も異常とするように変更
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function pfInstallSeitouseiChck(sInputPass As String) As Boolean
    Dim myLen As Long                        '文字列の長さ
    Dim lngSumRet As Long
    Dim i As Integer
    Dim bRet As Boolean
    Dim lngCnt As Long
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
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    Else
     '「自改ﾊﾞｰｼﾞｮﾝ：ファイルリスト無し」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      pfInstallSeitouseiChck = False
      Set objFso = Nothing
      Exit Function
    End If

   '｢ワーク｣フォルダからファイルリストを取得する
    bRet = fReadFileList(sInputPass & MN_FILELIST)

    'サム値チェック
    For lngCnt = 0 To UBound(FileList) - 1
        If pfFileSumChk(sInputPass & FileList(lngCnt), lngSumRet) <> True Then
            pfInstallSeitouseiChck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    'ファイル数最大チェック
    If UBound(FileList) > FILECNT_MAX Then
      pfInstallSeitouseiChck = False

      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)

      Exit Function
    End If
    For i = 0 To UBound(FileList) - 1
       '取得ファイル名のサイズを取得
       myLen = LenB(StrConv(Trim(FileList(i)), vbFromUnicode))                                              '半角換算のバイト数を取得
       If FILE_NAME_MAX_SIZE < myLen Then
          '13バイト以上の場合
           bRet = False
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
           Exit For
       End If
    Next

' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD START
    If bRet = False Then
        pfInstallSeitouseiChck = bRet
        Exit Function
    End If

    For i = 0 To UBound(FileList) - 1
        ' ファイルリスト内の種別を抽出
        'intCheckKind = CInt(Left$(FileListType(i), 1))         'EG20 V30.1.0.1 DEL
        intCheckKind = Asc(Left$(FileListType(i), 1))   'EG20 V30.1.0.1 ADD
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
            ' エラーログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_PRGKIND_ERROR, 0)
            Exit For
        End If
    Next
' EG20 V6.9.0.1【量産対応：種別チェック機能追加】ADD END

    pfInstallSeitouseiChck = bRet
Exit Function

FileGetError:
    pfInstallSeitouseiChck = False
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： fNewVersion
'//  機能名称    ： 最新バージョン処理
'//  機能概要    ： 最新(ワーク)バージョンを、実行(実行)バージョンに登録
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V5.0.2.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【TR-51修正対応】
'//  REVISIONS   ：(EG20 V30.3.0.1) 2014-11-11  CODED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function fNewVersion(nCorner As Integer) As Boolean
    Dim bRet As Boolean                      '戻り値
    Dim sSrcFileName            As String    'ワークフォルダ内ファイルリスト
    Dim lngErrCode As Long                   'エラーコード
    Dim iKansiAplChk As Integer              'アプリ起動チェック戻り値　'V1.6.0.1 ADD
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD
    
    On Error Resume Next
    
    '｢ワーク｣フォルダのファイルリストを検索する
    'ワークフォルダ内ファイル名を作成
    'ワークフォルダ内のディレクトリの名前を表示します。（コーナ単位）
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\" & MN_FILELIST
    
    'ファイルの検索をする
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    Else
        'ファイルが存在しない
        '「自改ﾊﾞｰｼﾞｮﾝ：ファイルリスト無し」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

        fNewVersion = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function
    End If
    
    bRet = pfSeitouseiChck(nCorner)    'V1.4.0.1　ADD
    '｢ワーク｣フォルダからファイルリストより、登録ファイル数をカウントする
    If bRet = True Then
       bRet = fReadFileList(sSrcFileName)
    End If

    If bRet = True Then
        '｢旧｣フォルダ内のファイルを全て削除する
        If sOldFolderRemove(nCorner) <> True Then
'            Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加    EG20 V3.6.0.1削除
            Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1追加
            fNewVersion = False
            Exit Function
        End If

        '｢旧｣フォルダ内のファイルを｢実行｣フォルダの内容に置換える
        If sCopyNOWtoOLD(nCorner) <> True Then
'            Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加    EG20 V3.6.0.1削除
            Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1追加
            fNewVersion = False
            Exit Function
        End If

        '｢実行｣フォルダ内のファイルを｢ワーク｣フォルダの内容に置換える
        If sCopyWRKtoNOW(nCorner) <> True Then
            Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加
            fNewVersion = False
            Exit Function
        End If
    
        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
'        '自改バージョン情報更新要求メールを管理プロセスへ送信する。
'        '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
'        iKansiAplChk = CheckAppStart(PROC_KANRI)
'        If iKansiAplChk <> 0 Then
'            '監視盤起動時：管理プロセスに自改バージョン情報更新要求メールを送信する。
''            psVersionUpdateReqest (ML_REQUEST_EG20GATE)
'            frmVerUpdateIkkatsu.Show vbModal
'        Else
'            '監視盤未起動時：改札機バージョン更新処理結果に正常を設定する。
'            gintGateVerInfUpdRes = MailSts.stsNormal
'        End If
'
'        '改札機バージョン更新処理結果
'        If gintGateVerInfUpdRes = MailSts.stsNormal Then
'            '正常
'            fNewVersion = True
'        Else
'            '異常
'            fNewVersion = False
'        End If
        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
        
        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
        If nCorner <> NGATE_00 Then
            '自改バージョン情報更新要求メールを管理プロセスへ送信する。
            '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
            iKansiAplChk = CheckAppStart(PROC_KANRI)
            If iKansiAplChk <> 0 Then
                '監視盤起動時：管理プロセスに自改バージョン情報更新要求メールを送信する。
    '            psVersionUpdateReqest (ML_REQUEST_EG20GATE)
                frmVerUpdateIkkatsu.Show vbModal
            Else
                '監視盤未起動時：改札機バージョン更新処理結果に正常を設定する。
                gintGateVerInfUpdRes = MailSts.stsNormal
            End If
        
            '改札機バージョン更新処理結果
            If gintGateVerInfUpdRes = MailSts.stsNormal Then
                '正常
                fNewVersion = True
            Else
                '異常
                fNewVersion = False
            End If
        Else
            fNewVersion = True
        End If
        'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END
  
'        fNewVersion = True             ' EG20 V5.0.2.1削除
    Else
        fNewVersion = False
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： fOldVersion
'//  機能名称    ： 旧バージョン処理
'//  機能概要    ： 一世代前のバージョンを実行(実行)バージョンに返す。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V5.0.2.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【TR-51修正対応】
'//  REVISIONS   ：(EG20 V30.3.0.1) 2014-11-11  CODED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function fOldVersion(nCorner As Integer) As Boolean
    Dim bRet As Boolean                     '戻り値
    Dim sSrcFileName            As String   '旧フォルダ内ファイルリスト
    Dim lngErrCode              As Long     'エラーコード
    Dim iKansiAplChk As Integer              'アプリ起動チェック戻り値　'V1.6.0.1 ADD

    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD
    
    On Error Resume Next
 
   '旧フォルダ内のファイルリストを検索する。
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    Else                                'ファイルが存在しない
        '「自改ﾊﾞｰｼﾞｮﾝ：ファイルリスト無し」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)
 
        fOldVersion = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '処理を終了する
    End If
    
    '｢旧｣フォルダからファイルリストを取得する
    bRet = fReadFileList(sSrcFileName)

' EG20 V3.6.0.1 【統合TR-No.260】追加開始
    bRet = fDataFileCheck(sSrcFileName)
    If bRet = False Then
       Call pubfuncErrorOccur(MN_FOLD_OLD)
       fOldVersion = False
       Exit Function
    End If
' EG20 V3.6.0.1 【統合TR-No.260】追加終了

  ' EG20 V3.0.0.2追加開始
    ' 改札機共通判定処理
    bRet = pubfuncCommonGateCheck(nCorner, MN_FOLD_OLD)
    If bRet = False Then
        fOldVersion = False
       Exit Function
    End If
  ' EG20 V3.0.0.2追加終了

    '｢実行｣フォルダ内のファイルを全て削除する
    If sNowFolderRemove(nCorner) <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.6.0.1追加
        fOldVersion = False
        Exit Function
    End If
    
    '｢旧｣フォルダ内のファイルを｢実行｣フォルダの内容に置換える
    If sCopyOLDtoNOW(nCorner) <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.6.0.1追加
        fOldVersion = False
        Exit Function
    End If
    
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
'    '自改バージョン情報更新要求メールを管理プロセスへ送信する。
'    '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
'     iKansiAplChk = CheckAppStart(PROC_KANRI)
'     If iKansiAplChk <> 0 Then
'        '監視盤起動時：管理プロセスに自改バージョン情報更新要求メールを送信する。
''         psVersionUpdateReqest (ML_REQUEST_EG20GATE)
'        frmVerUpdateIkkatsu.Show vbModal
'    Else
'        '監視盤未起動時：改札機バージョン更新処理結果に正常を設定する。
'        gintGateVerInfUpdRes = MailSts.stsNormal
'    End If
'
'     '改札機バージョン更新処理異常
'    If gintGateVerInfUpdRes = MailSts.stsNormal Then
'        '正常
'        fOldVersion = True
'    Else
'        '異常
'        fOldVersion = False
'    End If
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
    
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD START
    If nCorner <> NGATE_00 Then
        '自改バージョン情報更新要求メールを管理プロセスへ送信する。
        '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
         iKansiAplChk = CheckAppStart(PROC_KANRI)
         If iKansiAplChk <> 0 Then
            '監視盤起動時：管理プロセスに自改バージョン情報更新要求メールを送信する。
    '         psVersionUpdateReqest (ML_REQUEST_EG20GATE)
            frmVerUpdateIkkatsu.Show vbModal
        Else
            '監視盤未起動時：改札機バージョン更新処理結果に正常を設定する。
            gintGateVerInfUpdRes = MailSts.stsNormal
        End If
         
         '改札機バージョン更新処理異常
        If gintGateVerInfUpdRes = MailSts.stsNormal Then
            '正常
            fOldVersion = True
        Else
            '異常
            fOldVersion = False
        End If
    Else
        fOldVersion = True
    End If
    'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 ADD END

'    fOldVersion = True                 ' EG20 V5.0.2.1削除
End Function


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sOldFolderRemove
'//  機能名称    ： 旧フォルダ内ファイル削除処理
'//  機能概要    ： 旧フォルダ内のファイルを削除する。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function sOldFolderRemove(nCorner As Integer) As Boolean
    Dim MyName As String                'ファイル名
    Dim lngErrCode As Long              'エラーコード
    
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    
    On Error GoTo ErrorHandler          'エラーハンドルの登録
   
   '戻り値初期化
    sOldFolderRemove = True
 
    '「実行」フォルダ内のディレクトリの名前を表示します。
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\"
    
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
    
    Exit Function           '処理を終了する

ErrorHandler:   ' エラー処理ルーチン。
    '「自改ﾊﾞｰｼﾞｮﾝ：旧フォルダﾌｧｲﾙ削除異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLDFILE_DELETE_ERROR, lngErrCode)

    sOldFolderRemove = False
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sNowFolderRemove
'//  機能名称    ： 実行フォルダ内のファイル削除処理
'//  機能概要    ： 実行フォルダ内のファイルを削除する。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function sNowFolderRemove(nCorner As Integer) As Boolean
    Dim MyName As String                'ファイル名
    Dim lngErrCode As Long              'エラーコード
    
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト

    On Error GoTo ErrorHandler          'エラーハンドルの登録

    '初期値設定
    sNowFolderRemove = True
    
    '「実行」フォルダ内のディレクトリの名前を表示します。
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\"
    
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

    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    '「自改ﾊﾞｰｼﾞｮﾝ：実行フォルダﾌｧｲﾙ削除異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOWFILE_DELETE_ERROR, lngErrCode)

    sNowFolderRemove = False
    
    Set objFso = Nothing
    Set objFi = Nothing
    
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sCopyNOWtoOLD
'//  機能名称    ： 実行バージョン保存処理
'//  機能概要    ： 実行フォルダ内のファイルを、旧フォルダにコピーする。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function sCopyNOWtoOLD(nCorner As Integer) As Boolean
    Dim MyName As String                'ファイル名
    Dim sSrcFileName As String          'コピー元ファイルのフルパス名
    Dim sDstFileName As String          'コピー先ファイルのフルパス名
    
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    
    On Error GoTo ErrorHandler              'エラーハンドル設定
  
    '戻り値初期化
    sCopyNOWtoOLD = True
   
    '実行フォルダ内のディレクトリの名前を表示します。
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\"
    
    For Each objFi In objFso.GetFolder(gstrMyPath).files  'ループを開始
        If objFso.FileExists(objFi.Path) = True Then      'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます｡
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                '実行フォルダ内ファイル名を作成する
                sSrcFileName = gstrMyPath & MyName

                '旧フォルダ内ファイル名を作成する
                sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MyName

                'ワークフォルダ内のファイルを実行フォルダにコピーする
                FileCopy sSrcFileName, sDstFileName

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    
    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    
    sCopyNOWtoOLD = False
    
    Set objFso = Nothing
    Set objFi = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sCopyWRKtoNOW
'//  機能名称    ： 最新バージョンコピー
'//  機能概要    ： ワークフォルダ内のファイルを、実行フォルダにコピー。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応（PASSINFコピー対応）
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function sCopyWRKtoNOW(nCorner As Integer) As Boolean
    
    Dim i As Integer                'カウンタ
    Dim sSrcFileName As String      'コピー元ファイル名
    Dim sDstFileName As String      'コピー先ファイル名
    Dim bError As Boolean           'フラグ
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler      'エラーハンドルの登録
  
    '戻り値初期化
    sCopyWRKtoNOW = True
    
    '****************************
    '* ファイルリストをコピーする *
    '****************************
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\" & MN_FILELIST
                                    'ワークフォルダ内ファイル名を作成する
    sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '実行フォルダ内ファイル名を作成する
    If objFso.FileExists(sSrcFileName) = True Then     'ファイルの検索をする   'V1.20.0.1 ADD
        'ファイルリストを「ワーク」フォルダから「実行」フォルダにコピーする
        FileCopy sSrcFileName, sDstFileName
    Else                                'ファイルが存在しない
        sCopyWRKtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '処理を終了する
    End If

    bError = False                  'エラーフラグを「偽」にする
    For i = 0 To UBound(FileList) - 1
                                    'ファイルリスト一覧数分繰り返す
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\" & FileList(i)
                                    'ワークフォルダ内ファイル名を作成する
        sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)
                                    '実行フォルダ内ファイル名を作成する

        'ワークフォルダ内のファイルを実行フォルダにコピーする
        If objFso.FileExists(sSrcFileName) = True Then   'ファイルの検索をする   'V1.20.0.1 ADD
            'ファイルを「ワーク」フォルダから「実行」フォルダにコピーする
            FileCopy sSrcFileName, sDstFileName
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2追加開始
    If pfuncCopyPASSINF(nCorner, MN_FOLD_WRK) = False Then
        sCopyWRKtoNOW = False
    End If
' EG20 V3.0.0.2追加終了
    
    Exit Function                           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    sCopyWRKtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sCopyOLDtoNOW
'//  機能名称    ： 旧バージョンに戻す処理
'//  機能概要    ： 旧フォルダ内のファイルを、実行フォルダにコピーする。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応（PASSINFコピー対応）
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function sCopyOLDtoNOW(nCorner As Integer) As Boolean
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
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    'ワークフォルダ内ファイル名を作成する
    sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '実行フォルダ内ファイル名を作成する
    
    If objFso.FileExists(sSrcFileName) = True Then 'ファイルの検索をする   'V1.20.0.1 ADD
        'ファイルリストを「旧」フォルダから「実行」フォルダにコピーする
        FileCopy sSrcFileName, sDstFileName
    Else
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '処理を終了する
    End If

    bError = False                  'エラーフラグを「偽」にする
    For i = 0 To UBound(FileList) - 1
                                    'ファイルリスト数分繰り返す
        '旧フォルダ内ファイル名を作成する
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)

        '実行フォルダ内ファイル名を作成する
        sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)

        '旧フォルダ内のファイルを実行フォルダにコピーする
        If objFso.FileExists(sSrcFileName) = True Then 'ファイルの検索をする   'V1.20.0.1 ADD
            'ファイルを「旧」フォルダから「実行」フォルダにコピーする
            FileCopy sSrcFileName, sDstFileName
        Else                                'ファイルが存在しない
            bError = True                   'エラーフラグを「真」にする
        End If
    Next
    If bError = True Then
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function
    End If

    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2追加開始
    If pfuncCopyPASSINF(nCorner, MN_FOLD_OLD) = False Then
        sCopyOLDtoNOW = False
    End If
' EG20 V3.0.0.2追加終了
    
    Exit Function       '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    sCopyOLDtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： pfSeitouseiChck
'//  機能名称    ： プログラム判定データ正当性チェック処理
'//  機能概要    ： プログラム判定データ正当性チェック処理を行う。
'//
'//                 型        名称      意味
'//  引数        ： Integer   nCorner   コーナ番号（0～5）
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function pfSeitouseiChck(nCorner As Integer) As Boolean
    Dim bRet As Boolean
    
    Dim szTargetFolder As String            ' 対象フォルダ
    
    On Error Resume Next
    
    pfSeitouseiChck = True
    
    szTargetFolder = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
    '********************************
    '*プロ判正当性チェック
    '********************************
    '自改プログラム判定データ正当性チェックを行う(対象ファイル：HAN_KUKA.KUK)
    bRet = fDataFileCheck(szTargetFolder & MN_FILELIST)
    If bRet = False Then
'       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加     EG20 V3.6.0.1削除
       Call pubfuncErrorOccur(MN_FOLD_WRK)          ' EG20 V3.6.0.1追加
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2追加開始
    ' 改札機共通判定処理
    bRet = pubfuncCommonGateCheck(nCorner, MN_FOLD_WRK)
    If bRet = False Then
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2追加終了

    '機種正当性チェック(対象ファイル：XX_GATEY.VEF　XX:ユーザー名　Y：データ種別)
    bRet = fKishuCheck(szTargetFolder)
    If bRet = False Then
       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2追加
       pfSeitouseiChck = False
       Exit Function
    End If

    pfSeitouseiChck = bRet
Exit Function

FileGetError:
    pfSeitouseiChck = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： fDataFileCheck
'//  機能名称    ： 自改プログラム判定データ正当性チェック処理
'//  機能概要    ： 対象となるHAN_KUKA.KUK有無チェックを行う。
'//
'//                 型        名称      意味
'//  引数        ： String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
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
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
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
    
    'ファイル機種正当性チェックを行う。
    For i = 0 To iCnt - 1
         'チェック対象ファイルパス作成
        sChkFileName = sFileList & FileList(i)
    
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
            
            '文字抽出
            'sChkData = Left(sKisyu, Len(EG20_JIKAI_KISHU))      'EG20 V30.1.0.1 DEL
            sChkData = Left(sKisyu, Len(EG30_JIKAI_KISHU))       'EG20 V30.1.0.1 DEL
            'If EG20_JIKAI_KISHU = sChkData Then        'EG20 V30.1.0.1 DEL
            If EG30_JIKAI_KISHU = sChkData Then         'EG20 V30.1.0.1 ADD
                bRet = True  '機種正当性：正常
            Else
                bRet = False '機種正当性：異常
                fKishuCheck = bRet
                Set objFso = Nothing    'V1.20.0.1 ADD
                Exit Function
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：frmJVer.sSetChkFile流用
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：frmJVer.fSelectFile流用
'///////////////////////////////////////////////////////////////////
Private Function fSelectFile() As String
 
    fSelectFile = ""
    'バージョンチェックファイル名を設定する。
    Select Case FolderSyubetu
        'EG20 V30.1.0.1 DEL START
'       Case 0 '判定CPU-Pro
'            fSelectFile = EG20_HANTEI_CPU_CHK_FILE
'
'       Case 1 'メインCPU-Pro
'            fSelectFile = EG20_MAIN_CPU_CHK_FILE
'
'       Case 2 'サブCPU1-Pro
'            fSelectFile = EG20_SUB_CPU1_CHK_FILE
'
'       Case 3 'サブCPU2-Pro
'            fSelectFile = EG20_SUB_CPU2_CHK_FILE
'
'       Case 4 'サブCPU3-Pro
'            fSelectFile = EG20_SUB_CPU3_CHK_FILE
'
'       Case 5 'メインCPU-OS
'            fSelectFile = EG20_MAIN_OS_CHK_FILE
       'EG20 V30.1.0.1 DEL END
       'EG20 V30.1.0.1 ADD START
       Case 0 '判定CPU-Pro
            fSelectFile = EG30_HANTEI_CPU_CHK_FILE
       
       Case 1 'メインCPU-Pro
            fSelectFile = EG30_MAIN_CPU_CHK_FILE
       
       Case 2 'サブCPU-Pro
            fSelectFile = EG30_SUB_CPU_CHK_FILE
       
       Case 3 'メインCPU-OS
            fSelectFile = EG30_MAIN_OS_CHK_FILE
       'EG20 V30.1.0.1 ADD END
     
     End Select


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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：frmJVer.fReadFileList流用
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
            'EG20 V30.1.0.1 DEL START
'            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
'            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.1.0.1 ADD START
            'ファイル種別は大文字に変換せず、ファイル名だけを大文字に変換するようにする。（今までは種別が数字だったから問題なかった）
            FileListType(iListCnt - 1) = Trim$(Left$(sFileName, 18))
            FileList(iListCnt - 1) = UCase(Mid$(FileListType(iListCnt - 1), 3, 16))
            'EG20 V30.1.0.1 ADD　END
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_02_06】
'//  REVISIONS   ：(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：frmJVer.fHankukaChck流用
'///////////////////////////////////////////////////////////////////
Private Function fHankukaChck(sFilePath As String) As Boolean
    Dim iFileNumber As Integer           'ファイル番号
    Dim i As Integer
    Dim lSts As Long
    Dim sKeyName As String
    Dim lPos As Long                     'バージョン情報格納位置
    Dim lLen As Long                     'ファイルサイズ
    'Dim uFooter As MN_FOOT          'フッタ情報格納エリア      'EG20 V30.1.0.1 DEL
    Dim uFooter As MN_KAN_FOOT          'フッタ情報格納エリア   'EG20 V30.1.0.1 ADD
    Dim sDateTime As String
    Dim j As Integer
    Dim lngErrCode As Long          'エラーコード
    Dim uHeder As HAN_KUKA_KUK_HEADER       'ヘッダ情報格納エリア
    Dim sGetInfo As String * MAX_PATH_SIZE  'INIファイル取得用
    Dim sChkFileData As String
    Dim iMojisu As Integer
    
    Dim bChkSts As Boolean              'チェック結果フラグ
    Dim sChkData As String              '比較文字抽出
    
   '初期化：正常(ブランク）
    sNGSts = ""
    sNGKoumoku = ""
    'V1.4.0.1 ADD END
    Dim oFs As New FileSystemObject 'V2.5.0.1 ADD
    
    fHankukaChck = False
    
 'ファイル有無チェックを行う。
 If oFs.FileExists(sFilePath) = False Then
    'ファイルが無ければ正当性チェックを行わない。
    fHankukaChck = True
    Set oFs = Nothing
    Exit Function
 End If

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
      'EG20 V30.1.0.1 DEL START
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 DEL END
      'EG20 V30.1.0.1 ADD START
      lSts = GetPrivateProfileString(EG30_HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 ADD END
      If lSts = False Then
       
      Else
        HAN_KUKA_DATA.sHederKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      'ヘッダ：期待値ファイル名取得
      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
      'EG20 V30.1.0.1 DEL START
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 DEL END
      'EG20 V30.1.0.1 ADD START
      lSts = GetPrivateProfileString(EG30_HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 ADD END
      If lSts = False Then
        
      Else
         HAN_KUKA_DATA.sHederFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      'EG20 V30.1.0.1 DEL START（新幹線はフッタ無し）
      'フッタ：期待値機種名取得
'      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'
'      Else
'        HAN_KUKA_DATA.sFotterKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
'      End If
'      'フッタ：期待値ファイル名取得
'      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'
'      Else
'        HAN_KUKA_DATA.sFotterFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
'      End If
       'EG20 V30.1.0.1 DEL END
    Next i
    'V1.4.0.1 ADD END

    On Error GoTo ErrorHandler      'エラーハンドル設定
    
    'HAN_KUKA.KUKファイルサイズ取得
    lLen = FileLen(sFilePath)
    
    '未使用のファイル番号を取得する
    iFileNumber = FreeFile
    
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
   
   'ヘッダ情報：機種名チェック
   iMojisu = InStr(uHeder.sKisyuName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sKisyuName, 1)
   Else
     sChkFileData = Mid(uHeder.sKisyuName, 1, iMojisu)
   End If
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

   'ヘッダ情報：ファイル名チェック
   iMojisu = InStr(uHeder.sProgrumName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sProgrumName, 1)
   Else
     sChkFileData = Mid(uHeder.sProgrumName, 1, iMojisu)
   End If

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
    
   'EG20 V30.1.0.1 DEL START
'   'フッタ情報：機種名チェック
'   iMojisu = InStr(uFooter.sKisyu, " ") - 1
'   If iMojisu < 0 Then
'     sChkFileData = Mid(uFooter.sKisyu, 1)
'   Else
'     sChkFileData = Mid(uFooter.sKisyu, 1, iMojisu)
'   End If
'
'   bChkSts = False
'    For i = 0 To INI_MAX - 1
'       If HAN_KUKA_DATA.sFotterKisyu(i) <> "" Then
'          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterKisyu(i)))
'          If sChkData = HAN_KUKA_DATA.sFotterKisyu(i) Then
'             bChkSts = True
'           Exit For
'          End If
'      End If
'    Next
'    'チェック結果フラグ判定
'    If bChkSts = False Then
'       '機種名期待値全不一致：
'        sNGSts = ERROR_FOTTER
'        sNGKoumoku = KISHU_NAME_ERROR
'         GoTo ErrorHandler
'    End If
'
'   'フッタ情報：ファイル名チェック
'   iMojisu = InStr(uFooter.sFileName, " ") - 1
'   If iMojisu < 0 Then
'     sChkFileData = Mid(uFooter.sFileName, 1)
'   Else
'     sChkFileData = Mid(uFooter.sFileName, 1, iMojisu)
'   End If
'
'   bChkSts = False
'    For i = 0 To INI_MAX - 1
'       If HAN_KUKA_DATA.sFotterFile(i) <> "" Then
'          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterFile(i)))
'          If sChkData = HAN_KUKA_DATA.sFotterFile(i) Then
'             bChkSts = True
'           Exit For
'          End If
'      End If
'    Next
'    'チェック結果フラグ判定
'    If bChkSts = False Then
'       '機種名期待値全不一致：
'        sNGSts = ERROR_FOTTER
'        sNGKoumoku = FILE_NAME_ERRORE
'         GoTo ErrorHandler
'    End If
    'EG20 V30.1.0.1 DEL END
      
    'フッタ情報：作成日付が数値かどうか
     sDateTime = ""
     For j = 0 To 3
         sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
     Next
    'sDateTime = sDateTime & " " 'V1.4.0.1 DEL
     For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_FOTTER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    'EG20 V30.1.0.1 DEL START 新幹線のフッタ情報にはバージョンは存在しない
'    'バージョン値チェック
'    'フッタ情報：バージョン値が数値かどうか
'    If IsNumeric(uFooter.sVersion) = False Then
'       sNGSts = ERROR_FOTTER
'       sNGKoumoku = VERSION_ERROR
'       GoTo ErrorHandler
'       Exit Function
'    End If
'    'V1.4.0.1 ADD END
    'EG20 V30.1.0.1 DEL END
    
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
        AppActivate frmKansenGateVerUpdate.Caption, False
        pfFormActive (frmKansenGateVerUpdate.hwnd)
    End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pubfuncCommonGateCheck
'//  機能名称  : 改札機共通判定処理
'//  機能概要  : サム値チェック、ファイル数最大チェックの実行
'//
'//              型        名称             意味
'//  引数      ：Integer   nCorner   コーナ番号（0～5）
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
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pubfuncCommonGateCheck(nCorner As Integer, nKind As Integer) As Boolean

    Dim lngSumRet As Long
    Dim lngCnt As Long
    Dim lngFileListCnt As Long               'ファイルリスト数
    Dim i As Integer
    Dim strWork     As String                '作業エリア
    Dim iFileNumber As Integer               '未使用ファイル番号
    Dim bRet As Boolean
    Dim sGetFileListName As String           'ファイルリスト内記載ファイル名
    Dim myLen As Long                        '文字列の長さ
    Dim sSrcFileName            As String    'ワークフォルダ内ファイルリスト
    Dim lTotalCount As Long                  ' 結果件数

    Dim lngPgmHanteiRcvErrSts   As Long     'プログラム判定受信異常状態
    Dim lngPgmHanteiSndErrSts   As Long     'プログラム判定配信異常状態
    Dim lngPgmHanteiErrSts      As Long     'プログラム判定異常状態（実行）
    Dim lngPgmHanteiErrStsOld   As Long     'プログラム判定異常状態（旧）
    Dim lngPgmHanteiElseErrSts  As Long     'プログラム判定その他異常状態

    
    On Error Resume Next

    ' /////////////////////////////////////////////////////
    ' // サム値チェック
    For lngCnt = 0 To UBound(FileList) - 1
        
        '
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(nKind, FolderSyubetu) & "\" & FileList(lngCnt)
        If pfFileSumChk(sSrcFileName, lngSumRet) <> True Then
            
            '「プログラム判定受信異常状態」取得
            lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
        
            '監視設定エリア「プログラム判定受信異常状態」を更新
            Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_SumChk)
                    
            '監マプロセスに「状態変化通知」を送信
            If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_SumChk Then
                Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_SumChk)
            End If
            
' メッセージボックスは表示しない
'            'サム値異常
'            If lngSumRet = SUM_CHK.SumErr Then
'               MsgBox "サム値が異常です。" _
'                      & Chr(vbKeyReturn) & "データを確認してください。", _
'                      vbOKOnly + vbExclamation, _
'                      "自動改札機 バージョン管理"
'
'            'サム値異常以外異常
'            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
'               MsgBox "異常終了しました。", _
'                     vbOKOnly + vbExclamation, _
'                      "自動改札機 バージョン管理"
'            End If
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

' メッセージボックスは表示しない
'        MsgBox "ファイル数が上限を超えています。" _
'                & Chr(vbKeyReturn) & "データを確認してください。", _
'                vbOKOnly + vbExclamation, _
'                "自動改札機 バージョン管理"
        pubfuncCommonGateCheck = False

        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
        Exit Function
    End If

    'EG20 V30.1.0.1 DEL START 北陸新幹線では全種別の上限値を持っていないのでチェックは不要とする
'    ' /////////////////////////////////////////////////////
'    ' // 全ファイル数最大チェック（実行＋追加分）
'    bRet = True
'    lTotalCount = pfuncTotalListCount(nCorner)
'    lTotalCount = lTotalCount + UBound(FileList)
'    If lTotalCount > TOTALFILECNT_MAX Then
'        bRet = False
'    End If
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
'
'' メッセージボックスは表示しない
''        MsgBox "ファイル数が上限を超えています。" _
''                & Chr(vbKeyReturn) & "データを確認してください。", _
''                vbOKOnly + vbExclamation, _
''                "自動改札機 バージョン管理"
'        pubfuncCommonGateCheck = False
'
'        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
'        Exit Function
'    End If
    'EG20 V30.1.0.1 DEL END
    
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
'    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(nKind, FolderSyubetu) & "\" & MN_FILELIST
'    'ファイルリストをオープン。
'    Open sSrcFileName For Input As #iFileNumber
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
'' メッセージボックスは表示しない
''                MsgBox "ファイル名が異常です。" _
''                        & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
''                        vbOKOnly + vbExclamation, _
''                        "自動改札機 バージョン管理"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            'フォーマット異常
'            ElseIf " " <> Mid(strWork, 2, 1) Then
'              'ループ抜け
'' メッセージボックスは表示しない
''                MsgBox "ファイル名が異常です。" _
''                        & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
''                        vbOKOnly + vbExclamation, _
''                        "自動改札機 バージョン管理"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            ElseIf (InStr(strWork, ".") - 1) = -1 Then
'' メッセージボックスは表示しない
''                MsgBox "ファイル名が異常です。" _
''                        & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
''                        vbOKOnly + vbExclamation, _
''                        "自動改札機 バージョン管理"
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
'' メッセージボックスは表示しない
''                    MsgBox "ファイル名が異常です。" _
''                            & Chr(vbKeyReturn) & "ファイルリストを確認してください。", _
''                            vbOKOnly + vbExclamation, _
''                            "自動改札機 バージョン管理"
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
'//  引数      : Integer   nCorner   コーナ番号（0～5）
'//
'//              型        値               意味
'//  戻り値    : LONG      lResultCount     件数
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：frmJVer.fReadFileList流用
'///////////////////////////////////////////////////////////////////
Private Function pfuncTotalListCount(nCorner As Integer) As Long
    Dim lResultCount As Long                ' 結果件数
    Dim iLoop As Integer                    ' ループ
    
    Dim iFileNumber As Integer              'ファイル番号
    Dim sFileName As String                 'ファイル名
    Dim sSrcFileName As String              'ファイル名
    Dim iListCnt As Integer                 'ファイル格納数
    Dim objFso As New FileSystemObject      ' ファイルシステムオブジェクト

    On Error GoTo ErrorHandler      'エラーハンドル設定
    
    lResultCount = 0
    iFileNumber = FreeFile   '未使用のファイル番号を取得する
    For iLoop = 0 To 8
        
        iFileNumber = FreeFile   '未使用のファイル番号を取得する
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(0, iLoop) & "\" & MN_FILELIST
   
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

    On Error GoTo ErrorHandler              ' エラーハンドルの登録

    ' 対象が判定データの場合のみ処理を行う
    ' 上記に該当しない場合は正常終了
    If FolderSyubetu <> 0 Then
        pfuncCopyPASSINF = True
        Set objFso = Nothing
        Exit Function
    End If

    ' コピー元ファイル
    szSrcFile = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(nKind, 0) & "\" & "PASSINF"
    szDstFile = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, 0) & "\" & "PASSINF"

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


