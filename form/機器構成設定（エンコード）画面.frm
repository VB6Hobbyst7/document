VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmKikiDataSubGate 
   BorderStyle     =   0  'なし
   Caption         =   "機器構成設定"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmbCornerName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8280
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   15
      Top             =   700
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox cmbGoki 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9480
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "改札機画面へ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   7
      Left            =   7250
      TabIndex        =   13
      Top             =   7800
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "駅情報画面へ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   6
      Left            =   7250
      TabIndex        =   12
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   9120
      Top             =   2160
   End
   Begin VB.ComboBox CmbDummy 
      Height          =   345
      Left            =   4080
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   11
      Top             =   9720
      Width           =   2655
   End
   Begin VB.ListBox ListDummy 
      Height          =   510
      Left            =   120
      TabIndex        =   10
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   " 機器情報設定   画面へ戻る"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "媒体取外"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   5
      Left            =   4850
      TabIndex        =   7
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "一時保存"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "一時保存データ 取込"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   3
      Left            =   2450
      TabIndex        =   5
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "設定反映"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   1
      Left            =   2450
      TabIndex        =   3
      Top             =   7800
      Width           =   2175
   End
   Begin VB.TextBox txtDummy 
      Height          =   495
      IMEMode         =   3  'ｵﾌ固定
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   10425
      Width           =   3135
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "媒体入力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   4
      Left            =   4850
      TabIndex        =   2
      Top             =   7800
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   5730
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10107
      _Version        =   393216
      Rows            =   18
      Cols            =   8
      FixedCols       =   2
      RowHeightMin    =   350
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblEkiName 
      Caption         =   "駅名：○○○○○○○○○○○○○○○○○○○○"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   720
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "機器構成設定（エンコードコーナ号機情報定義）"
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
      TabIndex        =   8
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiDataSubGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：機器構成設定（エンコード）画面.frm
'//  パッケージ名：機器構成設定（エンコード）画面のフォームモジュール
'//
'//  概要：機器構成設定（エンコード）画面
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応
'//                 【HKRK_Kansi07_003_01】Sub_gate_kan.iniフォーマット見直し対応
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   'メイルタイマのインターバル値
Private Const TITOL_EKI_NAME = "駅名："                 '駅名タイトル       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
Private bScroll As Boolean

'設定反映フラグ
Private SetteiHaneiFlg As Boolean

'機器情報データ更新フラグ
Private KikiDataUpDateFlg As Boolean

'機器構成データ（エンコードコーナ号機情報定義）イメージファイル読取用の構造体
Private Type SUBGATE_IMAGE_FILE
    sType       As String                '種別
    sGoki       As String                '号機
    sNo         As String                '種別毎通番
    sCorner     As String                'コーナ        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    sTuuban     As String                '通番
    sKoumoku    As String                '項目
    sKubun      As String                '区分
    sSettei     As String                '設定値
    sSyosai     As String                '設定値詳細
End Type

'Private Const START_DATA_COL_INDEX = 1       '1行のデータ設定を開始するカラムインデックス  'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
Private Const START_DATA_COL_INDEX = 2       '1行のデータ設定を開始するカラムインデックス   'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
'Private Const MAX_DATA_COL_INDEX = 6         '1行の最大設定カラム数        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'Private Const MAX_DATA_COL_INDEX = 3         '1行の最大設定カラム数         ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加   'EG20 V30.1.0.1 DEL
'Private Const MAX_DATA_COL_INDEX = 6         '1行の最大設定カラム数         ' EG20 V30.1.0.1 ADD   'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
Private Const MAX_DATA_COL_INDEX = 7         '1行の最大設定カラム数         ' EG20 V30.1.0.1 ADD    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
Private Const FM_CORNER_COL = 2                 'FM券コーナ番号の列（COLプロパティ）
Private Const FM_GOKI_COL = 3                   'FM券号機番号の列（COLプロパティ）
Private Const SINKANSENIC_CORNER_COL = 4        '新幹線ICコーナ番号（COLプロパティ）
Private Const SINKANSENIC_GOKI_COL = 5          '新幹線IC号機番号（COLプロパティ）
Private Const NRZ_CORNER_COL = 6                'NRZ券コーナ番号(COLプロパティ)
Private Const NRZ_GOKI_COL = 7                  'NRZ券号機番号(COLプロパティ)
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】ADD END


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 機器情報設定（エンコード）画面(アクティブ時：イベントプロシージャ)
'//  機能概要  : 最前前表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'エラールーチンを宣言
    On Error Resume Next
    
    '自画面最前面表示処理を行う。
    pfFormActive (hwnd)
    
    'フォーカス位置を設定
    cmdCancel.SetFocus
    
    'タイマを起動する
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 機器情報設定（エンコード）画面(ディアクティブ時)
'//  機能概要  : メール受信用、タイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'タイマを停止する
    tmrMail.Enabled = False
End Sub
'EG20 V2.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 機器情報設定（エンコード）画面(ロード時：イベントプロシージャ)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-20  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード
    Dim iLoopCnt             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_START, 0)
    
    '----------------------------------------------------
    '画面初期値設定
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '機器情報データ更新フラグ設定（更新設定）
    KikiDataUpDateFlg = True
    
    '機器情報設定（エンコードコーナ号機情報定義）イメージファイル作成
    bRet = dllGetKikiIniData(2, 0, KIKI_DATA_SET_SUBGATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
    If bRet = False Then
        '機器情報設定（エンコードコーナ号機情報定義）イメージファイル削除
        Kill KIKI_DATA_SET_SUBGATE_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
    End If

    '号機コンボボックス初期値
    cmbGoki.Clear

    'For iLoopCnt = 0 To 15 'EG20 V30.1.0.1 DEL
    For iLoopCnt = 0 To 31  'EG20 V30.1.0.1 ADD
            cmbGoki.AddItem iLoopCnt + 1 & "号機"
    Next
    cmbGoki.ListIndex = 0

' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    'コーナ設定コンボボックスの初期化処理
    Call InitCornerComboBox
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

    '機器情報データ更新フラグ設定（通常設定）
    KikiDataUpDateFlg = False
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '設定反映フラグ（変更なし）
    SetteiHaneiFlg = False

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ処理（タイムアップ時：イベントプロシージャ）
'//  機能概要  : 汎用メイル受信処理を行う
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(V30.1.0.1) 2014-06-04 REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線改行対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    Dim intStatus As Integer         '受信メールチェック結果
    Dim iResponse As Integer
    
    On Error Resume Next
    
    'メールを受信する。
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
    '受信メールがあれば、メールＩＤ毎の処理をする。
        Select Case udtReadMail.udtlHeader.dwId        'メールＩＤ
            Case ML_ID_PROEND_ORD
                '「プロセス終了指示」を受信した場合、
                '「プロセス終了指示受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
                'プロセスの終了処理を行う
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示」を受信した場合
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '表示元画面（保守データ収集画面）をアクティブ表示する。
                'AppActivate frmInputMstData.Caption, False     'EG20 V30.1.0.1 DEL
                'EG20 V30.1.0.1 ADD START
                AppActivate frmKikiDataSubGate.Caption, False
                pfFormActive (frmKikiDataSubGate.hwnd)
                'EG20 V30.1.0.1 ADD END
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '「保守操作卓プログラム送信要求」を受信した場合
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
                    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理設定反映結果")
                Else
                    iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "反映処理設定反映結果")
                End If
                Call SetEnableTrue
            Case Else
                 'その他のメールを受信した場合
                 '「メールID不正」ログ出力
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    Dim iResponse           As Integer          'MsgBox戻り値
    
    'エラールーチンを宣言
    On Error Resume Next
    
    If SetteiHaneiFlg = True Then
        iResponse = MsgBox("画面表示中に設定されたデータが失われます。" & Chr(vbKeyReturn) & _
                            "よろしいですか？", _
                            vbYesNo + vbQuestion, _
                            "設定反映釦未押下")
        
        If iResponse = vbNo Then Exit Sub
    End If
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_END, 0)
    
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDisp
'//  機能名称  : 画面再描画処理
'//  機能概要  : 画面を再描画する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim strFileName          As String          'ファイル名
    Dim nCornerIndex         As Integer         ' コーナ選択状態
    
    ' EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
    Dim iLoopCnt            As Integer          'ループカウンタ（セル縦方向）
    Dim iLoopCnt2           As Integer          'ループカウンタ（セル横方向）
    ' EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END

    'エラールーチンを宣言
    On Error Resume Next

' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    If CmbCornerName.ListIndex < 0 Then
        Exit Sub
    End If
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

    '初期値設定
    strFileName = ""                            'ファイル名
    cmbGoki.Enabled = False                     '号機選択コンボボックス選択不可設定
    CmbCornerName.Enabled = False               ' コーナ選択部選択不可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    LblEkiName.Caption = TITOL_EKI_NAME         '駅名ラベル初期化           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    '----------------------------------------------------
    'グリッドタイトル設定
    '----------------------------------------------------
    Call sDispGridTitol
    
    '機器情報データ更新フラグチェック
    If KikiDataUpDateFlg = True Then
        Erase KikiDataTbl
        ReDim KikiDataTbl(0)
        Call pfKikiDataSet
    End If
    
    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear
        GridIni.Enabled = False
        
        '処理釦押下不可能設定
        CmdKikiSetMenu(0).Enabled = False           '機器構成項目設定反映
        CmdKikiSetMenu(1).Enabled = False           '機器構成項目媒体出力
        CmdKikiSetMenu(2).Enabled = False           '機器構成項目内部保存

        Exit Sub
        
    End If
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    '----------------------------------------------------
    '駅名ラベル更新
    '----------------------------------------------------
    LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo(NotEkiVer)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
    '機器構成情報（エンコードコーナ号機情報定義）イメージファイル検索
    strFileName = Dir(KIKI_DATA_SET_SUBGATE_FILE)
    
    'ファイルが存在した場合
    If strFileName <> "" Then
    
        'グリッドデータ部設定
'        Call sDispDataSet(cmbGoki.ListIndex + 1)                               ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'        nCornerIndex = CmbCornerName.ListIndex
'        Call sDispDataSet(cmbGoki.ListIndex + 1, nCornerIndex + 1)
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        '大分類:5の駅都度データはコーナ指定ではなくなったため、コーナは0固定とする。（駅都度のコーナは0で検索する）
        '1～32号機
        For iLoopCnt = 0 To 31
            '項目①～⑥
            For iLoopCnt2 = 0 To 5
                Call sDispDataSet(iLoopCnt + 1, 0, iLoopCnt2 + 1)
            Next
        Next
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_SUBGATE_IMAGE, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear
        GridIni.Enabled = False
        
        '処理釦押下不可能設定
        CmdKikiSetMenu(0).Enabled = False           '機器構成項目設定反映
        CmdKikiSetMenu(1).Enabled = False           '機器構成項目媒体出力
        CmdKikiSetMenu(2).Enabled = False           '機器構成項目内部保存

    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDispGridTitol
'//  機能名称  : グリッドタイトル部設定処理
'//  機能概要  : グリッドの初期値、タイトルを設定する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispGridTitol()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    Dim ColCount                As Integer         ' カラムカウンタ
    Dim RowCount                As Integer         'ループカウンタ

    GridIni.Visible = False             '設定中は非表示に設定
    
    'グリッドタイトル設定
    With GridIni
    
        '----------------------------------
        'グリッドの初期化
        '----------------------------------
        .Clear
        
        '----------------------------------
        'グリッドセル数設定
        '----------------------------------
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'        .Rows = 18
'        .Cols = 7
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'        .Rows = 5
'        .Cols = 4
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
' EG20 V30.1.0.1 DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】DEL START
'' EG20 V30.1.0.1 ADD START
'        .Rows = 2
'        .Cols = 7
'' EG20 V30.1.0.1 ADD END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】ADD START
        .Rows = 33
        .Cols = 8
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】ADD END
        For ColCount = 2 To (.Cols - 1)
            .ColWidth(ColCount) = 1748
        Next

        '----------------------------------
        'グリッド幅設定
        '----------------------------------
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000         'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
        
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'        For ColCount = 1 To (.Cols - 1)
'            .ColWidth(ColCount) = 1748
'        Next
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        For ColCount = 2 To (.Cols - 1)
            .ColWidth(ColCount) = 1548
        Next
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END

        '----------------------------------
        'タイトル設定
        '----------------------------------
        For RowCount = 1 To (.Rows - 1)
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
            '号機表示
            .Col = 0
            .Row = RowCount: .Text = RowCount & "号機"
            .CellAlignment = flexAlignLeftCenter
            
            '自社・他社表示（北陸新幹線では駅都度は自社のみ）
            .Col = 1
            .Row = RowCount: .Text = "自社"
            .CellAlignment = flexAlignLeftCenter
            
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'            If RowCount = 1 Then
'                '自社設定
'                .Col = 0
'                .Row = RowCount: .Text = "自社"
'                .CellAlignment = flexAlignLeftCenter
'
'            Else
'                '他社設定
'                .Col = 0
'                .Row = RowCount: .Text = "他社" & RowCount - 1
'                .CellAlignment = flexAlignLeftCenter
'            End If
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
        
        Next

        .RowHeight(0) = 500

    End With

    GridIni.Visible = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDispDataClear
'//  機能名称  : グリッドデータ部クリア処理
'//  機能概要  : グリッドデータ部をクリアする
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear()
    
    Dim iRowCnt             As Integer         'ループカウンタ
    Dim ColCount             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    GridIni.Visible = False             '設定中は非表示に設定
    
    'グリッド初期化
    With GridIni

        For iRowCnt = 1 To (.Rows - 1)
        
            .Row = iRowCnt

            '項目設定
            For ColCount = 2 To (.Rows - 1)
                .Col = ColCount
                .Text = ""
                .CellAlignment = flexAlignLeftCenter
            Next

        Next

    End With

    GridIni.Visible = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDispDataSet
'//  機能名称  : グリッドデータ部設定処理
'//  機能概要  : グリッドデータ部を設定する
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sDispDataSet(iGoki As Integer)                       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer)    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加 'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer, iKomoku As Integer)    ' EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
    
    Dim intFileNumber       As Integer                      ' ファイルポインタ
    Dim iKikiDataCnt        As Integer                      ' 機器情報データカウンタ
    Dim ColCount            As Integer                      ' カラムカウンタ
    Dim RowCount            As Integer                      ' 行カウンタ
    
    Dim strBunrui_Dai       As String                       ' 大分類
    Dim strBunrui_Tyu       As String                       ' 中分類
    Dim strBunrui_Sho       As String                       ' 小分類
    Dim strKomoku           As String                       ' 項目
    Dim strKubun            As String                       ' 区分
    Dim strData             As String                       ' 設定値
    Dim strSetShosai        As String                       ' 設定値詳細
    
    Dim strDispData         As String                       ' 表示データ
    Dim strCorner           As String                       ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim iCmpCorner          As Integer                      ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    'エラールーチンを宣言
    On Error Resume Next

    '初期値設定
    iKikiDataCnt = 0
    
    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイルをオープンする。
    Open KIKI_DATA_SET_SUBGATE_FILE For Input As #intFileNumber
    
    GridIni.Visible = False             '設定中は非表示に設定

    ColCount = START_DATA_COL_INDEX     'データ設定のスタートカラムインデックス
    RowCount = 1                        'データ設定のスタート行インデックス
    Do While Not EOF(intFileNumber)
        '１ 行読み込み
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, _
'                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strCorner, _
                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

        '機器情報データ更新フラグチェック
        If KikiDataUpDateFlg = False Then
            For iKikiDataCnt = 0 To UBound(KikiDataTbl)
        
                '該当データ検索
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'                If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(iKikiDataCnt).iBunrui_Dai) And _
'                   (CInt(strBunrui_Tyu) = KikiDataTbl(iKikiDataCnt).iBunrui_Tyu) And _
'                   (CInt(strBunrui_Sho) = KikiDataTbl(iKikiDataCnt).iBunrui_Sho) Then
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(iKikiDataCnt).iBunrui_Dai) And _
                   (CInt(strBunrui_Tyu) = KikiDataTbl(iKikiDataCnt).iBunrui_Tyu) And _
                   (CInt(strBunrui_Sho) = KikiDataTbl(iKikiDataCnt).iBunrui_Sho) And _
                   (CInt(strCorner) = KikiDataTbl(iKikiDataCnt).iBunrui_Corner) Then
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                  
                    strData = KikiDataTbl(iKikiDataCnt).strData
                    strData = StrConv(strData, vbUnicode)
                    
                End If
            Next
        End If
                
        '号機番号チェック
        If CStr(iGoki) = strBunrui_Tyu Then
            If iKomoku = CInt(strBunrui_Sho) Then       'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                ' コーナ判定追加
                ' 選択したコーナ、もしくはコーナ無関係のレコードは採用する
                iCmpCorner = CInt(strCorner)
                'If (iCorner = iCmpCorner) Then 'EG20 V30.1.0.1 DEL
                If ((iCorner = iCmpCorner) Or (iCmpCorner = 0)) Then 'EG20 V30.1.0.1 ADD
        
        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                
                    'グリッド設定
                    With GridIni
                
                        'カラムインデックス設定
                        '.Col = ColCount                    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
                        .Col = ColCount + (iKomoku - 1)     'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
                        
                       'タイトル設定
                        If (strKomoku <> "") Then
                            .Row = 0
                            .Text = strKomoku
                            .CellAlignment = flexAlignLeftCenter
                        End If
        
                        '項目設定
                        '.Row = RowCount        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
                        .Row = iGoki            'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
                        .Text = Format(pfDispIniData(.Text, strData, strKubun), "000")
                        .CellAlignment = flexAlignLeftCenter
                        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD START
                        '駅都度データ1レコード分の設定値をセルにセットしたので、一旦終わらす。
                        Exit Do
                        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD END
                        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL START 上記Exit Doにロジック変更したため不要になった。
        '                ColCount = ColCount + 1
        '                If ColCount > MAX_DATA_COL_INDEX Then
        '                 ColCount = START_DATA_COL_INDEX
        '                 RowCount = RowCount + 1
        '                End If
                       'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL END
        
                    End With
                
                End If          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            End If          'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】ADD
        End If
    
    Loop

    GridIni.Visible = True
    
    'ファイルをクローズする。
    Close #intFileNumber
    
    '号機選択コンボボックス選択可設定
    cmbGoki.Enabled = True
    CmbCornerName.Enabled = True                ' コーナ選択部選択可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

    '処理釦押下可能設定
    CmdKikiSetMenu(0).Enabled = True            '機器構成項目設定反映
    CmdKikiSetMenu(1).Enabled = True            '機器構成項目媒体出力
    CmdKikiSetMenu(2).Enabled = True            '機器構成項目内部保存

    Exit Sub

'エラー処理
Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
    
    'グリッドタイトル設定
    Call sDispGridTitol
    
    'グリッドデータ部クリア処理
    Call sDispDataClear
    GridIni.Enabled = False

    GridIni.Visible = True

    '処理釦押下不可能設定
    CmdKikiSetMenu(0).Enabled = False           '機器構成項目設定反映
    CmdKikiSetMenu(1).Enabled = False           '機器構成項目媒体出力
    CmdKikiSetMenu(2).Enabled = False           '機器構成項目内部保存

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : GridIni_Click
'//  機能名称  : グリッドを選択された時のイベントプロシージャ
'//  機能概要  : ダミーテキストのセット
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Click()
    
    Dim iLoopCnt As Integer
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'クリックされた位置にダミーテキストを移動し、フォーカスを合わせる
    With GridIni
        
        CmbDummy.Left = .Left + .CellLeft
        CmbDummy.Top = .Top + .CellTop
        CmbDummy.Width = .CellWidth
        CmbDummy.Height = .CellHeight
        CmbDummy.Text = .Text
        CmbDummy.Visible = True
        CmbDummy.SetFocus

        'ダミーコンボボックス初期値
        CmbDummy.Clear
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        'クリックされた列によってコンボボックスから選べる値を切り替える
        Select Case .Col
            'FM券コーナ番号、新幹線ICコーナ番号、NRZ券コーナ番号
            Case FM_CORNER_COL, SINKANSENIC_CORNER_COL, NRZ_CORNER_COL
                '入力値は00～99
                '駅都度仕様では00～99だが、例外的な値も考慮して255まで入れられるよう、とりあえずしておいて欲しいと東芝様からの指示により
                '000～255に変更
                For iLoopCnt = 0 To 255
                    CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
                    
                    'コンボボックスのインデックスを設定
                    If .Text <> "" Then
                        If iLoopCnt = CInt(.Text) Then
                            
                            '値が一致したらインデックス設定
                            CmbDummy.ListIndex = iLoopCnt
                            
                        End If
                    End If
                Next
            'FM券号機番号、新幹線IC号機番号、NRZ券号機番号
            Case FM_GOKI_COL, SINKANSENIC_GOKI_COL, NRZ_GOKI_COL
                '入力値は01～16
                '駅都度仕様では01～16だが、例外的な値も考慮して255まで入れられるよう、とりあえずしておいて欲しいと東芝様からの指示により
                '000～255に変更
                For iLoopCnt = 0 To 255
                    'CmbDummy.AddItem Format(CStr(iLoopCnt + 1), "000")
                    CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
                    
                    'コンボボックスのインデックスを設定
                    If .Text <> "" Then
                        'If (iLoopCnt + 1) = CInt(.Text) Then
                        If (iLoopCnt) = CInt(.Text) Then
                            
                            '値が一致したらインデックス設定
                            CmbDummy.ListIndex = iLoopCnt
                            
                        End If
                    End If
                Next
            Case Else
                For iLoopCnt = 0 To 255
                    CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
                    
                    'コンボボックスのインデックスを設定
                    If iLoopCnt = CInt(.Text) Then
                        
                        '値が一致したらインデックス設定
                        CmbDummy.ListIndex = iLoopCnt
                        
                    End If
                Next
            End Select
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'        For iLoopCnt = 0 To 255
'            CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
'
'            'コンボボックスのインデックスを設定
'            If iLoopCnt = CInt(.Text) Then
'
'                '値が一致したらインデックス設定
'                CmbDummy.ListIndex = iLoopCnt
'
'            End If
'        Next
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
    
    End With

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : GridIni_Scroll
'//  機能名称  : グリッドをスクロールした時のイベントプロシージャ
'//  機能概要  : ダミーテキストの非表示
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Scroll()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'グリッドがスクロールされた時、ダミーテキストを非表示にする
    If bScroll = False Then
        CmbDummy.Visible = False
        CmbDummy.Clear
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmbDummy_Click
'//  機能名称  : ダミーテキストが選択された時のイベントプロシージャ
'//  機能概要  : グリッドへの反映
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度修正対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_Click()

    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim iLoopCnt2           As Integer                      ' ループカウンタ
    Dim byBuff()            As Byte                         'バイトバッファ
    Dim iGoki               As Integer                      ' 号機番号
    Dim iBunrui_Sho         As Integer                      ' 小分類
    Dim iBunrui_Corner      As Integer                      ' コーナ分類            ' EG20 V3.0.0.2 （駅都度修正対応）追加

    'エラールーチンを宣言
    On Error Resume Next

    If GridIni.Text <> CmbDummy.Text Then
        '設定反映フラグ（変更あり）
        SetteiHaneiFlg = True
    End If

    GridIni.Text = CmbDummy.Text
    GridIni.CellAlignment = flexAlignLeftCenter
    
    'iGoki = cmbGoki.ListIndex + 1                                              'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
    iGoki = GridIni.Row                                                         'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
    'iBunrui_Sho = ((GridIni.Row - 1) * MAX_DATA_COL_INDEX) + GridIni.Col       'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
    iBunrui_Sho = GridIni.Col - 1                                               'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
    'iBunrui_Corner = CmbCornerName.ListIndex + 1                                    ' EG20 V3.0.0.2 （駅都度修正対応）追加 EG20 V30.3.0.1 【HKRK_Kansi07_003_01】DEL
    iBunrui_Corner = 0                                                           'EG20 V30.3.0.1【HKRK_Kansi07_003_01】 コーナ別ではないので0固定 ADD

    For iLoopCnt = 0 To UBound(KikiDataTbl)

        '該当データ検索
        If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
           (iGoki = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
           (iBunrui_Sho = KikiDataTbl(iLoopCnt).iBunrui_Sho) And _
           (iBunrui_Corner = KikiDataTbl(iLoopCnt).iBunrui_Corner) Then             ' EG20 V3.0.0.2 （駅都度修正対応）追加

            '機器構成情報データ保存
            byBuff = StrConv(GridIni.Text, vbFromUnicode)

            Erase KikiDataTbl(iLoopCnt).strData

            '動的配列の内容をログパラメータ構造体の静的配列に格納する。
            For iLoopCnt2 = 0 To UBound(KikiDataTbl(iLoopCnt).strData)
                'Null値になったら処理を抜ける。
                If byBuff(iLoopCnt2) = vbVEmpty Then Exit For

                KikiDataTbl(iLoopCnt).strData(iLoopCnt2) = byBuff(iLoopCnt2)

                '動的配列の最大要素になったら処理を抜ける
                If iLoopCnt2 = UBound(byBuff) Then Exit For
            Next

            Exit For

        End If

    Next

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmbDummy_KeyDown
'//  機能名称  : キーボード押下時のイベントプロシージャ
'//  機能概要  : ダミーテキストのセット
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '特殊キーを押下された時、下記の処理を行う
    bScroll = True
    
    With GridIni
    
        'ダミーテキストのセット
        CmbDummy.Left = .Left + .CellLeft
        CmbDummy.Top = .Top + .CellTop
        CmbDummy.Width = .CellWidth
        CmbDummy.Height = .CellHeight
        CmbDummy.Text = .Text
        CmbDummy.Visible = True
        CmbDummy.SetFocus

    End With
    bScroll = False

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmbDummy_LostFocus
'//  機能名称  : ダミーテキストからフォーカスが移動した時のイベントプロシージャ
'//  機能概要  : ダミーテキストを非表示にする
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_LostFocus()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'ダミーテキストを非表示にする
    CmbDummy.Visible = False
    CmbDummy.Clear

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmdKikiSetMenu_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦押下処理に従う
'//
'//              型        名称     　　　意味
'//  引数      : Integer　 Index          選択釦のインデックス
'//
'//              型        値        　　 意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmdKikiSetMenu_Click(Index As Integer)
    
    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bUnlock             As Boolean          ' ロック解除フラグ      ' EG20 V3.0.0.2 追加

    'エラールーチンを宣言
    On Error Resume Next
    
    '全ボタンを押下不可とする。
    Call SetEnableFalse
    
' EG20 V3.0.0.2 追加開始
' 押下した釦に応じてロック解除を制限する
' ※メール受信を待つため
    bUnlock = True
' EG20 V3.0.0.2 追加終了
    Select Case Index
        
        Case 0                                 ' 機器構成項目設定反映
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_INSTOL, 0)
            
            '機器構成項目設定反映処理
            Call sInstolKikiData
            bUnlock = False                     ' EG20 V3.0.0.2 追加
        Case 1                                 ' 機器構成項目媒体出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_OUTPUT, 0)
            
            '機器構成項目媒体出力処理
            Call sKikiDataOutPut
    
        Case 2                                 ' 機器構成項目内部保存
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_SAVE, 0)
            
            '機器構成項目内部保存処理
            Call sKikiDataSave
        
        Case 3                                 ' 機器構成設定データ選択
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_SELECT, 0)
            
            '機器構成設定データ選択処理
            Call sKikiDataSelect
    
        Case 4                                 ' 媒体入力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_MEDIUM_INPUT, 0)
            
            '媒体入力処理
            Call sInputMedium
    
        Case 5                                 ' 媒体取外
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
    
            '媒体取外処理
            Call pfRemove(Me)

        Case 6                                 ' 駅情報画面へ
            If SetteiHaneiFlg = True Then
                iResponse = MsgBox("画面表示中に設定されたデータが失われます。" & Chr(vbKeyReturn) & _
                                    "よろしいですか？", _
                                    vbYesNo + vbQuestion, _
                                    "設定反映釦未押下")
                If iResponse = vbNo Then
                    '全ボタンを押下可とする。
                    Call SetEnableTrue
                    Exit Sub
                End If
            End If
            '設定反映フラグ（変更なし）
            SetteiHaneiFlg = False
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIINFO_GAMEN_GO_BUTTOM, 0)
            Unload Me
            Load frmKikiData
            frmKikiData.Show 1
            Exit Sub

        Case 7                                 ' 自改画面へ
            If SetteiHaneiFlg = True Then
                iResponse = MsgBox("画面表示中に設定されたデータが失われます。" & Chr(vbKeyReturn) & _
                                    "よろしいですか？", _
                                    vbYesNo + vbQuestion, _
                                    "設定反映釦未押下")
                If iResponse = vbNo Then
                    '全ボタンを押下可とする。
                    Call SetEnableTrue
                    Exit Sub
                End If
            End If
            '設定反映フラグ（変更なし）
            SetteiHaneiFlg = False
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)

            '表示中画面アンロード
            Unload Me

            '自改画面表示
            Load frmKikiDataGate
            frmKikiDataGate.Show 1
            Exit Sub

        Case Else
            '処理なし
            
    End Select

    '全ボタンを押下可とする。
' EG20 V3.0.0.2 追加開始
    If bUnlock = True Then
        Call SetEnableTrue
    End If
' EG20 V3.0.0.2 削除終了
'    Call SetEnableTrue                 ' EG20 V3.0.0.2 削除

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sInstolKikiData
'//  機能名称  : 「機器構成項目設定反映」釦押下時処理
'//  機能概要  : 画面表示データをINIファイルへ反映する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.76修正対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sInstolKikiData()

    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bRet                As Boolean          '関数戻り値
    Dim lErrCode            As Long             'エラーコード
    Dim strFileName         As String           '媒体ファイル名
    
    Dim bData()             As Byte             'バイナリデータ
    Dim lLoopCnt            As Long             'ループカウンタ
    Dim lLoopCnt2           As Long             'ループカウンタ
    Dim bSysChange          As Boolean          'コンピュータ名、ネットワーク変更処理判定
    Dim byBuff()            As Byte             'バイトバッファ
    Dim strSetteiData       As String           ' 設定値

    Dim lResult             As Long             ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加

    'エラールーチンを宣言
    On Error Resume Next
    
    iResponse = MsgBox("機器構成データをインストールします。" & Chr(vbKeyReturn) & _
                        "よろしいですか？", _
                        vbOKCancel + vbExclamation, _
                        "設定反映確認")
    If iResponse = vbCancel Then
        Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
        Exit Sub
    End If
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '機器構成データテーブル（大分類：エンコードコーナ号機）の設定値を「999」の書式に変換
    For lLoopCnt = 0 To UBound(KikiDataTbl)
        '該当データ検索
        If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(lLoopCnt).iBunrui_Dai) Then
            
            strSetteiData = KikiDataTbl(lLoopCnt).strData
            strSetteiData = StrConv(strSetteiData, vbUnicode)
            strSetteiData = Format(strSetteiData, "000")
    
            '機器構成情報データ保存
            byBuff = StrConv(strSetteiData, vbFromUnicode)

            Erase KikiDataTbl(lLoopCnt).strData

            '動的配列の内容をログパラメータ構造体の静的配列に格納する。
            For lLoopCnt2 = 0 To UBound(KikiDataTbl(lLoopCnt).strData)
                'Null値になったら処理を抜ける。
                If byBuff(lLoopCnt2) = vbVEmpty Then Exit For

                KikiDataTbl(lLoopCnt).strData(lLoopCnt2) = byBuff(lLoopCnt2)

                '動的配列の最大要素になったら処理を抜ける
                If lLoopCnt2 = UBound(byBuff) Then Exit For
            Next
        End If
    Next

    '構造体配列をバイナリ配列に変換
    ReDim bData((UBound(KikiDataTbl) + 1) * Len(KikiDataTbl(0))) As Byte
    For lLoopCnt = 0 To UBound(KikiDataTbl)
          MoveMemory bData(lLoopCnt * Len(KikiDataTbl(0))), KikiDataTbl(lLoopCnt), Len(KikiDataTbl(lLoopCnt))
    Next
    
    '機器構成データインストール処理
    bRet = dllInstolKikiData(KIKI_DATA_FILE, EKI_SETTI_FILE, bData(0), UBound(KikiDataTbl) + 1, lErrCode)
    
    If bRet = False Then
        
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
        '異常終了
        iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理設定反映結果")
        Call SetEnableTrue              ' EG20 V3.0.0.2（メール送信しない場合のみロック解除対応）追加
    Else
        'コンピュータ名、ネットワーク変更処理
        
        bSysChange = pfNetWorkChng(Me)
        If bSysChange = False Then

            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
            '異常終了
             iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理設定反映結果")
            Call SetEnableTrue              ' EG20 V3.0.0.2（メール送信しない場合のみロック解除対応）追加
        Else
' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加開始
            ' //////////////////////////////////////////////
            ' // 操作卓プログラム処理
            ' //////////////////////////////////////////////
             lResult = pubfuncTakuProgramData(2, EKI_SETTI_FILE)
             If lResult = 0 Then
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                ' 異常終了
                iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理設定反映結果")
                Call SetEnableTrue
                Exit Sub
             ElseIf lResult = 1 Then
                ' メール送信中
                ' ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
                ' 設定反映フラグ（変更なし）
                SetteiHaneiFlg = False
                 
                Exit Sub
             End If
' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加終了
        
            'ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
            
            '正常終了
            iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "反映処理設定反映結果")
            
            '設定反映フラグ（変更なし）
            SetteiHaneiFlg = False
            Call SetEnableTrue              ' EG20 V3.0.0.2（メール送信しない場合のみロック解除対応）追加
        End If
    End If


End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sKikiDataOutPut
'//  機能名称  : 「機器構成項目媒体出力」釦押下時処理
'//  機能概要  : 機器構成データファイルを外部媒体に出力する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataOutPut()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
    Dim iResponse            As Integer         'MsgBox戻り値

    Dim iRet                 As Integer         'メッセージボックス戻り値
    Dim lSekuta              As Long            'セクタ（クラスタ当り）
    Dim lByte                As Long            'バイト数（セクタ当り）
    Dim lKurasuta            As Long            'フリークラスタ数
    Dim lDrive               As Long            'ドライブのクラスタ数（合計）
    Dim strDrive             As String          'ドライブ
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '機器構成データファイル検索
    '----------------------------------------------------
    strFileName = Dir(KIKI_DATA_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_KIKI_DATA, 0)
        
        '異常終了
        MsgBox "媒体出力するデータがありません。", _
                vbOKOnly + vbExclamation, _
                 "データ無警告"
        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '媒体出力処理
    '----------------------------------------------------
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
'        FileCopy KIKI_DATA_FILE, sWriteDir & Dir(KIKI_DATA_FILE)                                       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        FileCopy KIKI_DATA_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(KIKI_DATA_FILE)    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
        '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体出力結果")
    
    End If
  
  Exit Sub
 
COPY_ERROR:

    '異常ログ出力
    Select Case Err.Number
        Case 61 ' 媒体出力空き容量不足
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case 71 ' 媒体なし
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_NOT_DISK, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

    '異常終了
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体出力結果")

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sKikiDataSave
'//  機能名称  : 「機器構成項目内部保存」釦押下時処理
'//  機能概要  : 機器構成データファイルを指定フォルダに出力する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSave()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
    Dim sMyPath(1 To 3)      As String          'ファイルパス
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim iLoopCount           As Integer         'ループカウンタ
    Dim intFileNo            As Integer         'ファイル番号

    'エラールーチンを宣言
    On Error Resume Next
    
    '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '内部保存処理
    '----------------------------------------------------
    iResponse = MsgBox("機器構成設定を一時保存します。" & vbCrLf & "よろしいですか？", _
    vbOKCancel + vbQuestion, "一時保存確認")
    
    If iResponse = vbCancel Then Exit Sub
     
     'ファイル検索
    strFileName = Dir(KIKI_DATA_S_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then

        intFileNo = FreeFile                                        '未使用のファイル番号を取得する
        Open KIKI_DATA_S_FILE For Output Access Write As #intFileNo
        Close #intFileNo
    End If
    
    '一時保存ファイルを作成する
    Name KIKI_DATA_S_FILE As KIKI_DATA_S_TEMP_FILE
    
    'ファイル名取得
    sWriteDir = KIKI_DATA_S_FILE
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
        FileCopy KIKI_DATA_FILE, sWriteDir
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
        '一時保存ファイル削除
        Kill KIKI_DATA_S_TEMP_FILE
        
        '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "一時保存結果")
    
    End If
  
  Exit Sub
 
COPY_ERROR:

    '異常ログ出力
    Select Case Err.Number
        Case 61 ' 空き容量不足
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

        'ファイル検索
        strFileName = Dir(KIKI_DATA_S_FILE)
        If strFileName <> "" Then
            'ファイル削除
            Kill KIKI_DATA_S_FILE
        End If
        'ファイル名称を元に戻す
        Name KIKI_DATA_S_TEMP_FILE As KIKI_DATA_S_FILE
    
    '異常終了
     iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存結果")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sKikiDataSelect
'//  機能名称  : 「機器構成設定データ選択」釦押下時処理
'//  機能概要  : 機器構成データ内部保存ファイルを機器構成データファイルにコピーする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSelect()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
    Dim sMyPath(1 To 3)      As String          'ファイルパス
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim iLoopCount           As Integer         'ループカウンタ
    Dim intFileNo            As Integer         'ファイル番号
    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード

    'エラールーチンを宣言
    On Error Resume Next
    
    '----------------------------------------------------
    '機器構成データファイル更新処理
    '----------------------------------------------------
    iResponse = MsgBox("機器構成設定の一時保存データを取込みます。" & vbCrLf & "よろしいですか？", _
    vbOKCancel + vbQuestion, "一時保存データ取込確認")
    
    If iResponse = vbCancel Then Exit Sub
    
   'ファイル検索
    strFileName = Dir(KIKI_DATA_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then

        intFileNo = FreeFile                                        '未使用のファイル番号を取得する
        Open KIKI_DATA_FILE For Output Access Write As #intFileNo
        Close #intFileNo
    End If
    
    '一時保存ファイルを作成する
    Name KIKI_DATA_FILE As KIKI_DATA_BACKUP_FILE
    
    'ファイル名取得
    strFileName = Dir(KIKI_DATA_S_FILE)
    sWriteDir = strFileName
    If sWriteDir <> "" Then
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
        On Error GoTo COPY_ERROR
        'ファイルコピー
         FileCopy KIKI_DATA_S_FILE, KIKI_DATA_FILE
        
        '機器構成データ（エンコードコーナ号機情報定義）イメージファイル
        bRet = dllGetKikiIniData(2, 1, KIKI_DATA_SET_SUBGATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
            'ファイル削除
            Kill KIKI_DATA_FILE
            'ファイル名称を元に戻す
            Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
            '異常終了
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存データ取込結果")
            Exit Sub
        End If
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
        '一時保存ファイル削除
        Kill KIKI_DATA_BACKUP_FILE
        
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
        '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "一時保存データ取込結果")
    
        '機器情報データ更新フラグ設定（更新設定）
        KikiDataUpDateFlg = True
        '画面表示処理
        Call sDisp
        '機器情報データ更新フラグ設定（通常設定）
        KikiDataUpDateFlg = False
        
        '設定反映フラグ（変更あり）
        SetteiHaneiFlg = True
    End If
  
  Exit Sub
 
COPY_ERROR:

    '異常ログ出力
    Select Case Err.Number
        Case 61 ' 空き容量不足
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

   'ファイル検索
   strFileName = Dir(KIKI_DATA_FILE)
   If strFileName <> "" Then
    'ファイル削除
    Kill KIKI_DATA_FILE
   End If
   'ファイル名称を元に戻す
   Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
   
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
   '異常終了
   iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存データ取込結果")

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sInputMedium
'//  機能名称  : 「媒体入力」釦押下時処理
'//  機能概要  : 外部媒体を機器構成データファイルにコピーする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.77修正対応】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sInputMedium()

    Dim iResponse               As Integer      'MsgBox戻り値
    Dim bRet                    As Boolean      '関数戻り値
    Dim lErrCode                As Long         'エラーコード
    Dim strFileName             As String       '媒体ファイル名
    
    Dim iRet                    As Integer      'メッセージボックス戻り値
    Dim lSekuta                 As Long         'セクタ（クラスタ当り）
    Dim lByte                   As Long         'バイト数（セクタ当り）
    Dim lKurasuta               As Long         'フリークラスタ数
    Dim lDrive                  As Long         'ドライブのクラスタ数（合計）
    Dim strDrive                As String       'ドライブ
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    
    'エラールーチンを宣言
    On Error Resume Next
    
    iResponse = MsgBox("機器構成設定の媒体入力を行います。" & vbCrLf & "よろしいですか？", _
    vbOKCancel + vbQuestion, "媒体入力確認")
    
    If iResponse = vbCancel Then
        Set objFso = Nothing
        Exit Sub
    End If
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
    Set objFso = Nothing
    '拡張子を設定
'    CommonDialog1.Filter = "機器構成データファイル（KIKI_DATA.CSV）|KIKI_DATA.CSV|"    ' EG20 V5.0.2.1削除
    CommonDialog1.Filter = "機器構成データファイル（KIKI_DATA.CSV）|*KIKI_DATA.CSV|"    ' EG20 V5.0.2.1追加
    'ファイル選択画面を開く
    CommonDialog1.ShowOpen
    '選択したファイル名を取得
    strFileName = CommonDialog1.FileName
    
    Call ChDrive("D")

    'ファイル存在チェック
    If strFileName <> "" Then

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了

        On Error GoTo COPY_ERROR
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL Start
        'ファイルコピー
'        FileCopy strFileName, KIKI_DATA_FILE
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL End
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
        '一時保存フォルダにデータをコピーし読取専用を解除する
       If pfChangeAttrNormal(strFileName, PATH_HOSHUTMP_KIKI_DATA, KIKI_DATA_FILE) = False Then
          GoTo COPY_ERROR
       End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
        
        '機器情報設定（エンコードコーナ号機情報定義）イメージファイル作成
        bRet = dllGetKikiIniData(2, 1, KIKI_DATA_SET_SUBGATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
            
            '異常終了
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果")
            
            Exit Sub
       End If
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
        '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体入力結果")
    
        '機器情報データ更新フラグ設定（更新設定）
        KikiDataUpDateFlg = True
        '画面表示処理
        Call sDisp
        '機器情報データ更新フラグ設定（通常設定）
        KikiDataUpDateFlg = False
        
        '設定反映フラグ（変更あり）
        SetteiHaneiFlg = True
    End If

  Exit Sub
  
COPY_ERROR:

    '異常ログ出力
    Select Case Err.Number
        Case 61 ' 空き容量不足
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    '一時保存フォルダを削除する
    psDeleteFolder PATH_HOSHUTMP
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了

    '異常終了
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果")

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : SetEnableFalse
'//  機能名称  : 画面ロック処理
'//  機能概要  : 画面をロックする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    CmdKikiSetMenu(3).Enabled = False
    CmdKikiSetMenu(4).Enabled = False
    CmdKikiSetMenu(5).Enabled = False
    CmdKikiSetMenu(6).Enabled = False
    CmdKikiSetMenu(7).Enabled = False
    cmdCancel.Enabled = False
    
    'CmdKikiSetMenu(0)～(2)は条件によっては元々押下不可のため判定を行う
    If cmbGoki.Enabled = True Then
        cmbGoki.Enabled = False     '号機選択コンボボックス選択不可設定
        DoEvents
    End If
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    If CmbCornerName.Enabled = True Then
        CmbCornerName.Enabled = False
        DoEvents
    End If
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
    If CmdKikiSetMenu(0).Enabled = True Then
        CmdKikiSetMenu(0).Enabled = False
    End If
    
    If CmdKikiSetMenu(1).Enabled = True Then
        CmdKikiSetMenu(1).Enabled = False
    End If
    
    If CmdKikiSetMenu(2).Enabled = True Then
        CmdKikiSetMenu(2).Enabled = False
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    Dim strFileName          As String          'ファイル名
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下可とする。
    CmdKikiSetMenu(3).Enabled = True
    CmdKikiSetMenu(4).Enabled = True
    CmdKikiSetMenu(5).Enabled = True
    CmdKikiSetMenu(6).Enabled = True
    CmdKikiSetMenu(7).Enabled = True
    cmdCancel.Enabled = True
    
    'コンボボックスとCmdKikiSetMenu(0)～(2)は条件によっては元々押下不可のため、画面表示の有無で判定を行う
    strFileName = Dir(KIKI_DATA_SET_SUBGATE_FILE)
    'ファイルが存在する場合
    If strFileName <> "" Then
        cmbGoki.Enabled = True              '号機選択コンボボックス選択可設定
        CmbCornerName.Enabled = True        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        CmdKikiSetMenu(0).Enabled = True
        CmdKikiSetMenu(1).Enabled = True
        CmdKikiSetMenu(2).Enabled = True
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmbGoki_Click
'//  機能名称  : 号機選択処理
'//  機能概要  : グリッドデータを再設定する
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmbGoki_Click()
    
    Dim iIndex          As Integer                  'インデックス
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    Call SetEnableFalse
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_GOKI_SELECT, 0)
    
    '画面表示処理
    Call sDisp

    '全ボタンを押下可とする。
    Call SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : CmbCornerName_Click
'//  機能名称  : コーナ選択部選択処理
'//  機能概要  : グリッドデータを再設定する
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//                 cmbEkiInfo_Click()流用
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmbCornerName_Click()

    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    Call SetEnableFalse
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_GAMEN_CORNER_SELECT, 0)
    
    '画面表示処理
    Call sDisp

    '全ボタンを押下可とする。
    Call SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : InitCornerComboBox
'//  機能名称  : コーナ設定コンボボックスの初期化処理
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub InitCornerComboBox()

    Dim intLoop   As Integer            ' ループカウンタ
    Dim strCorner As String             ' 文字列格納エリア
    
    On Error Resume Next
    
    ' /////////////////////////////////////////////////////
    ' // 初期化処理
    ' /////////////////////////////////////////////////////
    ' コーナ名称設定処理
    Call gsGetCornerName
    
    CmbCornerName.Clear
    For intLoop = 0 To 5
    
        '設定ありのコーナを活性にする
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
            strCorner = gstrCornerName(intLoop)
            CmbCornerName.AddItem strCorner
        End If
    Next intLoop
    CmbCornerName.ListIndex = 0

End Sub

