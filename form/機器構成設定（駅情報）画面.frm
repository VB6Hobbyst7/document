VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmKikiData 
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "   ｴﾝｺｰﾄﾞｺｰﾅ    号機情報定義画面へ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   7
      Left            =   7200
      TabIndex        =   14
      Top             =   7800
      Width           =   2175
   End
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
      TabIndex        =   13
      Top             =   480
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Index           =   6
      Left            =   7200
      TabIndex        =   11
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   120
      Top             =   480
   End
   Begin VB.ComboBox cmbEkiInfo 
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
      TabIndex        =   9
      Top             =   960
      Width           =   2295
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
      TabIndex        =   8
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   7800
      Width           =   2175
   End
   Begin VB.TextBox txtDummy 
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10560
      Width           =   1095
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   6195
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10927
      _Version        =   393216
      Rows            =   12
      Cols            =   4
      FixedCols       =   2
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
      Height          =   390
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "機器構成設定（駅情報）"
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
      Height          =   403
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：機器情報設定（駅情報）画面.frm
'//  パッケージ名：機器情報設定（駅情報）画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　「自改画面へ」釦処理追加
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//                 コンピュータ名、ネットワーク変更処理追加
'//                 ディスク情報取得位置変更
'//                 ファイル検索処理削除
'//                 媒体ファイル名を固定名称に変更
'//                 画面ロック処理／画面ロック解除処理追加
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                 「一時保存データ取込」釦処理を修正
'//                  ボタン名称変更によるポップアップ変更
'//     REVISIONS :(1.17.0.1) 2009-12-24  REVISED BY [TCC] E.Watanabe
'//                 不具合修正
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                画面再前面表示修正(不具合修正)
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//                 ファイル選択画面をOS仕様に変更
'//                 カーソル移動の処理を削除
'//                 設定反映ボタンが押されずに画面遷移するときの警告表示を追加
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ポップアップ確認画面のタイトル修正
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(2.7.0.1) 2011-01-18  REVISED BY [TCC] S.Terao
'//                 ＥＧＲ_ＪＥ対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(EG20 V5.12.0.1) 2012-05-18  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   'メイルタイマのインターバル値
Private bScroll As Boolean
Private GamenDataTbl() As EKI_DATA_TBL                  '画面表示用データテーブル(配列の要素[0]は未使用)
Private KikiDataUpDateFlg As Boolean                    '機器情報データ更新フラグ

Private SetteiHaneiFlg As Boolean                       '設定反映フラグ     'V1.20.0.1 ADD

Private Const TITOL_EKI_NAME = "駅名："                 '駅名タイトル       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 機器情報設定（駅情報）画面(アクティブ時：イベントプロシージャ)
'//  機能概要  : 最前前表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.17.0.1) 2009-12-24  REVISED BY [TCC] E.Watanabe
'//                 不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'エラールーチンを宣言
    On Error Resume Next
    
    '自画面最前面表示処理を行う。
    pfFormActive (hwnd)
    
'V1.17.0.1 ADD START
    'フォーカス位置を設定
    cmdCancel.SetFocus
'V1.17.0.1 ADD END
    
    'タイマを起動する
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 機器情報設定（駅情報）画面(ディアクティブ時)
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
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 機器情報設定（駅情報）画面(ロード時：イベントプロシージャ)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                設定反映フラグ追加
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_START, 0)
    
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
    
    '機器情報設定（駅情報）イメージファイル作成
    bRet = dllGetKikiIniData(0, 0, KIKI_DATA_SET_EKI_INFO_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
    If bRet = False Then
        '機器情報設定（駅情報）イメージファイル削除
        Kill KIKI_DATA_SET_EKI_INFO_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
    End If
    
    '駅情報コンボボックス初期値設定
    cmbEkiInfo.Clear
    cmbEkiInfo.AddItem "監視"
    cmbEkiInfo.AddItem "ネットワーク"
    cmbEkiInfo.ListIndex = 0
    
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
    SetteiHaneiFlg = False     'V1.20.0.1 ADD
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                画面再前面表示修正(不具合修正)
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
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
'                AppActivate frmInputMstData.Caption, False     ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
                AppActivate frmKikiData.Caption, False          ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
                pfFormActive (frmKikiData.hwnd)                 ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                設定反映釦の未押下メッセージ追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    Dim iResponse           As Integer          'MsgBox戻り値    'V1.20.0.1 ADD
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'V1.20.0.1 ADD START
    If SetteiHaneiFlg = True Then
        iResponse = MsgBox("画面表示中に設定されたデータが失われます。" & Chr(vbKeyReturn) & _
                            "よろしいですか？", _
                            vbYesNo + vbQuestion, _
                            "設定反映釦未押下")
        
        If iResponse = vbNo Then Exit Sub
    End If
    'V1.20.0.1 ADD END
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_END, 0)
    
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim strFileName          As String          'ファイル名
    Dim nCornerIndex         As Integer         ' コーナ選択状態

    'エラールーチンを宣言
    On Error Resume Next

' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    If CmbCornerName.ListIndex < 0 Then
        Exit Sub
    End If
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

    '初期値設定
    strFileName = ""                            'ファイル名
    cmbEkiInfo.Enabled = False                  '駅情報コンボボックス選択不可設定
    CmbCornerName.Enabled = False               ' コーナ選択部選択不可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    LblEkiName.Caption = TITOL_EKI_NAME         '駅名ラベル初期化           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    '----------------------------------------------------
    'グリッドタイトル設定
    '----------------------------------------------------
    Call sDispGridTitol
    Erase GamenDataTbl
    ReDim GamenDataTbl(0)
    
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
        Call sDispDataClear(1, GridIni.Rows)

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
    
    '機器構成情報（駅情報）イメージファイル検索
    strFileName = Dir(KIKI_DATA_SET_EKI_INFO_FILE)
    
    'ファイルが存在した場合
    If strFileName <> "" Then
    
        'グリッドデータ部設定
'        Call sDispDataSet(pfGetCodeDaiBunrui(cmbEkiInfo))                              ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        nCornerIndex = CmbCornerName.ListIndex
        Call sDispDataSet(pfGetCodeDaiBunrui(cmbEkiInfo), nCornerIndex + 1)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
        cmbEkiInfo.Enabled = True                   '駅情報コンボボックス選択可設定
        CmbCornerName.Enabled = True                ' コーナ選択部選択不可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
        '処理釦押下可能設定
        CmdKikiSetMenu(0).Enabled = True            '機器構成項目設定反映
        CmdKikiSetMenu(1).Enabled = True            '機器構成項目媒体出力
        CmdKikiSetMenu(2).Enabled = True            '機器構成項目内部保存

    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKIINFO_IMAGE, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear(1, GridIni.Rows)
    
        '処理釦押下不可能設定
        CmdKikiSetMenu(0).Enabled = False           '機器構成項目設定反映
        CmdKikiSetMenu(1).Enabled = False           '機器構成項目媒体出力
        CmdKikiSetMenu(2).Enabled = False           '機器構成項目内部保存

    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispGridTitol()
    
    'エラールーチンを宣言
    On Error Resume Next

    'グリッドタイトル設定
    With GridIni
    
        '----------------------------------
        'グリッドの初期化
        '----------------------------------
        .Clear
        .Width = 11550
        
        '----------------------------------
        'グリッドセル数設定
        '----------------------------------
        .Rows = 9
        .Cols = 4
        
        '----------------------------------
        'グリッド幅設定
        '----------------------------------
        .ColWidth(0) = 400
        .ColWidth(1) = 3700
        .ColWidth(2) = 2500
'        .ColWidth(3) = 4825        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        .ColWidth(3) = 5000         ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        
        '----------------------------------
        'タイトル設定
        '----------------------------------
        '項目設定
        .Col = 1
        .Row = 0: .Text = "項目"
        .CellAlignment = flexAlignCenterCenter

        '設定値設定
        .Col = 2
        .Text = "設定値"
        .CellAlignment = flexAlignCenterCenter

        '詳細設定
        .Col = 3
        .Text = "設定値詳細"
        .CellAlignment = flexAlignCenterCenter
        
'        .RowHeight(0) = 700        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        .RowHeight(0) = 450         ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    End With
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sDispDataClear
'//  機能名称  : グリッドデータ部クリア処理
'//  機能概要  : グリッドデータ部をクリアする
'//
'//              型        名称         意味
'//  引数      : Integer   intStartRow  開始行位置
'//              Integer   intEndRow    終了行位置
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear(intStartRow As Integer, intEndRow As Integer)
    
    Dim iLoopCnt             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    'グリッド初期化
    With GridIni
            .Rows = intEndRow

        For iLoopCnt = intStartRow To intEndRow - 1

            '通番設定
            .Col = 0
            .Row = iLoopCnt: .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '項目設定
            .Col = 1
            .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '設定値設定
            .Col = 2
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
            .CellAlignment = flexAlignLeftCenter

            '詳細設定
            .Col = 3
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
            .CellAlignment = flexAlignLeftCenter

            .RowHeight(iLoopCnt) = 700
        Next

    End With
        
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sDispDataSet
'//  機能名称  : グリッドデータ部設定処理
'//  機能概要  : グリッドデータ部を設定する
'//
'//              型        名称         意味
'//  引数      : Integer   iBunrui_Dai  大分類
'//            : Integer   iCorner      コーナ  ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(2.7.0.1) 2011-01-18  REVISED BY [TCC] S.Terao
'//                 ＥＧＲ_ＪＥ対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sDispDataSet(iBunrui_Dai As Integer)                       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
Private Sub sDispDataSet(iBunrui_Dai As Integer, iCorner As Integer)    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    Dim intFileNumber       As Integer                      ' ファイルポインタ
    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim iKikiDataCnt        As Integer                      ' 機器情報データカウンタ
    Dim iRowCnt             As Integer                      ' 行カウンタ
    
    Dim strBunrui_Dai       As String                       ' 大分類
    Dim strBunrui_Tyu       As String                       ' 中分類
    Dim strBunrui_Sho       As String                       ' 小分類
    Dim strNo               As String                       ' 通番
    Dim strKomoku           As String                       ' 項目
    Dim strKubun            As String                       ' 区分
    Dim strData             As String                       ' 設定値
    Dim strSetShosai        As String                       ' 設定値詳細
                
    Dim byBuff()            As Byte                         ' バイトバッファ
    Dim iLoopCnt2           As Integer
    Dim iBuffCnt            As Integer                      'V2.7.0.1 ADD　エリア確保数
    
    Dim strCorner           As String                       ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim iCmpCorner          As Integer                      ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    'エラールーチンを宣言
    On Error Resume Next

    '初期値設定
    iKikiDataCnt = 0
    iLoopCnt = 1
    GridIni.Rows = 1
    
    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '機器構成情報（駅情報）イメージファイルをオープンする。
    Open KIKI_DATA_SET_EKI_INFO_FILE For Input As #intFileNumber
    
    GridIni.Visible = False
    Do While Not EOF(intFileNumber)
        
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'        '１ 行読み込み
'        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strNo, _
'                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        '１ 行読み込み
        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strCorner, strNo, _
                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        
        '機器情報データ更新フラグチェック
        If KikiDataUpDateFlg = False Then
            strData = KikiDataTbl(iKikiDataCnt).strData
            strData = StrConv(strData, vbUnicode)
        End If
        
        '機器情報データカウンタインクリメント
        iKikiDataCnt = iKikiDataCnt + 1
        
        If iBunrui_Dai = strBunrui_Dai Then
        
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        ' コーナ判定追加
        ' 選択したコーナ、もしくはコーナ無関係のレコードは採用する
        iCmpCorner = CInt(strCorner)
        If ((iCorner = iCmpCorner) Or (iCmpCorner = 0)) Then
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        
            With GridIni
                
                '最大行数インクリメント
                .Rows = iLoopCnt + 1
                
                '通番設定
                .Col = 0
                .Row = iLoopCnt: If .Text = "" Then .Text = CStr(iLoopCnt)
                .CellAlignment = flexAlignLeftCenter
    
                '項目設定
                .Col = 1
                If .Text = "" Then .Text = strKomoku
                .CellAlignment = flexAlignLeftCenter
    
                '設定値設定
                .Col = 2
                .Text = strData
                .CellAlignment = flexAlignLeftCenter
    
                '画面表示用データ保存
                ReDim Preserve GamenDataTbl(.Rows - 1)
                GamenDataTbl(.Rows - 1).iBunrui_Dai = CInt(strBunrui_Dai)     '大分類
                GamenDataTbl(.Rows - 1).iBunrui_Tyu = CInt(strBunrui_Tyu)     '中分類
                GamenDataTbl(.Rows - 1).iBunrui_Sho = CInt(strBunrui_Sho)     '小分類
                GamenDataTbl(.Rows - 1).iBunrui_Corner = iCmpCorner           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                
                '設定値
                If strData <> "" Then 'V2.7.0.1 ADD
                    byBuff = StrConv(strData, vbFromUnicode)
                'V2.7.0.1 ADD START
                Else
                   '半角スペースを代入
                   '半角スペースで変換をかけても、
                   '「設定反映」時は画面内容を見直されるため
                   '半角スペース変換結果(32)は(0）となる。
                    strData = "  "
                    byBuff = StrConv(strData, vbFromUnicode)
                End If
                'V2.7.0.1 ADD END
                '動的配列の内容をログパラメータ構造体の静的配列に格納する。
                For iLoopCnt2 = 0 To UBound(GamenDataTbl(.Rows - 1).strData)
                    'Null値になったら処理を抜ける。
                    If byBuff(iLoopCnt2) = vbVEmpty Then Exit For
                    
                    GamenDataTbl(.Rows - 1).strData(iLoopCnt2) = byBuff(iLoopCnt2)
                    
                    '動的配列の最大要素になったら処理を抜ける
                    If iLoopCnt2 = UBound(byBuff) Then Exit For
                Next
                 
                '詳細設定
                .Col = 3
                If .Text = "" Then .Text = strSetShosai
                .CellAlignment = flexAlignLeftCenter
    
                .RowHeight(iLoopCnt) = 700
        
                'インクリメント
                iLoopCnt = iLoopCnt + 1             '表示行数
            
                '表示データが１画面に表示しきれない場合
                If .Rows > 8 Then
                    'スクロールバー分、グリッドを広げる
                    .Width = 11775
                End If
                
            End With
        
        End If          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        End If
    
    Loop

    GridIni.Visible = True
    'ファイルをクローズする。
    Close #intFileNumber
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '表示行数に満たない場合データをクリアする
    If GridIni.Rows < 9 Then
        'グリッドデータ部クリア処理
'        Call sDispDataClear(GridIni.Rows - 1, 9) 'V1.8.0.1 DEL
         Call sDispDataClear(GridIni.Rows, 9) 'V1.8.0.1 ADD
    End If

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
'    Call sDispDataClear(1, GridIni.Rows - 1)   'V1.8.0.1 DEL
     Call sDispDataClear(1, GridIni.Rows)       'V1.8.0.1 ADD
     
    GridIni.Visible = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmbEkiInfo_Change
'//  機能名称  : 駅情報選択処理
'//  機能概要  : グリッドデータを再設定する
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmbEkiInfo_Click()
    
    Dim iIndex          As Integer                  'インデックス
    
    'エラールーチンを宣言
    On Error Resume Next

'V1.12.0.1 ADD START
    '全ボタンを押下不可とする。
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_EKIINFO_SELECT, 0)
    
    '画面表示処理
    Call sDisp

'V1.12.0.1 ADD START
    '全ボタンを押下可とする。
    Call SetEnableTrue
'V1.12.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Click()
    
    Dim iColSave        As Integer          '列保存エリア
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'クリックされた位置にダミーテキストを移動し、フォーカスを合わせる
    With GridIni
    
        '詳細設定値は処理しない
        If .Col = 3 Then Exit Sub
        
        iColSave = .Col
        .Col = 0
        If .Text = "" Then
            .Col = iColSave
            Exit Sub
        End If
        .Col = iColSave
        
        txtDummy.Left = .Left + .CellLeft
        txtDummy.Top = .Top + .CellTop
        txtDummy.Width = .CellWidth
        txtDummy.Height = .CellHeight
        txtDummy.Text = .Text
        txtDummy.Visible = True
        txtDummy.SetFocus
        
        'ダミーテキストの最終にフォーカス移動
        SendKeys "{END}"
    
    End With

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : GridIni_KeyPress
'//  機能名称  : グリッドを押下された時のイベントプロシージャ
'//  機能概要  : ダミーテキストのセット
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_KeyPress(KeyAscii As Integer)
    
    'エラールーチンを宣言
    On Error Resume Next

' EG20 V3.5.0.1追加開始
    If GridIni.Col = 3 Then
        Exit Sub
    End If
' EG20 V3.5.0.1追加終了

    'クリックされた位置にダミーテキストを移動し、フォーカスを合わせる
    With GridIni
        
        txtDummy.Left = .Left + .CellLeft
        txtDummy.Top = .Top + .CellTop
        txtDummy.Width = .CellWidth
        txtDummy.Height = .CellHeight
        If KeyAscii <> 13 Then
            txtDummy.Text = .Text & Chr(KeyAscii)
        Else
            txtDummy.Text = .Text
        End If
        txtDummy.Visible = True
        txtDummy.SetFocus
        
        'ダミーテキストの最終にフォーカス移動
        SendKeys "{END}"
    
    End With

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Scroll()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'グリッドがスクロールされた時、ダミーテキストを非表示にする
    If bScroll = False Then
        txtDummy.Visible = False
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtDummy_Change
'//  機能名称  : ダミーテキストが変更された時のイベントプロシージャ
'//  機能概要  : グリッドへの反映
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                設定反映フラグ追加
'//                空白セル入力時の回避処理を追加
'//     REVISIONS :(EG20 V5.12.0.1) 2012-05-18  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.4.0.1) 2012-06-17 REVISED BY [TCC] H.Sugimoto
'//                【総点検修正対応：半角スペースの入力を抑止する対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtDummy_Change()
    
    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim iLoopCnt2           As Integer                      ' ループカウンタ
    Dim byBuff()            As Byte                         ' バイトバッファ
    Dim szWork              As String                       ' ワーク    ' EG20 V6.4.0.1追加
    
    'エラールーチンを宣言
    On Error Resume Next
    
' EG20 V6.4.0.1追加開始
    If InStr(txtDummy.Text, " ") > 0 Then
        szWork = Replace(txtDummy.Text, " ", "")
        txtDummy.Text = szWork
        MsgBox "スペースの入力できません。" & vbCrLf & _
                "入力内容を確認してください。", vbOKOnly + vbCritical, "設定値入力異常"
        Exit Sub
    End If
' EG20 V6.4.0.1追加終了
    
    'V1.20.0.1 ADD START
    If GridIni.Text <> txtDummy.Text Then
        '設定反映フラグ（変更あり）
        SetteiHaneiFlg = True
    End If
    'V1.20.0.1 ADD END
    
    'グリッドに入力項目を反映させる
    GridIni.Text = txtDummy.Text
    
    'V1.20.0.1 ADD START
    'データがない行をクリックした場合
    If UBound(GamenDataTbl) < GridIni.Row Then
        Exit Sub
    End If
    'V1.20.0.1 ADD END
    
    '画面表示データ保存
    byBuff = StrConv(GridIni.Text, vbFromUnicode)
    '動的配列の内容をログパラメータ構造体の静的配列に格納する。
    Erase GamenDataTbl(GridIni.Row).strData
    For iLoopCnt2 = 0 To UBound(GamenDataTbl(GridIni.Row).strData)
        'Null値になったら処理を抜ける。
        If byBuff(iLoopCnt2) = vbVEmpty Then Exit For
        
        GamenDataTbl(GridIni.Row).strData(iLoopCnt2) = byBuff(iLoopCnt2)
        
        '動的配列の最大要素になったら処理を抜ける
        If iLoopCnt2 = UBound(byBuff) Then Exit For
    Next
    
    For iLoopCnt = 0 To UBound(KikiDataTbl) - 1
    
        '画面情報と機器情報の大分類、中分類、小分類が一致した場合
' EG20 V5.12.0.1削除開始
'        If (GamenDataTbl(GridIni.Row).iBunrui_Dai = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
'           (GamenDataTbl(GridIni.Row).iBunrui_Tyu = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
'           (GamenDataTbl(GridIni.Row).iBunrui_Sho = KikiDataTbl(iLoopCnt).iBunrui_Sho) Then
' EG20 V5.12.0.1削除終了
' EG20 V5.12.0.1追加開始
        If (GamenDataTbl(GridIni.Row).iBunrui_Dai = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
           (GamenDataTbl(GridIni.Row).iBunrui_Tyu = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
           (GamenDataTbl(GridIni.Row).iBunrui_Sho = KikiDataTbl(iLoopCnt).iBunrui_Sho) And _
           (GamenDataTbl(GridIni.Row).iBunrui_Corner = KikiDataTbl(iLoopCnt).iBunrui_Corner) Then
' EG20 V5.12.0.1追加終了
            '機器構成情報データ保存
            '動的配列の内容をログパラメータ構造体の静的配列に格納する。
            Erase KikiDataTbl(iLoopCnt).strData
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
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtDummy_KeyDown
'//  機能名称  : キーボード押下時のイベントプロシージャ
'//  機能概要  : ダミーテキストのセット
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                 カーソル移動の処理を削除
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iColSave        As Integer          '列保存エリア
    Dim iRowSave        As Integer          '行保存エリア
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '特殊キーを押下された時、下記の処理を行う
    bScroll = True
    On Err GoTo ShoriErr
    
    With GridIni
        'V1.20.0.1 DEL START
'        '↑を押下された時
'        If KeyCode = 38 Then
'            If .Row <> 0 And .Row <> 1 Then
'                'セルを上に一つ移動
'                .Row = .Row - 1
'            End If
'        '↓、またはenterを押下された時
'        ElseIf KeyCode = 40 Or KeyCode = 13 Then
        'V1.20.0.1 DEL END
        If KeyCode = 13 Then    'V1.20.0.1 ADD

            iColSave = .Col
            iRowSave = .Row
            .Col = 0
            .Row = .Row + 1
            If .Text = "" Then
                .Col = iColSave
                .Row = iRowSave
                Exit Sub
            End If
            .Col = iColSave
            .Row = iRowSave
            
            If .Row <> .Rows - 1 Then
                'セルを下に一つ移動
                .Row = .Row + 1
            End If
        'V1.20.0.1 DEL START
'        '←、または→を押下された時
'        ElseIf KeyCode = 37 Or KeyCode = 39 Then
'            Exit Sub
        'V1.20.0.1 DEL END
        End If

        'ダミーテキストのセット
        txtDummy.Left = .Left + .CellLeft
        txtDummy.Top = .Top + .CellTop
        txtDummy.Width = .CellWidth
        txtDummy.Height = .CellHeight
        txtDummy.Text = .Text
        txtDummy.Visible = True
        txtDummy.SetFocus
    End With
    bScroll = False

ShoriErr:

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtDummy_Change
'//  機能名称  : ダミーテキストからフォーカスが移動した時のイベントプロシージャ
'//  機能概要  : ダミーテキストを非表示にする
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtDummy_LostFocus()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'ダミーテキストを非表示にする
    txtDummy.Visible = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　「自改画面へ」釦処理追加
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                設定反映釦の未押下メッセージ追加
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmdKikiSetMenu_Click(Index As Integer)
    Dim iResponse           As Integer          'MsgBox戻り値   'V1.20.0.1 ADD
    Dim bUnlock             As Boolean          ' ロック解除フラグ      ' EG20 V3.0.0.2 追加

    'エラールーチンを宣言
    On Error Resume Next
    
'V1.12.0.1 ADD START
    '全ボタンを押下不可とする。
    Call SetEnableFalse
'V1.12.0.1 ADD END
 ' EG20 V3.0.0.2 追加開始
' 押下した釦に応じてロック解除を制限する
' ※メール受信を待つため
    bUnlock = True
' EG20 V3.0.0.2 追加終了
   
    Select Case Index
        
        Case 0                                 ' 機器構成項目設定反映
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_INSTOL, 0)
            
            '機器構成項目設定反映処理
            Call sInstolKikiData
            bUnlock = False                     ' EG20 V3.0.0.2 追加

        Case 1                                 ' 機器構成項目媒体出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_OUTPUT, 0)
            
            '機器構成項目媒体出力処理
            Call sKikiDataOutPut
    
        Case 2                                 ' 機器構成項目内部保存
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_SAVE, 0)
            
            '機器構成項目内部保存処理
            Call sKikiDataSave
        
        Case 3                                 ' 機器構成設定データ選択
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_SELECT, 0)
            
            '機器構成設定データ選択処理
            Call sKikiDataSelect
    
        Case 4                                 ' 媒体入力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_MEDIUM_INPUT, 0)
            
            '媒体入力処理
            Call sInputMedium
    
        Case 5                                 ' 媒体取外
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
            
            '媒体取外処理
            Call pfRemove(Me)
'V1.4.0.1 ADD START
        Case 6                                 ' 自改画面へ
            'V1.20.0.1 ADD START
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
            'V1.20.0.1 ADD END
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)
            Unload Me
            Load frmKikiDataGate
            frmKikiDataGate.Show 1
            Exit Sub     'V1.20.0.1 ADD
'V1.4.0.1 ADD END
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        Case 7                                 ' エンコードコーナ号機画面へ
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
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_SUBGATE, 0)
            
            '表示中画面アンロード
            Unload Me
            
            'エンコードコーナ号機画面表示
            Load frmKikiDataSubGate
            frmKikiDataSubGate.Show 1
            Exit Sub
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

        Case Else
            '処理なし
            
    End Select
'V1.12.0.1 ADD START
    '全ボタンを押下可とする。
' EG20 V3.0.0.2 追加開始
    If bUnlock = True Then
        Call SetEnableTrue
    End If
' EG20 V3.0.0.2 追加終了
'    Call SetEnableTrue                 ' EG20 V3.0.0.2 削除
'V1.12.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.10.0.1) 2009-10-23   REVISED BY [TCC] D.Yamashita
'//                 フェーズ３残件項目対応　キャンセル不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-16   REVISED BY [TCC] C.Terui
'//                 コンピュータ名、ネットワーク変更処理追加
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                設定反映フラグ追加
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ポップアップ確認画面のタイトル修正
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.76修正対応】
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-08  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sInstolKikiData()

    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bRet                As Boolean          '関数戻り値
    Dim lErrCode            As Long             'エラーコード
    Dim strFileName         As String           '媒体ファイル名
    
    Dim bData()             As Byte             'バイナリデータ
    Dim iLoopCnt            As Integer          'ループカウンタ
    Dim bSysChange          As Boolean          'コンピュータ名、ネットワーク変更処理判定   'V1.12.0.1 ADD
    
    Dim lResult             As Long             ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加
' EG20 V30.0.3.1 追加開始（計算に利用する変数をLONG型に変更）
    Dim lLoop               As Long             ' ループカウンタ
    Dim lRecord             As Long             ' レコード
    Dim lIndex              As Long             ' インデックス
    Dim lSize               As Long             ' サイズ
' EG20 V30.0.3.1 追加終了（計算に利用する変数をLONG型に変更）
    
    'エラールーチンを宣言
    On Error Resume Next
'V1.8.0.1 DEL START
'    iResponse = MsgBox("機器構成データをインストールします。よろしいですか？" & Chr(vbKeyReturn) & _
'                        "反映は再起動後になります。", _
'                        vbYesNo + vbExclamation, _
'                        "媒体入力確認")
'V1.8.0.1 DEL END
'V1.8.0.1 ADD START
'V1.21.0.1 DEL START
'    iResponse = MsgBox("機器構成データをインストールします。よろしいですか？" & Chr(vbKeyReturn) & _
'                        "反映は再起動後になります。", _
'                        vbOKCancel + vbExclamation, _
'                        "媒体入力確認")
'V1.21.0.1 DEL END
'V1.21.0.1 ADD START
    iResponse = MsgBox("機器構成データをインストールします。よろしいですか？" & Chr(vbKeyReturn) & _
                        "反映は再起動後になります。", _
                        vbOKCancel + vbExclamation, _
                        "設定反映確認")
'V1.21.0.1 ADD END
'V1.8.0.1 ADD END
'    If iResponse = vbNo Then Exit Sub      'V1.10.0.1 DEL
    If iResponse = vbCancel Then
        Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
        Exit Sub   'V1.10.0.1 ADD
    End If
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
' EG20 V30.3.0.1 削除開始（計算に利用する変数をLONG型に変更）
'    '構造体配列をバイナリ配列に変換
'    ReDim bData((UBound(KikiDataTbl) + 1) * Len(KikiDataTbl(0))) As Byte
'    For iLoopCnt = 0 To UBound(KikiDataTbl)
'          MoveMemory bData(iLoopCnt * Len(KikiDataTbl(0))), KikiDataTbl(iLoopCnt), Len(KikiDataTbl(iLoopCnt))
'    Next
' EG20 V30.3.0.1 削除終了（計算に利用する変数をLONG型に変更）
' EG20 V30.3.0.1 追加開始（計算に利用する変数をLONG型に変更）
    lSize = Len(KikiDataTbl(0))
    lRecord = UBound(KikiDataTbl)
    ReDim bData((lRecord + 1) * lSize) As Byte
    For lLoop = 0 To lRecord
        lIndex = lLoop * lSize
        MoveMemory bData(lIndex), KikiDataTbl(lLoop), lSize
    Next
' EG20 V30.3.0.1 追加終了（計算に利用する変数をLONG型に変更）
    
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
        'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果") '1.21.0.1 DEL
        iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理設定反映結果") 'V1.21.0.1 ADD
        Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
    Else
'        'ログ出力
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'
'        '正常終了
'        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体入力結果")
'    End If
'V1.12.0.1 ADD START
            'コンピュータ名、ネットワーク変更処理
            bSysChange = pfNetWorkChng(Me)
            If bSysChange = False Then
                '異常ログ出力
                Call pfOutPutErrLog(lErrCode)

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
                Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
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
                'iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体入力結果")  'V1.21.0.1 DEL
                 iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "反映処理設定反映結果")  'V1.21.0.1 ADD
                '設定反映フラグ（変更なし）
                SetteiHaneiFlg = False      'V1.20.0.1 ADD
                Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
            End If
        End If
'V1.12.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 ディスク情報取得位置変更
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                  ボタン名称変更によるポップアップ変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
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
    Dim lDrive               As Long               'ドライブのクラスタ数（合計）
    Dim strDrive             As String       'ドライブ
    
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
    ''V1.20.0.1 DEL START
    ''ディスク情報を取得
''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD
    '
    'If lDrive = 0 Then
    '    strDrive = "d:"
    'Else
''        strDrive = "a:"    'V1.12.0.1 DEL
    '    strDrive = "H:"     'V1.12.0.1 ADD
    'End If
    'V1.20.0.1 DEL END
    
    'sWriteDir = pfDirSelection(strDrive, "機器構成ファイル書込み先のディレクトリ選択")    'V1.20.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
    If sWriteDir <> "" Then
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
'        FileCopy KIKI_DATA_FILE, sWriteDir & Dir(KIKI_DATA_FILE)                                       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        FileCopy KIKI_DATA_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(KIKI_DATA_FILE)    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '正常終了
        'iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "機器構成項目媒体出力結果") 'V1.13.0.1 DEL
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体出力結果")              'V1.13.0.1 ADD
    
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

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    '異常終了
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "機器構成項目媒体出力結果") 'V1.8.0.1 DEL
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "機器構成項目媒体出力結果") 'V1.8.0.1 ADD 'V1.13.0.1 DEL
     iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体出力結果") 'V1.13.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 ファイル検索処理削除
'//     REVISIONS :(1.13.0.1) 2009-11-19  REVISED BY [TCC] S.Terao
'//                 釦名変更による、ポップアップタイトル変更
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSave()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
'    Dim sMyPath(1 To 3)      As String          'ファイルパス      'V1.12.0.1 DEL
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim iLoopCount           As Integer         'ループカウンタ
    Dim intFileNo            As Integer         'ファイル番号

    'エラールーチンを宣言
    On Error Resume Next
    
    '初期値設定
    strFileName = ""

'V1.12.0.1 DEL START
'    '----------------------------------------------------
'    '機器構成データファイル検索
'    '----------------------------------------------------
'    strFileName = Dir(KIKI_DATA_FILE)
'
'    'ファイルが存在しない場合
'    If strFileName = "" Then
'
'        '異常ログ出力
'        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_KIKI_DATA, 0)
'
'        '異常終了
'        MsgBox "媒体出力するデータがありません。", _
'                vbOKOnly + vbExclamation, _
'                 "データ無警告"
'        Exit Sub
'
'    End If
'
'    '----------------------------------------------------
'    '内部保存ファイル検索
'    '----------------------------------------------------
'    For iLoopCount = 1 To 3
'
'        'ファイルパス取得
'        sMyPath(iLoopCount) = Replace(KIKI_DATA_S_FILE, "##", Format(iLoopCount, "0#"))
'
'        'ファイル検索
'        strFileName = Dir(sMyPath(iLoopCount))
'
'        'ファイルが存在しない場合
'        If strFileName = "" Then
'
'            intFileNo = FreeFile                                        '未使用のファイル番号を取得する
'            Open sMyPath(iLoopCount) For Output Access Write As #intFileNo
'            Close #intFileNo
'
'        End If
'
'    Next
'V1.12.0.1 DEL END

    '----------------------------------------------------
    '内部保存処理
    '----------------------------------------------------
'V1.12.0.1 ADD START
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
'V1.12.0.1 ADD END
    
    'ファイル名取得
'    sWriteDir = pfDispFileSelect("d:", FOLDER_KIKI_DATA, FILE_NAME_KIKI_DATA_S, "内部保存ﾌｧｲﾙ選択")    'V1.12.0.1 DEL
    sWriteDir = KIKI_DATA_S_FILE  'V1.12.0.1 ADD
    
    If sWriteDir <> "" Then
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
        FileCopy KIKI_DATA_FILE, sWriteDir
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
'V1.12.0.1 ADD START
        '一時保存ファイル削除
        Kill KIKI_DATA_S_TEMP_FILE
'V1.12.0.1 ADD END
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '正常終了
        'iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "機器構成項目内部保存結果")　　'V1.13.0.1 DEL
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "一時保存結果")     'V1.13.0.1 ADD
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

'V1.12.0.1 ADD START
        'ファイル検索
        strFileName = Dir(KIKI_DATA_S_FILE)
        If strFileName <> "" Then
            'ファイル削除
            Kill KIKI_DATA_S_FILE
        End If
        'ファイル名称を元に戻す
        Name KIKI_DATA_S_TEMP_FILE As KIKI_DATA_S_FILE
'V1.12.0.1 ADD END

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    '異常終了
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "機器構成項目内部保存結果") 'V1.8.0.1 DEL
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "機器構成項目内部保存結果")    'V1.8.0.1 ADD  'V1.13.0.1 DEL
     iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存結果")    'V1.13.0.1 ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 ファイル検索処理削除
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                 コピーファイルパス指定を修正
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                設定反映フラグ追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSelect()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
'    Dim sMyPath(1 To 3)      As String          'ファイルパス      'V1.12.0.1 DEL
    Dim sMyPath              As String          'ファイルパス       'V1.12.0.1 ADD
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim iLoopCount           As Integer         'ループカウンタ
    Dim intFileNo            As Integer         'ファイル番号
    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード

    'エラールーチンを宣言
    On Error Resume Next
    
'V1.12.0.1 DEL START
'    '----------------------------------------------------
'    '内部保存ファイル検索
'    '----------------------------------------------------
'    For iLoopCount = 1 To 3
'
'        'ファイルパス取得
'        sMyPath(iLoopCount) = Replace(KIKI_DATA_S_FILE, "##", Format(iLoopCount, "0#"))
'
'        '初期値設定
'        strFileName = ""
'
'        'ファイル検索
'        strFileName = Dir(sMyPath(iLoopCount))
'
'        'ファイルが存在しない場合
'        If strFileName = "" Then
'
'            intFileNo = FreeFile                                        '未使用のファイル番号を取得する
'            Open sMyPath(iLoopCount) For Output Access Write As #intFileNo
'            Close #intFileNo
'        End If
'
'    Next
'V1.12.0.1 DEL END

    '----------------------------------------------------
    '機器構成データファイル更新処理
    '----------------------------------------------------
'V1.12.0.1 ADD START
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
'V1.12.0.1 ADD END
    
    'ファイル名取得
'    sWriteDir = pfDispFileSelect("d:", FOLDER_KIKI_DATA, FILE_NAME_KIKI_DATA_S, "機器構成ﾌｧｲﾙ選択")    'V1.12.0.1 DEL
'V1.12.0.1 ADD START
    strFileName = Dir(KIKI_DATA_S_FILE)
    sWriteDir = strFileName
'V1.12.0.1 ADD START
    If sWriteDir <> "" Then

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
        'FileCopy sWriteDir, KIKI_DATA_FILE　'V1.13.0.1 DEL
         FileCopy KIKI_DATA_S_FILE, KIKI_DATA_FILE  'V1.13.0.1 ADD
        
        '機器情報設定（駅情報）イメージファイル作成
        bRet = dllGetKikiIniData(0, 1, KIKI_DATA_SET_EKI_INFO_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
'V1.12.0.1 ADD START
            'ファイル削除
            Kill KIKI_DATA_FILE
            'ファイル名称を元に戻す
            Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
'V1.12.0.1 ADD END

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
            
            '異常終了
            'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "媒体入力結果") 'V1.8.0.1 DEL
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果")  'V1.8.0.1 ADD
            Exit Sub
        End If
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'V1.12.0.1 ADD START
        '一時保存ファイルを削除する
        Kill KIKI_DATA_BACKUP_FILE
'V1.12.0.1 ADD END
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
        '正常終了
'        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "機器構成設定データ選択結果") 'V1.13.0.1 DEL
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "一時保存データ取込結果")      'V1.13.0.1 ADD
    
        '機器情報データ更新フラグ設定（更新設定）
        KikiDataUpDateFlg = True
        '画面表示処理
        Call sDisp
        '機器情報データ更新フラグ設定（通常設定）
        KikiDataUpDateFlg = False
        
        '設定反映フラグ（変更あり）
        SetteiHaneiFlg = True       'V1.20.0.1 ADD
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

'V1.12.0.1 ADD START
        'ファイル検索
        strFileName = Dir(KIKI_DATA_FILE)
        If strFileName <> "" Then
            'ファイル削除
             Kill KIKI_DATA_FILE
        End If
        'ファイル名称を元に戻す
        Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
'V1.12.0.1 ADD END
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '異常終了
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "機器構成設定データ選択結果") 'V1.8.0.1 DEL
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "機器構成設定データ選択結果") 'V1.8.0.1 ADD  'V1.13.0.1 DEL
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存データ取込結果")      'V1.13.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 媒体ファイル名を固定名称に変更
'//                 ディスク情報取得位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                設定反映フラグ追加
'//                ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.77修正対応】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
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
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト  'V1.20.0.1 ADD
    
    'エラールーチンを宣言
    On Error Resume Next
'V1.12.0.1 ADD START
    iResponse = MsgBox("機器構成設定の媒体入力を行います。" & vbCrLf & "よろしいですか？", _
    vbOKCancel + vbQuestion, "媒体入力確認")
    
    'V1.20.0.1 DEL START
    'If iResponse = vbCancel Then Exit Sub
''V1.12.0.1 ADD END
    '
    ''ディスク情報を取得
''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD
    '
    'If lDrive = 0 Then
    '    strDrive = "d:"
    'Else
''        strDrive = "a:"        'V1.12.0.1 DEL
    '    strDrive = "H:"         'V1.12.0.1 ADD
    'End If
    '
    ''媒体ファイル名取得
''    strFileName = pfFileSelection(strDrive, "*.csv", "媒体入力ﾌｧｲﾙ選択")          'V1.12.0.1 DEL
    'strFileName = pfFileSelection(strDrive, "KIKI_DATA.CSV", "媒体入力ﾌｧｲﾙ選択")   'V1.12.0.1 ADD
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
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
    'V1.20.0.1 ADD END
    
    Call ChDrive("D")  'V2.5.0.1 ADD
    
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
        
        '機器情報設定（駅情報）イメージファイル作成
        bRet = dllGetKikiIniData(0, 1, KIKI_DATA_SET_EKI_INFO_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
            
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
            
            '異常終了
            'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "媒体入力結果") 'V1.8.0.1 DEL
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果") 'V1.8.0.1 ADD
            
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
        SetteiHaneiFlg = True       'V1.20.0.1 ADD
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
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "媒体入力結果") 'V1.8.0.1 DEL
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果")  'V1.8.0.1 ADD

End Sub
'V1.12.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
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
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
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
    CmdKikiSetMenu(7).Enabled = False               ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    cmdCancel.Enabled = False
    
    'コンボボックス、CmdKikiSetMenu(0)～(2)は条件によっては元々押下不可のため判定を行う
    If cmbEkiInfo.Enabled = True Then
        cmbEkiInfo.Enabled = False
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
    
    DoEvents
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
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
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
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
    CmdKikiSetMenu(7).Enabled = True                ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    cmdCancel.Enabled = True
    
    'コンボボックスとCmdKikiSetMenu(0)～(2)は条件によっては元々押下不可のため、画面表示の有無で判定を行う
    strFileName = Dir(KIKI_DATA_SET_EKI_INFO_FILE)
    'ファイルが存在する場合
    If strFileName <> "" Then
        cmbEkiInfo.Enabled = True
        CmbCornerName.Enabled = True                ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        CmdKikiSetMenu(0).Enabled = True
        CmdKikiSetMenu(1).Enabled = True
        CmdKikiSetMenu(2).Enabled = True
    End If
    
    DoEvents
    
End Sub
'V1.12.0.1 ADD END

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

