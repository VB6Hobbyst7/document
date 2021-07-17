VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEkiData 
   BorderStyle     =   0  'なし
   Caption         =   "駅都度データ確認"
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
      TabIndex        =   12
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton CmdMoveSubGateGamen 
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
      Left            =   7080
      TabIndex        =   11
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMoveGateGamen 
      Caption         =   "改札機画面へ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7080
      TabIndex        =   10
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "テキスト媒体出力（駅情報）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   7800
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "駅設定出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "駅設定入力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
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
      Height          =   555
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   8880
      Top             =   960
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
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   10425
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "  機器情報設定    画面へ戻る"
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
      Left            =   9500
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   6195
      Left            =   120
      TabIndex        =   5
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
      FocusRect       =   0
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
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "駅都度データ確認（駅情報）"
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
      TabIndex        =   3
      Top             =   0
      Width           =   12015
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
      Height          =   388
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   8295
   End
End
Attribute VB_Name = "frmEkiData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：駅都度データ確認（駅情報）画面.frm
'//  パッケージ名：駅都度データ確認（駅情報）画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-04-06   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応
'//                   「駅設定出力」「駅設定入力」
'//　　　　　　　　　 「駅設定テキスト出力」「媒体取外」釦処理追加
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//                 駅設定ファイル書込み先ディレクトリ変更
'//                 ディスク情報取得位置変更
'//                 テキスト出力内容変更
'//                 画面ロック処理／画面ロック解除処理追加
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                フォルダ選択画面での「取消」釦押下処理追加
'//                「テキスト媒体出力(駅情報)」釦押下処理修正
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                画面再前面表示修正(不具合修正)
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.76修正対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   'メイルタイマのインターバル値
' Private Const TITOL_EKI_NAME = "駅名　　　："           '駅名タイトル     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
Private Const TITOL_EKI_NAME = "駅名："                 '駅名タイトル       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

'V1.12.0.1 ADD START
'機器構成データ（駅情報）イメージファイル読取用の構造体
Private Type EKIINFO_IMAGE_FILE
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
'V1.12.0.1 ADD END
Private gstrFileName        As String                       ' 出力ファイル名    ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 駅都度データ確認（駅情報）画面(アクティブ時：イベントプロシージャ)
'//  機能概要  : 最前前表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'エラールーチンを宣言
    On Error Resume Next
    
    '自画面最前面表示処理を行う。
    pfFormActive (hwnd)
    
    'タイマを起動する
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 駅都度データ確認（駅情報）画面(ディアクティブ時)
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
'//  機能名称  : 駅都度データ確認（駅情報）画面(ロード時：イベントプロシージャ)
'//  機能概要  : 初期処理を行う。
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
Private Sub Form_Load()

    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_EKIINFO_GAMEN_START, 0)
    
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
    
    '駅都度データ確認（駅情報）イメージファイル作成
    bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '駅都度データ確認（駅情報）イメージファイル削除
        Kill EKI_TUDO_CHK_EKI_INFO_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
    End If
    
    '駅情報コンボボックス初期値設定
    cmbEkiInfo.Clear
    cmbEkiInfo.AddItem "駅情報"
    cmbEkiInfo.AddItem "監視"
    cmbEkiInfo.AddItem "ネットワーク"
    cmbEkiInfo.AddItem "画面"                           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    cmbEkiInfo.ListIndex = 0
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    'コーナ設定コンボボックスの初期化処理
    Call InitCornerComboBox
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
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
                AppActivate frmEkiData.Caption, False           ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
                pfFormActive (frmEkiData.hwnd)                  ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '「保守操作卓プログラム送信要求」を受信した場合
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
                    'プログレスバーを消去する
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
                    Kill gstrFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
                    Call SetEnableTrue
                Else
                    Call pfuncInstallEkiSettei
                End If
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_EKIINFO_GAMEN_END, 0)
    
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
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim strFileName          As String          'ファイル名
    Dim iLoopCnt             As Integer         'ループカウンタ
    Dim bRet                 As Boolean         '関数戻り値
    Dim strKubun             As String          '区分
    Dim strIniData           As String          'INIファイル設定値
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
    LblEkiName.Caption = TITOL_EKI_NAME         '駅名ラベル初期化
    
    '----------------------------------------------------
    'グリッドタイトル設定
    '----------------------------------------------------
    Call sDispGridTitol
    
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

        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '駅名ラベル更新
    '----------------------------------------------------
'   LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo              'V2.1.0.1 DEL
    LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo(NotEkiVer)   'V2.1.0.1 ADD
    
    '駅都度データ確認（駅情報）イメージファイル検索
    strFileName = Dir(EKI_TUDO_CHK_EKI_INFO_FILE)
    
    'ファイルが存在した場合
    If strFileName <> "" Then
    
        'グリッドデータ部設定
'        Call sDispDataSet(cmbEkiInfo.ListIndex + 1)                    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        nCornerIndex = CmbCornerName.ListIndex
        Call sDispDataSet(pfGetCodeDaiBunrui(cmbEkiInfo), nCornerIndex + 1)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        
        'INIデータチェック
        With GridIni
            For iLoopCnt = 1 To .Rows - 1
            
                .Row = iLoopCnt
'V1.11.0.1 DEL START
'                .Col = 3
'
'                'データチェック
'                bRet = pfDispDataChk(.Text)
'                If bRet = False Then
'                    'データ不一致の場合、セルの背景色を赤色にする
'                    .CellBackColor = QBColor(12)
'                End If
'V1.11.0.1 DEL END
                'V1.11.0.1 ADD START
                .Col = 0
                If .Text <> "" Then
                    .Col = 3
                    
                    'データチェック
                    bRet = pfDispDataChk(.Text)
                    If bRet = False Then
                        'データ不一致の場合、セルの背景色を赤色にする
                        .CellBackColor = QBColor(12)
                    End If
                End If
                'V1.11.0.1 ADD END
            Next
        End With
    
        cmbEkiInfo.Enabled = True                  '駅情報コンボボックス選択可設定
        CmbCornerName.Enabled = True               ' コーナ選択部選択不可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKIINFO_IMAGE, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear(1, GridIni.Rows)
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
        .Cols = 5
        
        '----------------------------------
        'グリッド幅設定
        '----------------------------------
        .ColWidth(0) = 500
        .ColWidth(1) = 3800
        .ColWidth(2) = 730
        .ColWidth(3) = 2700
        .ColWidth(4) = 3700
        
        '----------------------------------
        'タイトル設定
        '----------------------------------
        '項目設定
        .Col = 1
        .Row = 0: .Text = "項目"
        .CellAlignment = flexAlignCenterCenter

        '区分設定
        .Col = 2
        .Text = "区分"
        .CellAlignment = flexAlignCenterCenter

        '設定値設定
        .Col = 3
        .Text = "設定値"
        .CellAlignment = flexAlignCenterCenter

        '詳細設定
        .Col = 4
        .Text = "設定値詳細"
        .CellAlignment = flexAlignCenterCenter
        
'        .RowHeight(0) = 700        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        .RowHeight(0) = 440         ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    End With
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sDispDataClear
'//  機能名称  : グリッドデータ部クリア処理
'//  機能概要  : グリッドデータ部をクリアする
'//
'//              型        名称      意味
'//  引数      : Integer   intStartRow  開始行位置
'//              Integer   intEndRow    終了行位置
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear(intStartRow As Integer, intEndRow As Integer)
    
    Dim iLoopCnt             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    'グリッド初期化
    With GridIni

        .Rows = intEndRow   'V1.11.0.1 ADD
        For iLoopCnt = intStartRow To intEndRow - 1

            '通番設定
            .Col = 0
            .Row = iLoopCnt: .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '項目設定
            .Col = 1
            .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '区分設定
            .Col = 2
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'            .Text = "" & vbCrLf & _
'                    "" & vbCrLf & _
'                    ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
            
            .CellAlignment = flexAlignCenterCenter

            '設定値設定
            .Col = 3
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'            .Text = "" & vbCrLf & _
'                    "" & vbCrLf & _
'                    ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
            .CellAlignment = flexAlignLeftCenter

            '詳細設定
            .Col = 4
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'            .Text = "" & vbCrLf & _
'                    "" & vbCrLf & _
'                    ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
            .CellAlignment = flexAlignLeftCenter

            .RowHeight(iLoopCnt) = 938
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
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応
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
    Dim iRowCnt             As Integer                      ' 行カウンタ
    Dim iBunrui_Sho()      As Integer                       ' 小分類テーブル
    
    Dim strBunrui_Dai       As String                       ' 大分類
    Dim strBunrui_Tyu       As String                       ' 中分類
    Dim strBunrui_Sho       As String                       ' 小分類
    Dim strNo               As String                       ' 通番
    Dim strKomoku           As String                       ' 項目
    Dim strKubun            As String                       ' 区分
    Dim strData             As String                       ' 設定値
    Dim strSetShosai        As String                       ' 設定値詳細
    Dim strCorner           As String                       ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim iCmpCorner          As Integer                      ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '初期化
    ReDim iBunrui_Sho(0)

    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '駅都度データ確認（駅情報）イメージファイルをオープンする。
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber
    
    GridIni.Visible = False
    iLoopCnt = 1
    GridIni.Rows = 1
    Do While Not EOF(intFileNumber)
    
        '１ 行読み込み
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strNo, _
'                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strCorner, strNo, _
                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        
        If iBunrui_Dai = strBunrui_Dai Then

' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        ' コーナ判定追加
        ' 選択したコーナ、もしくはコーナ無関係のレコードは採用する
        iCmpCorner = CInt(strCorner)
        If ((iCorner = iCmpCorner) Or (iCmpCorner = 0)) Then
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        
            'グリッド初期化
            With GridIni
            
                '項目検索
                For iRowCnt = 0 To iLoopCnt - 2
                    If CStr(iBunrui_Sho(iRowCnt)) = strBunrui_Sho Then Exit For
                Next
                
                '項目が見つからなかった場合
                If iRowCnt = .Rows - 1 Then
                    
                    '小分類テーブル登録
                    ReDim Preserve iBunrui_Sho(.Rows - 1)
                    iBunrui_Sho(iRowCnt) = CInt(strBunrui_Sho)
                    
                    '表示行数をインクリメント
                    iLoopCnt = iLoopCnt + 1
                    .Rows = iLoopCnt
                    
                End If
            
                '表示データが１画面に表示しきれない場合
                If .Rows > 6 Then
                    'スクロールバー分、グリッドを広げる
                    .Width = 11775
                End If

                '通番設定
                .Col = 0
                .Row = iLoopCnt - 1: If .Text = "" Then .Text = CStr(iLoopCnt - 1)
                .CellAlignment = flexAlignLeftCenter
    
                '項目設定
                .Col = 1
                If .Text = "" Then .Text = strKomoku
                .CellAlignment = flexAlignLeftCenter
    
                '区分設定
                .Col = 2
                .Text = pfDispAplBunrui(.Text, strKubun)
                .CellAlignment = flexAlignCenterCenter
    
                '設定値設定
                .Col = 3
                .Text = pfDispIniData(.Text, strData, strKubun)
                .CellAlignment = flexAlignLeftCenter
    
                '詳細設定
                .Col = 4
                If .Text = "" Then .Text = strSetShosai
                .CellAlignment = flexAlignLeftCenter
    
                .RowHeight(iLoopCnt - 1) = 938
        
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
    If GridIni.Rows < 7 Then
        'グリッドデータ部クリア処理
'        Call sDispDataClear(GridIni.Rows - 1, 9)   'V1.11.0.1 DEL
        Call sDispDataClear(GridIni.Rows, 7)        'V1.11.0.1 ADD
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
    Call sDispDataClear(1, GridIni.Rows - 1)

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
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_GAMEN_EKIINFO_SELECT, 0)
    
    '画面表示処理
    Call sDisp

'V1.12.0.1 ADD START
    '全ボタンを押下可とする。
    Call SetEnableTrue
'V1.12.0.1 ADD END

End Sub

'V1.4.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : CmdMenu_Click
'//  機能名称  : 「駅設定出力」「駅設定入力」「駅設定テキスト出力」
'//              「媒体取外」釦押下処理
'//  機能概要  : 各釦名称処理を行う。
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmdMenu_Click(Index As Integer)
  
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
       Case 0                                  '駅設定出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_OUTPUT, 0)
            
            '駅設定出力処理
            Call sEkiSetteiOutPut
        
        Case 1                                  '駅設定入力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_INPUT, 0)
            
            '駅設定入力処理
            Call sInstolEkiSettei
            bUnlock = False                     ' EG20 V3.0.0.2 追加
        
        Case 2                                  '駅設定テキスト出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_DISP_TEXT, 0)
            
            '駅設定テキスト出力処理
            Call sDispTextEkiDataNow
        
        Case 3                                  '媒体取外
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
            
            '媒体取外処理
            Call pfRemove(Me)
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
'//  関数名称  : sEkiSetteiOutPut
'//  機能名称  : 「駅設定出力」釦押下時処理
'//  機能概要  : 現在駅設定ファイルを外部媒体に出力する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 駅設定ファイル書込み先ディレクトリ変更
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
Private Sub sEkiSetteiOutPut()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
    Dim iResponse            As Integer         'MsgBox戻り値

    'エラールーチンを宣言
    On Error Resume Next
'V1.8.0.1 DEL START
'    iResponse = MsgBox("選択されている駅の現在の駅都度データ１駅分を出力します。" & Chr(vbKeyReturn) & _
'                        "よろしいですか？", _
'                        vbYesNo + vbQuestion, _
'                        "駅設定出力確認")
'V1.8.0.1 DEL END
'V1.8.0.1 ADD START
    iResponse = MsgBox("選択されている駅の現在の駅都度データ１駅分を出力します。" & Chr(vbKeyReturn) & _
                        "よろしいですか？", _
                        vbOKCancel + vbQuestion, _
                        "駅設定出力確認")
'V1.8.0.1 ADD END
'    If iResponse = vbNo Then Exit Sub              'V1.12.0.1 DEL
    If iResponse = vbCancel Then Exit Sub           'V1.12.0.1 DEL
    
'初期値設定
    strFileName = ""

    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '異常終了
        MsgBox "媒体出力するデータがありません。", _
                vbOKOnly + vbExclamation, _
                 "データ無警告"
        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '媒体出力処理
    '----------------------------------------------------
'    sWriteDir = pfDirSelection("a:", "駅設定ファイル書込み先のディレクトリ選択")   'V1.12.0.1 DEL
    'sWriteDir = pfDirSelection("H:", "駅設定ファイル書込み先のディレクトリ選択")    'V1.12.0.1 ADD     'V1.20.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
    If sWriteDir <> "" Then
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
'        FileCopy EKI_SETTI_FILE, sWriteDir & Dir(EKI_SETTI_FILE)       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        FileCopy EKI_SETTI_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(EKI_SETTI_FILE)        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
       '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定出力結果")
    
    End If
    
  Exit Sub
 
COPY_ERROR:

    Select Case Err.Number
        Case 61 ' 媒体出力空き容量不足
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case 71 ' 媒体なし
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_NOT_DISK, 0)
        Case Else
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    iResponse = MsgBox("異常終了しました", vbOKOnly + vbCritical, "駅設定出力結果")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sInstolEkiSettei
'//  機能名称  : 「駅設定入力」釦押下時処理
'//  機能概要  : 外部媒体から現在駅設定ファイルインストールする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 ディスク情報取得位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.76修正対応】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sInstolEkiSettei()

    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bRet                As Boolean          '関数戻り値
    Dim strFileName         As String           '媒体ファイル名

    Dim objFso As New FileSystemObject          'ファイルシステムオブジェクト

    Dim lResult             As Long             ' 処理結果

    'エラールーチンを宣言
    On Error Resume Next
    iResponse = MsgBox("駅都度データ１駅分をインストールします。" & Chr(vbKeyReturn) & _
                        "よろしいですか？", _
                        vbOKCancel + vbQuestion, _
                        "駅設定入力確認")
    If iResponse = vbCancel Then
        Set objFso = Nothing
        Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
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
    ' 拡張子を設定
    CommonDialog1.Filter = "ＣＳＶ（カンマ区切り）(*.csv)|*.csv|"
    ' ファイル選択画面を開く
    CommonDialog1.ShowOpen
    ' 選択したファイル名を取得
    strFileName = CommonDialog1.FileName
    
    Call ChDrive("D")  'V2.5.0.1 ADD

    'ファイル存在チェック
    If strFileName <> "" Then
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL Start
'        ' 出力先ファイル名を保存
'        gstrFileName = strFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL End
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
        ' 出力先ファイル名を保存
        gstrFileName = PATH_HOSHUWRK_EKI_INFO
        '一時保存フォルダにデータをコピーし読取専用を解除する
        If pfChangeAttrNormal(strFileName, PATH_HOSHUTMP_EKI_INFO, PATH_HOSHUWRK_EKI_INFO) = False Then
            Kill gstrFileName
            '一時保存フォルダを削除する
            psDeleteFolder PATH_HOSHUTMP
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' 異常終了
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
            Call SetEnableTrue
            Exit Sub
        End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End

        ' //////////////////////////////////////////////
        ' // 操作卓プログラム処理
        ' //////////////////////////////////////////////
        lResult = pubfuncTakuProgramData(2, gstrFileName)
        If lResult = 0 Then
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
            Kill gstrFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' 異常終了
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
            Call SetEnableTrue
            Exit Sub
        ElseIf lResult = 1 Then
            ' メール送信中
            ' ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
            Exit Sub
        End If


        ' //////////////////////////////////////////////
        ' // 統合監視盤非動作中のためメール応答を待たずに
        ' // 即時更新
        ' //////////////////////////////////////////////
        bRet = pfuncInstallEkiSettei

    End If
    Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
End Sub

' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]削除開始（全体見直し）
'Private Sub sInstolEkiSettei()
'
'    Dim iResponse           As Integer          'MsgBox戻り値
'    Dim bRet                As Boolean          '関数戻り値
'    Dim lErrCode            As Long             'エラーコード
'    Dim strFileName         As String           '媒体ファイル名
'
'    Dim iRet                    As Integer      'メッセージボックス戻り値
'    Dim lSekuta                 As Long         'セクタ（クラスタ当り）
'    Dim lByte                   As Long         'バイト数（セクタ当り）
'    Dim lKurasuta               As Long         'フリークラスタ数
'    Dim lDrive                  As Long         'ドライブのクラスタ数（合計）
'    Dim strDrive                As String       'ドライブ
'    Dim bSysChange              As Boolean      'システム設定処理戻り値　'V1.8.0.1　ADD
'    Dim bUpData                 As Boolean      '画面更新処理戻り値　　　'V1.8.0.1　ADD
'
'    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト  'V1.20.0.1 ADD
'
'    Dim lResult             As Long             ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加
'
'    'エラールーチンを宣言
'    On Error Resume Next
''V1.8.0.1 DEL START
''    iResponse = MsgBox("駅都度データ１駅分をインストールします。" & Chr(vbKeyReturn) & _
''                        "よろしいですか？", _
''                        vbYesNo + vbQuestion, _
''                        "駅設定入力確認")
''V1.8.0.1 DEL END
''V1.8.0.1 ADD START
'    iResponse = MsgBox("駅都度データ１駅分をインストールします。" & Chr(vbKeyReturn) & _
'                        "よろしいですか？", _
'                        vbOKCancel + vbQuestion, _
'                        "駅設定入力確認")
''V1.8.0.1 ADD END
'    'V1.20.0.1 DEL START
'''    If iResponse = vbNo Then Exit Sub              'V1.12.0.1 DEL
'    'If iResponse = vbCancel Then Exit Sub           'V1.12.0.1 ADD
'    '
'    ''ディスク情報を取得
'''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)  'V1.12.0.1 DEL
'    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)   'V1.12.0.1 ADD
'    '
'    'If lDrive = 0 Then
'    '    strDrive = "d:"
'    'Else
'''        strDrive = "a:"    'V1.12.0.1 DEL
'    '    strDrive = "H:"     'V1.12.0.1 ADD
'    'End If
'    '
'    ''媒体ファイル名取得
'    'strFileName = pfFileSelection(strDrive, "*.csv", "駅設定ﾌｧｲﾙ選択")
'    'V1.20.0.1 DEL END
'    'V1.20.0.1 ADD START
'    If iResponse = vbCancel Then
'        Set objFso = Nothing
'        Exit Sub
'    End If
'    '取得ファイル名を初期化
'    CommonDialog1.FileName = ""
'    '初期ディレクトリを設定
'    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'フォルダ選択画面デフォルトパス１が存在するか
'        '存在するため、デフォルトパス１（H:）を設定
'        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
'    Else
'        '存在しないため、デフォルトパス２（C:）を設定
'        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
'    End If
'    Set objFso = Nothing
'    '拡張子を設定
'    CommonDialog1.Filter = "ＣＳＶ（カンマ区切り）(*.csv)|*.csv|"
'    'ファイル選択画面を開く
'    CommonDialog1.ShowOpen
'    '選択したファイル名を取得
'    strFileName = CommonDialog1.FileName
'    'V1.20.0.1 ADD END
'
'    Call ChDrive("D")  'V2.5.0.1 ADD
'
'    'ファイル存在チェック
'    If strFileName <> "" Then
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'        'プログレスバーを表示する
'        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'        '現在駅設定データインストール処理
'        bRet = dllInstolEkiDataNow(strFileName, EKI_SETTI_FILE, lErrCode)
'
'        If bRet = False Then
'
'            '異常ログ出力
'            Call pfOutPutErrLog(lErrCode)
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'            'プログレスバーを消去する
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'            '異常終了
'            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
'
'        Else
''V1.8.0.1 ADD START
'            '----------------------------------------------------
'            'コンピュータ名、ネットワーク変更処理
'            '----------------------------------------------------
'            'Call pfNetWorkChng(Me)
'            bSysChange = True
'            bUpData = True
'            bSysChange = pfNetWorkChng(Me)
''V1.8.0.1 ADD END
'             'ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
''V1.8.0.1 ADD START
'           '駅都度データ確認（駅情報）イメージファイル作成
'            bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
'            If bRet = False Then
'                '駅都度データ確認（駅情報）イメージファイル削除
'                Kill EKI_TUDO_CHK_EKI_INFO_FILE
'                '異常ログ出力
'                Call pfOutPutErrLog(lErrCode)
'                bUpData = False
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'                'プログレスバーを消去する
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'            End If
'
'' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加開始
'            ' //////////////////////////////////////////////
'            ' // 操作卓プログラム処理
'            ' //////////////////////////////////////////////
'             lResult = pubfuncTakuProgramData(2)
'             If lResult = 0 Then
'                'プログレスバーを消去する
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'                ' 異常終了
'                iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "設定反映結果")
'                Call SetEnableTrue
'                Exit Sub
'             ElseIf lResult = 1 Then
'                ' メール送信中
'                ' ログ出力
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'
'                Exit Sub
'             End If
'' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加終了
'
'            '駅情報コンボボックス初期値設定
'            cmbEkiInfo.Clear
'            cmbEkiInfo.AddItem "駅情報"
'            cmbEkiInfo.AddItem "監視"
'            cmbEkiInfo.AddItem "ネットワーク"
'            cmbEkiInfo.AddItem "画面"                           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'            cmbEkiInfo.ListIndex = 0
'
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'            'コーナ設定コンボボックスの初期化処理
'            Call InitCornerComboBox
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'            'プログレスバーを消去する
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'            If bSysChange = True And bUpData = True Then
''V1.8.0.1 ADD END
'            '正常終了
'            iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定入力結果")
'            End If         'V1.8.0.1 ADD
'        End If
'    End If
'
'End Sub
' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]削除終了（全体見直し）

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sDispTextEkiDataNow
'//  機能名称  : 「駅設定テキスト出力」釦押下時処理
'//  機能概要  : 現在駅設定ファイルをテキスト表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 テキスト出力内容変更
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                フォルダ選択画面での「取消」釦押下処理追加
'//                「テキスト媒体出力(駅情報)」釦押下処理修正
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(EG20 V6.6.0.1)  2012-06-20  CODED BY  [TCC] H.Sugimoto
'//                 【選択コーナ別の出力対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispTextEkiDataNow()

    Dim strFileName          As String          'ファイル名
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim lRetVal              As Long            '戻り値
    Dim sCommand             As String          'コマンド文字列
'V1.12.0.1 ADD START
    Dim sWriteDir            As String              '書き込み先フォルダ名
    Dim intFileNumber        As Integer             'ファイルポインタ
    Dim strLineCount         As String              '行数カウンタ
    Dim i                    As Integer             'ループカウンタ１
    Dim j                    As Integer             'ループカウンタ２
    Dim k                    As Integer             'ループカウンタ３
    Dim ReadFileSettei()     As EKIINFO_IMAGE_FILE  'ファイル読込用構造体
    Dim fso         As New FileSystemObject         'ファイルシステムオブジェクト
    Dim FsoTS As TextStream

    Set fso = CreateObject("Scripting.FileSystemObject")
'V1.12.0.1 ADD END
'V1.13.0.1 ADD START
    Dim skansi               As String  '前回区分判定（監視用）
    Dim sidu                 As String  '前回区分判定（IDU用）
    Dim sldu                 As String  '前回区分判定（LDU用）
    Dim sTaku                As String  '前回区分判定（操作卓用）   ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim sSyoKoumoku          As String  '前回小項目判定
    Dim nProcMode            As Integer ' 現在処理分類              ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    Dim szCornerName         As String          ' コーナ名称        ' EG20 V6.6.0.1追加
    Dim nNullIndex           As Integer         ' 文字数ワーク      ' EG20 V6.6.0.1追加
    Dim nCornerIndex         As Integer         ' コーナ選択状態    ' EG20 V6.6.0.1追加
    Dim strSaveFileName      As String          ' 保存ファイル名    ' EG20 V6.6.0.1追加
    
    '初期化
    skansi = ""
    sidu = ""
    sldu = ""
'V1.13.0.1 ADD END

    'エラールーチンを宣言
    On Error Resume Next
'V1.8.0.1 DEL START
'    iResponse = MsgBox("選択されている駅の現在の駅都度データ１駅分をテキスト表示します。" & Chr(vbKeyReturn) & _
'                        "よろしいですか？", _
'                        vbYesNo + vbQuestion, _
'                        "駅設定テキスト出力確認")
'V1.8.0.1 DEL END
'V1.12.0.1 DEL START
''V1.8.0.1 ADD START
'    iResponse = MsgBox("選択されている駅の現在の駅都度データ１駅分をテキスト表示します。" & Chr(vbKeyReturn) & _
'                        "よろしいですか？", _
'                        vbOKCancel + vbQuestion, _
'                        "駅設定テキスト出力確認")
''V1.8.0.1 ADD END
''    If iResponse = vbNo Then Exit Sub
'V1.12.0.1 DEL START
    
'V1.12.0.1 ADD START
    '書き込み先ファイル選択
    'sWriteDir = pfDirSelection("H:", "機器構成ファイル書込み先のディレクトリ選択")         'V1.20.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
'V1.12.0.1 ADD START
'V1.13.0.1 ADD START
    If sWriteDir = "" Then
       'フォルダ選択画面「取消」釦押下時は処理終了
       Exit Sub
    End If
'V1.13.0.1 ADD END
    
   '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
'    strFileName = Dir(EKI_SETTI_FILE)              'V1.12.0.1 DEL
    strFileName = Dir(EKI_TUDO_CHK_EKI_INFO_FILE)   'V1.12.0.1 ADD

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '異常終了
        MsgBox "テキスト表示するデータがありません。", _
                vbOKOnly + vbExclamation, _
                 "データ無警告"
        Exit Sub
        
    End If
'V1.12.0.1 ADD START
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    On Error GoTo OUTPUT_ERROR
    
    'ファイル番号取得
    intFileNumber = FreeFile
    
    'CSVファイルオープン
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber
    
    'CSVファイル行数カウント（ファイル終端までループを繰り返す）
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1削除
        Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1追加
            Line Input #intFileNumber, strLineCount
            j = j + 1
        Loop
    
    'CSVファイルクローズ
    Close #intFileNumber

    'ファイル番号取得
    intFileNumber = FreeFile
    
    '再設定
    ReDim ReadFileSettei(j) As EKIINFO_IMAGE_FILE        'ファイル読込用エリア
        
    'CSVファイルオープン
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber

    'リスト表示分読み込み（ファイル終端までループを繰り返す）
        For i = 0 To j - 1
            Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, _
            ReadFileSettei(i).sCorner, ReadFileSettei(i).sTuuban, ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, _
            ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
        Next

    'CSVファイルクローズ
    Close #intFileNumber
    
    '一時ファイルを作る
    Set FsoTS = fso.CreateTextFile(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, True)
       
'    FsoTS.Write ("設置駅：" & Trim(pfGetEkiNameInfo(NotEkiVer)) & vbCrLf & vbCrLf)     ' EG20 V6.6.0.1削除
' EG20 V6.6.0.1追加開始
    FsoTS.Write ("設置駅　　：" & Trim(pfGetEkiNameInfo(NotEkiVer)) & vbCrLf)
    ' コーナ名称の付加
    nNullIndex = InStr(gstrCornerName(CmbCornerName.ListIndex), Chr(0))
    If nNullIndex <> 0 Then
        szCornerName = Left(gstrCornerName(CmbCornerName.ListIndex), nNullIndex - 1)
    Else
        szCornerName = gstrCornerName(CmbCornerName.ListIndex)
    End If
    FsoTS.Write ("設置コーナ：" & szCornerName & vbCrLf & vbCrLf)
    nCornerIndex = CmbCornerName.ListIndex + 1
' EG20 V6.6.0.1追加終了
       
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]コメント追加開始
' 従来小項目通番は大項目単位連番に対して１から連番であることが
' 前提であったが前提はなくなった。
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]コメント追加終了
    
    nProcMode = 0       ' 現在処理分類              ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    For k = 0 To j - 1
        
' EG20 V6.6.0.1条件追加開始
      If ((ReadFileSettei(k).sCorner = nCornerIndex) Or (ReadFileSettei(k).sCorner = 0)) Then
' EG20 V6.6.0.1条件追加終了
        
        '項目
        If ReadFileSettei(k).sType = 1 Then
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sTuuban = 1 Then        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
            If nProcMode <> 1 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                'タイトル表示処理（駅情報）
                FsoTS.Write ("【駅情報】" & vbCrLf & "項目,区分,設定値" & vbCrLf)
                nProcMode = 1                                                           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            End If
'            FsoTS.Write (ReadFileSettei(k).sKoumoku & ",")     'V1.13.0.1 DEL
        
        ElseIf ReadFileSettei(k).sType = 2 Then
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sNo <> ReadFileSettei(k - 1).sNo Then   ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
            If nProcMode <> 2 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                'タイトル表示処理（監視）
                FsoTS.Write (vbCrLf & "【監視】" & vbCrLf & "項目,区分,設定値" & vbCrLf)
                nProcMode = 2                                                           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            End If
'            FsoTS.Write (ReadFileSettei(k).sKoumoku & ",")     'V1.13.0.1 DEL
        
'        Else                                       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        ElseIf ReadFileSettei(k).sType = 3 Then     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sNo <> ReadFileSettei(k - 1).sNo Then   ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
            If nProcMode <> 3 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                'タイトル表示処理
                FsoTS.Write (vbCrLf & "【ネットワーク】" & vbCrLf & "項目,区分,設定値" & vbCrLf)
                nProcMode = 3                                                           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            End If
'            FsoTS.Write (ReadFileSettei(k).sKoumoku & ",")     'V1.13.0.1 DEL
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        Else
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sNo <> ReadFileSettei(k - 1).sNo Then   ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
            If nProcMode <> 7 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                'タイトル表示処理
                FsoTS.Write (vbCrLf & "【画面】" & vbCrLf & "項目,区分,設定値" & vbCrLf)
                nProcMode = 7                                                           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            End If
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        End If

'V1.13.0.1 ADD START
        '今回と前回の小項目が同じかどうか判定する
        If ReadFileSettei(k).sNo <> sSyoKoumoku Then
                '現在の小項目を保存する（区分別）
'                If ReadFileSettei(k).sKubun = "監視" Then      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
                If ReadFileSettei(k).sKubun = "統合" Then       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                    skansi = ReadFileSettei(k).sNo
                ElseIf ReadFileSettei(k).sKubun = "IDU" Then
                    sidu = ReadFileSettei(k).sNo
                ElseIf ReadFileSettei(k).sKubun = "LDU" Then
                    sldu = ReadFileSettei(k).sNo
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                ElseIf ReadFileSettei(k).sKubun = "操卓" Then
                    sTaku = ReadFileSettei(k).sNo
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                End If
                '現在の小項目を保存する（全体）
                sSyoKoumoku = ReadFileSettei(k).sNo
                'ファイルに出力する
                FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                ReadFileSettei(k).sSettei & vbCrLf)
        Else
            '小項目が同じだった場合、区分が同じかどうか確認する。同じであれば出力しない。
            Select Case ReadFileSettei(k).sKubun
'                Case "監視"            ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
                Case "統合"             ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                    If ReadFileSettei(k).sNo = skansi Then
                        '処理なし
                    Else
                        'ファイルに出力する
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
                Case "IDU"
                    If ReadFileSettei(k).sNo = sidu Then
                        '処理なし
                    Else
                        'ファイルに出力する
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
                Case "LDU"
                    If ReadFileSettei(k).sNo = sldu Then
                        '処理なし
                    Else
                        'ファイルに出力する
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                Case "操卓"
                    If ReadFileSettei(k).sNo = sTaku Then
                        '処理なし
                    Else
                        'ファイルに出力する
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
            End Select
        End If
      End If              ' EG20 V6.6.0.1追加
'V1.13.0.1 ADD END
'V1.13.0.1 DEL START
'            '区分
'            FsoTS.Write (ReadFileSettei(k).sKubun & ",")
'
'            '設定値
'            FsoTS.Write (ReadFileSettei(k).sSettei & vbCrLf)
'V1.13.0.1 DEL END
    Next
    
    'ファイルをクローズする。
    FsoTS.Close
        
    '一時ファイルを媒体にコピーする
'    Call FileCopy(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, sWriteDir & EKI_SETTI_EKI_INFO_FILE)        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
' EG20 V6.6.0.1削除開始
'    Call FileCopy(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & EKI_SETTI_EKI_INFO_FILE)        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
' EG20 V6.6.0.1削除終了
' EG20 V6.6.0.1追加開始
    strSaveFileName = sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Replace(szCornerName, " ", "") & "_" & EKI_SETTI_EKI_INFO_FILE
    Call FileCopy(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, strSaveFileName)
' EG20 V6.6.0.1追加終了
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '正常終了
    iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定テキスト出力結果")
'V1.12.0.1 ADD END

'V1.12.0.1 DEL START
'    sCommand = MN_EXE_MEMO & EKI_SETTI_FILE         'メモ帳実行コマンドを作成する
'    lRetVal = Shell(sCommand, vbMaximizedFocus)     'ノートパッドを起動する
'    AppActivate lRetVal, True                       'アクティブ（前面表示）にする
'
'    SendKeys "{LEFT}", True
'V1.12.0.1 DEL END

'V1.12.0.1 ADD START

    Exit Sub
    
OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, CREATE_FILE_ERROR, 0)

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '異常終了
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定テキスト出力結果")
'V1.12.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfStartUpProc
'//  機能名称  : ファイル選択画面処理
'//  機能概要  : ファイル選択画面を表示し、選択されたファイル名を返す。
'//
'//              型        名称      意味
'//  引数      : String　　sDrive　　[IN]初期表示ドライブ名
'//  　　      : String　　sPattern　[IN]選択対象ファイル拡張子
'//  　　      : String　　sTitle　　[IN]画面表示ラベル
'//
'//              型        値        意味
'//  戻り値    :String　　　　　　　 [OUT]戻り値
'//                                      選択されたファイルパス:正常　""：エラー
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function pfFileSelection(sDrive As String, _
                                sPattern As String, _
                                sTitle As String) As String
                                
    Dim sWorkDrive As String                    'ワーク用初期表示ドライブ名

    'ドライブ異常処理を定義する。
    On Error GoTo Drive_Error
    
    sWorkDrive = sDrive                         '初期表示ドライブ名をワーク用にセットする。
    frmFil.filSelection.Pattern = sPattern      '選択対象拡張子をセットする。
    frmFil.lblFileSelection = sTitle            'サブタイトルをセットする。

Retry:
    frmFil.drvSelection.Drive = sWorkDrive      'ドライブをセットする。
    frmFil.dirSelection.Path = sWorkDrive & "\" 'ディレクトリをセットする。
    
    'ファイル選択画面を表示する。
    frmFil.Show 1
    
    '選択されたファイル名を返す。
    pfFileSelection = gstrMyPath
    
    Exit Function

'**ドライブ指定異常処理**
Drive_Error:

'    If Left$(sWorkDrive, 1) = "a" Then     'V1.12.0.1 DEL
    If Left$(sWorkDrive, 1) = "H" Then      'V1.12.0.1 ADD
        'a:ドライブが異常なら、カレントドライブを表示させる。
        sWorkDrive = Left$(App.Path, 2)
        GoTo Retry
    End If
    
    'その他のドライブなら、ファイル選択なしで戻る。
    pfFileSelection = ""

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : CmdMoveGateGamen_Click
'//  機能名称  : 「自改画面へ」釦押下処理
'//  機能概要  : 駅都度データ確認(自改)画面を表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveGateGamen_Click()
   
'V1.12.0.1 ADD START
    '全ボタンを押下不可とする。
    Call SetEnableFalse
    DoEvents
'V1.12.0.1 ADD END
   
   '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)
    Unload Me
    Load frmEkiDataGate
    frmEkiDataGate.Show 1
    
End Sub
'V1.4.0.1 ADD END
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
    cmbEkiInfo.Enabled = False
    CmdMenu(0).Enabled = False
    CmdMenu(1).Enabled = False
    CmdMenu(2).Enabled = False
    CmdMenu(3).Enabled = False
    CmdMoveGateGamen.Enabled = False
    cmdCancel.Enabled = False
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    CmdMoveSubGateGamen.Enabled = False
    CmbCornerName.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

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
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下可とする。
    cmbEkiInfo.Enabled = True
    CmdMenu(0).Enabled = True
    CmdMenu(1).Enabled = True
    CmdMenu(2).Enabled = True
    CmdMenu(3).Enabled = True
    CmdMoveGateGamen.Enabled = True
    cmdCancel.Enabled = True
 
 ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    CmdMoveSubGateGamen.Enabled = True
    CmbCornerName.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

    DoEvents
    
End Sub
'V1.12.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmdMoveGateGamen_Click
'//  機能名称  : エンコードコーナ設定画面切替
'//  機能概要  : 駅都度データ確認（エンコードコーナ設定）画面を表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EGR HK1.1.0.1) 2011-05-11  CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//                 EGR HK流用
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveSubGateGamen_Click()
    
    '全ボタンを押下不可とする。
    Call SetEnableFalse
    DoEvents
   
   '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SUBGATE_GAMEN_GO_BUTTOM, 0)

    '表示中画面アンロード
    Unload Me
                
    Load frmEkiDataSubGate
    frmEkiDataSubGate.Show 1

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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pfuncInstallEkiSettei
'//  機能名称  : 駅設定インストール処理
'//  機能概要  :
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
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfuncInstallEkiSettei() As Boolean

    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bRet                As Boolean          '関数戻り値
    Dim lErrCode            As Long             'エラーコード

    Dim bSysChange              As Boolean      'システム設定処理戻り値
    Dim bUpData                 As Boolean      '画面更新処理戻り値

    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    Call SetEnableFalse

    pfuncInstallEkiSettei = True

    '現在駅設定データインストール処理
    bRet = dllInstolEkiDataNow(gstrFileName, EKI_SETTI_FILE, lErrCode)
    
    If bRet = False Then
            
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
            
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        pfuncInstallEkiSettei = False
        '異常終了
        iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
            
    Else
        
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        '----------------------------------------------------
        'コンピュータ名、ネットワーク変更処理
        '----------------------------------------------------
        bUpData = True
        bSysChange = True
        bSysChange = pfNetWorkChng(Me)
         'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
            
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
            
        '駅都度データ確認（駅情報）イメージファイル作成
        bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
        If bRet = False Then
            '駅都度データ確認（駅情報）イメージファイル削除
            Kill EKI_TUDO_CHK_EKI_INFO_FILE
               
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
            bUpData = False
            pfuncInstallEkiSettei = False
        End If

        '駅情報コンボボックス初期値設定
        cmbEkiInfo.Clear
        cmbEkiInfo.AddItem "駅情報"
        cmbEkiInfo.AddItem "監視"
        cmbEkiInfo.AddItem "ネットワーク"
        cmbEkiInfo.AddItem "画面"
        cmbEkiInfo.ListIndex = 0

        'コーナ設定コンボボックスの初期化処理
        Call InitCornerComboBox
            
        '画面表示処理
        Call sDisp
            
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
        If bSysChange = True And bUpData = True Then
            
            '正常終了
            iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定入力結果")
        End If
    End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    Kill gstrFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
    gstrFileName = ""
    Call SetEnableTrue
End Function

