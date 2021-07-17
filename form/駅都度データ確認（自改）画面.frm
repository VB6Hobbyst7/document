VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEkiDataGate 
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   480
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdMoveGateGamen 
      Caption         =   "駅情報画面へ"
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
      TabIndex        =   14
      Top             =   8400
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
      TabIndex        =   13
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "テキスト媒体出力（改札機）"
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
      TabIndex        =   12
      Top             =   7800
      Width           =   2295
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
      TabIndex        =   11
      Top             =   7800
      Width           =   2175
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
      TabIndex        =   10
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   11040
      Top             =   7320
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
      Height          =   6710
      Left            =   9995
      TabIndex        =   9
      Top             =   960
      Width           =   300
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPageDown 
      Caption         =   "▼  ▼"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   7
      Top             =   4380
      Width           =   1215
   End
   Begin VB.CommandButton cmdPageUp 
      Caption         =   "▲ ▲"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   6
      Top             =   2820
      Width           =   1215
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "▲"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtDummy 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
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
      Height          =   6710
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   11827
      _Version        =   393216
      Rows            =   30
      Cols            =   17
      FixedCols       =   2
      RowHeightMin    =   300
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      GridLinesFixed  =   1
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
      Caption         =   "駅都度データ確認（改札機）"
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
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "frmEkiDataGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：駅都度データ確認（自改）画面.frm
'//  パッケージ名：駅都度データ確認（自改）画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応
'//                 「駅設定出力」「駅設定入力」
'//                 「駅設定テキスト出力」「媒体取外」釦処理追加
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//                 駅設定ファイル書込み先ディレクトリ位置変更
'//                 ディスク情報取得位置を変更
'//                 出力フォーマット変更
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                フォルダ選択画面での「取消」釦押下処理追加
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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(EG30 V33.2.0.1) 2017-10-05 REVISED BY  [TCC] T.Nakajima
'//                 2017年度施策 現地版対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   'メイルタイマのインターバル値
' Private Const TITOL_EKI_NAME = "駅名　　　："           '駅名タイトル     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
Private Const TITOL_EKI_NAME = "駅名："                 '駅名タイトル       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

Private strhairetu()        As String                       ' 表示データ
Private gstrFileName        As String                       ' 出力ファイル名    ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 駅都度データ確認（自改）画面(アクティブ時：イベントプロシージャ)
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
'//  機能名称  : 駅都度データ確認（自改）画面(ディアクティブ時)
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
'//  機能名称  : 駅都度データ確認（自改）画面(ロード時：イベントプロシージャ)
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_GATE_GAMEN_START, 0)
    
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
    
    '駅都度データ確認（自改）イメージファイル作成
    bRet = dllGetEkiIniData(1, EKI_TUDO_CHK_GATE_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '駅都度データ確認（自改）イメージファイル削除
        Kill EKI_TUDO_CHK_GATE_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
    End If
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    'コーナ設定コンボボックスの初期化処理
    Call InitCornerComboBox
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
    '画面表示処理
    Call sDisp

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
'                AppActivate frmInputMstData.Caption, False ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
                AppActivate frmEkiDataGate.Caption, False   ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
                pfFormActive (frmEkiDataGate.hwnd)          ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '「保守操作卓プログラム送信要求」を受信した場合
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
                    Kill gstrFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
                    'プログレスバーを消去する
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_GATE_GAMEN_END, 0)
    
    '画面消去
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
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG30 V33.2.0.1) 2017-10-05  CODED BY  [TCC] T.Nakajima
'//                 2017年度施策 現地版対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim iRowCnt              As Integer         '行数カウンタ
    Dim iColCnt              As Integer         '列数カウンタ
    Dim bRet                 As Boolean         '関数戻り値
    Dim strFileName          As String          'ファイル名
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
        Call sDispDataClear

        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '駅名ラベル更新
    '----------------------------------------------------
'   LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo              'V2.1.0.1 DEL
    LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo(NotEkiVer)   'V2.1.0.1 ADD
    
    '駅都度データ確認（自改）イメージファイル検索
    strFileName = Dir(EKI_TUDO_CHK_GATE_FILE)
    
    'ファイルが存在した場合
    If strFileName <> "" Then
    
        'グリッドデータ部設定
'        Call sDispDataSet                                              ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        nCornerIndex = CmbCornerName.ListIndex
        Call sDispDataSet(nCornerIndex + 1)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
        'INIデータチェック
        With GridIni
            For iRowCnt = 1 To .Rows - 1
                .Row = iRowCnt
                .Col = 0
                If (.Text <> "") Then
                For iColCnt = 2 To .Cols - 1
                
                    'セル指定
                    .Row = iRowCnt
                    .Col = iColCnt
                    
                    'データチェック
                    bRet = pfDispDataChk(.Text)
                    If bRet = False Then
                        'EG30 V33.2.0.1 ADD START
                        '通路種別(小分類8)は個別に設定可能なので、アプリ毎に設定値が異なっても赤表示にしない
                        If iColCnt <> Bunrui_Sho_Type.GATE_TYPE_TURO + 1 Then
                            .CellBackColor = QBColor(12)
                        End If
                        'EG30 V33.2.0.1 ADD END
                        'EG30 V33.2.0.1 DEL START
                        'データ不一致の場合、セルの背景色を赤色にする
                        '.CellBackColor = QBColor(12)
                        'EG30 V33.2.0.1 DEL END
                    End If
                    
                Next
                End If
            Next
        End With
        
        CmbCornerName.Enabled = True               ' コーナ選択部選択不可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_GATE_IMAGE, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear
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
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispGridTitol()
    
    Dim ColCount                As Integer         ' カラムカウンタ
    Dim RowCount                As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    'グリッドタイトル設定
    With GridIni
    
        '----------------------------------
        'グリッドの初期化
        '----------------------------------
        .Clear
        
        '----------------------------------
        'グリッドセル数設定
        '----------------------------------
        .Rows = 19
'        .Rows = 17
'        .Cols = 18 'V1.11.0.1 DEL
'        .Cols = 10  'V1.11.0.1 ADD             ' EG20 V3.0.0.2 （駅都度修正対応）削除
        .Cols = 10  'V1.11.0.1 ADD              ' EG20 V3.0.0.2 （駅都度修正対応）追加
        
        '----------------------------------
        'グリッド幅設定
        '----------------------------------
        .ColWidth(0) = 900
        .ColWidth(1) = 700
        
        For ColCount = 2 To (.Cols - 1)
            .ColWidth(ColCount) = 2050
        Next
        
        '----------------------------------
        'タイトル設定
        '----------------------------------
        '区分設定
        .Col = 1
        .Row = 0: .Text = "区分"
        .CellAlignment = flexAlignCenterCenter
        For RowCount = 1 To (.Rows - 1)
            .Row = RowCount
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'            .Text = "監視" & vbCrLf & _
'                    "IDU" & vbCrLf & _
'                    "LDU"
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
            .Text = "統合" & vbCrLf & _
                    "操卓" & vbCrLf & _
                    "IDU" & vbCrLf & _
                    "LDU"
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
            .CellAlignment = flexAlignCenterCenter
            .RowHeight(.Row) = 938          ' EG20 V3.0.0.2 （駅都度修正対応）追加
        Next

'        .RowHeight(0) = 500    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        .RowHeight(0) = 660     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
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
Private Sub sDispDataClear()
    
    Dim iLoopCnt             As Integer         'ループカウンタ
    Dim ColCount             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    'グリッド初期化
    With GridIni

        For iLoopCnt = 1 To (.Rows - 1)

            '号機設定
            .Col = 0
            .Row = iLoopCnt: .Text = iLoopCnt & "号機"
            .CellAlignment = flexAlignLeftCenter

            '項目設定
            For ColCount = 2 To (.Rows - 1)
                .Col = ColCount
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'                .Text = "" & vbCrLf & _
'                        "" & vbCrLf & _
'                        ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                .Text = "" & vbCrLf & _
                        "" & vbCrLf & _
                        "" & vbCrLf & _
                        ""
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                .CellAlignment = flexAlignLeftCenter
            Next

'            .RowHeight(iLoopCnt) = 700         ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
            .RowHeight(iLoopCnt) = 938          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
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
'//  引数      : Integer   iCorner   コーナ  ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sDispDataSet()                                 ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
Private Sub sDispDataSet(iCorner As Integer)                ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    Dim intFileNumber       As Integer                      ' ファイルポインタ
    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim ColCount            As Integer                      ' カラムカウンタ
    Dim iTuban              As Integer                      ' 通番
    
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

    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '駅都度データ確認（自改）イメージファイルをオープンする。
    Open EKI_TUDO_CHK_GATE_FILE For Input As #intFileNumber
    
    GridIni.Visible = False
    iTuban = 1
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
        
        If BUNRUI_DAI.DAI_Gate = strBunrui_Dai Then
        
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        ' コーナ判定追加
        ' 選択したコーナ、もしくはコーナ無関係のレコードは採用する
        iCmpCorner = CInt(strCorner)
        If ((iCorner = iCmpCorner) Or (iCmpCorner = 0)) Then
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        
            'グリッド初期化
            With GridIni
        
                '号機設定
                .Col = 0
                .Row = strBunrui_Tyu
                If (.Text = "") Then .Text = strBunrui_Tyu & "号機"
                .CellAlignment = flexAlignLeftCenter
                
                'V1.8.0.1 ADD START
                If .Cols <= strBunrui_Sho + 1 Then
                   '----------------------------------
                   'グリッドセル数設定
                   '----------------------------------
                   .Cols = strBunrui_Sho + 2
                   
                   '----------------------------------
                   'グリッド幅設定
                   '----------------------------------
                   .ColWidth(.Cols - 1) = 2050
                End If
                'V1.8.0.1 ADD END
                
                '項目設定
                .Col = strBunrui_Sho + 1
                .Text = pfDispIniData(.Text, strData, strKubun)
                .CellAlignment = flexAlignLeftCenter
'                .RowHeight(.Row) = 700     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
                .RowHeight(.Row) = 938      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            
                'タイトル設定
                .Col = strBunrui_Sho + 1
                .Row = 0
                If (.Text = "") Then
                    .Text = strKomoku
                    .CellAlignment = flexAlignLeftCenter
'                    .RowHeight(.Row) = 500     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
                    .RowHeight(.Row) = 660      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                End If
            
            End With
        
        End If          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        End If
    
    Loop

    GridIni.Visible = True
    
    'ファイルをクローズする。
    Close #intFileNumber

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

    GridIni.Visible = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdDown_Click
'//  機能名称  : 「▼」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
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
Private Sub cmdDown_Click()

    'エラールーチンを宣言
    On Error Resume Next
    
    'グリッドの変更
    With GridIni
        If .TopRow < 11 Then
            .TopRow = .TopRow + 1
        End If
    End With
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdPageDown_Click
'//  機能名称  : 「▼▼」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
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
Private Sub cmdPageDown_Click()

    'エラールーチンを宣言
    On Error Resume Next
    
    'グリッドの変更
    With GridIni
        If .TopRow < 11 Then
            .TopRow = .TopRow + 6
        Else
            .TopRow = 11
        End If
    End With

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdPageUp_Click
'//  機能名称  : 「▲▲」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
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
Private Sub cmdPageUp_Click()

    'エラールーチンを宣言
    On Error Resume Next
    
    'グリッドの変更
    With GridIni
        If .TopRow > 6 Then
            .TopRow = .TopRow - 6
        Else
            .TopRow = 1
        End If
    End With
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdUp_Click
'//  機能名称  : 「▲」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
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
Private Sub cmdUp_Click()

    'エラールーチンを宣言
    On Error Resume Next
    
    'グリッドの変更
    With GridIni
        If .TopRow > 1 Then
            .TopRow = .TopRow - 1
        End If
    End With
    
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
'//                 駅設定ファイル書込み先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10   REVISED BY [TCC] S.Yoshimori
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
'    If iResponse = vbNo Then Exit Sub          'V1.12.0.1 DEL
    If iResponse = vbCancel Then Exit Sub       'V1.12.0.1 ADD

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
'        FileCopy EKI_SETTI_FILE, sWriteDir & Dir(EKI_SETTI_FILE)                       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
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
'//                 ディスク情報取得位置を変更
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
'    Dim bSysChange              As Boolean      'システム設定処理戻り値　’V1.8.0.1　ADD
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
'''    If iResponse = vbNo Then Exit Sub          'V1.12.0.1 DEL
'    'If iResponse = vbCancel Then Exit Sub       'V1.12.0.1 ADD
'    '
'    ''ディスク情報を取得
'''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
'    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD
'    '
'    'If lDrive = 0 Then
'    '    strDrive = "d:"
'    'Else
'''        strDrive = "a:"    'V1.12.0.1 DEL
'    '    strDrive = "H:"    'V1.12.0.1 ADD
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
'            bUpData = True
'            bSysChange = True
'            bSysChange = pfNetWorkChng(Me)
''V1.8.0.1 ADD END
'             'ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
''V1.8.0.1 ADD START
'            '駅都度データ確認（自改）イメージファイル作成
'            bRet = dllGetEkiIniData(1, EKI_TUDO_CHK_GATE_FILE, EKI_SETTI_FILE, lErrCode)
'            If bRet = False Then
'               '駅都度データ確認（自改）イメージファイル削除
'               Kill EKI_TUDO_CHK_GATE_FILE
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'                'プログレスバーを消去する
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'               '異常ログ出力
'               Call pfOutPutErrLog(lErrCode)
'               bUpData = False
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
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'            'コーナ設定コンボボックスの初期化処理
'            Call InitCornerComboBox
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'            '画面表示処理
'            Call sDisp
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
'
'            End If 'V1.8.0.1 ADD
'        End If
'    End If
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
'//                 出力フォーマット変更
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                フォルダ選択画面での「取消」釦押下処理追加
'//     REVISIONS :(1.20.0.1) 2010-03-10   REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//                 ファイル読込みエラー時に正常になる不具合を修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-13  CODED BY  [TCC] H.Sugimoto
'//                 【コーナ名スペース除去対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispTextEkiDataNow()

    Dim strFileName          As String          'ファイル名
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim lRetVal              As Long            '戻り値
    Dim sCommand             As String          'コマンド文字列
'V1.12.0.1 ADD START
    Dim sWriteDir            As String          '書き込み先フォルダ名
    Dim intFileNumber        As Integer         'ファイルポインタ
    Dim ColCount             As Integer         'カラムカウンタ
    Dim RowCount             As Integer         'ループカウンタ
    Dim TypeCount            As Integer         'ループカウンタ
    Dim sData                As String          '入力用文字列
    Dim strData_Kansi()      As String          '監視盤情報保存配列
    Dim strData_Idu()        As String          'IDU情報保存配列
    Dim strData_Ldu()        As String          'LDU情報保存配列
    Dim iLength              As Integer         '改行コード検索用（長さ）
    Dim iLeft                As Integer         '改行コード検索用（先頭）
    Dim iRight               As Integer         '改行コード検索用（終端）
'V1.12.0.1 ADD END
    Dim strData_Taku()       As String          ' 操作卓情報保存配列    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim strSaveFileName      As String          ' 保存ファイル名        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim szCornerName         As String          ' コーナ名称            ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim nNullIndex           As Integer         ' 文字数ワーク          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim strWork              As String          ' ワーク                ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

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
'    If iResponse = vbNo Then Exit Sub
'V1.12.0.1 DEL END
   
'V1.12.0.1 ADD START
    '書き込み先ファイル選択
    'sWriteDir = pfDirSelection("H:", "機器構成ファイル書込み先のディレクトリ選択")    'V1.20.0.1 DEL
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
    strFileName = Dir(EKI_SETTI_FILE)

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
       
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
       
'V1.12.0.1 ADD START
    On Error GoTo OUTPUT_ERROR
    
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '現在駅設定ファイルをオープンする。
    Open PATH_WORK & EKI_SETTI_GATE_FILE For Output As #intFileNumber
    'タイトル表示
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    ' コーナ名称の付加
    nNullIndex = InStr(gstrCornerName(CmbCornerName.ListIndex), Chr(0))
    If nNullIndex <> 0 Then
        szCornerName = Left(gstrCornerName(CmbCornerName.ListIndex), nNullIndex - 1)
    Else
'        szCornerName = ""                                              ' EG20 V3.3.0.1削除
        szCornerName = gstrCornerName(CmbCornerName.ListIndex)          ' EG20 V3.3.0.1追加
    End If
    Print #intFileNumber, "設置駅　　：" & Trim(pfGetEkiNameInfo(NotEkiVer))
    Print #intFileNumber, "設置コーナ：" & szCornerName
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    Print #intFileNumber, "【自改】"
    
    'エラールーチンを宣言
'    On Error Resume Next       'V1.20.0.1 DEL

    'グリッドタイトル設定
    With GridIni
    
        '行数分ループさせる
        For RowCount = 0 To .Rows - 1
            'sData初期化
            sData = ""
            '各項目表示
            If RowCount = 0 Then
                .Col = 0
                sData = "項目,"
                For ColCount = 1 To .Cols - 1
                    .Col = ColCount
                    .Row = RowCount
                
                    If ColCount <> .Cols - 1 Then
                        sData = sData & .Text & ","
                    Else
                        sData = sData & .Text
                    End If
                Next
                Print #intFileNumber, sData
            Else
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                .Row = RowCount
                .Col = 0
                If (.Text <> "") Then
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                
                '再定義
                ReDim strData_Kansi(.Cols)
                ReDim strData_Idu(.Cols)
                ReDim strData_Ldu(.Cols)
                ReDim strData_Taku(.Cols)           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            
                '項目分ループする
                For ColCount = 1 To .Cols - 1
                    .Col = ColCount
                    .Row = RowCount
                    ' 初期化
                    strData_Kansi(ColCount) = ""
                    strData_Idu(ColCount) = ""
                    strData_Ldu(ColCount) = ""
                    strData_Taku(ColCount) = ""     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
                    
                    '改行コード検索
                    iLength = Len(.Text)
                    iLeft = InStr(.Text, vbCrLf)
                    iRight = InStrRev(.Text, vbCrLf)
                    
                    ' 監視設定値取得
                    strData_Kansi(ColCount) = Mid(.Text, 1, iLeft - 1)
                
                    ' ＩＤＵ設定値取得
                    strData_Idu(ColCount) = Mid(.Text, iLeft + 2, iRight - iLeft - 2)
                
                    ' ＬＤＵ設定値取得
                    strData_Ldu(ColCount) = Mid(.Text, iRight + 2, iLength - iRight - 1)
                        
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                    ' 操作卓とＩＤＵの設定値を分解
                    strWork = strData_Idu(ColCount)
                    iLeft = InStr(strWork, vbCrLf)
                    iLength = Len(strWork)
                    ' 操作卓設定値取得
                    strData_Taku(ColCount) = Mid(strWork, 1, iLeft - 1)
                    ' ＩＤＵ設定値取得
                    strData_Idu(ColCount) = ""
                    strData_Idu(ColCount) = Mid(strWork, iLeft + 2, iLength - iLeft - 1)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                Next
                '監視盤情報出力
                .Col = 0
                .Row = RowCount
                sData = sData & .Text & ","
                For ColCount = 1 To .Cols - 1
                    If ColCount <> .Cols - 1 Then
                        sData = sData & strData_Kansi(ColCount) & ","
                    Else
                        sData = sData & strData_Kansi(ColCount)
                    End If
                Next
                Print #intFileNumber, sData

' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                'sData初期化
                sData = ""
                
                ' 操作卓情報出力
                For ColCount = 0 To .Cols - 1
                    If ColCount <> .Cols - 1 Then
                        sData = sData & strData_Taku(ColCount) & ","
                    Else
                        sData = sData & strData_Taku(ColCount)
                    End If
                Next
                Print #intFileNumber, sData
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                
                'sData初期化
                sData = ""
                
                'IDU情報出力
                For ColCount = 0 To .Cols - 1
                    If ColCount <> .Cols - 1 Then
                        sData = sData & strData_Idu(ColCount) & ","
                    Else
                        sData = sData & strData_Idu(ColCount)
                    End If
                Next
                Print #intFileNumber, sData
                        
                'sData初期化
                sData = ""
                
                'LDU情報出力
                For ColCount = 0 To .Cols - 1
                    If ColCount <> .Cols - 1 Then
                        sData = sData & strData_Ldu(ColCount) & ","
                    Else
                        sData = sData & strData_Ldu(ColCount)
                    End If
                Next
                Print #intFileNumber, sData

                End If          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
            End If
                
        Next
    
    End With
    
    'ファイルをクローズする。
    Close #intFileNumber
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'    '一時ファイルを媒体にコピーする
'    Call FileCopy(PATH_WORK & EKI_SETTI_GATE_FILE, sWriteDir & EKI_SETTI_GATE_FILE)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    ' コーナ名称の付加
'    strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & szCornerName & "_" & EKI_SETTI_GATE_FILE       ' EG20 V6.1.0.1削除
' EG20 V6.1.0.1追加開始
    strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Replace(szCornerName, " ", "") & "_" & EKI_SETTI_GATE_FILE
' EG20 V6.1.0.1追加終了
    '一時ファイルを媒体にコピーする
    Call FileCopy(PATH_WORK & EKI_SETTI_GATE_FILE, sWriteDir & strSaveFileName)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
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
'    SendKeys "{LEFT}", True
'V1.12.0.1 DEL END
    
'V1.12.0.1 ADD START
    
    Exit Sub
    
OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
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
'//  機能名称  : 「駅情報画面へ」釦押下処理
'//  機能概要  : 駅都度データ確認(駅情報)画面を表示する。
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
'V1.12.0.1 ADD END
  
   '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIINFO_GAMEN_GO_BUTTOM, 0)
    Unload Me
    Load frmEkiData
    frmEkiData.Show 1
    
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

    '全釦を押下不可とする。
    CmdMenu(0).Enabled = False
    CmdMenu(1).Enabled = False
    CmdMenu(2).Enabled = False
    CmdMenu(3).Enabled = False
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    cmdPageUp.Enabled = False
    cmdPageDown.Enabled = False
    CmdMoveGateGamen.Enabled = False
    cmdCancel.Enabled = False
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    CmdMoveSubGateGamen.Enabled = False
    CmbCornerName.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
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

    '全釦を押下可とする。
    CmdMenu(0).Enabled = True
    CmdMenu(1).Enabled = True
    CmdMenu(2).Enabled = True
    CmdMenu(3).Enabled = True
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    cmdPageUp.Enabled = True
    cmdPageDown.Enabled = True
    CmdMoveGateGamen.Enabled = True
    cmdCancel.Enabled = True
        
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    CmdMoveSubGateGamen.Enabled = True
    CmbCornerName.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
        
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

    Dim bSysChange          As Boolean          'システム設定処理戻り値
    Dim bUpData             As Boolean          '画面更新処理戻り値

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
            
        '駅都度データ確認（自改）イメージファイル作成
        bRet = dllGetEkiIniData(1, EKI_TUDO_CHK_GATE_FILE, EKI_SETTI_FILE, lErrCode)
        If bRet = False Then
            '駅都度データ確認（自改）イメージファイル削除
            Kill EKI_TUDO_CHK_GATE_FILE
               
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
            bUpData = False
            pfuncInstallEkiSettei = False
        End If

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

