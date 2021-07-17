VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmKikiDataGate 
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   480
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
      Left            =   7200
      TabIndex        =   12
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   8520
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
      Height          =   495
      IMEMode         =   3  'ｵﾌ固定
      Left            =   120
      TabIndex        =   0
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   6600
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1125
      Width           =   11770
      _ExtentX        =   20770
      _ExtentY        =   11642
      _Version        =   393216
      Rows            =   17
      Cols            =   17
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
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
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "機器構成設定（改札機）"
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
      TabIndex        =   7
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiDataGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：機器情報設定（自改）画面.frm
'//  パッケージ名：機器情報設定（自改）画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　「駅情報画面へ」釦押下処理追加
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
'//     REVISIONS :(1.16.0.1) 2009-12-21  REVISED BY [TCC] E.Watanabe
'//                 内部保存エリアへの格納ループカウンタ最大値を修正
'//     REVISIONS :(1.17.0.1) 2009-12-24  REVISED BY [TCC] E.Watanabe
'//                 不具合修正
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                画面再前面表示修正(不具合修正)
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//                 ファイル選択画面をOS仕様に変更
'//                 カーソル移動の処理を削除
'//                 号機番号の入力桁数を制御する処理を追加
'//                 設定反映ボタンが押されずに画面遷移するときの警告表示を追加
'//                 通路種別と自改種別の正当性チェックを追加
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ポップアップ画面タイトル修正
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(EG30 V33.2.0.1) 2017-10-05 CODED BY  [TCC] T.Nakajima
'//                 2017年度施策 現地版対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   'メイルタイマのインターバル値
Private Const TITOL_EKI_NAME = "駅名："                 '駅名タイトル       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
Private bScroll As Boolean
Private strCombData() As String
Private iBunrui_Sho_Save() As Integer

'V1.20.0.1 ADD START
'入力桁数チェック用
Private Type KetaFileData
    sName As String             '正当性チェック対象項目名
    iKeta  As Integer           '桁数
End Type
'種別正当性チェック用
Private Type HikakuFileData
    sName1 As String             '正当性チェック対象項目名1
    sName2 As String             '正当性チェック対象項目名2
    sMoji1  As String            '文字1
    sMoji2  As String            '文字2
    iCol1 As Integer             '項目1のカラム番号
    iCol2 As Integer             '項目2のカラム番号
End Type

Private Const iModMax = 99       'ファイル読み込みMAX値
Private uKetaFileData() As KetaFileData
Private uHikakuFileData() As HikakuFileData

'設定反映フラグ
Private SetteiHaneiFlg As Boolean
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 機器情報設定（自改）画面(アクティブ時：イベントプロシージャ)
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
'//  機能名称  : 機器情報設定（自改）画面(ロード時：イベントプロシージャ)
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
'//                号機番号の入力桁数を制御するためINIファイル読み込み
'//                設定反映フラグ追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_GATE_GATE_GAMEN_START, 0)
    
    '----------------------------------------------------
    '画面初期値設定
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    ReDim strCombData(0)
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '機器情報設定（自改）イメージファイル作成
    bRet = dllGetKikiIniData(1, 0, KIKI_DATA_SET_GATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
    If bRet = False Then
        '機器情報設定（自改）イメージファイル削除
        Kill KIKI_DATA_SET_GATE_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
    End If
    
    '画面表示処理
    Call sDisp
    
    'V1.20.0.1 ADD START
    'INIファイルの読込み
    Call psGetFileChk
    'V1.20.0.1 ADD END
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
'V1.20.0.1 ADD START
    '設定反映フラグ（変更なし）
    SetteiHaneiFlg = False
'V1.20.0.1 ADD END

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
                AppActivate frmKikiDataGate.Caption, False  ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
                pfFormActive (frmKikiDataGate.hwnd)         ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
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
    Dim iResponse           As Integer          'MsgBox戻り値   'V1.20.0.1 ADD
    
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_GATE_GATE_GAMEN_END, 0)
    
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

    'エラールーチンを宣言
    On Error Resume Next

    '初期値設定
    strFileName = ""                            'ファイル名
    LblEkiName.Caption = TITOL_EKI_NAME         '駅名ラベル初期化           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    '----------------------------------------------------
    'グリッドタイトル設定
    '----------------------------------------------------
    Call sDispGridTitol
    Erase KikiDataTbl
    ReDim KikiDataTbl(0)
    Call pfKikiDataSet
    Erase iBunrui_Sho_Save
    ReDim iBunrui_Sho_Save(0)
    
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
    
    '機器構成情報（自改）イメージファイル検索
    strFileName = Dir(KIKI_DATA_SET_GATE_FILE)
    
    'ファイルが存在した場合
    If strFileName <> "" Then
    
        'グリッドデータ部設定
        Call sDispDataSet
    
        '処理釦押下可能設定
        CmdKikiSetMenu(0).Enabled = True            '機器構成項目設定反映
        CmdKikiSetMenu(1).Enabled = True            '機器構成項目媒体出力
        CmdKikiSetMenu(2).Enabled = True            '機器構成項目内部保存

    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_GATE_IMAGE, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear
        
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
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispGridTitol()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    Dim ColCount                As Integer         ' カラムカウンタ

    'グリッドタイトル設定
    With GridIni
    
        '----------------------------------
        'グリッドの初期化
        '----------------------------------
        .Clear
        
        '----------------------------------
        'グリッドセル数設定
        '----------------------------------
'        .Rows = 17                     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        .Rows = 33                      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'        .Cols = 17 'V1.11.0.1 DEL
'        .Cols = 10  'V1.11.0.1 ADD     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        .Cols = 14  'V1.11.0.1 ADD      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        
        '----------------------------------
        'グリッド幅設定
        '----------------------------------
        .ColWidth(0) = 1000
        .RowHeight(0) = 500
        For ColCount = 1 To (.Cols - 1)
            'グリッドの幅変更
            .ColWidth(ColCount) = 2100
        Next
        
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
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear()
    
    Dim iLoopCnt             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    'グリッド初期化
    With GridIni

        For iLoopCnt = 1 To (.Rows - 1)

            '号機設定
            .Col = 0
            .Row = iLoopCnt: .Text = iLoopCnt & "号機"
            .CellAlignment = flexAlignLeftCenter

            .RowHeight(iLoopCnt) = 365
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
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataSet()
    
    Dim intFileNumber       As Integer                      ' ファイルポインタ
    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim ColCount            As Integer                      ' カラムカウンタ
    
    Dim strBunrui_Dai       As String                       ' 大分類
    Dim strBunrui_Tyu       As String                       ' 中分類
    Dim strBunrui_Sho       As String                       ' 小分類
    Dim strKomoku           As String                       ' 項目
    Dim strKubun            As String                       ' 区分
    Dim strData             As String                       ' 設定値
    Dim strSetShosai        As String                       ' 設定値詳細
    
    Dim strDispData         As String                       ' 表示データ
    Dim byBuff()            As Byte                         ' バイトバッファ
    Dim iLoopCnt2           As Integer
    Dim strCorner           As String                       ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    
    'エラールーチンを宣言
    On Error Resume Next

    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '駅都度データ確認（自改）イメージファイルをオープンする。
    Open KIKI_DATA_SET_GATE_FILE For Input As #intFileNumber
    
    GridIni.Visible = False
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
        
            'グリッド初期化
            With GridIni
        
                '号機設定
                .Col = 0
                .Row = strBunrui_Tyu
                If (.Text = "") Then .Text = strBunrui_Tyu & "号機"
                .CellAlignment = flexAlignLeftCenter
                
                'V1.8.0.1 ADD START
                If .Cols <= strBunrui_Sho Then
                   '----------------------------------
                    'グリッドセル数設定
                    '----------------------------------
                    .Cols = strBunrui_Sho + 1
            
                    '----------------------------------
                    'グリッド幅設定
                    '----------------------------------
                    .ColWidth(.Cols - 1) = 2050
               End If
               'V1.8.0.1 ADD END
                
                '項目設定
                .Col = strBunrui_Sho
                .Text = strData
                .CellAlignment = flexAlignLeftCenter
                .RowHeight(.Row) = 365
                 
                 'タイトル設定
                .Col = strBunrui_Sho
                .Row = 0
                If (.Text = "") Then
                    .Text = strKomoku
                    .CellAlignment = flexAlignLeftCenter
                    .RowHeight(.Row) = 500
                    
                    ReDim Preserve iBunrui_Sho_Save((UBound(iBunrui_Sho_Save) + 1))
                    iBunrui_Sho_Save(.Col) = strBunrui_Sho

                End If
            
            End With
        
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
    
    Dim iLoopCnt As Integer
    
    'エラールーチンを宣言
    On Error Resume Next
    
    'クリックされた位置にダミーテキストを移動し、フォーカスを合わせる
    With GridIni
        
        If sInitCombDummy = False Then
            CmbDummy.Left = .Left + .CellLeft
            CmbDummy.Top = .Top + .CellTop
            CmbDummy.Width = .CellWidth
            CmbDummy.Height = .CellHeight
            CmbDummy.Text = .Text
            CmbDummy.Visible = True
            CmbDummy.SetFocus
            
        Else
            txtDummy.Left = .Left + .CellLeft
            txtDummy.Top = .Top + .CellTop
            txtDummy.Width = .CellWidth
            txtDummy.Height = .CellHeight
            txtDummy.Text = .Text
            txtDummy.Visible = True
            txtDummy.SetFocus
            
            'ダミーテキストの最終にフォーカス移動
            SendKeys "{END}"
        End If
    
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
        CmbDummy.Visible = False
        CmbDummy.Clear
        txtDummy.Visible = False
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.16.0.1) 2009-12-21  REVISED BY [TCC] E.Watanabe
'//                 内部保存エリアへの格納ループカウンタ最大値を修正
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                設定反映フラグ追加
'//     REVISIONS :(EG30 V33.2.0.1) 2017-10-05  CODED BY  [TCC] T.Nakajima
'//                 2017年度施策 現地版対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_Click()

    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim iLoopCnt2           As Integer                      ' ループカウンタ
    Dim byBuff()            As Byte                         'バイトバッファ

    'エラールーチンを宣言
    On Error Resume Next

    'グリッドに入力項目を反映させる
' EG20 V3.0.0.2 （駅都度修正対応）削除開始
'    If Bunrui_Sho_Type.GATE_TYPE_SHUBETU = iBunrui_Sho_Save(GridIni.Col) Or _
'       Bunrui_Sho_Type.GATE_TYPE_TURO = iBunrui_Sho_Save(GridIni.Col) Then
' EG20 V3.0.0.2 （駅都度修正対応）削除終了
' EG20 V3.0.0.2 （駅都度修正対応）追加開始
' EG30 V33.2.0.1 DEL START
'    If Bunrui_Sho_Type.GATE_TYPE_SHUBETU = iBunrui_Sho_Save(GridIni.Col) Or _
'       Bunrui_Sho_Type.GATE_TYPE_HANSOU = iBunrui_Sho_Save(GridIni.Col) Or _
'       Bunrui_Sho_Type.GATE_TYPE_TURO = iBunrui_Sho_Save(GridIni.Col) Then
' EG30 V33.2.0.1 DEL END
' EG20 V3.0.0.2 （駅都度修正対応）追加終了
' EG30 V33.2.0.1 ADD START
    If Bunrui_Sho_Type.GATE_TYPE_SHUBETU = iBunrui_Sho_Save(GridIni.Col) Or _
       Bunrui_Sho_Type.GATE_TYPE_HANSOU = iBunrui_Sho_Save(GridIni.Col) Or _
       Bunrui_Sho_Type.GATE_TYPE_TURO = iBunrui_Sho_Save(GridIni.Col) Or _
       Bunrui_Sho_Type.GATE_TYPE_ICMTURO = iBunrui_Sho_Save(GridIni.Col) Then
' EG30 V33.2.0.1 ADD END
        'V1.20.0.1 ADD START
        If GridIni.Text <> CmbDummy.Text Then
            '設定反映フラグ（変更あり）
            SetteiHaneiFlg = True
        End If
        'V1.20.0.1 ADD END

        GridIni.Text = CmbDummy.Text
    Else

        'V1.20.0.1 ADD START
        If GridIni.Text <> txtDummy.Text Then
            '設定反映フラグ（変更あり）
            SetteiHaneiFlg = True
        End If
        'V1.20.0.1 ADD END

        GridIni.Text = txtDummy.Text
    End If

'    For iLoopCnt = 0 To UBound(KikiDataTbl) - 1            'V1.16.0.1 DEL
    For iLoopCnt = 0 To UBound(KikiDataTbl)                 'V1.16.0.1 ADD

        '該当データ検索
        If (BUNRUI_DAI.DAI_Gate = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
           (GridIni.Row = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
           (GridIni.Col = KikiDataTbl(iLoopCnt).iBunrui_Sho) Then

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
'//     REVISIONS :(1.16.0.1) 2009-12-21  REVISED BY [TCC] E.Watanabe
'//                 内部保存エリアへの格納ループカウンタ最大値を修正
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                号機番号の入力桁数を制御
'//                設定反映フラグ追加
'//     REVISIONS :(EG20 V6.4.0.1) 2012-06-17 REVISED BY [TCC] H.Sugimoto
'//                【総点検修正対応：半角スペースの入力を抑止する対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtDummy_Change()
    
    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim iLoopCnt2           As Integer                      ' ループカウンタ
    Dim byBuff()            As Byte                         'バイトバッファ

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
    For iLoopCnt = 0 To UBound(uKetaFileData)
    
        With uKetaFileData(iLoopCnt)
        
            '行のタイトルとINIの項目名が一致していたとき
            'INIを読込んでいないときはチェックしない
            If "" <> .sName And GridIni.TextMatrix(0, GridIni.Col) = .sName Then
                
                '桁数オーバーしたとき
                If Len(GridIni.Text) > .iKeta Then
                    '古い桁は捨てて右二桁を切り取る
                    GridIni.Text = Right$(GridIni.Text, .iKeta)
                    
                    'ダミーテキストの最終にフォーカス移動
                    SendKeys "{END}"
                    
                    Exit For
                End If
            End If
        End With
    Next
    'V1.20.0.1 ADD END

'    For iLoopCnt = 0 To UBound(KikiDataTbl) - 1            'V1.16.0.1 DEL
    For iLoopCnt = 0 To UBound(KikiDataTbl)                 'V1.16.0.1 ADD

        '該当データ検索
        If (BUNRUI_DAI.DAI_Gate = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
           (GridIni.Row = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
           (GridIni.Col = KikiDataTbl(iLoopCnt).iBunrui_Sho) Then

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
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                 カーソル移動の処理を削除
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '特殊キーを押下された時、下記の処理を行う
    bScroll = True
    On Err GoTo ShoriErr
    
    With GridIni
        'V1.20.0.1 DEL START
'        '←を押下された時
'        If KeyCode = 37 Then
'            If .Col <> 1 Then
'                'セルを左に一つ移動
'                .Col = .Col - 1
'            End If
        'V1.20.0.1 DEL END
'        '↑を押下された時
'        ElseIf KeyCode = 38 Then
'            If .Row <> 1 Then
'                'セルを上に一つ移動
'                .Row = .Row - 1
'            End If
        '→、またはenterを押下された時
'        ElseIf KeyCode = 39 Or KeyCode = 13 Then   'V1.20.0.1 DEL
        If KeyCode = 13 Then                        'V1.20.0.1 ADD
            If .Col <> .Cols - 1 Then
                'セルを右に一つ移動
                .Col = .Col + 1
            End If
'        '↓を押下された時
'        ElseIf KeyCode = 40 Then
'            If .Row <> .Rows - 1 Then
'                'セルを下に一つ移動
'                .Row = .Row + 1
'            End If
        End If

        If sInitCombDummy = False Then
            'ダミーテキストのセット
            CmbDummy.Left = .Left + .CellLeft
            CmbDummy.Top = .Top + .CellTop
            CmbDummy.Width = .CellWidth
            CmbDummy.Height = .CellHeight
            CmbDummy.Text = .Text
            CmbDummy.Visible = True
            CmbDummy.SetFocus
        Else
            'ダミーテキストのセット
            txtDummy.Left = .Left + .CellLeft
            txtDummy.Top = .Top + .CellTop
            txtDummy.Width = .CellWidth
            txtDummy.Height = .CellHeight
            txtDummy.Text = .Text
            txtDummy.Visible = True
            txtDummy.SetFocus
        End If
    End With
    bScroll = False

ShoriErr:

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
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '特殊キーを押下された時、下記の処理を行う
    bScroll = True
    On Err GoTo ShoriErr
    
    With GridIni
        'V1.20.0.1 DEL START
'        '←を押下された時
'        If KeyCode = 37 Then
'            If .Col <> 1 Then
'                'セルを左に一つ移動
'                .Col = .Col - 1
'            End If
'        '↑を押下された時
'        ElseIf KeyCode = 38 Then
'            If .Row <> 1 Then
'                'セルを上に一つ移動
'                .Row = .Row - 1
'            End If
'        '→、またはenterを押下された時
'        ElseIf KeyCode = 39 Or KeyCode = 13 Then
        'V1.20.0.1 DEL END
        If KeyCode = 13 Then    'V1.20.0.1 ADD
            If .Col <> .Cols - 1 Then
                'セルを右に一つ移動
                .Col = .Col + 1
            End If
        'V1.20.0.1 DEL START
'        '↓を押下された時
'        ElseIf KeyCode = 40 Then
'            If .Row <> .Rows - 1 Then
'                'セルを下に一つ移動
'                .Row = .Row + 1
'            End If
        'V1.20.0.1 DEL START
        End If

        If sInitCombDummy = False Then
            'ダミーテキストのセット
            CmbDummy.Left = .Left + .CellLeft
            CmbDummy.Top = .Top + .CellTop
            CmbDummy.Width = .CellWidth
            CmbDummy.Height = .CellHeight
            CmbDummy.Text = .Text
            CmbDummy.Visible = True
            CmbDummy.SetFocus
        Else
            'ダミーテキストのセット
            txtDummy.Left = .Left + .CellLeft
            txtDummy.Top = .Top + .CellTop
            txtDummy.Width = .CellWidth
            txtDummy.Height = .CellHeight
            txtDummy.Text = .Text
            txtDummy.Visible = True
            txtDummy.SetFocus
        End If
    End With
    bScroll = False

ShoriErr:

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
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
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtDummy_LostFocus
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
'//                 フェーズ２対応　「駅情報画面へ」釦押下処理追加
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                設定反映釦の未押下メッセージ追加
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
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
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_GATE_GAMEN_KIKIDATA_INSTOL, 0)
            
            '機器構成項目設定反映処理
            Call sInstolKikiData
            bUnlock = False                     ' EG20 V3.0.0.2 追加

        Case 1                                 ' 機器構成項目媒体出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_GATE_GAMEN_KIKIDATA_OUTPUT, 0)
            
            '機器構成項目媒体出力処理
            Call sKikiDataOutPut
    
        Case 2                                 ' 機器構成項目内部保存
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_GATE_GAMEN_KIKIDATA_SAVE, 0)
            
            '機器構成項目内部保存処理
            Call sKikiDataSave
        
        Case 3                                 ' 機器構成設定データ選択
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_GATE_GAMEN_KIKIDATA_SELECT, 0)
            
            '機器構成設定データ選択処理
            Call sKikiDataSelect
    
        Case 4                                 ' 媒体入力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_GATE_GAMEN_MEDIUM_INPUT, 0)
            
            '媒体入力処理
            Call sInputMedium
    
        Case 5                                 ' 媒体取外
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
    
            '媒体取外処理
            Call pfRemove(Me)
'V1.4.0.1 ADD START
        Case 6                                 ' 駅情報画面へ
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
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIINFO_GAMEN_GO_BUTTOM, 0)
            Unload Me
            Load frmKikiData
            frmKikiData.Show 1
            Exit Sub         'V1.20.0.1 ADD
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
'//                種別正当性チェックを追加
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ポップアップ画面タイトル修正
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.76修正対応】
'//     REVISIONS :(EG20 V5.12.0.1) 2012-05-18  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.6.0.1)  2012-06-20  CODED BY  [TCC] H.Sugimoto
'//                 【項目が未入力時に設定反映を行わない対応】
'//     REVISIONS :(EG20 V6.7.0.1)  2012-06-28  CODED BY  [TCC] H.Sugimoto
'//                 【項目チェックの対象を改札機情報のみとする修正】
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
    Dim iLoopCnt2           As Integer          'ループカウンタ2     'V1.20.0.1 ADD
    Dim bSysChange          As Boolean          'コンピュータ名、ネットワーク変更処理判定   'V1.12.0.1 ADD
    
    Dim lResult             As Long             ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加
' EG20 V5.12.0.1追加開始（計算に利用する変数をLONG型に変更）
    Dim lLoop               As Long             ' ループカウンタ
    Dim lRecord             As Long             ' レコード
    Dim lIndex              As Long             ' インデックス
    Dim lSize               As Long             ' サイズ
' EG20 V5.12.0.1追加終了（計算に利用する変数をLONG型に変更）
    
    'エラールーチンを宣言
    On Error Resume Next

' EG20 V6.6.0.1追加開始
    lRecord = UBound(KikiDataTbl)
    For lLoop = 0 To lRecord
      If KikiDataTbl(lLoop).iBunrui_Dai = BUNRUI_DAI.DAI_Gate Then      ' EG20 V6.7.0.1追加
        If KikiDataTbl(lLoop).strData(0) = vbEmpty Then
            MsgBox "設定値の入力されていない項目があります。" & Chr(vbKeyReturn) & _
                    "設定内容を確認してください。", vbCritical, "設定反映チェック異常"
            Call SetEnableTrue
            Exit Sub
        End If
      End If                                                            ' EG20 V6.7.0.1追加
    Next lLoop
' EG20 V6.6.0.1追加終了
    
    'V1.20.0.1 ADD START
    '種別正当性チェック
    For iLoopCnt = 0 To UBound(uHikakuFileData)
        
        With uHikakuFileData(iLoopCnt)
        
            '行ごとに比較していく
            For iLoopCnt2 = 1 To GridIni.Rows - 1
                
                '表示内容とINIファイルの指定文字が同じ場合（2つの比較対象のどちらか一方でも同じ場合）。
                'INIを読込んでいないときはチェックしない
                If .sName1 <> "" And .sName2 <> "" And _
                   (GridIni.TextMatrix(iLoopCnt2, .iCol1) = .sMoji1 Or _
                    GridIni.TextMatrix(iLoopCnt2, .iCol2) = .sMoji2) Then
                    
                    '表示内容とINIファイルの指定文字が2つとも一致していなければならない
                    If GridIni.TextMatrix(iLoopCnt2, .iCol1) <> .sMoji1 Or _
                       GridIni.TextMatrix(iLoopCnt2, .iCol2) <> .sMoji2 Then
                        
                        MsgBox .sName1 & "と" & .sName2 & "の設定値が不正です。" & Chr(vbKeyReturn) _
                               & "正しい値を入力してください。", vbExclamation, "設定反映正当性チェック異常"
                        Call SetEnableTrue                      ' EG20 V5.0.4.1【結合TR-No.76修正対応】追加
                        Exit Sub
                    End If
                End If
            Next
            
        End With
    Next
    'V1.20.0.1 ADD END
    
'V1.8.0.1 DEL START
'    iResponse = MsgBox("機器構成データをインストールします。" & Chr(vbKeyReturn) & _
'                        "よろしいですか？", _
'                        vbYesNo + vbExclamation, _
'                        "媒体入力確認")
'V1.8.0.1 DEL END
'V1.8.0.1 ADD START
'V1.21.0.1 DEL START
'    iResponse = MsgBox("機器構成データをインストールします。" & Chr(vbKeyReturn) & _
'                        "よろしいですか？", _
'                        vbOKCancel + vbExclamation, _
'                        "媒体入力確認")
'V1.8.0.1 ADD END
'V1.21.0.1 DEL END
'V1.21.0.1 ADD START
    iResponse = MsgBox("機器構成データをインストールします。" & Chr(vbKeyReturn) & _
                        "よろしいですか？", _
                        vbOKCancel + vbExclamation, _
                        "設定反映確認")
'V1.21.0.1 ADD END
'    If iResponse = vbNo Then Exit Sub      'V1.10.0.1 DEL
    If iResponse = vbCancel Then
        Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
        Exit Sub   'V1.10.0.1 ADD
    End If
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '構造体配列をバイナリ配列に変換
' EG20 V5.12.0.1削除開始（計算に利用する変数をLONG型に変更）
'    ReDim bData((UBound(KikiDataTbl) + 1) * Len(KikiDataTbl(0))) As Byte
'    For iLoopCnt = 0 To UBound(KikiDataTbl)
'          MoveMemory bData(iLoopCnt * Len(KikiDataTbl(0))), KikiDataTbl(iLoopCnt), Len(KikiDataTbl(iLoopCnt))
'    Next
' EG20 V5.12.0.1削除終了（計算に利用する変数をLONG型に変更）
' EG20 V5.12.0.1追加開始（計算に利用する変数をLONG型に変更）
    lSize = Len(KikiDataTbl(0))
    lRecord = UBound(KikiDataTbl)
    ReDim bData((lRecord + 1) * lSize) As Byte
    For lLoop = 0 To lRecord
        lIndex = lLoop * lSize
        MoveMemory bData(lIndex), KikiDataTbl(lLoop), lSize
    Next
' EG20 V5.12.0.1追加終了（計算に利用する変数をLONG型に変更）
    
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
        'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果")　 'V1.21.0.1 DEL
        iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理設定反映結果") 'V1.21.0.1 ADD
        Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
    Else
'        'ログ出力
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'
'        '正常終了
'        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体入力結果")
'    End If
'V1.12.0.1 START ADD
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
            'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果") 'V1.21.0.1 DEL
             iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理設定反映結果") 'V1.21.0.1 ADD
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
            'iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体入力結果") 'V1.21.0.1 DEL
            iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "反映処理設定反映結果")  'V1.21.0.1 ADD
            
            '設定反映フラグ（変更なし）
            SetteiHaneiFlg = False      'V1.20.0.1 ADD
            Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
        End If
    End If
'V1.12.0.1 START END


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
    'V1.20.0.1 DEL START
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
    'V1.20.0.1 DEL END
    
    'sWriteDir = pfDirSelection(strDrive, "機器構成ファイル書込み先のディレクトリ選択") 'V1.20.0.1 DEL
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
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "機器構成項目媒体出力結果")  'V1.8.0.1 DEL
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "機器構成項目媒体出力結果")   'V1.8.0.1 ADD  'V1.13.0.1 DEL
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体出力結果")                'V1.13.0.1 ADD

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
'//     REVISIONS :(1.12.0.1) 2009-11-16   REVISED BY [TCC] C.Terui
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
    Dim sMyPath(1 To 3)      As String          'ファイルパス
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim iLoopCount           As Integer         'ループカウンタ
    Dim intFileNo            As Integer         'ファイル番号

    'エラールーチンを宣言
    On Error Resume Next
    
    '初期値設定
    strFileName = ""

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
'V1.12.0.1 DEL START
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
        'iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "機器構成項目内部保存結果")   'V1.13.0.1 DEL
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "一時保存結果")    'V1.13.0.1 ADD
    
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
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbInformation, "機器構成項目内部保存結果")  'V1.8.0.1 DEL
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "機器構成項目内部保存結果")   'V1.8.0.1 ADD 'V1.13.0.1 DEL
     iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存結果")   'V1.13.0.1 ADD
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
'//     REVISIONS :(1.12.0.1) 2009-11-16   REVISED BY [TCC] C.Terui
'//                 ファイル検索処理削除
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                 コピーファイルパス指定を修正
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                設定反映フラグ追加
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ポップアップ画面タイトル修正
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
'
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
'        FileCopy sWriteDir, KIKI_DATA_FILE   'V1.13.0.1 DEL
         FileCopy KIKI_DATA_S_FILE, KIKI_DATA_FILE   'V1.13.0.1 ADD
        
        '機器情報設定（自改）イメージファイル作成
        bRet = dllGetKikiIniData(1, 1, KIKI_DATA_SET_GATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
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
            'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体入力結果")  'V1.8.0.1 ADD
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存データ取込結果")  'V1.21.0.1 ADD
            Exit Sub
        End If
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
'V1.12.0.1 ADD START
        '一時保存ファイル削除
        Kill KIKI_DATA_BACKUP_FILE
'V1.12.0.1 ADD END
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
        '正常終了
'        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "機器構成設定データ選択結果")  'V1.13.0.1 DEL
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "一時保存データ取込結果")       'V1.13.0.1 ADD
    
        '画面表示処理
        Call sDisp
        
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
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "機器構成設定データ選択結果")    'V1.8.0.1 ADD  'V1.13.0.1 DEL
     iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "一時保存データ取込結果")        'V1.13.0.1 ADD

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
''        strDrive = "a:"    'V1.12.0.1 DEL
    '    strDrive = "H:"     'V1.12.0.1 ADD
    'End If
    '
    ''媒体ファイル名取得
''    strFileName = pfFileSelection(strDrive, "*.csv", "媒体入力ﾌｧｲﾙ選択")   'V1.12.0.1 DEL
    'strFileName = pfFileSelection(strDrive, "KIKI_DATA.CSV", "媒体入力ﾌｧｲﾙ選択")    'V1.12.0.1 ADD
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
        If pfChangeAttrNormal (strFileName, PATH_HOSHUTMP_KIKI_DATA, KIKI_DATA_FILE) = False Then
            GoTo COPY_ERROR
        End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
        
        '機器情報設定（自改）イメージファイル作成
        bRet = dllGetKikiIniData(1, 1, KIKI_DATA_SET_GATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
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
        
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
        '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体入力結果")
    
        '画面表示処理
        Call sDisp
        
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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sInitCombDummy
'//  機能名称  : コンボボックス初期値設定処理
'//  機能概要  : コンボボックスの初期値を設定する
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
'//     REVISIONS :(EG30 V33.2.0.1) 2017-10-05 CODED BY  [TCC] T.Nakajima
'//                 2017年度施策 現地版対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sInitCombDummy() As Integer

    With GridIni
        
        If Bunrui_Sho_Type.GATE_TYPE_SHUBETU = iBunrui_Sho_Save(.Col) Then
            CmbDummy.Clear
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'            CmbDummy.AddItem "Ｅ"
'            CmbDummy.AddItem "Ｎ"
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
            CmbDummy.AddItem "次"           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            CmbDummy.AddItem "＊"
            sInitCombDummy = False
' EG30 V33.2.0.1 DEL START
'        ElseIf Bunrui_Sho_Type.GATE_TYPE_TURO = iBunrui_Sho_Save(.Col) Then
' EG30 V33.2.0.1 DEL END
' EG30 V33.2.0.1 ADD START
        ElseIf Bunrui_Sho_Type.GATE_TYPE_TURO = iBunrui_Sho_Save(.Col)  Or _
                Bunrui_Sho_Type.GATE_TYPE_ICMTURO = iBunrui_Sho_Save(.Col) Then
' EG30 V33.2.0.1 ADD END
            CmbDummy.Clear
            CmbDummy.AddItem "両"
            CmbDummy.AddItem "集"
            CmbDummy.AddItem "改"
            CmbDummy.AddItem "＊"
            sInitCombDummy = False
' EG20 V3.0.0.2 （駅都度修正対応）追加開始
        ElseIf Bunrui_Sho_Type.GATE_TYPE_HANSOU = iBunrui_Sho_Save(.Col) Then
            CmbDummy.Clear
            CmbDummy.AddItem "両"           ' 両用機
            CmbDummy.AddItem "集"           ' 集札専用機
            CmbDummy.AddItem "改"           ' 改札専用機
            CmbDummy.AddItem "無"           ' 無し
            CmbDummy.AddItem "＊"           ' 未設置
            sInitCombDummy = False
' EG20 V3.0.0.2 （駅都度修正対応）追加終了
        Else
            sInitCombDummy = True
        End If
    End With
        
End Function
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
    CmdKikiSetMenu(7).Enabled = False       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    cmdCancel.Enabled = False
    
    'CmdKikiSetMenu(0)～(2)は条件によっては元々押下不可のため判定を行う
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
    CmdKikiSetMenu(7).Enabled = True        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    cmdCancel.Enabled = True
    
    'コンボボックスとCmdKikiSetMenu(0)～(2)は条件によっては元々押下不可のため、画面表示の有無で判定を行う
    strFileName = Dir(KIKI_DATA_SET_EKI_INFO_FILE)
    'ファイルが存在する場合
    If strFileName <> "" Then
        CmdKikiSetMenu(0).Enabled = True
        CmdKikiSetMenu(1).Enabled = True
        CmdKikiSetMenu(2).Enabled = True
    End If
    
End Sub
'V1.12.0.1 ADD END

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : psGetFileChk
'//  機能名称  : 正当性チェックINIファイル読込み
'//  機能概要  : INIファイル読込み関数をCALL
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-19  REVISED BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psGetFileChk()
    
    Dim bRet As Boolean     '関数戻り値
   
    'エラールーチンを宣言
    On Error Resume Next
    
    'INIファイル読込み
    bRet = pfGetFile_uKeta
    
    'INIファイル読込み
    If bRet = True Then
        bRet = pfGetFile_uHikaku
    End If
    
    'INIファイル取得異常なら
    If bRet = False Then
    
        'グリッドタイトル設定
        Call sDispGridTitol
        
        'グリッドデータ部クリア処理
        Call sDispDataClear
        
        '処理釦押下不可能設定
        CmdKikiSetMenu(0).Enabled = False
        CmdKikiSetMenu(1).Enabled = False
        CmdKikiSetMenu(2).Enabled = False
        CmdKikiSetMenu(3).Enabled = False
        CmdKikiSetMenu(4).Enabled = False
        CmdKikiSetMenu(5).Enabled = False
        
        'INIファイル有無チェック異常時：「ファイル異常」ポップアップを表示
        MsgBox "INIファイルの取得に失敗しました｡", vbCritical, "ファイル異常"
        
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : pfGetFile_uKeta
'//  機能名称  : 桁数チェックのためのINIファイル読込み
'//  機能概要  : INIファイルより正当性チェックに使用する情報を読込む
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    :Boolean
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-19  REVISED BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetFile_uKeta() As Boolean

    Dim iRet As Integer                 '関数の戻り値
    Dim sKeyName As String              'INIファイルキー名
    Dim iChar As Integer                '読み込み文字数
    Dim iWord As Integer                '読み込みワード数
    Dim iModCnt As Integer              '読み込み項目数（配列の要素数）
    Dim sIni_Data As String * 128       'INIファイルより1行分取得
    Dim iLoopCnt As Integer             'ループカウンタ
    Dim iLoopCnt2 As Integer            'ループカウンタ2
    Dim MyName As String                'INI有無チェック
    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
    Dim bTitleFlg As Boolean            'INIのタイトル名が列名として存在するか

    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード
    
    'エラールーチンを宣言
    On Error GoTo FileError
    
    '初期化
    pfGetFile_uKeta = False
    iModCnt = 0
    ReDim uKetaFileData(iModCnt)
    
    'ファイル有無チェック
    If fsoObj.FileExists(KIKI_KOUSEI_CEHK_FILE) = False Then
        GoTo FileError
    End If
    
    '------------------------------------------------
    '桁数チェック情報の読み込み
    '------------------------------------------------
    For iLoopCnt = 0 To iModMax
        sKeyName = KOUSEI_SEC1_KEY1 & Format(iLoopCnt, "00")
        iRet = GetPrivateProfileString(KOUSEI_SEC1, _
                                       sKeyName, _
                                       DEFAILT, sIni_Data, Len(sIni_Data), _
                                       KIKI_KOUSEI_CEHK_FILE)
        iChar = 1
        iWord = 1
        
        '読み込み正常のときだけエリアに格納
        If iRet > 0 Then
            ReDim Preserve uKetaFileData(iModCnt)
            Do
               'モジュール情報格納エリアに1行分のデータを保持させる。
                If Mid(sIni_Data, iChar, 1) = "," Or Mid(sIni_Data, iChar, 1) = vbNullChar Then
                    Select Case iWord
                        Case 1
                            uKetaFileData(iModCnt).sName = Left(sIni_Data, iChar - 1)
                            sIni_Data = Mid(sIni_Data, iChar + 1)
                            iChar = 0
                            iWord = iWord + 1
                        Case 2
                            uKetaFileData(iModCnt).iKeta = Left(sIni_Data, iChar - 1)
                            sIni_Data = Mid(sIni_Data, iChar + 1)
                            iChar = 0
                            iWord = iWord + 1
                    End Select
                End If
                iChar = iChar + 1
                If iChar > Len(sIni_Data) Then
                    Exit Do
                End If
            Loop
            iModCnt = iModCnt + 1
        End If
    Next
        
    Set fsoObj = Nothing
    pfGetFile_uKeta = True

    Exit Function

FileError:

    'ログ出力「INIファイル読込異常」
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_INI_READ, lngErrCode)
    'ログ出力　┗ファイル名
    Call psFileNameGet(KIKI_KOUSEI_CEHK_FILE, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    Set fsoObj = Nothing
    pfGetFile_uKeta = False
    
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : pfGetFile_uHikaku
'//  機能名称  : 桁数チェックのためのINIファイル読込み
'//  機能概要  : INIファイルより正当性チェックに使用する情報を読込む
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    :Boolean
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-19  REVISED BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetFile_uHikaku() As Boolean

    Dim iRet As Integer                 '関数の戻り値
    Dim sKeyName As String              'INIファイルキー名
    Dim iChar As Integer                '読み込み文字数
    Dim iWord As Integer                '読み込みワード数
    Dim iModCnt As Integer              '読み込み項目数（配列の要素数）
    Dim sIni_Data As String * 128       'INIファイルより1行分取得
    Dim iLoopCnt As Integer             'ループカウンタ
    Dim iLoopCnt2 As Integer            'ループカウンタ2
    Dim MyName As String                'INI有無チェック
    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
    Dim bTitleFlg As Boolean            'INIのタイトル名が列名として存在するか

    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード
    
    'エラールーチンを宣言
    On Error GoTo FileError
    
    '初期化
    pfGetFile_uHikaku = False
    iModCnt = 0
    ReDim uHikakuFileData(iModCnt)
    
    'ファイル有無チェック
    If fsoObj.FileExists(KIKI_KOUSEI_CEHK_FILE) = False Then
        GoTo FileError
    End If
    
    '------------------------------------------------
        '種別の正当性チェック情報読み込み
    '------------------------------------------------
    For iLoopCnt = 0 To iModMax
        sKeyName = KOUSEI_SEC1_KEY2 & Format(iLoopCnt, "00")
        iRet = GetPrivateProfileString(KOUSEI_SEC1, _
                                       sKeyName, _
                                       DEFAILT, sIni_Data, Len(sIni_Data), _
                                       KIKI_KOUSEI_CEHK_FILE)
        iChar = 1
        iWord = 1
        
        '読み込み正常のときだけエリアに格納
        If iRet > 0 Then
            ReDim Preserve uHikakuFileData(iModCnt)
            Do
               'モジュール情報格納エリアに1行分のデータを保持させる。
                If Mid(sIni_Data, iChar, 1) = "," Or Mid(sIni_Data, iChar, 1) = vbNullChar Then
                    Select Case iWord
                        Case 1
                            uHikakuFileData(iModCnt).sName1 = Left(sIni_Data, iChar - 1)
                            sIni_Data = Mid(sIni_Data, iChar + 1)
                            iChar = 0
                            iWord = iWord + 1
                        Case 2
                            uHikakuFileData(iModCnt).sName2 = Left(sIni_Data, iChar - 1)
                            sIni_Data = Mid(sIni_Data, iChar + 1)
                            iChar = 0
                            iWord = iWord + 1
                        Case 3
                            uHikakuFileData(iModCnt).sMoji1 = Left(sIni_Data, iChar - 1)
                            sIni_Data = Mid(sIni_Data, iChar + 1)
                            iChar = 0
                            iWord = iWord + 1
                        Case 4
                            uHikakuFileData(iModCnt).sMoji2 = Left(sIni_Data, iChar - 1)
                            sIni_Data = Mid(sIni_Data, iChar + 1)
                            iChar = 0
                            iWord = iWord + 1
                    End Select
                End If
                iChar = iChar + 1
                If iChar > Len(sIni_Data) Then
                    Exit Do
                End If
            Loop
            iModCnt = iModCnt + 1
        End If
    Next
    
    '----------------------------------------------------------
    '表示タイトルと比較し、必要な項目が何カラム目かを格納
    '----------------------------------------------------------
    For iLoopCnt = 0 To UBound(uHikakuFileData)
        For iLoopCnt2 = 1 To GridIni.Cols - 1
            '行のタイトルとINIの項目名が一致していたとき
            If GridIni.TextMatrix(0, iLoopCnt2) = uHikakuFileData(iLoopCnt).sName1 Then
                uHikakuFileData(iLoopCnt).iCol1 = iLoopCnt2
            End If
            '行のタイトルとINIの項目名が一致していたとき
            If GridIni.TextMatrix(0, iLoopCnt2) = uHikakuFileData(iLoopCnt).sName2 Then
                uHikakuFileData(iLoopCnt).iCol2 = iLoopCnt2
            End If
        Next
    Next
        
    Set fsoObj = Nothing
    pfGetFile_uHikaku = True

    Exit Function

FileError:

    'ログ出力「INIファイル読込異常」
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_INI_READ, lngErrCode)
    'ログ出力　┗ファイル名
    Call psFileNameGet(KIKI_KOUSEI_CEHK_FILE, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    Set fsoObj = Nothing
    pfGetFile_uHikaku = False

End Function
'V1.20.0.1 ADD END

