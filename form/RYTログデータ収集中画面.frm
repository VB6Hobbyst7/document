VERSION 5.00
Begin VB.Form frmRYTSyusyuCyu 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ｏ Ｋ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmRYTSyusyuCyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmRYTSyusyuCyu.frm
'//  パッケージ名：ＲＹＴログデータ収集中画面
'//
'//  概要：ＲＹＴログデータ収集中画面
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　ＲＹＴログデータ収集中画面追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ＲＹＴログデータ収集中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
      
    Dim uMail As MAIL_RYT_LOG_CMD           'メール
    Dim bRtn As Boolean                 'メールの戻り値
    Dim lExitCode As Long
      
    On Error Resume Next
   
    'ＲＳプロセスに「ＲＹＴログデータ収集要求ＣＭＤ」を送信する。
    uMail.mlHeader.dwId = ML_ID_RYT_LOG_CMD
    uMail.mlHeader.dwSize = MlSize.RYT_LOG_CMD
    uMail.mlHeader.dwProid = RHOSHU_ID
    uMail.mlHeader.dwSubArea = 0
    uMail.dwRequestType = MailRYTType.ML_DT_LOGDATA_ID
    bRtn = DssSendMail(MAIL_SLOT_RYT, MlSize.RYT_LOG_CMD, uMail.mlHeader)
    If bRtn <> 0 Then
       '「ＲＹＴログデータ収集要求ＣＭＤ送信正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, RYT_LOG_CMD_OK, 0)
         
       '収集中のガイドを表示する｡
       lblMessage(0) = "ＲＹＴログデータを収集中です。"
       lblMessage(1) = "しばらくお待ち下さい。"
       tmrMail.Enabled = True
    Else
       '「ＲＹＴログデータ収集要求ＣＭＤ送信異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, RYT_LOG_CMD_ERROR, 0)
        'ＲＹＴログデータ収集処理結果(異常終了)画面を表示
        sSyusyuEnd (1)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : ＲＹＴログデータ収集中画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
  
  On Error Resume Next
  
  cmdOK.Visible = False
  cmdOK.Enabled = False
  
  'メイル受信用のインタバルタイマ値を設定する。
  tmrMail.Interval = MN_MAIL_INTERVAL
  tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : ＲＹＴログデータ収集中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    
    On Error Resume Next
    
    'メール受信用タイマを止める
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdOK_Click
'//  機能名称  : 「OK」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    
    On Error Resume Next
    
    '自画面を消す。
    Unload Me
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
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    Dim intStatus As Integer         '受信メールチェック結果

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
                 'プロセスの終了処理を行う
                 pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                 '「保守画面アクティブ表示」を受信した場合
                 '「保守画面アクティブ表示要求受信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                 '表示元画面（保守データ収集画面）をアクティブ表示する。
                 AppActivate frmRYTSyusyuCyu.Caption, False
            Case ML_ID_RYT_LOG_RES
                 '「ＲＹＴログデータ収集要求RES」を受信した場合
                 '「ＲＹＴログデータ収集要求ＲＥＳ受信」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, RYT_LOG_RES_RECV, 0)
                '収集結果を表示する。
                If udtReadMail.lngData(0) = 0 Then
                   '収集結果：正常
                       sSyusyuEnd (0)
                Else
                   '収集結果：異常
                       sSyusyuEnd (1)
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
'//  関数名称  : sSyusyuEnd
'//  機能名称  : 収集結果表示処理
'//  機能概要  : RYTログデータ収集結果の結果文言を表示する。
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSyusyuEnd(iEnd As Integer)
    Dim i As Integer       'カウンタ
    Dim lngErrCode As Long 'エラーコード

    On Error Resume Next

    If iEnd = 0 Then
       '正常終了時の文言を表示する。
       lblMessage(0) = "正常終了しました。"
       lblMessage(1) = ""
       '「ＲＹＴログデータ収集処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RYT_LOG_KANRI_GAMEN_SYUSYU_OK, 0)
    Else
       '収集失敗時の文言を表示する。
       lblMessage(0) = "異常終了しました。"
       lblMessage(1) = ""
       '「ＲＹＴログデータ収集処理異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RYT_LOG_KANRI_GAMEN_SYUSYU_ERROR, lngErrCode)
    End If
    
    cmdOK.Visible = True
    cmdOK.Enabled = True
End Sub

