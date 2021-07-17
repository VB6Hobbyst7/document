VERSION 5.00
Begin VB.Form frmTomasDataOut 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Timer tmrOutput 
      Left            =   480
      Top             =   0
   End
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
      TabIndex        =   0
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmTomasDataOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmTomasDataOut.frm
'//  パッケージ名：TOMASデータ媒体出力画面
'//
'//  概要：バージョン管理画面
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//                 新規作成【フェーズ３ TOMAS対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////

Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
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
'//  関数名称  : Form_Activate
'//  機能名称  : TOMASデータ表示画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
    
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TOMAS_DATA_DISP)

    '処理中のガイドを表示する｡
    lblMessage(0) = "出力中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    DoEvents
    
    tmrMail.Enabled = True
    tmrOutput.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : TOMASデータ表示画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next
    
    'タイマを止める
    tmrMail.Enabled = False
    tmrOutput.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : TOMASデータ表示画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '出力処理開始用タイマの値を設定する。
    tmrOutput.Interval = 100
    tmrOutput.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
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
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'プロセスの終了処理を行う
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示」を受信した場合
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '表示元画面（保守データ収集画面）をアクティブ表示する。
'                AppActivate frmRenewOutput.Caption, False  ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
                AppActivate frmTomasDataOut.Caption, False  ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
                pfFormActive (frmTomasDataOut.hwnd)         ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
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
'//  関数名称  : tmrOutput_Timer
'//  機能名称  : 出力処理実行タイマ
'//  機能概要  : 媒体出力する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrOutput_Timer()

    On Error Resume Next
    
    tmrOutput.Enabled = False
    Call sOutput_Data
     
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sOutput_Data
'//  機能名称  : 設定値出力処理
'//  機能概要  : 設定値を編集して媒体出力する
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
Private Sub sOutput_Data()

    Dim bySyoAssort As Byte             'ログ用小分類
    Dim strFilePath As String           '出力ファイルパス
    Dim strCornerPath As String         '設定ファイルパス
    Dim strStationNm As String          '駅名
    Dim strCornerNm As String           'コーナ名
    Dim intCount As Integer             'カウンタ
    Dim intCount2 As Integer            'カウンタ
    Dim intOutFile As Integer           '出力ファイル番号
    Dim intTgtFileNo As Integer         '出力対象設定ファイル番号
    Dim strTgtFileName As String        '出力対象設定ファイル
    Dim strTargetFile() As String       '出力対象ファイル
    Dim intFileNum As Integer
    Dim strDefault As String
    Dim strRet As String * 32
    Dim lngRet As Long
    Dim strOutFileName As String
    Dim strFileName As String
    Dim strCabTarget As String
    Dim lngRetZip As Long
    Dim objFileObj As FileSystemObject  'ファイルシステムオブジェクト
    Const lngBufSize = 32
    
    On Error GoTo Err_Handler
    
    Set objFileObj = New FileSystemObject
    
    Select Case gintTomasDataDispDiv
    Case TOMAS_DISP_DIV.TOMAS_DATA_VERSION
        strFileName = TOMAS_FILE_VERINFO
        
    Case TOMAS_DISP_DIV.TOMAS_DATA_KIKI
        strFileName = TOMAS_FILE_KIKIINFO
        
    Case TOMAS_DISP_DIV.TOMAS_DATA_ERR
        strFileName = TOMAS_FILE_ERRINFO
    Case Else
    End Select
    
    strOutFileName = gstrOutPath & strFileName
    
    '出力対象設定ファイルが存在しない場合は異常終了
    If objFileObj.FileExists(PATH_WORK & strFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strFileName, 0)
        GoTo Err_Handler
    End If
    
    'TOMASデータテキストファイルをコピーする
    Call objFileObj.CopyFile(PATH_WORK & strFileName, strOutFileName, True)
    
    Set objFileObj = Nothing
    
    lblMessage(0).Caption = "正常終了しました。"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    DoEvents
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    Exit Sub
    
'エラー処理
Err_Handler:

    Set objFileObj = Nothing
    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, bySyoAssort, KAKARISET_OUTPUT_ERR, 0)
    
    lblMessage(0).Caption = "異常終了しました。"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
End Sub


