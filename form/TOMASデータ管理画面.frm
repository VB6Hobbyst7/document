VERSION 5.00
Begin VB.Form frmTomasDataMng 
   BorderStyle     =   0  'なし
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   11280
      Top             =   5760
   End
   Begin VB.CommandButton cmdEject 
      Caption         =   "媒体取外"
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
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdDispErrInfo 
      Caption         =   "障害発生情報表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdOutErrInfo 
      Caption         =   "障害発生情報媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdDispVerInfo 
      Caption         =   "バージョン情報表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdOutVerInfo 
      Caption         =   "バージョン情報媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdDispKikiInfo 
      Caption         =   "機器状態データ表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdOutKikiInfo 
      Caption         =   "機器状態データ媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " メンテナンス   画面へ戻る"
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "TOMASデータ管理"
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
      TabIndex        =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmTomasDataMng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmTomasDataMng.frm
'//  パッケージ名：TOMASデータ管理画面
'//
'//  概要：バージョン管理画面
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//                 新規作成【フェーズ３ TOMAS対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   CODED   BY [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】

'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////

Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdDispErrInfo_Click
'//  機能名称  : 「障害発生情報表示」釦押下時処理
'//  機能概要  : 障害発生情報表示を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdDispErrInfo_Click()

    Dim blnRet As Boolean
    Dim lngErrCode As Long
    Dim strFileName As String
    Dim strCommand As String
    Dim lRetVal As Long
    
    On Error Resume Next
    
    'データ種別：障害発生情報
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_ERR
    
    gblnTomasDispErr = False
    
    'TOMASデータ表示中フォームを、モーダルウィンドウで表示する。
    frmTomasDataDisp.Show vbModal
    
    'エラーの場合
    If gblnTomasDispErr = True Then
        Exit Sub
    End If
    
    '正常終了
    strFileName = TOMAS_FILE_ERRINFO
    
    strCommand = MN_EXE_MEMO & PATH_WORK & strFileName      '実行コマンドを作成する
    lRetVal = Shell(strCommand, vbMaximizedFocus)           'ノートパッドを起動する
    AppActivate lRetVal, True                               'アクティブ（前面表示）にする
    SendKeys "{LEFT}", True
    
    '正常終了時は障害発生情報媒体出力釦を活性にする。
    cmdOutErrInfo.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdDispKikiInfo_Click
'//  機能名称  : 「機器状態データ表示」釦押下時処理
'//  機能概要  : 機器状態データ表示を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdDispKikiInfo_Click()
    
    Dim strFileName As String
    Dim strCommand As String
    Dim lRetVal As Long
    Dim strRet As String
    Dim intRetCd As Integer
    
    On Error Resume Next
    
    '「TOMASデータ管理画面：機器状態データ表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TOMAS_DATA_KIKI_DISP, 0)
    
    'データ種別：機器状態データ
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_KIKI
    
    gblnTomasDispErr = False
    
    'TOMASデータ表示中フォームを、モーダルウィンドウで表示する。
    frmTomasDataDisp.Show vbModal
    
    
    '正常終了時は機器状態データ媒体出力釦を活性にする。
    If gblnTomasDispErr = False Then
        cmdOutKikiInfo.Enabled = True
        
        strCommand = MN_EXE_MEMO & PATH_WORK & TOMAS_FILE_KIKIINFO      '実行コマンドを作成する
        lRetVal = Shell(strCommand, vbMaximizedFocus)           'ノートパッドを起動する
        AppActivate lRetVal, True                               'アクティブ（前面表示）にする
        SendKeys "{LEFT}", True
    
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdDispVerInfo_Click
'//  機能名称  : 「バージョン情報表示」釦押下時処理
'//  機能概要  : バージョン情報表示を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdDispVerInfo_Click()
    
    Dim strCommand As String
    Dim lRetVal As Long
    Dim strRet As String
    Dim intRetCd As Integer
    
    On Error Resume Next
    
    '「TOMASデータ管理画面：バージョン情報表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TOMAS_DATA_VER_DISP, 0)
    
    'データ種別：バージョン情報
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION
    
    gblnTomasDispErr = False
    gblnRecvErr = False
    
    'TOMASデータ表示中フォームを、モーダルウィンドウで表示する。
    frmTomasDataDisp.Show vbModal
    
    'バージョン情報表示RESを正常受信した場合は、完了通知を送信する。
    If gblnRecvErr = False Then
        Call fSDATAMailSend_Commit
    End If
    
    '正常終了時はバージョン情報媒体出力釦を活性にする。
    If gblnTomasDispErr = False Then
        
        strCommand = MN_EXE_MEMO & PATH_WORK & TOMAS_FILE_VERINFO      '実行コマンドを作成する
        lRetVal = Shell(strCommand, vbMaximizedFocus)           'ノートパッドを起動する
        AppActivate lRetVal, True                               'アクティブ（前面表示）にする
        SendKeys "{LEFT}", True
        
        cmdOutVerInfo.Enabled = True
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : fSDATAMailSend_Commit
'//  機能名称  : バージョン取得完了通知送信
'//  機能概要  : 初期処理時：メールを送信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSDATAMailSend_Commit() As Boolean

    Dim udtMail As VERSION_DATA_CMT_TYPE
    Dim bRet As Boolean             '関数戻り値
    Dim lngErrCode As Long          'エラーコード
    
    On Error Resume Next
 
    fSDATAMailSend_Commit = True
    
    'TOMASデータ出力要求を送信する
    udtMail.mlHeader.dwId = ML_ID_TOMAS_VARSION_DSP_COMMIT
    udtMail.mlHeader.dwSize = MlSize.TOMAS_DATA_DSP_CMT
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    
    udtMail.dwSeqNo = gintSeqNo_Version                 'シーケンス番号
    udtMail.dwDenbunSize = 8                            '電文サイズ（固定）
    udtMail.byCmd(0) = &H7A
    udtMail.byCmd(1) = &H41
    udtMail.byCmd(2) = &H1
    udtMail.byCmd(3) = &H1
    udtMail.byCmd(4) = &H1
    udtMail.byCmd(5) = &H1
    udtMail.byCmd(6) = &H0
    udtMail.byCmd(7) = &H0
    
    'メール送信
    bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMail), udtMail.mlHeader)
    
    If bRet = False Then
        '「TOMASデータ表示画面：TOMASデータ完了通知」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, TOMAS_DATA_VER_COMMIT, lngErrCode)
        fSDATAMailSend_Commit = False
        Exit Function
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdEject_Click
'//  機能名称  : 媒体取外釦押下時処理
'//  機能概要  : 媒体取外処理を実行する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdEject_Click()
    
    On Error Resume Next
    
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
    
    '媒体取外処理
    Call pfRemove(Me)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdOutVerInfo_Click
'//  機能名称  : 「バージョン情報媒体出力」釦押下時処理
'//  機能概要  : バージョン情報テキストファイルを媒体に出力する。
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
Private Sub cmdOutVerInfo_Click()

    On Error GoTo Err_Handler
    
    '媒体出力処理を行う
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION
    Call sOutput
    
    Exit Sub

Err_Handler:
    
    gblnTomasDispErr = True
    frmTomasDataDisp.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdOutKikiInfo_Click
'//  機能名称  : 「機器状態データ媒体出力」釦押下時処理
'//  機能概要  : バージョン情報テキストファイルを媒体に出力する。
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
Private Sub cmdOutKikiInfo_Click()

    On Error GoTo Err_Handler
    
    '媒体出力処理を行う
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_KIKI
    Call sOutput
    
    Exit Sub

Err_Handler:
    
    gblnTomasDispErr = True
    frmTomasDataDisp.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdOutErrInfo_Click
'//  機能名称  : 「障害発生情報媒体出力」釦押下時処理
'//  機能概要  : バージョン情報テキストファイルを媒体に出力する。
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
Private Sub cmdOutErrInfo_Click()

    On Error GoTo Err_Handler
    
    '媒体出力処理を行う
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_ERR
    Call sOutput
    
    Exit Sub

Err_Handler:
    
    gblnTomasDispErr = True
    frmTomasDataDisp.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdDispVerInfo_Click
'//  機能名称  : 「バージョン情報表示」釦押下時処理
'//  機能概要  : バージョン情報表示を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sOutput()
    
    On Error Resume Next
    
    gstrOutPath = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    If gstrOutPath = "" Then
        Exit Sub  'ディレクトリが指定されなければ、処理終了
    End If
    
    gblnTomasDispErr = False
    frmTomasDataOut.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    On Error Resume Next
    
    '「TOMASデータ表示画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TOMAS_DATA_MNG_GAMEN_END, 0)
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : TOMASデータ管理画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    On Error Resume Next
    
    'メール受信用タイマを起動する
    tmrMail.Enabled = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : TOMASデータ管理画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'メール受信用タイマを止める
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : TOMASデータ管理画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '媒体出力釦は初期値非活性
    cmdOutVerInfo.Enabled = False
    cmdOutKikiInfo.Enabled = False
    cmdOutErrInfo.Enabled = False
    
    'シーケンス番号初期化
    gintSeqNo_Version = MIN_SEQ_VERSION
    gintSeqNo_KikiData = MIN_SEQ_KIKIDATA
    
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//     概要      : 「メール受信用タイマ」がタイムアップした時のイベントプロシージャ
'//     説明      : メール受信処理を行う。
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   CODED   BY [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
'        AppActivate frmTimeDataSettei.Caption, False   ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
        AppActivate frmTomasDataMng.Caption, False      ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
        pfFormActive (frmTomasDataMng.hwnd)             ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If

End Sub
