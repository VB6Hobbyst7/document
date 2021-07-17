VERSION 5.00
Begin VB.Form frmTomasDataDisp 
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
   Begin VB.Timer tmrErrDisp 
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
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
      TabIndex        =   2
      Top             =   360
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
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
End
Attribute VB_Name = "frmTomasDataDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmTomasDataMng.frm
'//  パッケージ名：TOMASデータ表示画面
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
    
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TOMAS_DATA_DISP)

    If gintTomasDataDispDiv <> TOMAS_DISP_DIV.TOMAS_DATA_ERR Then
        'TOMASデータ出力要求を監マへ送信する。
        If fSDATAMailSend = False Then
            lblMessage(0) = "異常終了しました。"
            lblMessage(1) = ""
            cmdOK.Enabled = True
            gblnTomasDispErr = True
            gblnRecvErr = True
            Exit Sub
        End If
    End If
    
    '処理中のガイドを表示する｡
    lblMessage(0) = "処理中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
    tmrErrDisp.Enabled = True
    
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next
    
    'メール受信用タイマを止める
    tmrMail.Enabled = False
    
    tmrErrDisp.Enabled = False
    
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    tmrErrDisp.Interval = MN_MAIL_INTERVAL
    tmrErrDisp.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : fSDATAMailSend
'//  機能名称  : TOMASデータ出力要求送信処理
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
Private Function fSDATAMailSend() As Boolean

    Dim udtMailVer As VERSION_DATA_DISP_TYPE
    Dim udtMailKiki As KIKI_DATA_DISP_TYPE
    Dim bRet As Boolean             '関数戻り値
    Dim lngErrCode As Long          'エラーコード
    Dim strLog As String
    
    On Error Resume Next
 
    fSDATAMailSend = True
    
    'バージョン情報表示の場合
    If gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION Then
        'シーケンス番号
        If gintSeqNo_Version = MAX_SEQ_VERSION Then
            gintSeqNo_Version = MIN_SEQ_VERSION + 1
        Else
            gintSeqNo_Version = gintSeqNo_Version + 1
        End If
        
        udtMailVer.mlHeader.dwId = ML_ID_TOMAS_VARSION_DSP_REQ_CMD
        udtMailVer.mlHeader.dwSize = MlSize.VERSION_DATA_DSP_CMD
        udtMailVer.mlHeader.dwProid = RHOSHU_ID
        udtMailVer.mlHeader.dwSubArea = 0
        
        udtMailVer.dwSeqNo = gintSeqNo_Version              'シーケンス番号
        udtMailVer.dwBlockCheck = 0                         'ブロック番号チェック確認
        udtMailVer.dwDenbunSize = 6                         '電文サイズ（固定）
        udtMailVer.byCmd(0) = &H78                          'コマンドコード
        udtMailVer.byCmd(1) = &H41                          'サブコード
        udtMailVer.byCmd(2) = &H1                           'コーナNo
        udtMailVer.byCmd(3) = &H1                           '号機No
        udtMailVer.byCmd(4) = &H1                           'ブロックNo
        udtMailVer.byCmd(5) = &H1                           '最終ブロックNo
        strLog = TOMAS_DATA_VER_REQ_SEND
        
        'メール送信
        bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMailVer), udtMailVer.mlHeader)
    
    '機器状態データ表示の場合
    Else
        'シーケンス番号
        If gintSeqNo_KikiData = MAX_SEQ_KIKIDATA Then
            gintSeqNo_KikiData = MIN_SEQ_KIKIDATA + 1
        Else
            gintSeqNo_KikiData = gintSeqNo_KikiData + 1
        End If
        
        udtMailKiki.mlHeader.dwId = ML_ID_TOMAS_KIKIDATA_DSP_REQ_CMD
        udtMailKiki.mlHeader.dwSize = MlSize.KIKIINF_DATA_DSP_CMD
        udtMailKiki.mlHeader.dwProid = RHOSHU_ID
        udtMailKiki.mlHeader.dwSubArea = 0
    
        udtMailKiki.dwSeqNo = gintSeqNo_KikiData                'シーケンス番号
        udtMailKiki.dwDenbunSize = 6                            '電文サイズ（固定）
        udtMailKiki.byCmd(0) = &H79                             'コマンドコード
        udtMailKiki.byCmd(1) = &H41                             'サブコード
        udtMailKiki.byCmd(2) = &H1                              'コーナNo
        udtMailKiki.byCmd(3) = &H1                              '号機No
        udtMailKiki.byCmd(4) = &H1                              'ブロックNo
        udtMailKiki.byCmd(5) = &H1                              '最終ブロックNo
        strLog = TOMAS_DATA_KIKI_REQ_SEND
        
        'メール送信
        bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMailKiki), udtMailKiki.mlHeader)
    
    End If
    
    If bRet = False Then
        '「TOMASデータ表示画面：TOMASデータ出力要求送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, strLog, lngErrCode)
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        fSDATAMailSend = False
        Exit Function
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : fReadMailCheck
'//  機能名称  : TOMASデータ出力通知メールチェック処理
'//  機能概要  : メール受信時：メールを受信する。
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
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim blnRet As Boolean
    Dim lngErrCode As Long
    Dim intRetCd As Integer
    Dim strRet As String
    
    On Error Resume Next
    
    'リターンコードを取り出す
    strRet = Format(Hex(udtReadMail.lngData(3)), "00000000")
    intRetCd = CInt(Mid(strRet, 3, 2))
    
    '処理結果が異常の場合
    If intRetCd > 0 Then
        fReadMailCheck = False
        gblnRecvErr = True
        Exit Function
    End If
    Dim strMessage As String
    'バージョン取得の場合、完了通知を出す
    If gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION Then
        blnRet = dllCreateDispVerInfoFile(gintSeqNo_Version, lngErrCode)
    '機器状態要求の場合、受信データからファイルを作成する
    Else
        If sMakeDataFile(udtReadMail) = False Then
            fReadMailCheck = False
            Exit Function
        End If
        blnRet = dllCreateDispkikiStsFile(lngErrCode)
    End If
    
    '異常終了の場合
    If blnRet = False Then
        fReadMailCheck = False
        Exit Function
    End If
    
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    fReadMailCheck = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sMakeDataFile
'//  機能名称  : 機器状態データ作成
'//  機能概要  : 受信メールから機器状態データを作成する。
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
Private Function sMakeDataFile(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim lngHandle As Long
    Dim strKikiData As String
    Dim bRet As Boolean
    Dim lngRet As Long
    
    On Error Resume Next
    
    sMakeDataFile = True
    
    strKikiData = PATH_WORK & TOMAS_FILE_KIKIINFO_DAT
    
    '機器情報ファイルをオープン
    lngHandle = CreateFile(strKikiData, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           CREATE_ALWAYS, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため更新異常
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        sMakeDataFile = False
        Exit Function
    End If
         
    '機器情報ファイルに書き込む
    bRet = WriteFile(lngHandle, udtReadMail.lngData(4), udtReadMail.udtlHeader.dwSize - 32, lngRet, 0)
    If bRet = False Then
       'ハンドルのクローズ
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       Call CloseHandle(lngHandle)
       sMakeDataFile = False
       Exit Function
    End If
    
    'ハンドルのクローズ
     Call CloseHandle(lngHandle)

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : tmrErrDisp_Timer
'//  機能名称  : 障害情報データ表示処理
'//  機能概要  : 障害情報データを表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2012-02-08   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrErrDisp_Timer()

    Dim blnRet As Boolean
    Dim lngErrCode As Long
    
    On Error Resume Next
    
    If gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_ERR Then
    
        '障害発生情報表示処理
        blnRet = dllCreateDispErrFile(lngErrCode)
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        If blnRet = False Then
            lblMessage(0) = "異常終了しました。"
            lblMessage(1) = ""
            cmdOK.Enabled = True
            gblnTomasDispErr = True
            tmrErrDisp.Enabled = False
        Else
            Unload Me
        End If
        
    End If
    
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   CODED   BY [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF     'メール受信エリア
    Dim lngLength As Long                       '受信メールバイトサイズ
    Dim intStatus As Integer                    '受信メールチェック結果
    Dim strLog As String
    
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
'                AppActivate frmRenewCyu.Caption, False         ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
                AppActivate frmTomasDataDisp.Caption, False     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
                pfFormActive (frmTomasDataDisp.hwnd)            ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD

            Case ML_ID_TOMAS_VARSION_DSP_REQ_RES, ML_ID_TOMAS_KIKIDATA_DSP_REQ_RES
                '「バージョン取得（RES）」、「機器状態要求（RES）」を受信した場合
                '「TOMASデータ出力要求受信正常」ログ出力
                If udtReadMail.udtlHeader.dwId = ML_ID_TOMAS_VARSION_DSP_REQ_RES Then
                    strLog = TOMAS_DATA_VER_REQ_RECV
                Else
                    strLog = TOMAS_DATA_KIKI_REQ_RECV
                End If
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, strLog, 0)
                '内容をチェックする。
                If fReadMailCheck(udtReadMail) = True Then
                    Unload Me
                Else
                    lblMessage(0) = "異常終了しました。"
                    lblMessage(1) = ""
                    'プログレスバーを消去する
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                    gblnTomasDispErr = True
                    cmdOK.Enabled = True
                End If
            Case Else
                 'その他のメールを受信した場合
                 '「メールID不正」ログ出力
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
    
End Sub

