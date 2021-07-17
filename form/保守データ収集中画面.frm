VERSION 5.00
Begin VB.Form frmSyusyuCyu 
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmSyusyuCyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmSyusyuCyu.frm
'//  パッケージ名：保守データ収集中画面
'//
'//  概要：保守データ収集中画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・京王より、自改保守データ収集中画面(frmSyusyuCyu.frm)を流用
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'保守データINDEX
Private Const SYUSYU_KADO = 1    '稼働データ
Private Const SYUSYU_MENTE = 2   'メンテデータ
Private Const SYUSYU_ERRLOG = 3  'エラーログデータ
'Dim lngDataSyu(SYUSYU_KADO To SYUSYU_ERRLOG) As Long       '収集データのデータ種。 'EG20 V2.1.0.1 DEL
Dim intSyusyuIni(SYUSYU_KADO To SYUSYU_ERRLOG)  As Integer '収集要否(1/0)HosyuApl.INI定義値。
'Dim intSyusyuIndex  As Integer   '収集中の保守データINDEX       'EG20 V2.1.0.1 DEL
Dim lngGateSts(1 To MAX_GATE_NO) As Long                '号機毎収集状態
Dim iErrSts As Integer                                  '0:INI定義異常　2：メール送信異常　'V1.7.0.1 ADD
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 保守データ収集中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim intCount As Integer
    Dim blnSelected As Boolean
    'EG20 V2.1.0.1 ADD END
    
    cmdOK.Enabled = False
    
    On Error Resume Next
    
    'V1.7.0.1 ADD　START
    '初期化
    iErrSts = 0
    'V1.7.0.1 ADD　END
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    blnSelected = False
    For intCount = 0 To UBound(gintStatus)
        If gintStatus(intCount) = TAG_STATUS.STS_SENTAKU Then
            blnSelected = True
        End If
    Next
    
    '指定号機なしの場合、メッセージボックスを表示する
    If blnSelected = False Then
        lblMessage(0) = "指定号機が選択されていません。"
        lblMessage(1) = "選択してください。"
        cmdOK.Enabled = True
        Exit Sub
    End If
    
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_KBN_KADO_MAINTE)
    'EG20 V2.1.0.1 ADD END
    
    '収集中の保守データINDEXを初期化する。
'    intSyusyuIndex = 0         'EG20 V2.1.0.1 DEL

    '最初の保守データについて、保守データ収集指示を監マへ送信する。
    If fHDATAMailSend = False Then
        'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'      If iErrSts = 0 Then 'V1.7.0.1 ADD INI定義異常のみ下記処理を行う。
'    '何も送信しなかった（全データともに収集不要）ならば、
'        'メッセージボックスを表示し、
'        MsgBox "HosyuApl.iniに、収集すべきデータが定義されていません。"
'        '「稼動・メンテデータ収集画面：収集データ未定義」ログ出力
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KADO_MENTE_SYUSYU_GAMEN_DATA_SETTEI_ERROR, 0)
'        '自画面を消す。
'        Unload Me
'        Exit Sub
'      'V1.7.0.1 ADD START
'      Else
'         '自画面を消す。
'         Unload Me
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
        cmdOK.Enabled = True
        'EG20 V2.1.0.1 ADD END
         Exit Sub
'      End If           'EG20 V2.1.0.1 DEL
      'V1.7.0.1 ADD END
    End If
    
'    収集中のガイドを表示する｡
    lblMessage(0) = "保守データを収集中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 保守データ収集中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
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
Private Sub Form_Deactivate()
On Error Resume Next
    'メール受信用タイマを止める
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 保守データ収集中画面(ロード時)
'//  機能概要  : 初期処理を行う。
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
Private Sub Form_Load()
    Dim strSyusyuKey(SYUSYU_KADO To SYUSYU_ERRLOG) As String 'HosyuApl.INIキー値
    Dim i As Integer 'カウンタ
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim intCount As Integer
    Dim intCount2 As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next
    
    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
    '収集データのデータ種をセットしておく。
'    lngDataSyu(SYUSYU_KADO) = ML_DT_W_KADO_H      '稼働データ
'    lngDataSyu(SYUSYU_MENTE) = ML_DT_W_MENTE_H    'メンテデータ
'    lngDataSyu(SYUSYU_ERRLOG) = ML_DT_W_ERRLOG_H  'エラーログデータ
    
    ' HosyuApl.iniから「保守データ収集」定義内容(収集要否)を取出す。
'    strSyusyuKey(SYUSYU_KADO) = PROFILE_KEY_NAME_HDATA_KADO     '稼働データ キー
'    strSyusyuKey(SYUSYU_MENTE) = PROFILE_KEY_NAME_HDATA_MENTE   'メンテデータ キー
'    strSyusyuKey(SYUSYU_ERRLOG) = PROFILE_KEY_NAME_HDATA_ERRLOG 'エラーログデータ キー
    'EG20 V2.1.0.1 DEL END
    For i = SYUSYU_KADO To SYUSYU_ERRLOG
        intSyusyuIni(i) = GetPrivateProfileInt(PROFILE_KEY_HOSHU_DATA, _
                                               strSyusyuKey(i), DEFAILT_Int, HOSHUAPL_FILE)
    Next
    
    For i = 1 To MAX_GATE_NO
    '号機毎収集状態を、正常終了で初期化する。
        lngGateSts(i) = ML_DT_SEIJO_SHUSHU
    Next
        
    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
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
                'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'EG20 V2.1.0.1 ADD END
                 'プロセスの終了処理を行う
                 pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                 '「保守画面アクティブ表示」を受信した場合
                 '「保守画面アクティブ表示要求受信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                 '表示元画面（保守データ収集画面）をアクティブ表示する。
                 AppActivate frmSyusyuCyu.Caption, False
                 pfFormActive (frmSyusyuCyu.hwnd)           ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_HDATA_ANS
                 '「締切開始要求RES」を受信した場合
                 '「保守データ収集通知受信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_DATA_SYUSYU_REQ_RECV, 0)
                'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'EG20 V2.1.0.1 ADD END
                '収集通知内容をチェックする。
                If fReadMailCheck(udtReadMail) = True Then
                '収集指示に対する収集通知であれば、
                    '次のデータ種の収集指示を監マへ送信する。
'                    If fHDATAMailSend = False Then      'EG20 V2.1.0.1 DEL 【Mainte_03_01】
                       '何も送信しなかった（全データともに収集済）ならば、
                       '収集終了状態を表示する。
                       sSyusyuEnd
'                    End If                              'EG20 V2.1.0.1 DEL 【Mainte_03_01】
                'EG20 V2.1.0.1 ADD START【Mainte_03_01】
                Else
                    lblMessage(0) = "異常終了しました。"
                    lblMessage(1) = ""
                    cmdOK.Enabled = True
                    'プログレスバーを消去する
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'EG20 V2.1.0.1 ADD END
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
'//  関数名称  : fHDATAMailSend
'//  機能名称  : 保守データ収集指示送信処理
'//  機能概要  : 初期処理時：メールを送信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fHDATAMailSend() As Boolean
    Dim udtMail As MAIL_HDATA_REQ  '保守データ収集指示メール送信エリア
    Dim lngRet As Long             '関数戻り値
    Dim lngErrCode As Long         'エラーコード
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCtlIndex As Integer
    Dim intDataIndex As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'    '今回収集指示するデータ種を確定する。
'    For intSyusyuIndex = intSyusyuIndex + 1 To SYUSYU_ERRLOG
'        If intSyusyuIni(intSyusyuIndex) = 1 Then '収集要のデータ種を探す。
'            Exit For
'        End If
'    Next
'
'    '全ての収集データに指示済であれば、
'    If intSyusyuIndex > SYUSYU_ERRLOG Then
'       '全ての収集データに指示済であれば、
'       '｢処理終了｣で戻る。
'        fHDATAMailSend = False
'        iErrSts = 0             'V1.7.0.1 ADD
'        Exit Function
'    End If
    'EG20 V2.1.0.1 DEL END
 
    '保守データ収集指示を集計に送信する。
    udtMail.mlHeader.dwId = ML_ID_HDATA_REQ
    udtMail.mlHeader.dwSize = MlSize.HOSHU_SYUSYU_CMD
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
'    udtMail.dwRequestType = lngDataSyu(intSyusyuIndex) '該当データのデータ種        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
    udtMail.dwRequestType = ML_DT_W_KADO_MAINTE_H       '収集・メンテデータ         'EG20 V2.1.0.1 ADD 【Mainte_03_01】
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    '号機別収集ステータスを設定する
    For intCount = 0 To 31
        If gintStatus(intCount) = TAG_STATUS.STS_SENTAKU Then
            udtMail.dwStatus(intCount) = TAG_STATUS.STS_SENTAKU
        Else
            udtMail.dwStatus(intCount) = TAG_STATUS.STS_MISENTAKU
        End If
    Next intCount
    'EG20 V2.1.0.1 ADD END
    
    lngRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
        '「稼動・メンテデータ収集画面：保守データ収集指示送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_DATA_SYUSYU_CMD_SEND, lngErrCode)
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
        iErrSts = 2                   'V1.7.0.1 ADD
        'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        'EG20 V2.1.0.1 ADD END
        Exit Function
    Else
       '「稼動・メンテデータ収集画面：保守データ収集指示送信正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_DATA_SYUSYU_CMD_SEND, 0)
    End If
        
    '｢送信中｣で戻る。
    fHDATAMailSend = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fReadMailCheck
'//  機能名称  : 保守データ収集通知メールチェック処理
'//  機能概要  : メール受信時：メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim iEnd As Integer      '
    Dim i    As Integer      'カウンタ
    Dim iErr As Integer      '未収集号機の有無（1/0）
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim intIndex As Integer
    'EG20 V2.1.0.1 ADD END
    On Error Resume Next
    
    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'    iEnd = 0
'    If intSyusyuIndex > SYUSYU_ERRLOG Then
'        iEnd = 1  '保守データ収集は既に終了している。
'    ElseIf udtReadMail.lngData(0) <> lngDataSyu(intSyusyuIndex) Then
    'EG20 V2.1.0.1 DEL END
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    If udtReadMail.lngData(0) <> ML_DT_W_KADO_MAINTE_H Then
        iEnd = 2  '指示したデータ種と異なる通知。
    End If
    'EG20 V2.1.0.1 ADD END

'    If iEnd = 1 Then
'       'データ種が保守データ収集指示のものと異なる場合、
'       '収集通知異常のログ出力を依頼する。
'       sLogRequest 2, udtReadMail
'       '｢指示に対する通知ではない｣として、戻る。
'       fReadMailCheck = False
'       Exit Function
'   End If
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    'ステータス、未収フラグチェック
    If udtReadMail.lngData(1) > 0 And iEnd = 0 Then
        iEnd = 1
    ElseIf udtReadMail.lngData(2) > 0 And iEnd = 0 Then
        iEnd = 1
    End If
    
    If iEnd = 2 Then
       'データ種が保守データ収集指示のものと異なる場合、
       '収集通知異常のログ出力を依頼する。
       sLogRequest iErr, udtReadMail
       '｢指示に対する通知ではない｣として、戻る。
       fReadMailCheck = False
       Exit Function
    End If
    'EG20 V2.1.0.1 ADD END
    
  
   '今回の収集状態を、号機毎収集状態にメモする。
   iErr = 0       '未収集号機 無し、としておく。
   'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'   For i = 1 To MAX_GATE_NO
'       If udtReadMail.lngData(i) = ML_DT_MISHUSHU Then
'          '｢未収集｣であれば、メモする。
'          lngGateSts(i) = ML_DT_MISHUSHU
'          iErr = 1
'       ElseIf udtReadMail.lngData(i) = ML_DT_GOUKI_NASI Then
'              '｢号機なし｣であれば、メモする。
'              lngGateSts(i) = ML_DT_GOUKI_NASI
'       End If
'    Next
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    For i = 3 To MAX_GATE_NO + 2
        intIndex = i - 3
        If gintStatus(intIndex) <> TAG_STATUS.STS_MISENTAKU Then
            Select Case udtReadMail.lngData(i)
            Case ML_DT_MISHUSHU, ML_DT_IJO_SHUSHU
                '｢未収集｣、「異常終了」であれば、メモする。
                lngGateSts(intIndex + 1) = udtReadMail.lngData(i)
                iErr = 1
                gintStatus(intIndex) = TAG_STATUS.STS_MISHUSHU
            Case ML_DT_GOUKI_NASI
                '｢号機なし｣であれば、メモする。
                lngGateSts(intIndex + 1) = ML_DT_GOUKI_NASI
                gintStatus(intIndex) = TAG_STATUS.STS_MISENTAKU
            Case ML_DT_SEIJO_SHUSHU
                '「正常終了」
                gintStatus(intIndex) = TAG_STATUS.STS_SHUSHU
            End Select
        End If
    Next
    'EG20 V2.1.0.1 ADD END
       
    '収集チェックを行う。
    sLogRequest iErr, udtReadMail
    
    If iEnd = 0 Then
        fReadMailCheck = True
    Else
        fReadMailCheck = False
    End If
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sLogRequest
'//  機能名称  : 収集結果チェック処理
'//  機能概要  : メール受信時：収集結果チェックを行う。
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
Private Sub sLogRequest(iErr As Integer, udtReadMail As ML_KYOTU_INF)
    Dim strDataSyu  As String    'データ種の表示文字列
    Dim strErrorLog As String    'ログトレース依頼ﾊﾟﾗﾒｰﾀ
    
    On Error Resume Next

    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
    '収集中のデータ種の表示文字列をセットする。
'    If intSyusyuIndex = SYUSYU_KADO Then
'        strDataSyu = "稼働データ収集"
'    ElseIf intSyusyuIndex = SYUSYU_MENTE Then
'        strDataSyu = "メンテデータ収集"
'    Else
'        strDataSyu = "エラーログ収集"
'    End If
    'EG20 V2.1.0.1 DEL END
    
    If iErr = 0 Then
     '該当データ種、正常終了の場合、
        strErrorLog = fLogStatusGet(udtReadMail)
    ElseIf iErr = 1 Then
    '該当データ種、未収集ありの場合、
        strErrorLog = fLogStatusGet(udtReadMail)
    ElseIf iErr = 2 Then
    '通知メール異常の場合、
        strErrorLog = Format$(udtReadMail.lngData(0), "000000")
    End If
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    'ログ出力を依頼する。
    Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_DATA_SYUSYU_REQ_RECV & ":" & strErrorLog, 0)
    'EG20 V2.1.0.1 ADD END
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fLogStatusGet
'//  機能名称  : 収集結果号機別状態文字列編集処理
'//  機能概要  : メール受信時：号機別表示文字を編集する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fLogStatusGet(udtReadMail As ML_KYOTU_INF)
    Dim i       As Integer     'カウンタ
    Dim strWork As String      '編集文字列

    On Error Resume Next
    
    strWork = ""
    '全号機について調べる。
    For i = 1 To MAX_GATE_NO
        If udtReadMail.lngData(i) <> ML_DT_GOUKI_NASI Then
           '未実装の号機に付いては、編集しない。
           '号機番号を書込む。
           strWork = strWork & "No" & Format$(i, "00")
           If udtReadMail.lngData(i) = ML_DT_SEIJO_SHUSHU Then
            '｢正常終了｣であれば、OKを書込む。
               strWork = strWork & "=OK,"
           Else
            '｢未収集｣であれば、NGを書込む。
               strWork = strWork & "=NG,"
           End If
        End If
    Next
    fLogStatusGet = strWork
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSyusyuEnd
'//  機能名称  : 収集終了状態表示処理
'//  機能概要  : メール受信時：保守データ収集結果文言を表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSyusyuEnd()
    Dim i As Integer       'カウンタ
    Dim iEnd As Integer    '終了状態
    Dim lngErrCode As Long 'エラーコード

    On Error Resume Next

    iEnd = 0
    For i = 0 To MAX_GATE_NO - 1
        '未収集の号機があったならば、
        If gintStatus(i) = TAG_STATUS.STS_MISHUSHU Then
           '保守データ収集は、収集失敗とする。
           iEnd = i
           Exit For
        End If
    Next
    If iEnd = 0 Then
       '正常終了時の文言を表示する。
       lblMessage(0) = "正常終了しました。"
       lblMessage(1) = ""
       '「保守データ収集処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_SYUSYU_OK, 0)
    Else
       '収集失敗時の文言を表示する。
'       lblMessage(0) = "収集失敗。" & str(iEnd) & "号機が未収集です。"      'EG20 V2.1.0.1 DEL 【Mainte_03_01】
       lblMessage(0) = "収集失敗。未収集号機があります。"                    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
     '  lblMessage(1) = "-- 内訳はﾛｸﾞﾄﾚｰｽ(監視盤)参照。 --" 'V1.8.0.1 DEL
'        lblMessage(1) = "-- 内訳は監視盤ログ管理参照。 --" 'V1.8.0.1 ADD   'EG20 V2.1.0.1 DEL 【Mainte_03_01】
        lblMessage(1) = "指定号機選択表示で確認してください。"              'EG20 V2.1.0.1 ADD 【Mainte_03_01】
       '「保守データ収集処理異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_SYUSYU_ERROR, lngErrCode)
    End If
    cmdOK.Enabled = True
End Sub
