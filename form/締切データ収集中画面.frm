VERSION 5.00
Begin VB.Form frmShimekiriCyu 
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
Attribute VB_Name = "frmShimekiriCyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmShimekiriCyu.frm
'//  パッケージ名：締切データ収集中画面
'//
'//  概要：締切データ収集中画面
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//                 ・収集・メンテデータ収集中画面(frmSyusyuCyu.frm)を流用
'//                 ・フェーズ２対応【Mainte_05_03】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////

Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値
Private lngGateSts(1 To MAX_GATE_NO) As Long                '号機毎収集状態


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 締切データ収集中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    cmdOK.Enabled = False
    
    On Error Resume Next
    
    '締切データ収集指示を集計へ送信する。
    If fSDATAMailSend = False Then
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
        cmdOK.Enabled = True
        Exit Sub
      
    End If
    
'    収集中のガイドを表示する｡
    lblMessage(0) = "締切データを収集中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 締切データ収集中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
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
'//  機能名称  : 締切データ収集中画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim i As Integer 'カウンタ
    Dim intCount As Integer
    Dim intCount2 As Integer
    
    On Error Resume Next
    
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_SHIMEKIRI)
    
    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()

    On Error Resume Next
    
    '自画面を消す。
    Unload Me
    
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
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
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'プロセスの終了処理を行う
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示」を受信した場合
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '表示元画面（保守データ収集画面）をアクティブ表示する。
                AppActivate frmShimekiriCyu.Caption, False
                pfFormActive (frmShimekiriCyu.hwnd)         ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_HDATA_ANS
                '「締切完了通知」を受信した場合
                '「締切完了通知受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, SHIMEKIRI_SHUSHU_REQ_RECV, 0)
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                If fReadMailCheck(udtReadMail) = True Then
                    lblMessage(0) = "正常終了しました。"
                    lblMessage(1) = ""
                End If
                cmdOK.Enabled = True
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
'//  関数名称  : fSDATAMailSend
'//  機能名称  : 締切データ収集指示送信処理
'//  機能概要  : 初期処理時：メールを送信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSDATAMailSend() As Boolean

    Dim udtMail As MAIL_HDATA_REQ  '保守データ収集指示メール送信エリア
    Dim lngRet As Long              '関数戻り値
    Dim lngErrCode As Long          'エラーコード
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCtlIndex As Integer
    
    On Error Resume Next
 
    fSDATAMailSend = True
    
    '締切データ収集指示を集計に送信する。
    udtMail.mlHeader.dwId = ML_ID_HDATA_REQ
    udtMail.mlHeader.dwSize = MlSize.HOSHU_SYUSYU_CMD
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequestType = ML_DT_W_SHIMEKIRI_H     '締切データ
    
    For intCount = 0 To 31
        If gintShimekiri(intCount) = TAG_STATUS.STS_SENTAKU Then
            udtMail.dwStatus(intCount) = TAG_STATUS.STS_SENTAKU
        Else
            udtMail.dwStatus(intCount) = TAG_STATUS.STS_MISENTAKU
        End If
    Next intCount
    
    lngRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
        '「締切画面：締切データ収集指示送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, SHIMEKIRI_SHUSHU_REQ_SEND, lngErrCode)
        fSDATAMailSend = False
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        Exit Function
    Else
       '「締切画面：締切データ収集指示送信正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, SHIMEKIRI_SHUSHU_REQ_SEND, 0)
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : fReadMailCheck
'//  機能名称  : 締切データ完了通知メールチェック処理
'//  機能概要  : メール受信時：メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim iEnd As Integer      '
    Dim i    As Integer      'カウンタ
    Dim iErr As Integer      '未収集号機の有無（1/0）
    Dim intIndex As Integer
    On Error Resume Next
    
    fReadMailCheck = True
    
    If udtReadMail.lngData(0) <> ML_DT_W_SHIMEKIRI_H Then
       '｢指示に対する通知ではない｣として、戻る。
        fReadMailCheck = False
        'クリア通知内容をチェックする。
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
        Exit Function
    End If

    'ステータス、未収フラグチェック
    If udtReadMail.lngData(1) > 0 And iErr = 0 Then
        iErr = 1  'ステータスが正常ではない。
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
        fReadMailCheck = False
        Exit Function
    ElseIf udtReadMail.lngData(2) > 0 And iErr = 0 Then
        iErr = 1  '未収あり
        lblMessage(0) = "未送の締切データが統合監視盤内にあります。"
        lblMessage(1) = "締切データの収集処理を開始できません。"
        fReadMailCheck = False
        Exit Function
    End If
  
   '今回の収集状態を、号機毎収集状態にメモする。
   iErr = 0       '未収集号機 無し、としておく。
   
    For i = 3 To MAX_GATE_NO + 2
        intIndex = i - 3
        If gintShimekiri(intIndex) <> TAG_STATUS.STS_MISENTAKU Then
            Select Case udtReadMail.lngData(i)
            Case ML_DT_MISHUSHU, ML_DT_IJO_SHUSHU
                '｢未収集｣、「異常終了」であれば、メモする。
                lngGateSts(intIndex) = ML_DT_MISHUSHU
                gintShimekiri(intIndex) = TAG_STATUS.STS_MISHUSHU
                If iErr < 2 Then
                    iErr = 1
                End If
            Case ML_DT_GOUKI_NASI
                '｢号機なし｣であれば、メモする。
                lngGateSts(intIndex) = ML_DT_GOUKI_NASI
                '送信時に対象としていた号機が対象外で返ってきた場合
                If gintShimekiri(intIndex) <> TAG_STATUS.STS_MISHUSHU Then
                    '通常ありえないので異常終了扱いにする。
                    iErr = 2  'ステータスが正常ではない。
                End If
                gintShimekiri(intIndex) = TAG_STATUS.STS_MISENTAKU
            Case ML_DT_SEIJO_SHUSHU
                '「正常終了」
                gintShimekiri(intIndex) = TAG_STATUS.STS_SHUSHU
            End Select
        End If
    Next
       
    If iErr = 1 Then
        lblMessage(0) = "収集失敗。未収集号機があります。"
        lblMessage(1) = "--内訳は画面参照。--"
        fReadMailCheck = False
    ElseIf iErr = 2 Then
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
        fReadMailCheck = False
    End If
    
End Function
