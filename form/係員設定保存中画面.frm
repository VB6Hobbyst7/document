VERSION 5.00
Begin VB.Form frmRenewSave 
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
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ｏ Ｋ"
      Enabled         =   0   'False
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
Attribute VB_Name = "frmRenewSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmRenewSave.frm
'//  パッケージ名：係員設定保存中画面
'//
'//  概要：係員設定保存中画面
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//                 ・収集・メンテデータ収集中画面(frmSyusyuCyu.frm)を流用
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////

Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 係員設定保存中画面(アクティブ時)
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
    
    '締切データ出力指示を集計へ送信する。
    If fSDATAMailSend = False Then
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
        cmdOK.Enabled = True
        Exit Sub
    End If
    
'    保存中のガイドを表示する｡
    lblMessage(0) = "設定値を保存中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 係員設定保存中画面(ディアクティブ時)
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
'//  機能名称  : 係員設定保存中画面(ロード時)
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
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_KAKARI_SAVE)
    
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
                AppActivate frmRenewSave.Caption, False
                pfFormActive (frmRenewSave.hwnd)        ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_KAKARIIN_SAVE_RES
                '「係員保存完了通知」を受信した場合
                '「係員保存完了通知受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KAKARIIN_SAVE_REQ_RECV, 0)
                'クリア通知内容をチェックする。
                If fReadMailCheck(udtReadMail) = True Then
                    lblMessage(0) = "正常終了しました。"
                    lblMessage(1) = ""
                Else
                    lblMessage(0) = "異常終了しました。"
                    lblMessage(1) = ""
                End If
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
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
'//  機能名称  : 係員設定保存指示送信処理
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

    Dim udtMail As MAIL_KAKARIIN_SETTEI
    Dim bRet As Boolean             '関数戻り値
    Dim lngErrCode As Long          'エラーコード
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCtlIndex As Integer
    Dim intDataIndex As Integer
    
    On Error Resume Next
 
    fSDATAMailSend = True
    
    '係員設定保存要求を送信する
    udtMail.mlHeader.dwId = ML_ID_KAKARIIN_SAVE_CMD
    udtMail.mlHeader.dwSize = MlSize.KAKARIIN_SETTEI
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    
    For intCount = 0 To UBound(glngTergetCorner)
        udtMail.dwCorner(intCount) = glngTergetCorner(intCount)
    Next intCount
    
    'メール送信
    bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMail), udtMail.mlHeader)
    
    If bRet = False Then
       '「係員設定保存復元画面：係員設定保存要求送信異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KAKARIIN_SAVE_REQ_SEND, lngErrCode)
       lblMessage(0) = "異常終了しました。"
       lblMessage(1) = ""
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
       fSDATAMailSend = False
       Exit Function
    Else
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KAKARIIN_SAVE_REQ_SEND, lngErrCode)
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : fReadMailCheck
'//  機能名称  : 係員設定保存完了通知メールチェック処理
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

    Dim i    As Integer      'カウンタ
    Dim iErr As Integer      '未収集号機の有無（1/0）
    Dim intIndex As Integer
    Dim intCount As Integer
    
    On Error Resume Next
    
    iErr = 0

    For intCount = 0 To UBound(glngTergetCorner)
        glngTergetCorner(intCount) = udtReadMail.lngData(intCount)
        'エラーありの場合は記録する
        If glngTergetCorner(intCount) > 0 Then
            iErr = 1
        End If
    Next intCount
    
    If iErr > 0 Then
        fReadMailCheck = False
        Exit Function
    End If
    
   fReadMailCheck = True
    
End Function
