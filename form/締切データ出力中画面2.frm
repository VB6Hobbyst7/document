VERSION 5.00
Begin VB.Form frmShimekiriOutPut2 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   2715
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
   ScaleHeight     =   2715
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   240
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
      TabIndex        =   2
      Top             =   1200
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
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmShimekiriOutPut2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmShimekiriOutPut2.frm
'//  パッケージ名：締切データ出力中画面
'//
'//  概要：締切データ出力中画面
'//     ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
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
'//  機能名称  : 締切データ出力中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

' EG20 V7.3.0.1【EG20_KANSI03_01】DEL START
'    cmdOK.Enabled = False
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL START
    On Error Resume Next

' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
    Call gsGetCornerName
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
    
    '締切データ出力指示を集計へ送信する。
    If fSDATAMailSend = False Then
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
        lblMessage(2) = gstrCornerName(frmKVer.miCornerNo)
        frmKVer.mbMisouResult = False
        '自画面を消す。
        Unload Me
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL START
'        cmdOK.Enabled = True
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL START
        Exit Sub
      
    End If
    
'    収集中のガイドを表示する｡
    lblMessage(0) = "締切データを出力中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
    lblMessage(2) = gstrCornerName(frmKVer.miCornerNo)
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL START
'    cmdOK.Enabled = False
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL END
    tmrMail.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 締切データ出力中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'/
'//     ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
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
'//  機能名称  : 締切データ出力中画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim i As Integer 'カウンタ
    Dim intCount As Integer
    Dim intCount2 As Integer
    
    On Error Resume Next
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_SHIMEKIRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
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
'//     ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
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
'//     ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
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
                'プロセスの終了処理を行う
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示」を受信した場合
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '表示元画面（保守データ収集画面）をアクティブ表示する。
                AppActivate frmShimekiriOutPut2.Caption, False
                pfFormActive (frmShimekiriOutPut2.hwnd)     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_SHIMEKIRI_OUT_RES
                '「締切出力完了通知」を受信した場合
                '「締切出力完了通知受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, SHIMEKIRI_OUTPUT_REQ_RECV, 0)
                
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
                
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
                Call gsGetCornerName
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
                'クリア通知内容をチェックする。
                If fReadMailCheck(udtReadMail) = True Then
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
'                    frmShimekiriData.gbShimekiriResult = True
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
                    lblMessage(0) = "正常終了しました。"
                    lblMessage(1) = ""
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
                    lblMessage(2) = gstrCornerName(frmKVer.miCornerNo)
                    frmKVer.mbMisouResult = True
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
                Else
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
'                    frmShimekiriData.gbShimekiriResult = False
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
                    lblMessage(0) = "異常終了しました。"
                    lblMessage(1) = ""
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
                    lblMessage(2) = gstrCornerName(frmKVer.miCornerNo)
                    frmKVer.mbMisouResult = False
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
                End If
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL START
'                cmdOK.Enabled = True
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL END
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
                Unload Me   '自画面を消す。
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
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
'//  機能名称  : 締切データ出力指示送信処理
'//  機能概要  : 初期処理時：メールを送信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSDATAMailSend() As Boolean

    Dim udtMail As MAIL_HDATA_REQ  '保守データ収集指示メール送信エリア
    Dim lngRet As Long              '関数戻り値
    Dim lngErrCode As Long          'エラーコード
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCtlIndex As Integer
    Dim intDataIndex As Integer
    
    On Error Resume Next
 
    fSDATAMailSend = True
    
    '締切データ出力指示を集計に送信する。
    udtMail.mlHeader.dwId = ML_ID_SHIMEKIRI_OUT_CMD
    udtMail.mlHeader.dwSize = MlSize.SHIMEKIRI_OUTPUT_CMD
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequestType = ML_DT_W_SHIMEKIRI_H     '締切データ
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
'    udtMail.dwStatus(0) = frmShimekiriData.SSTab1.Tab + 1  ' コーナ   ' EG20 V6.3.0.1
    udtMail.dwStatus(0) = frmKVer.miCornerNo + 1           ' コーナ
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END
    
    lngRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       '「締切画面：締切データ出力指示送信異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, SHIMEKIRI_OUTPUT_REQ_SEND, lngErrCode)
       
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
       'プログレスバーを消去する
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       fSDATAMailSend = False
       Exit Function
    Else
       '「締切画面：締切データ出力指示送信正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, SHIMEKIRI_OUTPUT_REQ_SEND, 0)
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
'//     REVISIONS :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//                 ・締切データ出力中画面(frmSimekiriOutPut.frm)を流用
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim i    As Integer      'カウンタ
    Dim iErr As Integer      '未収集号機の有無（1/0）
    Dim intIndex As Integer
    On Error Resume Next
    
    iErr = 0
    If udtReadMail.lngData(0) <> ML_DT_W_SHIMEKIRI_H Then
        '指示したデータ種と異なる通知。
        fReadMailCheck = False
        Exit Function
    End If

    'ステータス、未収フラグチェック
    If udtReadMail.lngData(1) > 0 And iErr = 0 Then
        iErr = 1  'ステータスが正常ではない。
    ElseIf udtReadMail.lngData(2) > 0 And iErr = 0 Then
        iErr = 1  '未収あり
    End If
    
    If iErr > 0 Then
       'データ種が保守データ収集指示のものと異なる場合
        '指示したデータ種と異なる通知。
        fReadMailCheck = False
        Exit Function
    End If
    
   fReadMailCheck = True
    
End Function
