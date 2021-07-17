VERSION 5.00
Begin VB.Form frmSystemSetteiMenu 
   BorderStyle     =   0  'なし
   Caption         =   "リモートメンテナンス"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer tmrMail 
      Left            =   6600
      Top             =   6240
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "一定期間情報設定"
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
      Index           =   5
      Left            =   6360
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "IDサーバ始終業設定"
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
      Index           =   4
      Left            =   2040
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "状態監視機能設定"
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
      Index           =   3
      Left            =   6360
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "時間帯別データ設定"
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
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "システム日付設定"
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
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "リブート設定"
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
      Index           =   1
      Left            =   6360
      TabIndex        =   2
      Top             =   960
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
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "システム設定"
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
End
Attribute VB_Name = "frmSystemSetteiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmSystemSetteiMenu.frm
'//  パッケージ名：システム設定メニュー画面
'//  概要        ：ログ管理メニュー画面
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Option Explicit



Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：Form_Activate
'//  機能名称    ：システム設定メニュー画面(アクティブ時)
'//  機能概要    ：画面再表示処理を行う。
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：Form_Deactivate
'//  機能名称    ：システム設定メニュー画面(ディアクティブ時)
'//  機能概要    ：メール受信用のタイマ停止
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：Form_Load
'//  機能名称    ：Form_Load時処理
'//  機能概要    ：Form_Load時処理を行う。
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    '「システム設定メニュー画面 表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    ' /////////////////////////////////////////////////////
    ' IDU縮退チェック
    psIDUCheck

    If pbIDUSts = 1 Then
       cmdFixedExe(3).Visible = False       ' 状態監視機能設定
       cmdFixedExe(4).Visible = False       ' IDサーバ始終業設定
       cmdFixedExe(5).Visible = False       ' 一定期間情報設定釦
    End If
   
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdFixedExe_Click
'//  機能名称    ：各釦押下処理
'//  機能概要    ：釦押下により画面遷移する。
'//
'//                 型          名称            意味
'//  引数        ： Integer     Index           押下釦インデックス値
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
  
   Dim udtMail As ML_DISP_INF          '画面表示要求
   Dim iResponse As Integer            'メッセージボックス戻り値

   On Error Resume Next
  
  
    Select Case Index
        Case 0                                 ' システム日付設定
            '「システム設定メニュー画面：システム日付設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_SYSDATE_BUTTOM, 0)
            Load frmSystemDateSettei
            frmSystemDateSettei.Show 1
        Case 1                                 ' リブート時刻設定
            '「システム設定メニュー画面：リブート時刻設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_REBOOT_BUTTOM, 0)
            Load frmRebootTimeSettei
            frmRebootTimeSettei.Show 1
        Case 2                                 ' 時間帯別データ設定
            '「システム設定メニュー画面：時間帯別データ設定設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_TIMEDATA_BUTTOM, 0)
            Load frmTimeDataSettei
            frmTimeDataSettei.Show 1
        Case 3                                 ' 状態監視機能設定
            '「システム設定メニュー画面：状態監視機能設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_IDCONDITION_BUTTOM, 0)
        
            '画面表示要求（状態監視機能設定）をID制御に送信する
            If (SendMessageDispInfo(ML_DT_IDU_SYSTEM_CONDITION) = False) Then
         
                iResponse = MsgBox("状態監視機能設定釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "状態監視機能画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
            End If
        
        Case 4                                 ' IDサーバ始終業設定
            '「システム設定メニュー画面：IDサーバ始終業設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_IDJOBTIME_BUTTOM, 0)
        
            '画面表示要求（IDサーバ始終業設定）をID制御に送信する
            If (SendMessageDispInfo(ML_DT_IDU_SYSTEM_JOBTIME) = False) Then
         
                iResponse = MsgBox("IDサーバ始終業設定釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "IDサーバ始終業設定画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
            End If
        Case 5                                 ' 一定期間情報配信設定
            '「システム設定メニュー画面：一定期間情報配信設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_IDPERIOD_BUTTOM, 0)
            
            '画面表示要求（一定期間情報配信設定）をID制御に送信する
            If (SendMessageDispInfo(ML_DT_IDU_SYSTEM_PERIOD) = False) Then
         
                iResponse = MsgBox("一定期間情報配信設定釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "一定期間情報配信設定画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
            End If
    End Select

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdReturn_Click
'//  機能名称    ：戻る釦押下処理
'//  機能概要    ：戻る釦押下時処理処理を行う。
'//
'//                 型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '「システム設定メニュー画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：SendMessageDispInfo
'//  機能名称    ：画面表示状態通知
'//  機能概要    ：画面表示状態通知を行う。
'//
'//                 型      名称                意味
'//  引数         : Long    lDispInfo           画面要求種別
'//
'//  戻り値       : TRUE    （正常）
'//                 FALSE   （異常）
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考         :
'///////////////////////////////////////////////////////////////////////////////////////////////
Private Function SendMessageDispInfo(ByVal lDispInfo As Long) As Boolean

    Dim udtMail As ML_DISP_INF          '画面表示要求
    Dim bRet As Boolean                 'メール送信処理戻り値
    Dim lngErrCode As Long              'エラーコード
    
    '画面表示要求をID制御に送信する
    udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
    udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
    udtMail.udtlHeader.dwProid = RHOSHU_ID
    udtMail.udtlHeader.dwSubArea = 0
    udtMail.dwDisp_Type = lDispInfo
    bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
    If bRet = False Then
        '「画面表示要求メール送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
    Else
   
        '「画面表示要求メール送信正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
    End If
    
    SendMessageDispInfo = bRet

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信タイマ、タイムアップ処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmSystemSetteiMenu.Caption, False
        pfFormActive (frmSystemSetteiMenu.hwnd)
    End If
End Sub
