VERSION 5.00
Begin VB.Form frmJVerUpdate 
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Timer tmrMail 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '中央揃え
      Caption         =   "しばらくお待ち下さい。"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblMessage 
      Caption         =   "改札機用のバージョン情報を更新中です。"
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
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
End
Attribute VB_Name = "frmJVerUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmJVerUpData.frm
'//  パッケージ名：改札機バージョン更新中画面(EG-R自改/NEG自改用)
'//
'//  概要：改札機バージョン更新中画面(EG-R自改/NEG自改用)
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 改札機バージョン更新中画面(アクティブ時)
'//  機能概要  : メール受信用のタイマ起動
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
Private Sub Form_Activate()
    'メイル受信用のタイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 改札機バージョン更新中画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ起動
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
    'メイル受信用のタイマを止める。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 改札機バージョン更新中画面(ロード時)
'//  機能概要  : メール受信用のタイマ起動
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
    Dim udtMail As MAIL_GATE_VER_UPD_REQ  '自改バージョン情報更新要求メール送信エリア
    Dim lngRet As Long                    '関数戻り値
  
    On Error Resume Next

    '「改札機バージョン更新中」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_UPDATA, 0)

    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '自改バージョン情報更新要求メールを管理プロセスへ送信する。
    udtMail.mlHeader.dwId = ML_ID_GATE_VER_UPD_REQ
    udtMail.mlHeader.dwSize = MlSize.GATE_VER_UPD_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequest = gintVerJikai  '自改バージョン情報更新要求の種別
    lngRet = DssSendMail(MAIL_SLOT_KANRI, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       gintGateVerInfUpdRes = MailSts.stsErr
       Unload Me
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ処理
'//  機能概要  : メールを受信する。
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
Private Sub tmrMail_Timer()
    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    
    On Error Resume Next

    'メールを受信する。
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
   '受信メールがあれば、メールＩＤ毎の処理をする。
        Select Case udtReadMail.udtlHeader.dwId        'メールＩＤ
            Case ML_ID_PROEND_ORD
                 '「プロセス終了指示受信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                 'プロセスの終了処理を行う
                 pfAbortProc
                
            Case ML_ID_HOSHU_ACTIVE_REQ
                 '「保守画面アクティブ表示要求受信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                 '表示元画面（EG-R自改/NEG自改画面）をアクティブ表示する。
                 If gintVerJikai = ML_REQUEST_NGATE Then
                    gStrCurrentForm = sFormName_NJVer
                    AppActivate frmJVer.Caption, False
                    pfFormActive (frmJVerUpdate.hwnd)
                 Else
                    gStrCurrentForm = sFormName_EJVer
                    AppActivate frmJVer.Caption, False
                 End If
                
            Case ML_ID_GATE_VER_UPD_INF
                 '「自改ﾊﾞｰｼﾞｮﾝ情報更新通知受信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GATE_VERSIONINFO_UPDATA_REQ_RECV, 0)
                 '表示元画面（EG-R自改/NEG自改画面）をアクティブ表示する。
                 If gintVerJikai = ML_REQUEST_NGATE Then
                    gStrCurrentForm = sFormName_NJVer
                    AppActivate frmJVer.Caption, False
                    pfFormActive (frmJVerUpdate.hwnd)
                 'EG20 V30.1.0.1 ADD START
                 ElseIf gintVerJikai = ML_REQUEST_EG20GATE Then
                    gStrCurrentForm = sFormName_EG20JVer
                    AppActivate frmGateVerKanri.Caption, False
                    pfFormActive (frmGateVerKanri.hwnd)
                 ElseIf gintVerJikai = ML_REQUEST_EG30GATE Then
                    gStrCurrentForm = sFormName_EG30JVer
                    AppActivate frmKansenGateVerKanri.Caption, False
                    pfFormActive (frmKansenGateVerKanri.hwnd)
                'EG20 V30.1.0.1 ADD END
                 Else
                    gStrCurrentForm = sFormName_EJVer
                    AppActivate frmJVer.Caption, False
                 End If
                 gintGateVerInfUpdRes = udtReadMail.lngData(1)
                 
                 '本画面を終了する。
                 Unload Me
                
            Case Else
                '「メールID不正」ログ出力
                Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Sub
