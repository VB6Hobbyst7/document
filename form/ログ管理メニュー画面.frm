VERSION 5.00
Begin VB.Form frmLogMenu 
   BorderStyle     =   0  'なし
   Caption         =   "ログ管理"
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "操作卓"
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
   Begin VB.Timer tmrMail 
      Left            =   6840
      Top             =   6360
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＬＤＵ"
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
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＩＣＭ"
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
      TabIndex        =   4
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＩＤＵ"
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
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "改札機"
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
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "統合監視盤"
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
      Caption         =   "ログ管理"
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
      TabIndex        =   6
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmLogMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmLogMenu.frm
'//  パッケージ名：ログ管理メニュー画面
'//
'//  概要：ログ管理メニュー画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　ＲＹＴログ管理画面追加
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ログ管理メニュー画面(アクティブ時)
'//  機能概要  : 画面再表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : ログ管理メニュー画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : ログ管理メニュー画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    '「ログ管理メニュー画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'IDU縮退チェック
    psIDUCheck

    If pbIDUSts = 1 Then
      'IDU業務非表示
       cmdFixedExe(1).Visible = False
       cmdFixedExe(4).Visible = False
    End If
   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 釦押下により画面遷移する
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　押下釦インデックス値
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-23  CODED BY  [TCC] M.Matsumoto
'//                 EG20フェーズ２対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
   Dim udtMail As ML_DISP_INF          '画面表示要求
   Dim iResponse As Integer            'メッセージボックス戻り値
   Dim bRet As Boolean                 'メール送信処理戻り値
   Dim lngErrCode As Long              'エラーコード

   On Error Resume Next

 Select Case Index
        Case 0                                 'ログ管理(監視盤)
           '「ログ管理メニュー画面：監視盤釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_KANSI_BUTTOM, 0)
           Load frmKansiLogKanri
           frmKansiLogKanri.Show 1
        Case 1                                 'ログ管理(IDU)
           '「ログ管理メニュー画面：ＩＤＵ釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_IDU_BUTTOM, 0)
           Load frmIDULogkanri
           frmIDULogkanri.Show 1
        Case 2                                 'ログ管理(LDU)
           '「ログ管理メニュー画面：ＬＤＵ釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_LDU_BUTTOM, 0)
           Load frmLDULogkanri
           frmLDULogkanri.Show 1
        Case 3                                 'ログ管理(改札機)
          '「ログ管理メニュー画面：改札機釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_KAISATUKI_BUTTOM, 0)

           '画面表示要求(改札機)をLD制御に送信する
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_KAISATUKI_LOG
            bRet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
              '「ログ管理メニュー画面：画面表示要求メール送信異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
                iResponse = MsgBox("改札機釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "改札機ログ管理画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
                Exit Sub
            End If
              '「ログ管理メニュー画面：画面表示要求メール送信正常」ログ出力
               Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        Case 4                                 'ログ管理(判定IC-M)
          '「ログ管理メニュー画面：判定ＩＣ−Ｍ釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_ICM_BUTTOM, 0)
            
            '画面表示要求(判定IC-M)をID制御に送信する
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_HANTEI_LOG
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
                '「ログ管理メニュー画面：画面表示要求メール送信異常」ログ出力
                lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                 Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
         
                iResponse = MsgBox("判定IC-M釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "判定IC-Mログ管理画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
                Exit Sub
            End If
            '「ログ管理メニュー画面：画面表示要求メール送信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'    'V1.6.0.1 ADD START
'      Case 5                                 'ログ管理(RYT)
'           '「ログ管理メニュー画面：ＲＹＴ釦押下」ログ出力
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_RYT_BUTTOM, 0)
'           Load frmRYTSyusyu
'           frmRYTSyusyu.Show 1
'    'V1.6.0.1 ADD END
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
      Case 5                                 'ログ管理（操作卓）
           '「ログ管理メニュー画面：操作卓釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_TAKU_BUTTOM, 0)
           'EG20 V2.1.0.1 DEL START 【フェーズ２対応】
'           Load frmRYTSyusyu
'           frmRYTSyusyu.Show 1
           'EG20 V2.1.0.1 DEL END
           'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
           Load frmTakuLogKanri
           frmTakuLogKanri.Show 1
           'EG20 V2.1.0.1 ADD START
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      :なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
   On Error Resume Next
   
   '「ログ管理メニュー画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'V1.3.0.1 ADD START
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmLogMenu.Caption, False
        pfFormActive (frmLogMenu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
