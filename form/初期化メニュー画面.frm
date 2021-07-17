VERSION 5.00
Begin VB.Form frmSysformatMenu 
   BorderStyle     =   0  'なし
   Caption         =   "初期化"
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
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "設置時初期化"
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
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   5760
      Top             =   6600
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
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
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
      Index           =   2
      Left            =   2040
      TabIndex        =   3
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
      Index           =   1
      Left            =   6360
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "一括システム出荷時初期化"
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
      Caption         =   "システム初期化"
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
      TabIndex        =   5
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmSysformatMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmSysformatMenu.frm
'//  パッケージ名：システム初期化画面
'//
'//  概要：システム初期化画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  量産対応【設置時初期化機能追加】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 初期化メニュー画面(アクティブ時)
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
'//  機能名称  : 初期化メニュー画面(ディアクティブ時)
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
'//  機能名称  : 初期化メニュー画面(ロード時)
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

  '「システム初期化画面：表示」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_MENU_GAMEN_START, 0)

   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
   
   'IDU縮退チェック
   psIDUCheck
   
   If pbIDUSts = 1 Then
     'IDU業務非表示
      cmdFixedExe(2).Visible = False
   End If
   'V1.3.0.1 ADD START
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称の画面へ遷移する。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  量産対応【設置時初期化機能追加】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
  
  On Error Resume Next
    
    Select Case Index
        Case 0                                  '一括初期化
           '「システム初期化画面：一括初期化釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_MENU_GAMEN_ALLSYS_BUTTOM, 0)
           Load frmALLSysformat
           frmALLSysformat.Show 1
        Case 1                                 '監視盤初期化
           '「システム初期化画面：監視盤釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_MENU_GAMEN_KANSISYS_BUTTOM, 0)
           Load frmKansiSysformat
           frmKansiSysformat.Show 1
        Case 2                                  'ＩＤＵ初期化
           '「システム初期化画面：ＩＤＵ釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_MENU_GAMEN_IDUSYS_BUTTOM, 0)
           Load frmIDUSysformat
           frmIDUSysformat.Show 1
        Case 3                                  'ＬＤＵ初期化
           '「システム初期化画面：ＬＤＵ釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_MENU_GAMEN_LDUSYS_BUTTOM, 0)
           Load frmLDUSysformat
           frmLDUSysformat.Show 1
' EG20 V6.9.0.1 【量産対応：設置時初期化機能追加】ADD START
        Case 4                                  ' 設置時初期化
           '「システム初期化画面：設置時初期化釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_MENU_GAMEN_INSTALL_BUTTOM, 0)
           Load frmInstallformat
           frmInstallformat.Show 1
' EG20 V6.9.0.1 【量産対応：設置時初期化機能追加】ADD END
    End Select

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
  On Error Resume Next
   '「システム初期化画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_MENU_GAMEN_END, 0)
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
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmSysformatMenu.Caption, False
        pfFormActive (frmSysformatMenu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
