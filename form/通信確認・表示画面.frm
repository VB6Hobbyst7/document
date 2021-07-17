VERSION 5.00
Begin VB.Form frmTusinMenu 
   BorderStyle     =   0  'なし
   Caption         =   "通信確認・表示"
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
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "折り返しテスト"
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
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5640
      Top             =   6240
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "上位通信状態確認"
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
      Caption         =   "通信接続・切断"
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
      Caption         =   "通信確認"
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
      Caption         =   "通信確認・表示"
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
      TabIndex        =   4
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmTusinMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmTusinMenu.frm
'//  パッケージ名：通信確認・表示画面
'//
'//  概要：通信確認・表示画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 通信確認・表示画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
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
Private Sub Form_Activate()
On Error Resume Next
    'メール受信用タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 通信確認・表示画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
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
    'メール受信用タイマを止める
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 通信確認・表示画面(ロード時)
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
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
   Dim sGetIniData As String * ORIKAESI_TYPE    '折り返しテスト有無取得用　'V1.10.0.1 ADD
   Dim lSts As Long                 'V1.10.0.1 ADD
   
    On Error Resume Next

    '「通信確認・表示画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_MENU_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'V1.3.0.1 ADD START
    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    '1.3.0.1 ADD END
'V1.10.0.1 ADD START

    '折り返しテスト釦とりあえず非表示
    cmdFixedExe(3).Visible = False
    
    lSts = False

    'HOSHU.INIから折り返しテスト有無を取得
    lSts = GetPrivateProfileString(ORI_TEST_SEC, _
                                   ORI_TEST_KEY, _
                                   DEFAILT, _
                                   sGetIniData, _
                                   Len(sGetIniData), _
                                   HOSHU_FILE)
    If lSts = False Then
        '読み込み失敗は処理終了
        Exit Sub
    End If

    '取得した値がテスト有りの場合のみ表示する
    If Int(sGetIniData) = 1 Then
        cmdFixedExe(3).Visible = True
    End If
        
'V1.10.0.1 DEL
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称画面に遷移する。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
   
   On Error Resume Next
   
    Select Case Index
        Case 0                                 '通信確認
            '「通信確認・表示画面：通信確認釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_MENU_CONECT_BUTTOM, 0)
            Load frmPing
            frmPing.Show 1
        Case 1                                 '通信接続切断
             gStrCurrentForm = sFormName_ConectSts
            '「通信確認・表示画面：通信接続・切断釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_MENU_CONECT_START_END_BUTTOM, 0)
            Load frmConectSts
            frmConectSts.Show 1
        Case 2                                 '上位通信状態確認
             '「通信確認・表示画面：上位通賃状態確認釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_MENU_OVER_CONECT_BUTTOM, 0)
            Load frmOverConectSts
            frmOverConectSts.Show 1
'V1.10.0.1 ADD START
        Case 3                                 '折り返しテスト
             '「通信確認・表示画面：上位通賃状態確認釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_MENU_ORI_TEST_BUTTOM, 0)
            Load frmOriTest
            frmOriTest.Show 1
'V1.10.0.1 ADD END
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
    
    '「通信確認・表示画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_MENU_GAMEN_END, 0)
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
'//     ORIGINAL  :(1.3.0.1) 2009-13-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmTusinMenu.Caption, False
        pfFormActive (frmTusinMenu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
