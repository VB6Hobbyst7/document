VERSION 5.00
Begin VB.Form frmKansiSettei 
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "駅名データ媒体投入"
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
   Begin VB.Timer tmrMail 
      Left            =   840
      Top             =   2640
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "取扱券種モード設定"
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
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "監視設定"
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
      TabIndex        =   2
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "監視盤設定"
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
Attribute VB_Name = "frmKansiSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmKansiSettei.frm
'//  パッケージ名：監視盤設定画面
'//
'//  概要：監視盤設定画面
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//                 ・フェーズ３対応　新規追加画面
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
Private Const DEFAILT_HYOUJI_UMU = 0    '「駅名データ媒体投入」釦のデフォルト表示     'V2.7.0.1 ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 監視盤設定画面(アクティブ時)
'//  機能概要  : メール受信用のタイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
     
    On Error Resume Next
  
    'メイル受信用のタイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 監視盤設定画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    
   On Error Resume Next
    
    'メイル受信用のタイマを止める。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 監視盤設定画面(ロード時)
'//  機能概要  : 監視盤設定画面の初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
  
   Dim lSts As Long             '関数戻り値      'V2.7.0.1 ADD
   
   On Error Resume Next

    lSts = 0    '変数の初期化 'V2.7.0.1 ADD

    '「監視盤設定画面 表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSIBAN_SETTEI_GAMEN_START, 0)

    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

    'V2.7.0.1  ADD START
    'HOSHU.INIより、「駅名データ媒体投入」釦の表示有無を取得する。
    lSts = GetPrivateProfileInt(KANSI_EKIMEI_DATA_SEC, _
                                   KANSI_EKIMEI_DATA_KEY, _
                                   DEFAILT_HYOUJI_UMU, _
                                   HOSHU_FILE)
    If lSts = 1 Then
        cmdFixedExe(2).Visible = True
    Else
        cmdFixedExe(2).Visible = False
    End If
    'V2.7.0.1  ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
  
  On Error Resume Next

  Select Case Index
        Case 0                                 '取り扱い券種
            '「取扱券種モード設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSIBAN_SETTEI_KENSHUMODE_BUTTOM, 0)
            Load frmToriatukaiKenshuModeSettei
            frmToriatukaiKenshuModeSettei.Show 1
        Case 1                                 '監視設定
            '「監視設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSIBAN_SETTEI_KANSISETTEI_BUTTOM, 0)

            Load frmKansiSetteiSub
            frmKansiSetteiSub.Show 1
        'V2.7.0.1 ADD START
         Case 2                                 '駅名データ媒体入力
            '「駅名データ媒体入力釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSIBAN_SETTEI_EKIMEIDATA_INPUT_BUTTOM, 0)

            Load frmEkimeiFD
            frmEkimeiFD.Show 1
        'V2.7.0.1 ADD END
   End Select

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next

    '「監視盤設定画面 消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSIBAN_SETTEI_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : タイムアップ時処理
'//  機能概要  : メール受信タイムアップ時処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  
    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmKansiSettei.Caption, False
    End If

End Sub


