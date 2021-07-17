VERSION 5.00
Begin VB.Form frmVerChang 
   BorderStyle     =   0  'なし
   Caption         =   "バージョン管理"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   14.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
   Begin VB.Timer tmrMail 
      Left            =   5160
      Top             =   1080
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "磁気運賃"
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
      Index           =   7
      Left            =   4440
      TabIndex        =   21
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＰＡＳＭＯ運賃"
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
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "EG-R自改"
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4095
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   27
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "・バージョンチェックファイル："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label lblVerName 
         Caption         =   "・予備2："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "・メインCPU-OS："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1380
         Width           =   2370
      End
      Begin VB.Label lblVerName 
         Caption         =   "・メインCPU-Pro："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   705
         Width           =   2205
      End
      Begin VB.Label lblVerName 
         Caption         =   "・予備１："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "・サブCPU-Pro："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1035
         Width           =   2175
      End
      Begin VB.Label lblVerName 
         Caption         =   "・判定CPU-Pro： "
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2190
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   11
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3480
         TabIndex        =   10
         Top             =   705
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   9
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   8
         Top             =   1035
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "判定ＩＣ－Ｍ"
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
      Left            =   8400
      TabIndex        =   3
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＥＧ－Ｒ自改"
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
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＮＥＧ自改"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "   メニュー     画面へ戻る"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblVerName 
      Caption         =   "Z9"
      Height          =   375
      Index           =   22
      Left            =   7320
      TabIndex        =   29
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "・磁気運賃： "
      Height          =   495
      Index           =   21
      Left            =   4440
      TabIndex        =   28
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   20
      Left            =   7320
      TabIndex        =   25
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   19
      Left            =   7320
      TabIndex        =   24
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "Z9"
      Height          =   375
      Index           =   18
      Left            =   7320
      TabIndex        =   23
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "・NEG自改："
      Height          =   375
      Index           =   15
      Left            =   4440
      TabIndex        =   22
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "・ＩＣ共通運賃： "
      Height          =   495
      Index           =   17
      Left            =   4440
      TabIndex        =   20
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "バージョン切替"
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
   Begin VB.Label lblVerName 
      Caption         =   "・ＩＣ－Ｍ："
      Height          =   375
      Index           =   16
      Left            =   4440
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "frmVerChang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmVerChang.frm
'//  パッケージ名：バージョン切替画面
'//
'//  概要：バージョン切替画面
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//                 フェーズ２対応
'//     REVISIONS :(1.0.6.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 フェーズ1不具合対応
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
Private Const APL_INTERVAL = 390000     'アプリ起動タイマデフォルト値 'V1.6.0.1 ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン切替画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
   On Error Resume Next
    
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : バージョン切替画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    On Error Resume Next
   
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub

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
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  On Error Resume Next
 
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmVerChang.Caption, False
        pfFormActive (frmVerChang.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : バージョン切替画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

   On Error Resume Next
 
   '「バージョン切替画面：表示」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_GAMEN_START, 0)

   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
      
   '磁気運賃対応チェック
   psJikiCheck
    
   'IDU縮退チェック
   psIDUCheck
    
   If pbIDUSts = 1 Then
     'IDU業務非表示
      cmdFixedExe(5).Visible = False
      cmdFixedExe(6).Visible = False
   End If
   
   'バージョン取得処理
   psGetVersion

   'メール受信用のタイマ値を設定する。
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
   
   'V1.6.0.1 ADD START
   'INIファイルよりアプリ起動タイマ値を取得
   frmChangeVer.lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      APL_INTERVAL, HOSHU_FILE)
   '取得値が0の場合、デフォルト値を設定
   If frmChangeVer.lngMAX_Time = 0 Then
      frmChangeVer.lngMAX_Time = APL_INTERVAL
   End If
   'タイマ値設定
   frmChangeVer.tmrAplCheck.Interval = MN_MAIL_INTERVAL
   frmChangeVer.tmrAplCheck.Enabled = False
   'V1.6.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 各釦対象のバージョン切替処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　 Index    [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    
   On Error Resume Next
    
    Dim iRet As Integer
    Dim strWord As String
    Dim bRet As Boolean
    
    '切替対象を変数に設定する。
    Change_Version = Index
    
    Select Case Index
       Case EGR_CHANGE_VER
         '「バージョン切替画面：EG-R自改釦押下」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EGRJIKAI_BUTTOM, 0)
       Case NEG_CHANGE_VER
         '「バージョン切替画面：NEG自改釦押下」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_NEGJIKAI_BUTTOM, 0)
       Case ICM_CHANGE_VER
         '「バージョン切替画面：判定IC-M釦押下」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_ICM_BUTTOM, 0)
       Case PASMO_CHANGE_VER
         '「バージョン切替画面：PASMO運賃釦押下」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_PASMO_BUTTOM, 0)
       Case JIKIUNCHIN_CHANGE_VER
         '「バージョン切替画面：磁気運賃釦押下」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_JIKIUNCHIN_BUTTOM, 0)
  End Select
    
    strWord = cmdFixedExe(Index).Caption & "のバージョンを切り替えます。" & vbCrLf & "よろしいですか？"
    
    iRet = MsgBox(strWord, vbQuestion + vbOKCancel, "バージョン切替確認")
    
    If iRet = vbOK Then
       Load frmChangeVer
       frmChangeVer.lblMessage(0).Caption = "バージョン切替中です。"
       frmChangeVer.lblMessage(1).Caption = "しばらくお待ち下さい。"
       frmChangeVer.Show 1
    End If
'V1.8.0.1 ADD START
    'バージョン取得処理
    psGetVersion
'V1.8.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '「バージョン切替画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psGetVersion
'//  機能名称  : バージョン取得処理
'//  機能概要  : バージョン取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.4.0.1) 2009-03-17    CODED BY [TCC] S.Terao
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psGetVersion()

  On Error Resume Next
 
  Dim sVersion  As String
  Dim sGetJikiVer As String     'V1.10.0.1 ADD
  
 'EG-R自改バージョン取得
  '判定CPU
  sVersion = psEGRJVersion(HANTEI_CPU)
  lblVerName(8).Caption = sVersion
  'メインCPU
  sVersion = psEGRJVersion(MAIN_CPU)
  lblVerName(9).Caption = sVersion
 'サブCPU
  sVersion = psEGRJVersion(SUB_CPU)
  lblVerName(10).Caption = sVersion
 'メインOS
  sVersion = psEGRJVersion(MAIN_OS)
  lblVerName(11).Caption = sVersion
 '予備１
  sVersion = psEGRJVersion(YOBI1)
  lblVerName(12).Caption = sVersion
 '予備２
  sVersion = psEGRJVersion(YOBI2)
  lblVerName(13).Caption = sVersion
 'バージョンチェック
  sVersion = psEGRJVersion(VER_CHK)
  lblVerName(14).Caption = sVersion
  
 'NEG自改バージョン取得
  sVersion = psNEGJVersion
  lblVerName(18).Caption = sVersion

 'IC-Mバージョン取得
 If pbIDUSts = 1 Then
    'IDUバージョン非表示
    lblVerName(16).Enabled = False
    lblVerName(19).Caption = ""
 Else
    '非縮退時は表示処理
    sVersion = psICMGetVersion
    lblVerName(19).Caption = sVersion
 End If
 
 '共通運賃バージョン取得
 If pbIDUSts = 1 Then
    'IDUバージョン非表示
    lblVerName(17).Enabled = False
    lblVerName(20).Caption = ""
 Else
    '非縮退時は表示処理
    sVersion = psICUnchinGetVersion
    lblVerName(20).Caption = sVersion
 End If
 
'V1.10.0.1 ADD START
 '磁気運賃読み込み
 sGetJikiVer = psJikiUnchinVersion
 lblVerName(22).Caption = CStr(sGetJikiVer)
'V1.10.0.1 ADD END

 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psJikiCheck
'//  機能名称  : 磁気運賃対応ユーザチェック処理
'//  機能概要  : HOSHU.INIより、磁気運賃対応ユーザであるかどうかチェックする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psJikiCheck()
    Dim iFlag As Integer '取得ユーザフラグ
 
    On Error Resume Next
 
  ' HOSHU.INIより磁気運賃対応ユーザフラグを取得する。
    iFlag = GetPrivateProfileInt(KANS_JIKI, _
                                 KANSI_JIKI_FLAG, _
                                 DEFAILT_Int, _
                                 HOSHU_FILE)
     If iFlag = 0 Then
      'フラグが0の場合「磁気運賃」釦は非表示
      cmdFixedExe(7).Visible = False
     End If
End Sub
