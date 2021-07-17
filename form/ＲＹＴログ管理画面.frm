VERSION 5.00
Begin VB.Form frmRYTSyusyu 
   BorderStyle     =   0  'なし
   Caption         =   "稼働・メンテデータ収集（次世代自動改札機）"
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
   Begin VB.CommandButton cmdInstall 
      Caption         =   "媒体取外"
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
      Left            =   8520
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   7200
      Top             =   5760
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "      メニュー       画面へ戻る"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdSyusyu 
      Caption         =   " 収集 "
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
      Left            =   1080
      TabIndex        =   0
      Top             =   3315
      Width           =   2175
   End
   Begin VB.CommandButton cmdFDWrite 
      Caption         =   "媒体出力"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   3315
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "ＲＹＴログ管理"
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
Attribute VB_Name = "frmRYTSyusyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmRYTSyusyu.frm
'//  パッケージ名：ＲＹＴログ管理画面
'//
'//  概要：ＲＹＴログ管理画面
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　ＲＹＴログ管理画面追加
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : ＲＹＴログ管理画面(ロード時)
'//  機能概要  : 初期処理を行う。
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
Private Sub Form_Load()

    On Error Resume Next
    
    'メイル受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
              
    '「ＲＹＴログ管理画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_START, 0)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ＲＹＴログ管理画面(アクティブ時)
'//  機能概要  : メールタイマを起動する。
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
Private Sub Form_Activate()
  
    On Error Resume Next
    
    'メール受信用タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : ＲＹＴログ管理画面(ディアクティブ時)
'//  機能概要  : メールタイマを停止する。
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
Private Sub Form_Deactivate()
  
    On Error Resume Next
    
    'メール受信用タイマを止める
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
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
Private Sub cmdReturn_Click()
    
    On Error Resume Next
    
    '「ＲＹＴログ管理画面：消去」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_END, 0)

    '自画面を消す。
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdSyusyu_Click
'//  機能名称  : 「収集」釦押下時処理
'//  機能概要  : 「収集」釦押下時処理を行う。
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
Private Sub cmdSyusyu_Click()
    Dim iResponse As Integer   'MsgBox戻り値
    
    On Error Resume Next

    '「収集釦押下」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_SYUSYU_BUTTOM, 0)

     'ＲＹＴログデータ収集ポップアップ画面表示
    iResponse = MsgBox("ＲＹＴログデータを収集しますがよろしいですか？", _
                       vbOKCancel, "ＲＹＴログデータ管理")
    If iResponse = vbOK Then
       'OK釦が押されたら、ＲＹＴログデータ収集中画面表示
       'RYTログデータ収集中フォームを、モーダルウィンドウで表示する。
         frmRYTSyusyuCyu.Show vbModal
    Else
       'キャンセル釦押下
       '「収集処理未実行」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_MISHORI, 0)
       Exit Sub
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFDWrite_Click
'//  機能名称  : 「媒体出力」釦押下時処理
'//  機能概要  : 「媒体出力」釦押下時処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFDWrite_Click()
    
   Dim fso         As New FileSystemObject 'ファイルシステムオブジェクト
   Dim sWriteDir As String
   
   On Error Resume Next
  
   '「媒体出力釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_OUTPUT_BUTTOM, 0)
  
    'フォルダ選択ポップアップ画面表示
'    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")     'V1.12.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)      'V1.12.0.1 ADD

    '指定フォルダなし
    If Len(sWriteDir) = 0 Then
       '「媒体出力処理未実行」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_OUTPUT_MISHORI, 0)
       Exit Sub
    End If
    
    Set fso = Nothing

    m_sCopySaki = sWriteDir & "\" & RYT_LOG_FILE

    'コピー元パス作成
    m_sCopyMoto = E_FIRMWARE_LOG & RYT_LOG_FILE
    
    frmRYTSyusyuOutPut.Show vbModal
          
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdInstall_Click
'//  機能名称  : 「媒体取外」釦押下時処理
'//  機能概要  : 媒体の取り外しを行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdInstall_Click()
   
   On Error Resume Next
   
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)
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
        AppActivate frmRYTSyusyu.Caption, False
    End If

End Sub

