VERSION 5.00
Begin VB.Form frmSyorityu 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "媒体出力中"
   ClientHeight    =   2955
   ClientLeft      =   3420
   ClientTop       =   4800
   ClientWidth     =   6030
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   600
      Top             =   600
   End
   Begin VB.Label lblLogMessage 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "媒体出力中"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2130
      TabIndex        =   0
      Top             =   1200
      Width           =   1755
   End
End
Attribute VB_Name = "frmSyorityu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmSyorityu.frm
'//  パッケージ名：処理中のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'ダイアログ表示位置
Private Const DIALOGTOP     As Integer = 3495
Private Const DIALOGLEFT    As Integer = 2985
Private Const DIALOGHEIGHT  As Integer = 3375
Private Const DIALOGWIDTH   As Integer = 6165

Private Const MN_MAIL_INTERVAL = 100       'タイマのインターバル値
Private Const MN_MAIL_INTERVAL2 = 1000     'タイマのインターバル値 ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 処理中画面(アクティブ時：イベントプロシージャ)
'//  機能概要  : 処理起動タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
    
    'タイマを起動する
    tmrMail.Enabled = True
    tmrMail2.Enabled = True     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 処理中画面(ディアクティブ時)
'//  機能概要  : メール受信用、タイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'タイマを停止する
    tmrMail.Enabled = False
    tmrMail2.Enabled = False     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 処理中画面(ロード時：イベントプロシージャ)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    '配置設定
    Me.Top = DIALOGTOP
    Me.Left = DIALOGLEFT
    Me.Height = DIALOGHEIGHT
    Me.Width = DIALOGWIDTH
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    tmrMail2.Interval = MN_MAIL_INTERVAL2
    tmrMail2.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : タイマ処理（タイムアップ時：イベントプロシージャ）
'//  機能概要  : タイムアウト処理を行う
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : Long　 　 サイズ         メール送信サイズ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    Dim bRet As Boolean '関数戻り値

    On Error Resume Next
    
    'タイマを停止する
    tmrMail.Enabled = False
    
    If glShoriNo = SHORI_NO.NO_MEDIUM_OUT Then
        
        'ＵＳＢ取り外し処理
        bRet = dllEjectUsbDevice(glErrsts)
    ElseIf glShoriNo = SHORI_NO.NO_INSTOL Then
    
        '駅都度データ媒体インストール
        Call pfTgEkiDataInstol
    End If
        
    '自画面を消す。
    Unload Me

End Sub

' EG20 V8.1.0.1【EG20_KANSI05_01】ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail2_Timer
'//  機能名称  : タイマ処理（タイムアップ時：イベントプロシージャ）
'//  機能概要  : タイムアウト処理を行う
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V8.1.0.1) 2014-06-05  CODED   BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()

    On Error Resume Next

    ' 汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSyorityu.Caption, False
        pfFormActive (frmSyorityu.hwnd)
    End If
    
End Sub
' EG20 V8.1.0.1【EG20_KANSI05_01】ADD END
