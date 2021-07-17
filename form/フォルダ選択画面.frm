VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ディレクトリ選択画面"
   ClientHeight    =   2355
   ClientLeft      =   3960
   ClientTop       =   4425
   ClientWidth     =   5790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   2355
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTorikesi 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   1920
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.DirListBox dirSelection 
      Height          =   1770
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.DriveListBox drvSelection 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmDir.frm
'//  パッケージ名：フォルダ(ディレクトリ)選択画面
'//
'//  概要：フォルダ(ディレクトリ)選択画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・EG10より、フォルダ選択画面流用。
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : フォルダ(ディレクトリ)選択画面(アクティブ時)
'//  機能概要  : メール受信用タイマ起動
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
On Error Resume Next
    'タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : フォルダ(ディレクトリ)選択画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマ停止
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
On Error Resume Next
    'タイマを停止する
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : フォルダ(ディレクトリ)選択画面(ロード時)
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
    '「フォルダ選択画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DIR_GAMEN_START, 0)

    '表示位置を設定する
    Me.Move Screen.Width - Me.Width, 0

     ' ディレクトリパスの設定。
    dirSelection.Path = Left(App.Path, 3)
    
    'V1.3.0.1 ADD START
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdKakutei_Click
'//  機能名称  : 「確認」釦押下時処理
'//  機能概要  : 選択された内容を保存し、画面消去。
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
Private Sub cmdKakutei_Click()
On Error Resume Next
    '選択ディレクトリを設定する。
    '「c:\」や「d:\」といったルートドライブを設定した場合は￥マークを削除
    gstrMyPath = IIf(Len(dirSelection.Path) = 3, dirSelection.Path, _
                 dirSelection.Path + "\")
     '「フォルダ選択画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DIR_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdTorikesi_Click
'//  機能名称  : 「取消」」釦押下時処理
'//  機能概要  : 選択ディレクトリなし状態を保存し、画面消去。
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
Private Sub cmdTorikesi_Click()
On Error Resume Next
    '選択ディレクトリなし、を設定する。
    gstrMyPath = ""
     '「フォルダ選択画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DIR_GAMEN_END, 0)
    '自画面を消す。
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : drvSelection_Change
'//  機能名称  : ドライブリストボックスの内容変更時処理
'//  機能概要  : リストボックスの内容を更新する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub drvSelection_Change()
On Error GoTo Drive_Error
     ' ディレクトリパスの設定。
    dirSelection.Path = Left$(drvSelection.Drive, 2) & "\"
    Exit Sub
Drive_Error:
'    If Left$(drvSelection.Drive, 1) = "a" Then     'V1.12.0.1 DEL
    If Left$(drvSelection.Drive, 1) = "H" Then      'V1.12.0.1 ADD
    'a:ドライブが異常なら、カレントドライブを表示させる。
        drvSelection.Drive = Left$(App.Path, 2)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
On Error Resume Next
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmDir.Caption, False
        pfFormActive (frmDir.hwnd)
    End If
End Sub
