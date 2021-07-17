VERSION 5.00
Begin VB.Form frmFil 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ファイル選択画面"
   ClientHeight    =   2715
   ClientLeft      =   3795
   ClientTop       =   4860
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   2715
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelected 
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
      Height          =   1095
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.FileListBox filSelection 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   2520
      Pattern         =   "*.exe;*.com;*.bat;*.cmd"
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelected 
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
      Height          =   1095
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.DirListBox dirSelection 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.DriveListBox drvSelection 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblFileSelection 
      Caption         =   "実行ファイル選択"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmFil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmFil.frm
'//  パッケージ名：ファイル選択画面
'//
'//  概要：ファイル選択画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値 'V1.3.0.1 ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ファイル選択画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
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
'//  機能名称  : ファイル選択画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ停止
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
'//  機能名称  : ファイル選択画面(ロード時)
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
   '「ファイル選択画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, FIR_GAMEN_START, 0)

    lblFileSelection.Caption = "○○○○○○○○"
    Me.filSelection.Pattern = "*.XXX"

    '表示位置を設定する
    Me.Move Screen.Width - Me.Width, 0
    
    dirSelection.Path = drvSelection.Drive
    
    'V1.3.0.1 ADD START
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdSelected_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称の処理を行う。「確定」「取消」
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdSelected_Click(Index As Integer)
        
On Error Resume Next
    Select Case Index
    Case 0
    '「確定」ボタン押下の場合
        '表示ファイル指定のチェックを行う
        If filSelection.ListIndex = -1 Then
            'エラーメッセージを表示する
            MsgBox "ファイルが選択されていません。" _
                   & Chr(vbKeyReturn) & "選択してください。", _
                   vbOKOnly + vbExclamation, _
                   "ファイル選択"                  '実行ファイル→ファイル。
            Exit Sub
        End If
        'ファイル名をグローバルエリアにセットする
        gstrMyPath = IIf(Len(dirSelection.Path) = 3, dirSelection.Path _
                 & filSelection.List(filSelection.ListIndex), dirSelection.Path & "\" _
                 & filSelection.List(filSelection.ListIndex))
    Case 1
    '「取消」ボタン押下の場合
        'ファイル名なしをセットする。
        gstrMyPath = ""
    End Select
   
    '「ファイル選択画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, FIR_GAMEN_END, 0)

    '自画面を消す。
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : dirSelection_Change
'//  機能名称  : リストボックス内容更新処理①
'//  機能概要  : リストボックス内の更新を行う。
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
Private Sub dirSelection_Change()
On Error Resume Next
    ' ファイルパスを設定する。
    filSelection.Path = dirSelection.Path
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : dirSelection_Change
'//  機能名称  : リストボックス内容更新処理②
'//  機能概要  : リストボックス内の更新を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub drvSelection_Change()
On Error GoTo Drive_Error
    ' ディレクトリパスを設定する。
    dirSelection.Path = Left$(drvSelection.Drive, 2) & "\"
    Exit Sub
Drive_Error:
'    If Left$(drvSelection.Drive, 1) = "a" Then         'V1.12.0.1 DEL
    If Left$(drvSelection.Drive, 1) = "H" Then          'V1.12.0.1 ADD
    'a:ドライブが異常なら、カレントドライブを表示させる。
        drvSelection.Drive = Left$(App.Path, 2)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : dirSelection_Change
'//  機能名称  : メール受信用タイマがタイムアップ処理
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
'///////////////////////////////////////////////////////////////////*
Private Sub tmrMail_Timer()
On Error Resume Next
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmFil.Caption, False
    End If
End Sub
