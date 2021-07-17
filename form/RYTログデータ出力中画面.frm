VERSION 5.00
Begin VB.Form frmRYTSyusyuOutPut 
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ｏ Ｋ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '中央揃え
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '中央揃え
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmRYTSyusyuOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmRYTSyusyuOutPut.frm
'//  パッケージ名：ＲＹＴログデータ出力中画面
'//
'//  概要：ＲＹＴログデータ出力中画面
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//                 ・フェーズ３対応　ＲＹＴログデータ出力中画面追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ＲＹＴログデータ出力中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
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
Private Sub Form_Activate()
    
    On Error Resume Next
  
    '出力中のガイドを表示する｡
    lblMessage(0) = "ＲＹＴログデータを出力中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    tmrMail.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : ＲＹＴログデータ出力中画面(ロード時)
'//  機能概要  : 初期処理を行う。
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
Private Sub Form_Load()
  
  On Error Resume Next
  
  cmdOK.Visible = False
  cmdOK.Enabled = False
  
  'メイル受信用のインタバルタイマ値を設定する。
  tmrMail.Interval = MN_MAIL_INTERVAL
  tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : ＲＹＴログデータ出力中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
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
Private Sub Form_Deactivate()
    
    On Error Resume Next
    
    'メール受信用タイマを止める
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdOK_Click
'//  機能名称  : 「OK」釦押下時処理
'//  機能概要  : 自画面を消去する。
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
Private Sub cmdOK_Click()
    
    On Error Resume Next
    
    '自画面を消す。
    Unload Me
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
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    On Error Resume Next
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSyusyuOutPut.Caption, False
    End If
    
    '出力ファイル作成処理を行う。
    sOutPutRYTLog_Data

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sOutPutRYTLog_Data
'//  機能名称  : 保守データコピー処理
'//  機能概要  : 保守データコピーを行う。
'//
'//              型        名称      意味
'//  引数      : String　sOutPath　[IN]出力先パス
'//              String  sFromPath [IN]コピー元パス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function sOutPutRYTLog_Data()
   
   On Error GoTo Error_RYT_OutPut
 
   FileCopy m_sCopyMoto, m_sCopySaki
   
   sOutPutEnd (0)
  
  Exit Function
   
Error_RYT_OutPut:
 'ファイルコピー異常
 sOutPutEnd (1)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sOutPutEnd
'//  機能名称  : 出力結果表示処理
'//  機能概要  : ＲＹＴログデータ出力結果の結果文言を表示する。
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sOutPutEnd(iEnd As Integer)
    Dim i As Integer       'カウンタ
    Dim lngErrCode As Long 'エラーコード

    On Error Resume Next
        
    If iEnd = 0 Then
       '正常終了時の文言を表示する。
       lblMessage(0) = "正常終了しました。"
       lblMessage(1) = ""
       '「ＲＹＴログデータ出力処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RYT_LOG_KANRI_GAMEN_OUTPUT_OK, 0)
    Else
       '収集失敗時の文言を表示する。
       lblMessage(0) = "異常終了しました。"
       lblMessage(1) = ""
       '「ＲＹＴログデータ出力理異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RYT_LOG_KANRI_GAMEN_OUTPUT_ERROR, lngErrCode)
    End If
    cmdOK.Visible = True
    cmdOK.Enabled = True
End Sub

