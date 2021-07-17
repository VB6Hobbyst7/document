VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRMente 
   BorderStyle     =   0  'なし
   Caption         =   "ログトレース（EG-R自動改札機）"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdRemove 
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
      Height          =   735
      Left            =   9500
      TabIndex        =   11
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "ファイル削除"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   9500
      TabIndex        =   8
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "圧縮媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   9500
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.ListBox lstTraceFile 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   240
      MultiSelect     =   2  '拡張
      TabIndex        =   5
      Top             =   1080
      Width           =   9135
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "圧縮結果確認"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   9500
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "   ファイル     媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   9500
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "表示更新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   9500
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "データ収集 "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   9500
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "リモートメンテナンス画面へ戻る"
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
      Left            =   9500
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   9120
      Top             =   8040
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "自動改札機リモートメンテナンス"
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
      TabIndex        =   12
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblListItem 
      BorderStyle     =   1  '実線
      Caption         =   "    トレースファイル名"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   7455
   End
   Begin VB.Label lblListItem 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "バイトサイズ"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Caption         =   "自動改札機  トレースファイル"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   450
      Width           =   4335
   End
End
Attribute VB_Name = "frmRMente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmRMente.frm
'//  パッケージ名：自動改札機リモートメンテナンス画面
'//
'//  概要：自動改札機リモートメンテナンス画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-07-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 トレースファイル書込み先ディレクトリ位置変更
'//                 トレースファイル圧縮書込み先ディレクトリ位置変更
'//                 圧縮ファイル選択先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.272修正対応】
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 【圧縮フォルダ指定対応】
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

'リストボックスに関する値
Private Const LIST_FILE_SIZE_LENGTH = 11   'ﾊﾞｲﾄｻｲｽﾞ欄の文字数
Private Const LIST_FILE_ELIMITTER = " -- " 'ﾊﾞｲﾄｻｲｽﾞとﾄﾚｰｽﾌｧｲﾙ間の区切文字列
Private Const LIST_HEDDER_LENGTH = LIST_FILE_SIZE_LENGTH + 4 '上記、２つの文字数合計
Private sTOOLPass As String
'Private sHyoujiGoukiNo(0 To 18) As String         '表示号機番号格納エリア          ' EG20 V3.6.0.1【統合TR-No.272修正対応】削除
Private sHyoujiGoukiNo(0 To 31) As String         '表示号機番号格納エリア           ' EG20 V3.6.0.1【統合TR-No.272修正対応】追加
Private Const TITLENAME_CORNER = "コーナ#"        ' コーナ名                        ' EG20 V6.6.0.1追加
Private sRonriCornerNo(0 To 31) As String         '論理コーナ番号格納エリア         ' EG20 V6.6.0.1追加

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : CmdRemove_Click
'//  機能名称  : 「媒体取外」釦押下時処理
'//  機能概要  : 媒体の取り外しを行う。
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
Private Sub cmdRemove_Click()
   On Error Resume Next
   
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 自動改札機リモートメンテナンス(アクティブ時)
'//  機能概要  : メール受信用タイマ、起動
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
'//  機能名称  : 自動改札機リモートメンテナンス(ディアクティブ時)
'//  機能概要  : メール受信用タイマ、停止
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

   If blnCabfrmOpenFlg = True Then
      Call fnTsbCabCallDiverge
     Exit Sub
   End If

    'タイマを止める
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 自動改札機リモートメンテナンス(ロード時)
'//  機能概要  : 初期処理を行う。
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
Private Sub Form_Load()
    Dim iRet As Integer
    
On Error Resume Next
    '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_START, 0)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

   'GLTファイルを作成し、内容を更新する。
    iRet = fMakeGLTFile
    
    If iRet = 0 Then
        'リストボックスにトレースファイル名を表示する。
        fListDisplay
    End If
    
    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdTraceFile_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称の処理を行う。
'//              「データ収集」「表示更新」「ファイル媒体出力」
'//              「圧縮媒体出力」「圧縮結果確認」「ファイル削除」
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-07-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdTraceFile_Click(Index As Integer)
    Dim lRetVal As Double      'Shell関数戻り値
    Dim iResponse As Integer   'MsgBox戻り値
    Dim sWriteDir As String    'トレースファイル書込み先のディレクトリ
    Dim lngErrCode As Long   'エラーコード
   
   On Error Resume Next

    Select Case Index
    Case 0
       '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：データ収集釦押下」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_DATA_SHUSHU_BUTTOM, 0)
        '自駅.GLTファイルへ自改情報を書込む。
        fMakeGLTFile
        '自改SW保守データ作成処理を行う。
        'If sSWFileCopy > 0 Then   'V1.6.0.1 DEL
        sSWFileCopy  'V1.6.0.1 ADD
          'リモートメンテツールを起動する。
          psGATERMenteTool
          '自動改札機ツール起動
          lRetVal = Shell(sTOOLPass, vbNormalFocus)
          If 0 = lRetVal Then
             GoTo ERROR_MSG_RMENTE
          End If
          'リモートメンテツールをアクティブ（前面表示）にする
        '  AppActivate lRetVal, True 'V1.7.0.1 DEL
        'V1.6.0.1 DEL START
        'Else
        '  '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：自改保守SWデータファイルコピー異常」ログ出力
        '  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        '  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
        'End If
        'V1.6.0.1 DEL END
    Case 1
      '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：表示更新釦押下」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
       'リストボックスにトレースファイル名を表示する。
       fListDisplay
    Case 2
      '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイル媒体出力釦押下」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_FILE_OUTPUT_BUTTOM, 0)
      sCopyTraceFile
    Case 3
      '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：圧縮媒体出力釦押下」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_LZH_OUTPUT_BUTTOM, 0)
      sLzhFileWrite
    Case 4
      '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：圧縮結果確認釦押下」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_LZH_KAKUNIN_BUTTOM, 0)
      '圧縮ファイルの内容を表示する。
      sLzhFileDisplay
    Case 5
      '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイル削除釦押下」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_FILE_DELETE_BUTTOM, 0)
       '選択中ファイルを削除する。
        If fSelectedFilesDelete = True Then
            '削除ファイルがあったなら、リストボックスを表示更新する。
            fListDisplay
        End If
    Case Else
 End Select

 Exit Sub

ERROR_MSG_RMENTE:
'===トレースデータ収集エラーの場合、
    '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：リモートメンテツール起動異常」ログ出力
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, RMENTE_GAMEN_KIDOU_ERROR, 0)
    '「リモートメンテツール起動異常」ポップアップを表示する。
    iResponse = MsgBox("リモートメンテツール（R_Mente.exe）を起動できません。", _
                vbYes, _
               "リモートメンテツール実行エラー")
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
On Error Resume Next
    '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_END, 0)
    '自画面を消す。
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMakeGLTFile
'//  機能名称  : 自駅.GLTファイルへの自改情報を書き込み処理
'//  機能概要  : GATE.INIを参照し、自駅.GLTファイルへ、
'//              号機番号、表示文字、IPアドレスを書き込む。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.6.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.7.0.1)  2012-06-28  CODED BY  [TCC] H.Sugimoto
'//                 【項目チェックの対象を改札機情報のみとする修正】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeGLTFile() As Integer
    Dim lngRet As Long          '関数の返り値
    Dim iGate As Integer        '自改INDEX
    Dim j As Integer            'ワークINDEX
    Dim sGoukiNo As String      'GLTファイルレコードデータ(号機番号表示文字)
    Dim cWork As Byte           'ワークエリア
    Dim lngErrCode As Long      'エラーコード
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim intGLTFileNo As Integer     'ﾌｧｲﾙ番号
    Dim szCorner As String      ' コーナ番号
    Dim szTitleName As String                       ' タイトル名                    ' EG20 V6.7.0.1追加
    Dim fso As New FileSystemObject                 'ファイルシステムオブジェクト   ' EG20 V6.7.0.1追加

    On Error Resume Next
    MkDir PATH_RMENTE_GATE_DEN   '自改用電鉄フォルダを作成する。（GLTファイル用）
    
' EG20 V6.7.0.1追加開始
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH_RMENTE_GATE_DEN_JIEKI) = False Then
        'コピー先フォルダ作成
        fso.CreateFolder (PATH_RMENTE_GATE_DEN_JIEKI)
    End If
    Set fso = Nothing
' EG20 V6.7.0.1追加終了
    
    
    'GLTファイルを開く。ファイルが存在しなければ新規に作成される。
    On Error GoTo ErrorHandlerGLTFile
    intGLTFileNo = FreeFile        ' 未使用のファイル番号を取得する。
    Open GATE_GLT_FILE_FULLPASS For Output As #intGLTFileNo     ' GLTファイルを開く。

    For iGate = CNT_MIN To MAX_GATE_NO - 1
      '自動改札機情報取得
      sKeyName = "gate" & Format(iGate + 1, "00")
      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                     sKeyName, _
                                     DEFAILT, sGateData, Len(sGateData), _
                                     PATH_GATE_FILE)
      If iRet = 0 Then
         '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：自動改札機INIファイル読込異常」ログ出力
         Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
         Exit Function
      End If
        
      If Len(sGateData) <> 0 Then
         'データの取得
         ReDim sFData(15)
         iFCnt = 1
            
         For iFLoop = 1 To Len(sGateData)
             If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
                iFLoop2 = iFLoop
                Do
                  iFLoop2 = iFLoop2 + 1
                  If iFLoop2 > Len(sGateData) Then
                     sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                     iFCnt = iFCnt + 1
                     If iFCnt >= 16 Then
                         Exit For
                     End If
                     
                     iFLoop = iFLoop2
                     Exit Do
                  End If
                       
                  If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
                     sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                     iFCnt = iFCnt + 1
                     If iFCnt >= 16 Then
                           Exit For
                     End If
                     
                     iFLoop = iFLoop2
                     Exit Do
                  End If
                 Loop
             End If
         Next
      End If
      
      If Len(Trim(sFData(1))) = 1 Then
         '号機番号が１桁ならば、先頭に０を付加する。
'         sGoukiNo = "0" & Trim(sFData(1)) & "号機"                                 ' EG20 V6.7.0.1削除
         sGoukiNo = "0" & Trim(sFData(1))                                           ' EG20 V6.7.0.1追加
      Else
'         sGoukiNo = Trim(sFData(1)) & "号機"                                       ' EG20 V6.7.0.1削除
         sGoukiNo = Trim(sFData(1))                                                 ' EG20 V6.7.0.1追加
      End If
        
' EG20 V6.6.0.1 【号機番号にコーナ番号を付加する対応】追加開始
'        szCorner = Replace(TITLENAME_CORNER, "#", Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))) ' EG20 V6.7.0.1削除
        szCorner = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))                                  ' EG20 V6.7.0.1追加
        sRonriCornerNo(iGate) = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.6.0.1 【号機番号にコーナ番号を付加する対応】追加終了
' EG20 V6.7.0.1 【号機番号にコーナ番号を付加する対応】追加開始
        ' コーナ番号変換
        szTitleName = Replace(RMENTE_GOKITITLENAME, "$", szCorner)
        ' 号機番号変換
        szTitleName = Replace(szTitleName, "##", sGoukiNo)
' EG20 V6.7.0.1 【号機番号にコーナ番号を付加する対応】追加開始
      
      If Trim(sFData(4)) <> "＊" Then
         'Gate.iniファイルの号機番号表示文字、IPアドレスをGLTファイルに書き込む。
'         Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(5))                     ' EG20 V6.6.0.1削除
'         Print #intGLTFileNo, szCorner & "_" & sGoukiNo & "," & Trim(sFData(5))    ' EG20 V6.6.0.1追加 ' EG20 V6.7.0.1削除
         Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))                   ' EG20 V6.7.0.1追加
      End If
      
      '表示号機番号
      If Len(Trim(sFData(1))) = 1 Then
         '号機番号が１桁ならば、先頭に０を付加する。
         sHyoujiGoukiNo(iGate) = "0" & Trim(sFData(1))
      Else
         sHyoujiGoukiNo(iGate) = Trim(sFData(1))
      End If
    
    Next
    
    'GLTファイルを閉じる。
    Close #intGLTFileNo
    
    fMakeGLTFile = 0    '正常終了
    Exit Function

ErrorHandlerGateIni:
   '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイルアクセス異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 1
   'GLTファイルを閉じる。
   Close #intGLTFileNo
   Exit Function
ErrorHandlerGLTFile:
   '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイルアクセス異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 2
   'GLTファイルを閉じる。
   Close #intGLTFileNo

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSWFileCopy
'//  機能名称  : 自改保守SW設定データファイル作成処理
'//  機能概要  : 自改保守SW設定データを、自改保守SWデータファイルに
'//              コピーする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.6.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sSWFileCopy() As Integer

     Dim iCnt As Integer                     'カウンター
     Dim sSWDataPath As String               '自改保守SWデータファイル
     Dim sMyPath As String                   '自改保守SW設定データ
     
     On Error Resume Next
   
     sSWFileCopy = 0                         'ファイル存在数
    
    '自改最大数分ループする。
    For iCnt = 1 To MAX_GATE_NO
     '「GATE_SW##.dat」の「##」を01～16に変換する。
     sMyPath = Replace(GATE_SW_FILE, "##", Format(iCnt, "0#"))
     '自改保守SW設定データの検索を行う。
     If Dir(sMyPath) <> "" Then
        '自改保守SWデータファイルのパスを作成する。
        sSWDataPath = PATH_RMENTE_GATE_DEN_JIEKI_GOUKI
' EG20 V6.6.0.1追加開始
        '「コーナ$」の「$」を1～6に変換する。
        sSWDataPath = Replace(sSWDataPath, "$", sRonriCornerNo(iCnt - 1))
' EG20 V6.6.0.1追加終了
        '「##号機」の「##」を01～16に変換する。
        sSWDataPath = Replace(sSWDataPath, "##", Format(sHyoujiGoukiNo(iCnt - 1), "0#"))
        'フォルダ作成
        MkDir sSWDataPath
        sSWDataPath = sSWDataPath & TOOL_SW_File
        
        '自改保守SWデータを自改保守SWデータファイルにコピーする。
        FileCopy sMyPath, sSWDataPath
        sSWFileCopy = sSWFileCopy + 1
     End If
   Next
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fListDisplay
'//  機能名称  : リストボックスの内容を表示更新する。
'//  機能概要  : リストボックスの表示内容を消去後、
'//              最新のトレースファイル名を表示する。
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
Private Function fListDisplay()
    Dim sInFolder(1) As String  'トレースデータフォルダ名

    On Error Resume Next

    'リストボックスを空にする。
    lstTraceFile.Clear
    'トレースデータフォルダ以下のファイルをリストボックスに表示する。
    sInFolder(0) = PATH_RMENTE_GATE_DEN_JIEKI  '本電鉄フォルダから開始する。
    sFileDisplay 1, sInFolder()
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sFileDisplay
'//  機能名称  : リストボックス表示処理
'//  機能概要  : 指定フォルダ直下のファイル名をリストボックスに表示する。
'//              最新のトレースファイル名を表示する。
'//
'//              型        名称      意味
'//  引数      : String　　sFolder
'//        　　: Integer 　iFolderNo
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sFileDisplay(iFolderNo As Integer, sFolder() As String)
    Dim iInFileNo As Integer   '検索対象フォルダ直下のファイルの個数
    Dim sInFile() As String    '  同上  ファイル名（フルパス）
    Dim iInFolderNo As Integer '検索対象フォルダ直下のファイルの個数
    Dim sInFolder() As String  '  同上  フォルダ名（フルパス：最終文字は￥）
    Dim i     As Integer       'ワークカウンタ
    Dim j     As Integer       'ワークカウンタ
    Dim sFileSize As String * LIST_FILE_SIZE_LENGTH  '表示ファイルのバイトサイズ
    Dim sDisplay As String     'リストボックスへ表示する１行分の文字列

    On Error Resume Next

    '指定されたフォルダの全てについて実施する。
    For i = CNT_MIN To iFolderNo - 1
        '検索対象フォルダ直下のファイル・フォルダを取得する。
        psFolderSearch sFolder(i), iInFileNo, sInFile(), iInFolderNo, sInFolder()
        '検索対象フォルダ直下のファイルをリストボックスへ表示する。
        For j = 0 To iInFileNo - 1
            'ﾌｧｲﾙｻｲｽﾞは右詰め、３桁のカンマ区切りで表示する。
            RSet sFileSize = Format$(FileLen(sInFile(j)), "#,###")
            'ファイル名は、・・\自電鉄\自駅\までのフォルダ:RMENTE_DIR_TRACEは表示しない。
            '            （先頭に区切り文字:LIST_FILE_ELIMITTERを付ける。）
            sDisplay = sFileSize & LIST_FILE_ELIMITTER & _
                       Right(sInFile(j), Len(sInFile(j)) - Len(PATH_RMENTE_GATE_DEN_JIEKI))
            lstTraceFile.AddItem sDisplay
        Next
        '検索対象フォルダ直下のフォルダ以下のファイルをリストボックスに表示する。
        sFileDisplay iInFolderNo, sInFolder()
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sCopyTraceFile
'//  機能名称  : 「ファイル媒体出力」釦押下時処理
'//  機能概要  : ファイルを指定ディレクトリに出力する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 トレースファイル書込み先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sCopyTraceFile()
    Dim iLine As Integer         'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行ｲﾝﾃﾞｯｸｽ
    Dim iMaxLine As Integer      'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数
    Dim iFlag As Integer         '選択中ファイル有無（１／０）
    Dim iResponse As Integer     'MsgBoxボタンコード
    Dim sFullPass As String      'コピー元ファイルフルパス名
    Dim sFileName As String      'コピー元ファイル名
    Dim sCopyDir As String       'コピー先ディレクトリ
    Dim sCopyFileName As String  'コピー先ファイル名
    Dim lSts As Long             'ワーク（戻り値）
    Dim sWork As String          'ワーク
    Dim i As Integer             'ワーク
    Dim j As Integer             'ワーク
    Dim lngErrCode As Long       'エラーコード
    Dim iFileCounter As Integer  '対象ﾌｧｲﾙ数カウンタ    ' EG20 V5.9.0.1【ログ選択上限対応】ADD

On Error GoTo COPY_ERROR
    iFlag = 0   '選択中ファイル無としておく。
    'リストボックス表示中の全行について以下を実施する。
    iMaxLine = lstTraceFile.ListCount  'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数を得る。
    
' EG20 V5.9.0.1【ログ選択上限対応】ADD START
    iFileCounter = 0
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
            iFileCounter = iFileCounter + 1
        End If
    Next

    If iFileCounter > LOG_FILECNT_MAX Then
        ' 警告文言表示
        MsgBox "選択されたファイル数が上限を超えました。" _
               & Chr(vbKeyReturn) & "選択できるファイル数は[" & LOG_FILECNT_MAX & "]件までです。", _
               vbOKOnly + vbCritical, _
               "ファイル指定異常"
        Exit Sub
    End If
' EG20 V5.9.0.1【ログ選択上限対応】ADD END
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        '選択された行ならば、
            If iFlag = 0 Then
                ' 取出し先ディレクトリを選択する
'                sCopyDir = pfDirSelection("a:", "トレースファイル書込み先のディレクトリ選択")  'V1.12.0.1 DEL
                'sCopyDir = pfDirSelection("H:", "トレースファイル書込み先のディレクトリ選択")   'V1.12.0.1 ADD　'V1.20.0.1 DEL
                sCopyDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER) 'V1.20.0.1 ADD
                If sCopyDir = "" Then
                'ディレクトリ指定がなければ、 処理を終える。
                    Exit Sub
                End If
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
                'プログレスバーを表示する
                Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
            End If
            iFlag = 1  '選択中ファイル有りとする。
            'コピー元ファイル名表示内容をセットする。（sWork←ﾊﾞｲﾄｻｲｽﾞ--01号機\CTRC2000xxx.xxx）
            sWork = lstTraceFile.List(iLine)
            '先頭からﾊﾞｲﾄｻｲｽﾞ文字（"ﾊﾞｲﾄｻｲｽﾞ--" 長さ=LIST_HEDDER_LENGTH）を除外する。
            '                                     （sFileName←01号機\CTRC2000xxx.xxx）
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            'コピー元ファイル名フルパスをセットする。（sFullPass←C:\tool\R_Mente\DATA\本電鉄\自駅\01号機\CTRC2000xxx.xxx）
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            '書込み先ディレクトリ＋ファイル（コピー 元と同じ）名をセットする。
            '                                 （sCopyFileName←a:\01号機\CTRC2000xxx.xxx）
            sCopyFileName = sCopyDir & sFileName
            'コピー先ディレクトリにフォルダを作成する。
            On Error Resume Next
            i = 1
            sWork = sCopyDir
            Do
                j = InStr(i, sFileName, "\")
                If j = 0 Then Exit Do
                j = j + 1
                sWork = sWork & Mid$(sFileName, i, j - i)
                MkDir sWork
                i = j
            Loop
            'ログトレースファイルを指定ディレクトリに書き出す。
            On Error GoTo COPY_ERROR
            FileCopy sFullPass, sCopyFileName
        End If
    Next
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    If iFlag = 0 Then
    'ファイルが選択されていなければ、エラーメッセージを表示し、処理を終了する。
        MsgBox "取出しファイルが選択されていません。" _
               & Chr(vbKeyReturn) & "選択してください。", _
               vbOKOnly + vbExclamation, _
                "リモートメンテナンス（自動改札機）"
        Exit Sub
    End If
    
    '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイル媒体出力処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RMENTE_GAMEN_FILE_OUTPUT_OK, 0)
    Exit Sub

COPY_ERROR:
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    Select Case Err.Number
        Case 61 ' コピー先空き容量不足
            iResponse = MsgBox("受け側のドライブのディスクがいっぱいです。" _
               & Chr(vbKeyReturn) & "新しいディスクを挿入してください。", _
               vbOKOnly, _
               "ログファイル取出し")
        Case 70 ' ライトプロテクト
            lSts = CopyFile(sFullPass, sCopyFileName, 0)
            If (lSts = 0) Then
                iResponse = MsgBox("ファイルを作成または置換できません。このディスクはライトプロテクトされてます。" _
                   & Chr(vbKeyReturn) & "ライトプロテクトを解除するか　別のディスクを使ってください。", _
                   vbOKOnly, _
                   "ログファイル取出し")
            End If
        Case 71 ' ディスクを挿入してください
            iResponse = MsgBox("ドライブにディスクが入ってません。" _
               & Chr(vbKeyReturn) & "ディスクを挿入してからやり直してください。", _
               vbOKOnly, _
               "ログファイル取出し")
         Case Else
            iResponse = MsgBox("予期せぬエラーが発生しました。" _
               & Chr(vbKeyReturn) & "操作をやり直してください。", _
               vbOKOnly, _
               "ログファイル取出し")
    End Select
    
    '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイル媒体出力処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RMENTE_GAMEN_FILE_OUTPUT_ERROR, lngErrCode)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sLzhFileWrite
'//  機能名称  : 「圧縮媒体出力」釦押下時処理
'//  機能概要  : リストボックスで指定されたファイルを圧縮し、
'//              指定ディレクトリに出力する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 トレースファイル圧縮書込み先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ログ媒体出力時、上限を５１２件とする
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sLzhFileWrite()
    Dim iLine As Integer         'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行ｲﾝﾃﾞｯｸｽ
    Dim iMaxLine As Integer      'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数
    Dim iFlag As Integer         '選択中ファイル有無（１／０）
    Dim iResponse As Integer     'MsgBoxボタンコード
    Dim sFullPass As String      '圧縮元ファイルフルパス名
    Dim sFileName As String      '圧縮元ファイル名
    Dim sLzhDirName As String    '.LZHﾌｧｲﾙ格納ディレクトリ名
    Dim sLzhFileName As String   '.LZHﾌｧｲﾙ名
    Dim iSts As Integer          '関数戻り値
    Dim sWork As String          'ワーク
    Dim i As Integer             'ワーク
    Dim j As Integer             'ワーク
    Dim lngErrCode As Long       'エラーコード
    Dim nIndex As Integer        ' 文字数                    ' EG20 V5.6.0.1追加
    Dim iFileCounter As Integer  '対象ﾌｧｲﾙ数カウンタ    ' EG20 V5.9.0.1【ログ選択上限対応】ADD
    
On Error GoTo WRITE_ERROR
    iFlag = 0   '選択中ファイル無としておく。
    'リストボックス表示中の全行について以下を実施する。
    iMaxLine = lstTraceFile.ListCount  'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数を得る。
    
' EG20 V5.9.0.1【ログ選択上限対応】ADD START
    iFileCounter = 0
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
            iFileCounter = iFileCounter + 1
        End If
    Next

    If iFileCounter > LOG_FILECNT_MAX Then
        ' 警告文言表示
        MsgBox "選択されたファイル数が上限を超えました。" _
               & Chr(vbKeyReturn) & "選択できるファイル数は[" & LOG_FILECNT_MAX & "]件までです。", _
               vbOKOnly + vbCritical, _
               "ファイル指定異常"
        Exit Sub
    End If
' EG20 V5.9.0.1【ログ選択上限対応】ADD END
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        '選択された行ならば、
            If iFlag = 0 Then
                ' 取出し先ディレクトリを選択する
'                sLzhDirName = pfDirSelection("a:", "トレースファイル圧縮書込み先のディレクトリ選択")   'V1.12.0.1 DEL
                'sLzhDirName = pfDirSelection("H:", "トレースファイル圧縮書込み先のディレクトリ選択")    'V1.12.0.1 ADD 'V1.20.0.1 DEL
                sLzhDirName = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
                If sLzhDirName = "" Then
                'ディレクトリ指定がなければ、 処理を終える。
                    Exit Sub
                End If
' EG20 V5.6.0.1【圧縮フォルダ指定対応】追加開始
                ' 出力フォルダに半角スペースが含まれている場合、圧縮で異常が発生してしまうため
                ' 圧縮前にチェックして異常を表示する。
                nIndex = InStr(sLzhDirName, " ")
                If nIndex <> 0 Then
                    ' 警告ポップアップウィンドウを表示する。
                    Call MsgBox(CABFOLDERSELECT_ERRORMESSAGE, vbCritical, CABFOLDERSELECT_ERRORTITLE)
                    Exit Sub  'ディレクトリが指定されなければ、処理終了
                End If
' EG20 V5.6.0.1【圧縮フォルダ指定対応】追加終了

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
                'プログレスバーを表示する
                Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
            
            End If
            iFlag = 1  '選択中ファイル有りとする。
            '圧縮元ファイル名表示内容をセットする。（sWork←ﾊﾞｲﾄｻｲｽﾞ--01号機\CTRC2000xxx.xxx）
            sWork = lstTraceFile.List(iLine)
            '先頭からﾊﾞｲﾄｻｲｽﾞ文字（"ﾊﾞｲﾄｻｲｽﾞ--" 長さ=LIST_HEDDER_LENGTH）を除外する。
            '                                  （sFileName←01号機\CTRC2000xxx.xxx）
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            '圧縮元ファイル名フルパスをセットする。（sFullPass←C:\tool\R_Mente\DATA\本電鉄\自駅\01号機\CTRC2000xxx.xxx）
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            '書込み先ディレクトリ＋ファイル（圧縮元と同じ）名をセットし、拡張子に、.CABを付加する。
            '                                 （sLzhFileName←a:\01号機\CTRC2000xxx.xxx.CAB）
            sLzhFileName = sLzhDirName & sFileName & ".CAB"
            '圧縮先ディレクトリにフォルダを作成する。
            On Error Resume Next
            i = 1
            sWork = sLzhDirName
            Do
                j = InStr(i, sFileName, "\")
                If j = 0 Then Exit Do
                j = j + 1
                sWork = sWork & Mid$(sFileName, i, j - i)
                MkDir sWork
                i = j
            Loop
            On Error GoTo WRITE_ERROR
            '対象ファイルを、圧縮し.CABファイルに格納する。
            Call psCabReqest(CABREQEST.CAB_COMPRESSION, sLzhFileName, sFullPass)
        End If
    Next
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    If iFlag = 0 Then
    'ファイルが選択されていなければ、エラーメッセージを表示し、処理を終了する。
        MsgBox "取出しファイルが選択されていません。" _
               & Chr(vbKeyReturn) & "選択してください。", _
               vbOKOnly + vbExclamation, _
               "リモートメンテナンス（自動改札機）"
        Exit Sub
    End If
    
    '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：圧縮媒体出力処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RMENTE_GAMEN_LZH_OUTPUT_OK, 0)
  
    Exit Sub

WRITE_ERROR:
    '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：圧縮媒体出力処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RMENTE_GAMEN_LZH_OUTPUT_ERROR, lngErrCode)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sLzhFileDisplay
'//  機能名称  : 「圧縮結果確認」釦押下時処理
'//  機能概要  : 指定された圧縮ファイルの内容を取得し、メモ帳表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 圧縮ファイル選択先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sLzhFileDisplay()
    Dim sLzhFileName As String   '.LZHﾌｧｲﾙ名
    Dim sLzhDataFile As String   '.LZHﾌｧｲﾙ内容書込みファイル名（ﾌﾙﾊﾟｽ）
    Dim sCommand As String
    Dim lRetVal As Long
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト  'V1.20.0.1 ADD
    
    On Error Resume Next

    '圧縮ファイル選択画面を表示し、圧縮ファイルを選択させる。
'    sLzhFileName = pfCabFileSelection("a:")        'V1.12.0.1 DEL
    'sLzhFileName = pfCabFileSelection("H:")         'V1.12.0.1 ADD 'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    '取得ファイル名を初期化
    CommonDialog1.FileName = ""
    '初期ディレクトリを設定
    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'フォルダ選択画面デフォルトパス１が存在するか
        '存在するため、デフォルトパス１（H:）を設定
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
    Else
        '存在しないため、デフォルトパス２（C:）を設定
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
    End If
    Set objFso = Nothing
    '拡張子を設定
    CommonDialog1.Filter = "圧縮ファイル（*.cab）|*.cab|"
    'ファイル選択画面を開く
    CommonDialog1.ShowOpen
    '選択したファイル名を取得
    sLzhFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
    If sLzhFileName = "" Then Exit Sub   'ファイルが選択されなければ、戻る。
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '選択された圧縮ファイルの内容を取得する。
    Call psCabReqest(CABREQEST.CAB_DRAFT, sLzhFileName, vbNullString)
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'ファイル内容取得値代入
    sLzhDataFile = gstrCabErrCd
    If sLzhDataFile = "" Then Exit Sub   'ファイル内容取得エラーであれば、戻る。
    'メモ帳の実行コマンドを作成する
    sCommand = MN_EXE_MEMO & sLzhDataFile
    lRetVal = Shell(sCommand, vbMaximizedFocus)
    'メモ帳をアクティブ（前面表示）にする
    AppActivate lRetVal, True
    SendKeys "{LEFT}", True
    
    Call ChDrive("D")  'V2.5.0.1 ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fSelectedFilesDelete
'//  機能名称  : 「ファイル削除」釦押下時処理
'//  機能概要  : 選択中のファイルを削除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//                                   True:ファイル削除　FALSE：ファイル未削除
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     ORIGINAL  :(1.1.0.2) 2009-02-XX   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSelectedFilesDelete() As Boolean
    Dim iLine As Integer         'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行ｲﾝﾃﾞｯｸｽ
    Dim iMaxLine As Integer      'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数
    Dim iDelLine As Integer      'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの選択行数
    Dim iResponse As Integer     'MsgBoxボタンコード
    Dim sFullPass As String      '削除対象ファイルフルパス名
    Dim sFileName As String      '削除対象ファイル名
    Dim sWork As String          'ワーク

On Error GoTo ErrorDeleteFile
    
    'ファイル削除なしとしておく。
    fSelectedFilesDelete = False
    iDelLine = 0
    'リストボックス表示中の全行について以下を実施する。
    iMaxLine = lstTraceFile.ListCount  'ﾄﾚｰｽﾌｧｲﾙﾘｽﾄﾎﾞｯｸｽの行数を得る。
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        '選択された行ならば、
            If iDelLine = 0 Then
                '削除確認メッセージを表示する。
                iResponse = MsgBox("選択中のファイルを削除します。" _
                                    & Chr(vbKeyReturn) & " よろしいですか？", _
                                    vbYesNo + vbExclamation, _
                                    "トレースファイルの削除")
                If iResponse = vbNo Then
                ' [いいえ] ボタンを選択した場合、削除せず終了する。
                    Exit Function
                End If
            End If
            '該当行ファイル名表示内容をセットする。（sWork←ﾊﾞｲﾄｻｲｽﾞ--01号機\CTRC2000xxx.xxx）
            sWork = lstTraceFile.List(iLine)
            '先頭からﾊﾞｲﾄｻｲｽﾞ文字（"ﾊﾞｲﾄｻｲｽﾞ--" 長さ=LIST_HEDDER_LENGTH）を除外する。
            '                                   （sFileName←01号機\CTRC2000xxx.xxx）
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            'コピー元ファイル名フルパスをセットする。（sFullPass←:\tool\R_Mente\DATA\本電鉄\自駅\01号機\CTRC2000xxx.xxx）
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            '該当行のファイルを削除する。
            Kill sFullPass
            iDelLine = iDelLine + 1
            'ファイルを削除した。
            fSelectedFilesDelete = True
            '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイル削除」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, FILE_DELETE, 0)
        End If
    Next
Exit Function

ErrorDeleteFile:

    MsgBox "ファイルの削除でエラーが発生しました。", _
           vbOKOnly + vbExclamation, _
           "トレースファイルの削除"

    fSelectedFilesDelete = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ時処理
'//  機能概要  : メール受信処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
On Error Resume Next
    '汎用メール受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmRMente.Caption, False
        pfFormActive (frmRMente.hwnd)           ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psGATERMenteTool
'//  機能名称  : 自動改札機のリモートメンテナンスツールパスを取得処理
'//  機能概要  : 自動改札機リモートメンテナンスツールパスを取得を行う。
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
Public Sub psGATERMenteTool()
 
    Dim sPath As String * MAX_PATH_SIZE
    Dim iRet As Integer
    
    On Error Resume Next

    ' HOSHU.INIより自動改札機ツールパスを取得する。
    iRet = GetPrivateProfileString(KANSI_HOSHU_GATE_RMENTE_SEC, _
                                    KANSI_HOSHU_GATE_RMENTE_KEY, _
                                    DEFAILT, sPath, Len(sPath), _
                                    HOSHU_FILE)

      If iRet = 0 Then
        sTOOLPass = ""
      Else
        sTOOLPass = sPath
      End If
      
End Sub


