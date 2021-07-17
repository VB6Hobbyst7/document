VERSION 5.00
Begin VB.Form frmALLSysformat 
   BorderStyle     =   0  'なし
   Caption         =   "システム初期化機能(一括初期化)"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrLogTimer 
      Left            =   9840
      Top             =   3360
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   9840
      Top             =   2880
   End
   Begin VB.Timer tmrMail 
      Left            =   9840
      Top             =   2280
   End
   Begin VB.CommandButton cmdZikko 
      Caption         =   "初期化実行"
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
      Left            =   9120
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.ListBox LstStatus 
      Height          =   6360
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   8415
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   15000
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "システム初期化  画面へ戻る"
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
      Left            =   9120
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblHelp 
      Caption         =   "・Ｌ Ｄ Ｕ ：アプリケーションログ、保守プログラムログ、改札機ログ、その他データ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   8300
      Width           =   8895
   End
   Begin VB.Label lblHelp 
      Caption         =   $"一括初期化画面.frx":0000
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   7890
      Width           =   8895
   End
   Begin VB.Label lblHelp 
      Caption         =   $"一括初期化画面.frx":0097
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   7500
      Width           =   9015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "一括システム出荷時初期化"
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
      TabIndex        =   7
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblHelp 
      Caption         =   $"一括初期化画面.frx":012C
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7050
      Width           =   8895
   End
   Begin VB.Label lblKekka 
      BorderStyle     =   1  '実線
      Caption         =   "初期化は成功しました。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   5
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Caption         =   "初期化結果"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "frmALLSysformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmALLSysformat.frm
'//  パッケージ名：一括システム初期化画面
'/
'//  概要：一括システム初期化画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//         フェーズ２対応　設定ファイル(保存用）を追加
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                         保守総点検結果修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 リストボックスのスクロール処理追加
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加、確認ポップアップ修正
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-240】
'//
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private bChk() As Boolean

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000     'アプリ起動タイマデフォルト値
Dim lngMAX_Time As Long                    'INI取得設定値
Dim lngtime     As Long                    '現在タイマ値
'V1.5.0.1 ADD END
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        'ログ起動タイマデフォルト値(30秒)
Dim lngLogMAX_Time As Long                'INI取得設定値(ログ）
'V1.20.0.1 ADD END
'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 一括システム初期化画面(アクティブ時)
'//  機能概要  : 最前面表示を行う。
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
    pfFormActive (hwnd)
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 一括システム初期化画面(ディアクティブ時)
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
'//  機能名称  : 一括システム初期化画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.4.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    On Error Resume Next

    '「一括システム初期化：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ALL_SYSFORMAT_GAMEN_START, 0)

    '配置設定
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

    '初期化
    LstStatus.Clear
    lblKekka.Caption = ""
    
   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
   
   'V1.5.0.1 ADD START
   'INIファイルよりアプリ起動タイマ値を取得
   lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      APL_INTERVAL, HOSHU_FILE)
   '取得値が0の場合、デフォルト値を設定
   If lngMAX_Time = 0 Then
      lngMAX_Time = APL_INTERVAL
   End If
   
   'V1.20.0.1 ADD START
   'INIファイルよりログ起動タイマ値を取得
   lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
   '取得値が0の場合、デフォルト値を設定
   If lngLogMAX_Time = 0 Then
      lngLogMAX_Time = LOG_INTERVAL
   End If
   'V1.20.0.1 ADD END
   
   'タイマ値設定
   tmrAplTimer.Interval = MN_MAIL_INTERVAL
   tmrAplTimer.Enabled = False
   'V1.5.0.1 ADD END
   
   'V1.20.0.1 ADD START
   tmrLogTimer.Interval = MN_MAIL_INTERVAL
   tmrLogTimer.Enabled = False
   'V1.20.0.1 ADD END
  
   'V1.4.0.1 ADD START
   'IDU縮退チェック
    psIDUCheck
   'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdZikko_Click
'//  機能名称  : 「初期化実行」釦押下時処理
'//  機能概要  : システム初期化を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//         フェーズ２対応　設定ファイル(保存用）を追加
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                         保守総点検結果修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加、確認ポップアップ修正
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-23  REVISED BY [TCC] T.Koyama
'//                 EG20フェーズ２対応【残件№54】
'//                 ・保守ログファイルＣＬＯＳＥ処理追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZikko_Click()

    Dim iRet As Integer
    Dim sDBFormat As String
    Dim sLine As String
    Dim sExecName As String
    Dim sDbInitCmd As String
    Dim bRtn1 As Boolean
    Dim lngErrCode As Long              'エラーコード
    Dim uMail As ML_KYOTU_INF           'メール
    Dim bRet As Boolean
    Dim iKansiApp As Integer            '監視盤アプリ起動フラグ
    Dim iRetIDULog As Integer           'IDUログ起動フラグ
    Dim iRetLDULog As Integer           'IDUログ起動フラグ
    Dim bKansiDB_Code As Boolean
    Dim bIDUDB_Code As Boolean
    Dim lExitCode As Long
    Dim iTargetDB As Integer            '対象DB値
    ReDim bChk(9)
    Dim i As Integer                    'カウンター
    Dim bRtn As Boolean
    'V1.5.0.1  ADD START
    Dim bKansiRet As Boolean            '監視盤アプリ処理結果
    Dim bIDURet   As Boolean            'IDUアプリ処理結果
    Dim bLDURet   As Boolean            'LDUアプリ処理結果
   
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1  ADD END
    On Error GoTo ERR_SPACE

    '「一括システム初期化画面：初期化実行釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_START_BUTTOM, 0)

    '表示の初期化
    LstStatus.Clear
    lblKekka.Caption = ""
    
    iKansiApp = 1
    iRetIDULog = 1
    iRetLDULog = 1
    
    '「初期化確認」ポップアップを表示
    iRet = MsgBox("初期化処理を行います。よろしいですか？", vbExclamation + vbOKCancel, "初期化確認")
    If iRet = vbOK Then
         cmdZikko.Enabled = False  '「初期化実行」釦押下不可
         cmdCancel.Enabled = False '「メニュー画面へ戻る」釦押下不可

         On Error GoTo ERR_SPACE2
         
         '監視盤(管理プロセス)が起動しているかどうかチェックする。
         If CheckAppStart(PROC_KANRI) <> 0 Then
           'V1.20.0.1 DEL START
           'iRet = MsgBox("監視盤、ＩＤ中継ユニット、ＬＤユーティリティアプリケーションを終了します。" & Chr(vbKeyReturn) & _
           '               "よろしいですか？", vbQuestion + vbOKCancel, "終了確認")
           'If iRet = vbOK Then
           'V1.20.0.1 DEL END
              'アプリ終了要求を管理に送信する
               uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
               uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
               uMail.udtlHeader.dwProid = RHOSHU_ID
               uMail.udtlHeader.dwSubArea = 0
               'V1.5.0.1 DEL START
               'bRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               'If bRet = 0 Then
               'V1.5.0.1 DEL END
               'V1.5.0.1 ADD START
               bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               If bKansiRet = 0 Then
                  'V1.5.0.1 ADD END
                  '「一括システム初期化画面：メール送信異常」ログ出力
                  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                  Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lngErrCode)
                  GoTo ERR_SPACE2:
               Else
                  '「一括システム初期化画面：メール送信正常」ログ出力
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                  '監視盤アプリ終了確認
                  'iKansiApp = CheckAppEndComplete(PROC_KANRI, lExitCode)            'V1.5.0.1 DEL
               End If

           'V1.20.0.1 DEL START
'
'               'ログプロセス起動チェック
'               If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'
'                  'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
'                  iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
'                  If iRet = vbOK Then
'                     'IDUログ終了要求CMD送信
'                     'V1.5.0.1 DEL START
'                     'bRtn = EndIDULog
'                     'If bRtn = False Then
'                     'V1.5.0.1 DEL END
'                     'V1.5.0.1 ADD START
'                     bIDURet = EndIDULog
'                     If bIDURet = False Then
'                     'V1.5.0.1 ADD END
'                        '送信異常処理
'                        lblKekka.ForeColor = SYSFORMAT_ERROR
'                        lblKekka.Caption = "初期化に失敗しました"
'                        cmdZikko.Enabled = True
'                        cmdCancel.Enabled = True
'                        '処理を抜ける
'                        Exit Sub
'                     End If
'
'                     'IDUログプロセス終了確認
'                     'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) 'V1.5.0.1 DEL
'                  'V1.7.0.1 ADD START
'                  Else
'                    'ログプロセス終了メッセージ「キャンセル」釦押下時処理
'                    GoTo ERR_SPACE3
'                  'V1.7.0.1 ADD END
'                  End If
'               'V1.5.0.1 ADD START
'               Else
'                 bIDURet = True
'               'V1.5.0.1 ADD END
'               End If
'
'               'LDUログプロセス起動チェック
'               If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'
'                  'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
'                  iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
'
'                  If iRet = vbOK Then
'                     'LDUログ終了要求CMD送信
'                     'V1.5.0.1 DEL START
'                     'bRtn = EndLDULog
'                     'If bRtn = False Then
'                     'V1.5.0.1 DEL END
'                     'V1.5.0.1 ADD START
'                     bLDURet = EndLDULog
'                     If bLDURet = False Then
'                     'V1.5.0.1 ADD END
'                        '送信異常処理
'                        lblKekka.ForeColor = SYSFORMAT_ERROR
'                        lblKekka.Caption = "初期化に失敗しました"
'                        cmdZikko.Enabled = True
'                        cmdCancel.Enabled = True
'                        '処理を抜ける
'                        Exit Sub
'                     End If
'
'                     'LDUログプロセス終了確認
'                     'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) 'V1.5.0.1 DEL
'                  'V1.7.0.1 ADD START
'                  Else
'                    'ログプロセス終了メッセージ「キャンセル」釦押下時処理
'                    GoTo ERR_SPACE3
'                  'V1.7.0.1 ADD END
'                  End If
'               'V1.5.0.1 ADD START
'                   Else
'                bLDURet = True
'               'V1.5.0.1 ADD END
'               End If
'           Else
'             '「キャンセル釦押下」
'             '「一括システム初期化画面：初期化処理未実行」ログ出力
'             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'             cmdZikko.Enabled = True  '「初期化実行」釦押下可
'             cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
'             '処理を抜ける
'             Exit Sub
'           End If
'        'V1.5.0.1 ADD START
        'V1.20.0.1 DEL END
         Else
            bKansiRet = True
        'V1.5.0.1 ADD END
        ' End If  'V1.5.0.1 DEL
          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
             
             'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
             'V1.20.0.1 DEL START
             'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
             'If iRet = vbOK Then
             'V1.20.0.1 DEL END
             
                'IDU/LDUログ終了要求CMD送信
                'V1.5.0.1 DEL START
                'bRet = EndIDULog
                'If bRet = False Then
                'V1.5.0.1 DEL END
                'V1.5.0.1 ADD START
                bIDURet = EndIDULog
                If bIDURet = False Then
                'V1.5.0.1 ADD END
                 '送信異常
                 lblKekka.ForeColor = SYSFORMAT_ERROR
                 lblKekka.Caption = "初期化に失敗しました"
                 cmdZikko.Enabled = True
                 cmdCancel.Enabled = True
                 '処理を抜ける
                 Exit Sub
               End If
               
               'IDUログプロセス終了確認
               'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) 'V1.5.0.1 DEL
             'V1.7.0.1 ADD START
          'V1.20.0.1 DEL START
'             Else
'              'ログプロセス終了メッセージ「キャンセル」釦押下時処理
'              GoTo ERR_SPACE3
'             'V1.7.0.1 ADD END
'             End If
'         'V1.5.0.1 ADD START
          'V1.20.0.1 DEL END
          Else
            bIDURet = True
          'V1.5.0.1 ADD END
          End If
         
          If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
             
             'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
             'V1.20.0.1 DEL START
             'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
             '
             'If iRet = vbOK Then
             'V1.20.0.1 DEL END
             'IDU/LDUログ終了要求CMD送信
             'V1.5.0.1 DEL START
             'bRet = EndLDULog
             'If bRet = False Then
             'V1.5.0.1 DEL END
             'V1.5.0.1 ADD START
             bLDURet = EndLDULog
             If bLDURet = False Then
             'V1.5.0.1 ADD END
                '送信異常
                lblKekka.ForeColor = SYSFORMAT_ERROR
                lblKekka.Caption = "初期化に失敗しました"
                cmdZikko.Enabled = True
                cmdCancel.Enabled = True
                '処理を抜ける
                Exit Sub
              End If
              
              'LDUログプロセス終了確認
              'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)  'V1.5.0.1 DEL
            'V1.7.0.1 ADD START
          'V1.20.0.1 DEL START
'             Else
'              'ログプロセス終了メッセージ「キャンセル」釦押下時処理
'              GoTo ERR_SPACE3
'            'V1.7.0.1 ADD END
'             End If
'         'V1.5.0.1 ADD START
          'V1.20.0.1 DEL END
          Else
            bLDURet = True
          'V1.5.0.1 ADD END
         End If
       End If        'V1.5.0.1 ADD
'V1.5.0.1 ADD START
       '監視盤、IDU、LDUアプリのメール送信処理が全て正常だった場合のみ、アプリ起動タイマを起動させ、
       'アプリ起動チェックによりアプリの起動/未起動を判断する。
       'If (bKansiRet = True) And (bIDURet = True) And (bLDURet = True) Then  'V1.20.0.1 DEL
       If (bKansiRet = True) Then  'V1.20.0.1 ADD
           lngtime = 0
           lngtime = MN_MAIL_INTERVAL
           tmrAplTimer.Enabled = True
       Else
          '監視盤、IDU、LDUアプリのメール送信にてひとつでも異常があった場合、初期化処理を異常終了とする。
          '「一括システム初期化画面：システム初期化処理異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "初期化に失敗しました"
          cmdZikko.Enabled = True
          cmdCancel.Enabled = True
          '処理を抜ける
          Exit Sub
       End If
   End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'         'アプリまたはログプロセスで終了処理に失敗した場合
'         If (iKansiApp <> 1) Or (iRetIDULog <> 1) Or (iRetLDULog <> 1) Then
'            '「一括システム初期化画面：システム初期化処理異常」ログ出力
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "初期化に失敗しました"
'           cmdZikko.Enabled = True
'           cmdCancel.Enabled = True
'           '処理を抜ける
'           Exit Sub
'         End If
'
'        'V1.4.0.1 ADD START
'        If sCreateShokiFile = False Then
'           '「一括システム初期化画面：システム初期化処理異常」ログ出力
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "初期化に失敗しました"
'           cmdZikko.Enabled = True
'           cmdCancel.Enabled = True
'           '処理を抜ける
'           Exit Sub
'        End If
'        'V1.4.0.1 ADD END
'
'        'システムファイルの削除処理
'        bRtn1 = sSysFileDelete()
'
'        'システムファイル削除処理成功した場合、
'        'フォルダ、ファイルの削除処理を行う
'        If bRtn1 = True Then
'
'            '監視盤システム初期化
'            For i = 1 To 6
'               bChk(i) = True
'            Next
'            bChk(5) = False
'
'            If sFileDelete(stsKansi, KANSI_SYSTEMFILE) = False Then
'              '「一括システム初期化画面：システム初期化処理異常」ログ出力
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "初期化に失敗しました"
'              cmdZikko.Enabled = True  '「初期化実行」釦押下可
'              cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
'              Exit Sub
'           End If
'
'           'IDUシステム初期化
'           For i = 2 To 8
'               bChk(i) = True
'           Next
'           bChk(1) = False
'           If sFileDelete(stsIDU, PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE) = False Then
'              '「一括システム初期化画面：システム初期化処理異常」ログ出力
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "初期化に失敗しました"
'              cmdZikko.Enabled = True  '「初期化実行」釦押下可
'              cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
'              Exit Sub
'           End If
'
'           'LDUシステム初期化
'           For i = 2 To 9
'               bChk(i) = True
'           Next
'           bChk(1) = False
'           If sFileDelete(stsLDU, PATH_LDU_APP & PATH_LDU_DATA & PATH_LDU_SYSTEMFILE) = False Then
'              '「一括システム初期化画面：システム初期化処理異常」ログ出力
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "初期化に失敗しました"
'              cmdZikko.Enabled = True  '「初期化実行」釦押下可
'              cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
'              Exit Sub
'           End If
'
'           '監視盤：一件明細
'           Me.LstStatus.AddItem "DB初期化:" & "集計関連データ"
'           DoEvents
'           iTargetDB = stsKansiMeisai
'           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'           DoEvents
'           Me.Refresh
'           If bKansiDB_Code = True Then
'                '監視盤：別集札
'                iTargetDB = stsKansiBetu
'                '監視盤DB初期化処理
'                bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'
'           'IDUDB初期化処理
'            Me.LstStatus.AddItem "DB初期化:" & "DBデータ"
'            DoEvents
'            Me.Refresh
'
'           'IDUDB初期化処理
'           'IDU:DBデータ
'           iTargetDB = stsIDUMeisai
'           bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'           DoEvents
'           Me.Refresh
'           If bIDUDB_Code = True Then
'                Me.LstStatus.AddItem "DB初期化:" & "アプリケーションログ"
'                DoEvents
'                Me.Refresh
'                'IDU：アプリケーションログ
'                iTargetDB = stsIDUAPLlog
'                'IDU：アプリDB初期化処理
'               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'               DoEvents
'               Me.Refresh
'           End If
'           If bIDUDB_Code = True Then
'                Me.LstStatus.AddItem "DB初期化:" & "保守プログラム"
'                DoEvents
'                Me.Refresh
'                'IDU：保守ログ
'                iTargetDB = stsIDUMentelog
'                'IDU：保守DB初期化処理
'                bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'           If bIDUDB_Code = True Then
'                Me.LstStatus.AddItem "DB初期化:" & "判定ICモジュールログ"
'                DoEvents
'                Me.Refresh
'                'IDU：判定IC-Mモジュールログ
'                iTargetDB = stsIDUICM
'                'IDU：判定IC-MDB初期化処理
'                bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'           If bIDUDB_Code = True Then
'                'IDU：ネガリスト
'                iTargetDB = stsIDUNega
'                'IDU：ネガリストDB初期化処理
'                bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'
'           If bKansiDB_Code = True And bIDUDB_Code = True Then
'                '「一括システム初期化画面：システム初期化処理正常」ログ出力
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
'                lblKekka.ForeColor = SYSFORMAT_OK
'                lblKekka.Caption = "初期化は成功しました"
'           Else
'                '「一括システム初期化画面：DB初期化処理異常」ログ出力
'                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
'                lblKekka.ForeColor = SYSFORMAT_ERROR
'                lblKekka.Caption = "初期化に失敗しました"
'           End If
'       Else
'         '「一括システム初期化画面：システム初期化処理異常」ログ出力
'         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'         lblKekka.ForeColor = SYSFORMAT_ERROR
'         lblKekka.Caption = "初期化に失敗しました"
'       End If
'  End If
'
'  '初期化処理終了
'  cmdZikko.Enabled = True  '「初期化実行」釦押下可
'  cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
'V1.5.0.1 DEL END
Exit Sub

'V1.7.0.1 ADD START
ERR_SPACE3:
 '「キャンセル釦押下」
 '「一括システム初期化画面：初期化処理未実行」ログ出力
 Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
 cmdZikko.Enabled = True  '「初期化実行」釦押下可
 cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
 '処理を抜ける
 Exit Sub
'V1.7.0.1 ADD END

ERR_SPACE2:
  'エラー発生時の処理
  cmdZikko.Enabled = True  '「初期化実行」釦押下可
  cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
  '「一括システム初期化画面：システム初期化処理異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
  lblKekka.ForeColor = SYSFORMAT_ERROR
  lblKekka.Caption = "初期化に失敗しました"
ERR_SPACE:

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdCancel_Click
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
Private Sub cmdCancel_Click()
    On Error Resume Next

    '「一括システム初期化画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ALL_SYSFORMAT_GAMEN_END, 0)
    frmALLSysformat.ZOrder
    Unload Me
End Sub

'V1.4.0.1　ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sCreateShokiFile
'//  機能名称  : 保存ファイルを作成する。
'//  機能概要  : 各設定ファイルの保存用を作成する。
'//
'//              型        名称        意味
'//   引数     :なし
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         フェーズ２対応　設定ファイル(保存用）を追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sCreateShokiFile() As Boolean

   Dim NameChk As String        'ファイル有無チェック戻り値
   Dim lngErrCode As Long       'エラーコード
    
    sCreateShokiFile = False
    
    On Error GoTo ERR_SPACE
        
    '//////////////////////////////////////////////
    '自改設定、監視設定の保存用ファイルを作成する。
    '//////////////////////////////////////////////
    '自改設定ファイル有無チェック
    NameChk = Dir(G_SETTEI_FILE, vbNormal)
    If NameChk <> "" Then
       FileCopy G_SETTEI_FILE, SHOKI_G_SETTEI_FILE
    End If
    
    '監視設定ファイル有無チェック
    NameChk = Dir(K_SETTEI_FILE, vbNormal)
    If NameChk <> "" Then
       FileCopy K_SETTEI_FILE, SHOKI_K_SETTEI_FILE
    End If
    
    '///////////////////////////////////////////////////////////
    'IDU縮退チェック＆IDUファイル関連の保存用ファイルを作成する。
    '///////////////////////////////////////////////////////////
    'ファイル有無チェック
    If pbIDUSts = 1 Then
       sCreateShokiFile = True
       '「一括システム初期化画面：保存用設定ファイル作成正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
       Exit Function
    End If
    
   'IC_M設定ファイル有無チェック
    NameChk = Dir(PATH_IDU_APP & PATH_ICM_SETTEI, vbNormal)
    If NameChk <> "" Then
       FileCopy PATH_IDU_APP & PATH_ICM_SETTEI, PATH_IDU_APP & PATH_SHOKI_ICM_SETTEI
    End If
    
    'ID中継ユニット設定ファイル有無チェック
    NameChk = Dir(PATH_IDU_APP & PATH_IDU_SETTEI, vbNormal)
    If NameChk <> "" Then
       FileCopy PATH_IDU_APP & PATH_IDU_SETTEI, PATH_IDU_APP & PATH_SHOKI_IDU_SETTEI
    End If

    sCreateShokiFile = True
    '「一括システム初期化画面：保存用設定ファイル作成正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
    
    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「一括システム初期化画面：保存用設定ファイル作成異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SHOKI_CREATE_ERROR, lngErrCode)
    sCreateShokiFile = False
End Function
'V1.4.0.1　ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSysFileDelete
'//  機能名称  : システムファイル削除処理
'//  機能概要  : イベントログ、ワトソンログ、メモリダンプファイルを削除する
'//
'//              型        名称        意味
'//   引数     :なし
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 リストボックスのスクロール処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sSysFileDelete()
   Dim iRet As Integer           '削除処理戻り値
    Dim NameChk As String        'ファイル有無チェック戻り値
    Dim lhEventLog As Long       'イベントログのハンドル。
    Dim lReturn As Long          '関数戻り値
    Dim fs As Object
    Dim lngErrCode As Long       'エラーコード
    
    sSysFileDelete = False
    
    On Err GoTo ERR_SPACE
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    '/////////////////////////////
    'メモリダンプファイルの削除
    '/////////////////////////////
    'ファイル有無チェック
    NameChk = Dir(PATH_INS & MEMORYLOG, vbNormal)
    If NameChk <> "" Then
       iRet = fs.DeleteFile(PATH_INS & MEMORYLOG)
       If iRet <> 0 Then
           GoTo ERR_SPACE
       End If
       LstStatus.AddItem "削除したファイル - " & PATH_INS & MEMORYLOG
       LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD

    End If
    
    '/////////////////////////////
    'ワトソンログファイルの削除
    '/////////////////////////////
    'ファイル有無チェック
    NameChk = Dir(SYSDRWATSON_LOG, vbNormal)
    If NameChk <> "" Then
       iRet = fs.DeleteFile(SYSDRWATSON_LOG)
       If iRet <> 0 Then
          GoTo ERR_SPACE
       End If
       LstStatus.AddItem "削除したファイル - " & SYSDRWATSON_LOG
       LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
    End If
    
    Set fs = Nothing
    
    '/////////////////////////////
    'イベントログのクリア
    '/////////////////////////////
    ' イベントログ（アプリケーション）をクリアする。
    lhEventLog = OpenEventLog(vbNullString, "Application")
    lReturn = ClearEventLog(lhEventLog, vbNullString)
    lReturn = CloseEventLog(lhEventLog)

    ' イベントログ（システム）をクリアする。
    lhEventLog = OpenEventLog(vbNullString, "System")
    lReturn = ClearEventLog(lhEventLog, vbNullString)
    lReturn = CloseEventLog(lhEventLog)

    ' イベントログ（セキュリティ）をクリアする。
    lhEventLog = OpenEventLog(vbNullString, "Security")
    lReturn = ClearEventLog(lhEventLog, vbNullString)
    lReturn = CloseEventLog(lhEventLog)

    sSysFileDelete = True
    
    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「一括システム初期化画面：システムファイル削除異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SYSFILE_DELETE_ERROR, lngErrCode)
    Set fs = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sFileDelete
'//  機能名称  : ファイル・フォルダ削除処理
'//  機能概要  : 削除対象ファイル、削除対象フォルダの削除を行う。
'//
'//              型        名称        意味
'//   引数     :なし
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//             　　フェーズ１不具合対応　「DoEvents」にて画面の描写を行う。
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 リストボックスのスクロール処理追加
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-19  REVISED BY [TCC] M.Matsumoto
'//                 【統-313対応】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-240】
'//     REVISIONS :(EG20 V5.3.0.1) 2012-03-16  CODED BY  [TCC] H.Sugimoto
'//                 EG20【5002P2 TR-19】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sFileDelete(iKikitype As Integer, Sys_FilePath As String)
    Dim iFileNo As Integer
    Dim sFileData As String
    Dim iMozi, iKbn As Integer
    Dim sShubetu As String
    Dim sRoot As String
    Dim sPass As String
    Dim sType As String
    Dim sKomoku As String
    Dim bSyori As Boolean
    Dim fs As Object
    Dim MyName As String
    Dim i As Integer
    Dim sChkPass As String
    Dim iRet As Integer
    Dim sDeletePath As String    '削除対象フルパス
    Dim lngErrCode As Long       'エラーコード
    Dim lBool As Boolean                      ' EG20 V3.3.0.1【結合TR-240】追加

    sFileDelete = False

    On Error GoTo ERR_SPACE
    
    'ファイル有無チェック
    MyName = Dir(Sys_FilePath, vbNormal)
    If MyName = "" Then
        GoTo ERR_SPACE
    End If

' EG20 V3.3.0.1【結合TR-240】追加開始（位置移動）
    ' 保守ログファイルCLOSE
    lBool = dllCloseHoshuLogFile()
' EG20 V3.3.0.1【結合TR-240】追加終了（位置移動）

    iFileNo = FreeFile                                              '未使用のファイル番号を取得する。
    Open Sys_FilePath For Input As #iFileNo                         'システム初期化設定ファイルを開く。
    Line Input #iFileNo, sFileData                                  ' １行目は全体バージョンなので読飛ばす。
    Do While Not EOF(iFileNo)
        Line Input #iFileNo, sFileData                              ' １行分読込む。
        sFileData = Trim(sFileData)
        'データがなければ
        If Len(sFileData) = 0 Then
            Exit Do
        End If

        '作業用変数の初期化
        iMozi = 1
        iKbn = 1
        bSyori = False

        'ファイル内容取得
        Do
            If Mid(sFileData, iMozi, 1) = "," Or iMozi = Len(sFileData) Then
                Select Case iKbn
                    '種別
                    Case 1
                        sShubetu = Trim(Left(sFileData, iMozi - 1))
                        If sShubetu <> "2" And sShubetu <> "3" Then
                            Exit Do
                        End If
                    'ルートフォルダ
                    Case 2
                         sRoot = Trim(Left(sFileData, iMozi - 1))
                    'パス
                    Case 3
                         sPass = Trim(Left(sFileData, iMozi - 1))
                    '項目
                    Case 4
                        sKomoku = Trim(sFileData)
                        If bChk(Int(sKomoku)) = False Then
                           Exit Do
                        End If
                        bSyori = True
                        Exit Do
                End Select
                sFileData = Trim(Mid(sFileData, iMozi + 1))
                iMozi = 0
                iKbn = iKbn + 1
            End If
            iMozi = iMozi + 1
        Loop

        '取得データの処理の有無
        If bSyori = True Then
            
            'パスの取得
            Select Case iKikitype
                Case stsKansi
                     Select Case sRoot
                     Case 1  'アプリルート
                        sPass = PATH_KANSI & sPass
                     Case 2  'バックアップルート
                       If sPass = "" Then
                          sPass = Mid(PATH_FKANSI, 1, Len(PATH_FKANSI) - 2)
                       Else
                           sPass = PATH_FKANSI & sPass
                       End If
                     Case 4  'ログルート
                        sPass = PATH_EKANSI & sPass
' EG20 V5.3.0.1追加開始
                     Case 5  ' パス指定無し（フルパス）
                        ' パス種別の明示化 sPass = sPass
' EG20 V5.3.0.1追加終了
                     End Select
                
                Case stsIDU
                     Select Case sRoot
                        Case 1
                          'アプリルート
                          sPass = PATH_IDU_APP & "\\" & sPass
                        Case 2
                          'バックアップルート
                          sPass = PATH_BUC & "\\" & sPass
                        Case 4
                          'ログルート
                          sPass = PATH_IDU_LOG & "\\" & sPass
                     End Select
                     
                Case stsLDU
                   Select Case sRoot
                      Case 1
                        'アプリルート
                        sPass = PATH_LDU_APP & "\\" & sPass
                      Case 4
                        'ログルート
                         sPass = PATH_LDU_LOG & "\\" & sPass
                    End Select
             End Select
                    
            'ファイル有無チェック
            If sShubetu = 3 Then
                MyName = Dir(sPass, vbDirectory)
            Else
                MyName = Dir(sPass, vbNormal)
            End If

            '処理実行
            If MyName <> "" Then
                Set fs = CreateObject("Scripting.FileSystemObject")
                  Select Case sShubetu
                      'ファイル削除
                      Case 2:
                           iRet = fs.DeleteFile(sPass)
                          If iRet <> 0 Then
                              GoTo ERR_SPACE
                          End If
                          LstStatus.AddItem "削除したファイル - " & sPass
                          DoEvents      'V1.5.0.1　ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
                      'フォルダの削除／作成
                      Case 3:
                          fs.DeleteFolder (sPass), True
                          fs.CreateFolder (sPass)
                          LstStatus.AddItem "削除／作成したフォルダ - " & sPass
                          DoEvents      'V1.5.0.1　ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                  End Select
                'オブジェクト解放
                Set fs = Nothing
            Else
                '指定ＰＡＳＳナシ
                Select Case sShubetu
                   Case 2:
                       LstStatus.AddItem "指定ファイルなし - " & sPass
                       DoEvents      'V1.5.0.1　ADD
                       LstStatus.Selected(LstStatus.ListCount - 1) = True           'V1.12.0.1 ADD
                   Case 3:
                       Set fs = CreateObject("Scripting.FileSystemObject")
                       'ファイル有無チェック
'                       For i = 0 To Len(sPass)          'EG20 V2.1.0.1 DEL 【統-313対応】
                       For i = 0 To Len(sPass) - 1       'EG20 V2.1.0.1 ADD 【統-313対応】
                           If Mid(sPass, Len(sPass) - i, 1) = "\" Then
                               sChkPass = Left(sPass, Len(sPass) - i - 1)
                               Exit For
                           End If
                       Next
                       MyName = Dir(sChkPass, vbDirectory)
                       If MyName = "" Then
                           LstStatus.AddItem "フォルダ作成失敗 - " & sPass
                          DoEvents      'V1.5.0.1　ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                       Else
                           fs.CreateFolder (sPass)
                           LstStatus.AddItem "作成したフォルダ - " & sPass
                          DoEvents      'V1.5.0.1　ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                    End If
                       'オブジェクト解放
                       Set fs = Nothing
                End Select
            End If
        End If
    Loop
    Close #iFileNo

    sFileDelete = True
    
    Exit Function

ERR_SPACE:
    'V1.21.0.1 ADD  START
     If iFileNo > 0 Then
        Close #iFileNo
    End If
    'V1.21.0.1 ADD  END
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「一括システム初期化画面：ファイル・フォルダ初期化異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TARGET_FILE_FOLDER_DELETE_ERROR, lngErrCode)
    Set fs = Nothing
End Function

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmALLSysformat.Caption, False
        pfFormActive (frmALLSysformat.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
'V1.5.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrAplTimer_Timer
'//  機能名称  : アプリ起動チェックタイマ、タイムアップ処理
'//  機能概要  : タイムアップ毎にアプリ起動状態をチェックする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplTimer_Timer()
  'V1.20.0.1 ADD START
  Dim bLDURet As Boolean  'LDUログフラグ
  Dim bIDURet As Boolean  'IDUログフラグ
  'V1.20.0.1 ADD END
  
  On Error Resume Next

  '待ち時間がINI定義を超えたかどうかチェック
  If lngtime <= lngMAX_Time Then
    'アプリ起動チェックを行う。全アプリが終了したときのみ、初期化処理を行う。
    'If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then　'V1.20.0.1 DEL
    If CheckAppStart(PROC_KANRI) = 0 Then 'V1.20.0.1 ADD
      'アプリ起動チェックタイマを停止する。
      tmrAplTimer.Enabled = False
      'V1.20.0.1 DEL START
'      '初期化処理
'      DeleteFile_Folder
      'V1.20.0.1 DEL END
      'V1.20.0.1  ADD START
      If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
         bIDURet = EndIDULog 'IDUログ起動時はIDUログに対してログ終了要求CMD送信
      Else
         bIDURet = True
      End If
      If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
         bLDURet = EndLDULog  'LDUログ起動時はLDUログに対してログ終了要求CMD送信
      Else
         bLDURet = True
      End If
      
      If bIDURet = True And bLDURet = True Then
         lngtime = 0
         lngtime = MN_MAIL_INTERVAL
         tmrLogTimer.Enabled = True
      Else
         '「一括システム初期化画面：システム初期化処理異常」ログ出力
         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "初期化に失敗しました"
         cmdZikko.Enabled = True
         cmdCancel.Enabled = True
         Exit Sub
      End If
      'V1.20.0.1  ADD END
    Else
    '起動アプリ有りの場合、タイマを張り直す
      tmrAplTimer.Interval = MN_MAIL_INTERVAL
    '合計経過待ち時間をアップ
     lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI定義値を超えた場合、初期化処理異常とする。
    '「一括システム初期化画面：システム初期化処理異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "初期化に失敗しました"
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
    'アプリ起動チェックタイマを停止する。
    tmrAplTimer.Enabled = False
  End If
End Sub
'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrLogTimer_Timer
'//  機能名称  : ログ起動チェックタイマ、タイムアップ処理
'//  機能概要  : タイムアップ毎にログ起動状態をチェックする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加、確認ポップアップ修正
'//    REVISIONS :(EG20 V2.0.1.1) 2011-11-23  REVISED BY [TCC] T.Koyama
'//               ＥＧ２０フェーズ２対応【残件№54】
'//               ・保守ログファイルＣＬＯＳＥ処理追加
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-240】
'//    REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()

  Dim lBool As Boolean                      ' EG20 V2.0.1.1【残件№54】ADD

  On Error Resume Next

  '待ち時間がINI定義を超えたかどうかチェック
  If lngtime <= lngLogMAX_Time Then
    'ログ起動チェックを行う。全て終了したときのみ、初期化処理を行う。
    If CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
      'ログ起動チェックタイマを停止する。
      tmrLogTimer.Enabled = False
      
' EG20 V3.3.0.1【結合TR-240】削除開始（位置移動）
'      ' EG20 V2.0.1.1【残件№54】ADD START
'      ' 保守ログファイルCLOSE
'       lBool = dllCloseHoshuLogFile()
'      ' EG20 V2.0.1.1【残件№54】ADD START
' EG20 V3.3.0.1【結合TR-240】削除終了（位置移動）
      
      '初期化処理
      DeleteFile_Folder
    Else
    '起動ログ有り有りの場合、タイマを張り直す
      tmrLogTimer.Interval = MN_MAIL_INTERVAL
    '合計経過待ち時間をアップ
     lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI定義値を超えた場合、初期化処理異常とする。
    '「一括システム初期化画面：システム初期化処理異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "初期化に失敗しました"
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
    'ログ起動チェックタイマを停止する。
    tmrLogTimer.Enabled = False
  End If
End Sub
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : DeleteFile_Folder
'//  機能名称  : ファイル、フォルダ、DB初期化処理
'//  機能概要  : 初期化処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 リストボックスのスクロール処理追加
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-240】
'//     REVISIONS :(EG20 V7.5.0.1) 2013-12-07  CODED BY  [TCC] H.Kondoh
'//                 最大接続試験影響範囲確認不具合対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub DeleteFile_Folder()
    
    Dim iRet As Integer
    Dim sDBFormat As String
    Dim sLine As String
    Dim sExecName As String
    Dim sDbInitCmd As String
    Dim bRtn1 As Boolean
    Dim lngErrCode As Long              'エラーコード
    Dim bRet As Boolean
    Dim bKansiDB_Code As Boolean
    Dim bIDUDB_Code As Boolean
    Dim lExitCode As Long
    Dim iTargetDB As Integer            '対象DB値
    ReDim bChk(9)
    Dim i As Integer                    'カウンター
    Dim bRtn As Boolean
    'EG20 V2.1.0.1 ADD START 【統-313対応】
    Dim intLoop As Integer
    Dim lSts As Long
    'EG20 V2.1.0.1 ADD END
         
    
    On Error GoTo ERR_SPACE
   
    '監視盤、IDUアプリ、各設定ファイル(保存用)作成処理
    If sCreateShokiFile = False Then
       '「一括システム初期化画面：システム初期化処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "初期化に失敗しました"
        cmdZikko.Enabled = True
        cmdCancel.Enabled = True
        '処理を抜ける
         Exit Sub
    End If
   
    'システムファイルの削除処理
    bRtn1 = sSysFileDelete()
     
    'システムファイル削除処理成功した場合、
    'フォルダ、ファイルの削除処理を行う
    If bRtn1 = True Then

      '監視盤システム初期化
      For i = 1 To 6
          bChk(i) = True
      Next

'      bChk(5) = False                  ' EG20 V3.3.0.1【結合TR-240】削除
      bChk(6) = False                   ' EG20 V3.3.0.1【結合TR-240】追加

      If sFileDelete(stsKansi, KANSI_SYSTEMFILE) = False Then
         '「一括システム初期化画面：システム初期化処理異常」ログ出力
         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "初期化に失敗しました"
         cmdZikko.Enabled = True  '「初期化実行」釦押下可
         cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
         Exit Sub
       End If
           
       'IDUシステム初期化
       For i = 2 To 8
           bChk(i) = True
       Next

       bChk(1) = False

       If sFileDelete(stsIDU, PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE) = False Then
          '「一括システム初期化画面：システム初期化処理異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "初期化に失敗しました"
          cmdZikko.Enabled = True  '「初期化実行」釦押下可
          cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
          Exit Sub
       End If
           
       'LDUシステム初期化
       For i = 2 To 9
           bChk(i) = True
       Next

       bChk(1) = False

       If sFileDelete(stsLDU, PATH_LDU_APP & PATH_LDU_DATA & PATH_LDU_SYSTEMFILE) = False Then
          '「一括システム初期化画面：システム初期化処理異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "初期化に失敗しました"
          cmdZikko.Enabled = True  '「初期化実行」釦押下可
          cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
          Exit Sub
       End If
 
       '監視盤：一件明細
        Me.LstStatus.AddItem "DB初期化:" & "集計関連データ"
        DoEvents
        LstStatus.Selected(LstStatus.ListCount - 1) = True              'V1.12.0.1 ADD
        
        '監視盤：一件明細
        Me.LstStatus.AddItem "一件明細コーナ１　DB初期化開始"
        DoEvents
        iTargetDB = stsKansiMeisai
        bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
        Me.LstStatus.AddItem "一件明細コーナ１　DB初期化終了"
        DoEvents
        
        If bKansiDB_Code = True Then
           '監視盤：一件明細（コーナ２）
           Me.LstStatus.AddItem "一件明細コーナ２　DB初期化開始"
           DoEvents
           iTargetDB = stsKansiMeisai2
           'DB初期化処理
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "一件明細コーナ２　DB初期化終了"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '監視盤：一件明細（コーナ３）
           Me.LstStatus.AddItem "一件明細コーナ３　DB初期化開始"
           DoEvents
           iTargetDB = stsKansiMeisai3
           'DB初期化処理
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "一件明細コーナ３　DB初期化終了"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '監視盤：一件明細（コーナ４）
           Me.LstStatus.AddItem "一件明細コーナ４　DB初期化開始"
           DoEvents
'           bKansiDB_Code = stsKansiMeisai4     'EG20 V7.5.0.1 DEL
           iTargetDB = stsKansiMeisai4          'EG20 V7.5.0.1 ADD
           'DB初期化処理
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "一件明細コーナ４　DB初期化終了"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '監視盤：一件明細（コーナ５）
           Me.LstStatus.AddItem "一件明細コーナ５　DB初期化開始"
           DoEvents
           iTargetDB = stsKansiMeisai5
           'DB初期化処理
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "一件明細コーナ５　DB初期化終了"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '監視盤：一件明細（コーナ６）
           Me.LstStatus.AddItem "一件明細コーナ６　DB初期化開始"
           DoEvents
           iTargetDB = stsKansiMeisai6
           'DB初期化処理
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "一件明細コーナ６　DB初期化終了"
           DoEvents
        End If
            
        If bKansiDB_Code = True Then
           '監視盤：別集札
           Me.LstStatus.AddItem "別集札　DB初期化開始"
           DoEvents
           iTargetDB = stsKansiBetu
           '監視盤DB初期化処理
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "別集札　DB初期化終了"
           DoEvents
        End If
           
        'EG20 V2.1.0.1 ADD START 【統-313 START】
        For intLoop = 1 To 6
            If intLoop = 1 Then
                lSts = WritePrivateProfileString(SHKEI_EKITUDO_INI_SECTION, _
                       SHKEI_EKITUDO_INI_CNGFLG_KEY, "1", SHUKEI_EKITUDO_FILE)
            Else
                lSts = WritePrivateProfileString(SHKEI_EKITUDO_INI_SECTION & CStr(intLoop), _
                       SHKEI_EKITUDO_INI_CNGFLG_KEY, "1", SHUKEI_EKITUDO_FILE)
            End If
        Next intLoop
        'EG20 V2.1.0.1 ADD END
            
        bIDUDB_Code = False
        
        If bKansiDB_Code = True Then
            'IDUDB初期化処理
            Me.LstStatus.AddItem "DB初期化:" & "DBデータ"
            DoEvents
            LstStatus.Selected(LstStatus.ListCount - 1) = True          'V1.12.0.1 ADD
           
            'IDUDB初期化処理
            'IDU:DBデータ
            iTargetDB = stsIDUMeisai
            bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
            DoEvents
            If bIDUDB_Code = True Then
               Me.LstStatus.AddItem "DB初期化:" & "アプリケーションログ"
               DoEvents
               LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
               'IDU：アプリケーションログ
               iTargetDB = stsIDUAPLlog
               'IDU：アプリDB初期化処理
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
            If bIDUDB_Code = True Then
               Me.LstStatus.AddItem "DB初期化:" & "保守プログラム"
               DoEvents
               LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
               'IDU：保守ログ
               iTargetDB = stsIDUMentelog
               'IDU：保守DB初期化処理
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
            If bIDUDB_Code = True Then
               Me.LstStatus.AddItem "DB初期化:" & "判定ICモジュールログ"
               DoEvents
               LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
               'IDU：判定IC-Mモジュールログ
               iTargetDB = stsIDUICM
               'IDU：判定IC-MDB初期化処理
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
            If bIDUDB_Code = True Then
               'IDU：ネガリスト
               iTargetDB = stsIDUNega
               'IDU：ネガリストDB初期化処理
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
        End If
        If bKansiDB_Code = True And bIDUDB_Code = True Then
           '「一括システム初期化画面：システム初期化処理正常」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
           lblKekka.ForeColor = SYSFORMAT_OK
           lblKekka.Caption = "初期化は成功しました"
        Else
           '「一括システム初期化画面：DB初期化処理異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
           lblKekka.ForeColor = SYSFORMAT_ERROR
           lblKekka.Caption = "初期化に失敗しました"
        End If
    Else
     '「一括システム初期化画面：システム初期化処理異常」ログ出力
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
     lblKekka.ForeColor = SYSFORMAT_ERROR
     lblKekka.Caption = "初期化に失敗しました"
    End If
 
  '初期化処理終了
  cmdZikko.Enabled = True  '「初期化実行」釦押下可
  cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
  
Exit Sub

ERR_SPACE2:
  'エラー発生時の処理
  cmdZikko.Enabled = True  '「初期化実行」釦押下可
  cmdCancel.Enabled = True '「メニュー画面へ戻る」釦押下可
  '「一括システム初期化画面：システム初期化処理異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
  lblKekka.ForeColor = SYSFORMAT_ERROR
  lblKekka.Caption = "初期化に失敗しました"
ERR_SPACE:

End Sub
'V1.5.0.1 ADD　END
