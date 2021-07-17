VERSION 5.00
Begin VB.Form frmHoshu 
   BorderStyle     =   0  'なし
   Caption         =   "保守"
   ClientHeight    =   9000
   ClientLeft      =   555
   ClientTop       =   2325
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   9.75
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
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer tmrMail 
      Left            =   1200
      Top             =   7680
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   19
      Left            =   9000
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   18
      Left            =   9000
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   17
      Left            =   9000
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   16
      Left            =   9000
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   15
      Left            =   9000
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "15.稼働Ver一覧表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   14
      Left            =   6120
      TabIndex        =   16
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "14.TOMASデータ管理"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   13
      Left            =   6120
      TabIndex        =   15
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "13.システム設定　 "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   12
      Left            =   6120
      TabIndex        =   14
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "12.ＬＤＵ業務    "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   11
      Left            =   6120
      TabIndex        =   13
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "11.ＩＤＵ業務    "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   10
      Left            =   6120
      TabIndex        =   12
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "9.パスワード設定  "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   3240
      TabIndex        =   11
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "10.ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ    "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   3240
      TabIndex        =   10
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5.通信確認・表示 "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8.アプリ起動・終了 "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   3240
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1.バージョン管理 "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2.機器情報設定   "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6.システム初期化　"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7.ﾕｰﾃｨﾘﾃｨ起動    "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3.ログ管理       "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "監視画面へ戻る"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      TabIndex        =   0
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4.データ収集・出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "メンテナンスメニュー"
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
End
Attribute VB_Name = "frmHoshu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmHoshu.frm
'//  パッケージ名：メンテナンスメニュー画面
'//
'//  概要：メンテナンスメニュー画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-12   REVISED   BY [TCC] C.Terui
'//                 ・管理プロセスが起動していない場合の
'//                   「監視画面に戻る」ボタン押下処理
'//                 ・保守画面がアンロードされた時のイベント追加
'//     REVISIONS :(1.6.0.1) 2009-04-11   REVISED   BY [TCC] S.Terao
'//                 ・フェーズ３対応　監視盤設定釦追加
'//     REVISIONS :(1.10.0.1)2009-09-25   REVISED   BY [TCC] T.Furuya
'//                 ・KK対応
'//     REVISIONS :(2.2.0.1)  2010-09-11  REVISED BY [TCC] S.Terao
'//                 ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

Dim iGamenType As Integer '取得画面タイプ

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : メンテナンスメニュー画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.10.0.1)2009-09-25   REVISED   BY [TCC] T.Furuya
'//                 ・KK対応　戻る釦文言変更
'//     REVISIONS :(EG20 V6.2.0.1) 2012-06-15  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    Dim iKansiAplChk As Integer     'アプリ起動チェック戻り値用　'V1.10.0.1 ADD
    
    On Error Resume Next
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
    
' EG20 V6.2.0.1 追加開始
    If pubGetTomasFunction() = True Then
        cmdMenue(13).Enabled = True
    Else
        cmdMenue(13).Enabled = False
    End If
' EG20 V6.2.0.1 追加終了
    
'V1.10.0.1 ADD START
    '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '監視盤起動時：戻る釦の文言「監視画面へ戻る」
        cmdReturn.Caption = "監視画面へ戻る"
    Else
        '監視未起動時：戻る釦の文言「Windowsに戻る」
        cmdReturn.Caption = "Windowsに戻る"
    End If
'V1.10.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : メンテナンスメニュー画面(ディアクティブ時)
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
Private Sub Form_Deactivate()
    On Error Resume Next
    
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : メンテナンスメニュー画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.2.0.1) 2010-09-13   REVISED   BY [TCC] S.Terao
'//                 ・ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    On Error Resume Next
   
    '「メンテナンスメニュー画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_GAMEN_START, 0)

    pbAbortFlag = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
       
    'IDU/LDU縮退チェック
    psIDUCheck
    psLDUCheck

    If pbIDUSts = 1 Then
      'IDU業務非表示
       cmdMenue(10).Visible = False
    End If

    If pbLDUSts = 1 Then
      'LDU業務非表示
       cmdMenue(11).Visible = False
    End If
'V2.2.0.1 ADD START
    psUnchin_Dll
    psEki_Type iGamenType
'V2.2.0.1 ADD END

    Call gsGetGateInfo      'EG20 V2.1.0.1 ADD
    
    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「監視画面へ戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                 ・管理が起動していない状態の処理を追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
   On Error Resume Next
  
   '「メンテナンスメニュー画面：消去」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_GAMEN_END, 0)
' V1.3.0.1 ADD START
    If CheckAppStart(PROC_KANRI) = 0 Then
    '管理プロセスが起動していない場合
        psEndHoshuProc
    Else
' V1.3.0.1 ADD END
        '終了処理を行う
        psEndProc
    End If          ' V1.3.0.1 ADD
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdMenue_Click
'//  機能名称  : 各画面遷移釦押下時処理
'//  機能概要  : 釦名称画面に遷移する。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-04-11   REVISED   BY [TCC] S.Terao
'//                 ・フェーズ３対応　監視盤設定釦追加
'//     REVISIONS :(2.2.0.1) 2010-09-13   REVISED   BY [TCC] S.Terao
'//                 ・ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(EG20 V4.1.0.1) 2011-12-27   REVISED   BY [TCC] M.Matsumoto
'//                 ・【フェーズ３TOMAS対応】
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.58修正対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdMenue_Click(Index As Integer)
    Dim bRet As Boolean            'メール送信処理の戻り値
    Dim iResponse As Integer       'メッセージの戻り値
    Dim udtMail As ML_DISP_INF     '画面表示要求
    Dim lngErrCode As Long         'エラーコード
    
    On Error Resume Next
   
    Select Case Index
        Case 0                                 'バージョン管理
           '「メンテナンスメニュー画面：バージョン管理釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_VERSION_BUTTOM, 0)
            Load frmVersion
            frmVersion.Show 1
        Case 1                                 '機器情報設定
           '「メンテナンスメニュー画面：機器情報設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_KIKIINFOSETTEI_BUTTOM, 0)
            Load frmKikiSettei
            frmKikiSettei.Show 1
        Case 2                                 'ログ管理
           '「メンテナンスメニュー画面：ログ管理釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_LOG_BUTTOM, 0)
            Load frmLogMenu
            frmLogMenu.Show 1
        Case 3                                 '自改保守データ
           '「メンテナンスメニュー画面：自改保守データ釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_JIKAIHOSHU_BUTTOM, 0)
            Load frmGateHoshu
            frmGateHoshu.Show 1
        Case 4                                 '通信確認・表示
           '「メンテナンスメニュー画面：通信確認・表示釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_CONECT_BUTTOM, 0)
            Load frmTusinMenu
            frmTusinMenu.Show 1
        Case 5                                 'システム初期化
           '「メンテナンスメニュー画面：システム初期化釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_SYSFORMAT_BUTTOM, 0)
            Load frmSysformatMenu
            frmSysformatMenu.Show 1
        Case 6                                 'ﾕｰﾃｨﾘﾃｨ起動
           '「メンテナンスメニュー画面：ﾕｰﾃｨﾘﾃｨ起動釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_UTILITY_BUTTOM, 0)
            If pbUserLevel = 0 Then
                Load frmUtilityUSR             '一般メンテナンス
                frmUtilityUSR.Show 1
            Else
                Load frmUtility                '特権メンテナンス
                frmUtility.Show 1
            End If
        Case 7                                 'アプリ終了
           '「メンテナンスメニュー画面：アプリ起動・終了釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_APLSTART_END_BUTTOM, 0)
            Load frmAppConfig
            frmAppConfig.Show 1
        Case 8                                 'パスワード設定
           '「メンテナンスメニュー画面：パスワード設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_PASS_SETTEI_BUTTOM, 0)
            Load frmPassSet
            frmPassSet.Show 1
        Case 9                                 'ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ
            '「メンテナンスメニュー画面：ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_RMENTE_BUTTOM, 0)
            Load frmRmenteMenu
            frmRmenteMenu.Show 1
        Case 10                                'ＩＤＵ業務
           '「メンテナンスメニュー画面：IDU業務釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_IDU_BUTTOM, 0)
           '画面表示要求(IDU業務画面)をID制御に送信する
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_IDU_GAMEN
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
               '「メンテナンスメニュー画面：画面表示要求メール送信異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
' EG20 V5.2.0.1【結合TR-No.58修正対応】削除開始
'               '起動失敗ポップアップ表示
'               iResponse = MsgBox("IDU業務釦、定義エラー。" & _
'                                  Chr(vbKeyReturn) & _
'                                  "ID中継ユニット業務画面を起動できません。", _
'                                  vbOKOnly, _
'                                  "画面起動エラー")
' EG20 V5.2.0.1【結合TR-No.58修正対応】削除終了
' EG20 V5.2.0.1【結合TR-No.58修正対応】追加開始
               '起動失敗ポップアップ表示
               iResponse = MsgBox("IDU業務釦、定義エラー。" & _
                                  Chr(vbKeyReturn) & _
                                  "ＩＤＵ業務画面を起動できません。", _
                                  vbOKOnly, _
                                  "画面起動エラー")
' EG20 V5.2.0.1【結合TR-No.58修正対応】追加終了
               Exit Sub
            End If
            '「メンテナンスメニュー画面：画面表示要求メール送信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        Case 11                                'ＬＤＵ業務
            '「メンテナンスメニュー画面：LDU業務釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_LDU_BUTTOM, 0)
            '画面表示要求(LDU業務画面)をLD制御に送信する
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_LDU_GAMEN
            bRet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
              '「メンテナンスメニュー画面：画面表示要求メール送信異常」ログ出力
              lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
               '起動失敗ポップアップ表示
                iResponse = MsgBox("LDU業務釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "LDユーティリティ業務画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
               Exit Sub
            End If
            '「メンテナンスメニュー画面：画面表示要求メール送信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
'EG20 V2.1.0.1 ADD START
        Case 12                                 '監視盤設定
            '「メンテナンスメニュー画面：システム設定釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_SYSTEM_SETTEI_BUTTOM, 0)
            Load frmSystemSetteiMenu
            frmSystemSetteiMenu.Show 1
'EG20 V2.1.0.1 ADD END
'V1.6.0.1 ADD START
        Case 13                                 '監視盤設定
        'EG20 V4.1.0.1 DEL START 【フェーズ３TOMAS対応】
'            '「メンテナンスメニュー画面：監視盤設定釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_KANSIBAN_SETTEI_BUTTOM, 0)
'            Load frmKansiSettei
'            frmKansiSettei.Show 1
        'EG20 V4.1.0.1 DEL END
        'EG20 V4.1.0.1 ADD START 【フェーズ３TOMAS対応】
            '「メンテナンスメニュー画面：TOMASデータ管理釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_TOMAS_DATA_BUTTOM, 0)
            Load frmTomasDataMng
            frmTomasDataMng.Show 1
        'EG20 V4.1.0.1 ADD END
'V1.6.0.1 ADD END
'V2.2.0.1 ADD START
        'EG20 V5.2.0.1 DEL START 【稼働バージョン管理画面追加】
'       Case 14
'           '「メンテナンスメニュー画面：運賃データDLL釦押下」ログ出力
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_UNCHINDATA_DLL_BUTTOM, 0)
'
'           If iGamenType = 1 Then
'              'メトロ画面
'              '運賃データDLL釦押下時
'              Load frmICUnkai_Type1  '運賃データDLL画面をロードする
'              '運賃データDLL画面を表示する
'              frmICUnkai_Type1.Show 1
'           End If
'           If iGamenType <> 1 Then
'              '舞浜・相鉄画面
'              '運賃データDLL釦押下時
'              Load frmICUnkai_Type2  '運賃データDLL画面をロードする
'              '運賃データDLL画面を表示する
'              frmICUnkai_Type2.Show 1
'           End If
        'EG20 V5.2.0.1 DEL END
'V2.2.0.1 ADD END
        'EG20 V5.2.0.1 ADD START 【稼働バージョン管理画面追加】
        Case 14
            '「稼働Ver一覧表示画面釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_KADOVER_BUTTOM, 0)
            gStrCurrentForm = sFormName_KadoVerKanri
            Load frmKadoVerKanri
            frmKadoVerKanri.Show 1
        'EG20 V5.2.0.1 ADD END
    End Select
End Sub
' V1.3.0.1 ADD START
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'     概要      : 保守画面がアンロードされた時のイベントプロシージャ
'     説明      : 保守画面プロセスを終了する。
'
'     ORIGINAL  :(1.3.0.1) '09-03-12   CODED   BY [TCC] C.Terui
'     REVISIONS :(X.X.X.X) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    End   '保守画面プロセスを終了(Exit)する。
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
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmHoshu.Caption, False
        pfFormActive (frmHoshu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
'V2.2.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psUnchin_Dll
'//  機能名称  : 運賃データDLL釦表示／非表示ユーザチェック処理
'//  機能概要  : HOSHU.INIより、磁気運賃対応ユーザであるかどうかチェックする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(2.2.0.1) 2010-09-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psUnchin_Dll()

    Dim iFlag As Integer '取得ユーザフラグ
 
    On Error Resume Next
 
  ' HOSHU.INIより運賃データＤＬＬ対応ユーザフラグを取得する。
  '(デフォルトは非表示)
    iFlag = GetPrivateProfileInt(KANSI_UNCHIN_DLL_SEC, _
                                 KANSI_UNCHIN_DLL_KEY, _
                                 DEFAILT_Int, _
                                 HOSHU_FILE)
     If iFlag = 1 Then
      'フラグが1の場合「磁気運賃」釦は表示
      cmdMenue(14).Visible = True
     End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psEki_Type
'//  機能名称  : 駅都度によって運賃データDLL画面タイプ判断処理
'//  機能概要  : UnchinDLL_Sts.iniより、運賃データDLL画面タイプを判断する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(2.2.0.1) 2010-09-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psEki_Type(iType As Integer)

    Dim iFlag As Integer '取得フラグ
    
    On Error Resume Next
 
  ' UnchinDLL_Sts.iniより運賃データＤＬＬ画面タイプを判断する。
  '(デフォルトは舞浜)
    iFlag = GetPrivateProfileInt(UNCHIN_DLL_STS_SEC, _
                                    UNCHIN_DLL_STS_KEY, _
                                    DEFAILT_Int, _
                                    UNCHIN_DLL_STS_FILE)

    If iFlag = 0 Then
      '取得異常場合「舞浜運賃データDLL」画面を表示
      iFlag = 2
      iType = iFlag
     Else
      iType = iFlag
     End If
End Sub
'V2.2.0.1 ADD END
