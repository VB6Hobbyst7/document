VERSION 5.00
Begin VB.Form frmIDUSysformat 
   BorderStyle     =   0  'なし
   Caption         =   "                                                                    システム初期化機能"
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
      Left            =   11400
      Top             =   6720
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   8640
      Top             =   7920
   End
   Begin VB.Timer tmrMail 
      Left            =   8640
      Top             =   6120
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
      TabIndex        =   18
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ListBox LstStatus 
      Height          =   2985
      Left            =   120
      TabIndex        =   3
      Top             =   5640
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
   Begin VB.Frame frmSentaku 
      Caption         =   "初期化項目指定"
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   11775
      Begin VB.Frame FraKoumoku 
         Height          =   615
         Left            =   1200
         TabIndex        =   31
         Top             =   240
         Width           =   10455
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "全て初期化（ＤＬＬデータ含む）"
            Height          =   375
            Index           =   2
            Left            =   5160
            TabIndex        =   34
            Top             =   200
            Width           =   4215
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "項目選択"
            Height          =   375
            Index           =   1
            Left            =   2640
            TabIndex        =   33
            Top             =   200
            Width           =   1575
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "出荷時初期化"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   200
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   8
         Left            =   8160
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   30
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   7
         Left            =   8160
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   29
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   6
         Left            =   4080
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   28
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   5
         Left            =   4080
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   27
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   4
         Left            =   4080
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   26
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   3
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   25
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   2
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   24
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   23
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   0
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.Frame frmKoumoku 
         Caption         =   "項目"
         Height          =   3855
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   11535
         Begin VB.Frame FraShosai 
            Caption         =   "項目詳細"
            Height          =   1695
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   11295
            Begin VB.Label LblShosai 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1330
               Left            =   100
               TabIndex        =   21
               Top             =   220
               Width           =   11050
            End
         End
         Begin VB.CheckBox chkSonota 
            Caption         =   "その他データ"
            Height          =   255
            Left            =   8880
            TabIndex        =   15
            Top             =   1560
            Value           =   1  'ﾁｪｯｸ
            Width           =   2175
         End
         Begin VB.Frame frmDLL 
            Caption         =   "ＤＬＬデータ"
            Height          =   975
            Left            =   7920
            TabIndex        =   13
            Top             =   360
            Width           =   3135
            Begin VB.CheckBox chkDLL 
               Height          =   375
               Left            =   960
               TabIndex        =   14
               Top             =   360
               Width           =   2055
            End
         End
         Begin VB.Frame frmLog 
            Caption         =   "ログデータ"
            Height          =   1575
            Left            =   3840
            TabIndex        =   9
            Top             =   360
            Width           =   4035
            Begin VB.CheckBox chkLog 
               Caption         =   "判定ＩＣモジュールログ"
               DataField       =   "3"
               Height          =   375
               Index           =   2
               Left            =   960
               TabIndex        =   12
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Width           =   2955
            End
            Begin VB.CheckBox chkLog 
               Caption         =   "保守プログラムログ"
               DataField       =   "2"
               Height          =   375
               Index           =   1
               Left            =   960
               TabIndex        =   11
               Top             =   720
               Value           =   1  'ﾁｪｯｸ
               Width           =   2535
            End
            Begin VB.CheckBox chkLog 
               Caption         =   "アプリケーションログ"
               DataField       =   "1"
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   10
               Top             =   360
               Value           =   1  'ﾁｪｯｸ
               Width           =   2715
            End
         End
         Begin VB.Frame frmMeisai 
            Caption         =   "ＩＣ一件明細"
            Height          =   1575
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   3615
            Begin VB.CheckBox chkMeisai 
               Caption         =   "再送データ"
               Height          =   375
               Index           =   2
               Left            =   960
               TabIndex        =   8
               Top             =   1080
               Value           =   1  'ﾁｪｯｸ
               Width           =   2535
            End
            Begin VB.CheckBox chkMeisai 
               Caption         =   "バックアップデータ"
               Height          =   375
               Index           =   1
               Left            =   960
               TabIndex        =   7
               Top             =   720
               Value           =   1  'ﾁｪｯｸ
               Width           =   2535
            End
            Begin VB.CheckBox chkMeisai 
               Caption         =   "ＤＢデータ"
               DataField       =   "0"
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   6
               Top             =   360
               Value           =   1  'ﾁｪｯｸ
               Width           =   2535
            End
         End
      End
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
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C000&
      Caption         =   "IDUアプリケーションシステム初期化"
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
      TabIndex        =   19
      Top             =   0
      Width           =   12000
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
      TabIndex        =   17
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Caption         =   "初期化結果"
      Height          =   255
      Left            =   8760
      TabIndex        =   16
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "frmIDUSysformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmIDUSysformat.frm
'//  パッケージ名：システム初期化(IDU)画面
'/
'//  概要：システム初期化(IDU)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　保存用設定ファイル作成処理追加
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 リストボックスのスクロール処理追加
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加、確認ポップアップ修正
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'Private bChk() As Boolean              'V1.5.0.1 DEL

'初期化実行フラグ
Private bSysFormat As Boolean

Private ShosaiMoji(0 To 8) As String '詳細文言格納エリア
Private Const SYSMOJI_SIZE = 500
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000     'アプリ起動タイマデフォルト値
Dim lngMAX_Time As Long                 'INI取得設定値
Dim lngtime     As Long                 '現在タイマ値
Private bChk(8) As Boolean
'V1.5.0.1 ADD END
'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000     'ログ起動タイマデフォルト値(30秒)
Dim lngLogMAX_Time As Long                'INI取得設定値(ログ）
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : システム初期化(IDU)画面(アクティブ時)
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
'//  機能名称  : システム初期化(IDU)画面(ディアクティブ時)
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
'//  機能名称  : システム初期化(IDU)画面(ロード時)
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
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//             　　フェーズ２対応　IDU縮退機能チェック追加
'//     REVISIONS  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim ii  As Integer
    
    On Error Resume Next
    
    '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_SYSFORMAT_GAMEN_START, 0)
    
    gStrCurrentForm = sFormName_IDUSys
    
    '「詳細」釦押下文言取得処理
    ShosaiMongonGet

    '初期化
    OptShosai(0).Value = True   '初期化項目指定：詳細釦押下
    LstStatus.Clear             '削除ファイル表示部クリア
    OptKoumoku(0).Value = True  '初期化項目指定：「出荷時初期化」指定有り選択
    chkMeisai(0).Value = 1      'IC一件明細：DBデータチェック有り
    chkMeisai(1).Value = 1      'IC一件明細：バックアップデータチェック有り
    chkMeisai(2).Value = 1      'IC一件明細：再送データチェック有り
    chkLog(0).Value = 1         'ログデータ：アプリケーションログ
    chkLog(1).Value = 1         'ログデータ：保守プログラムログ
    chkLog(2).Value = 1         'ログデータ：判定ICモジュールログ
    chkSonota.Value = 1         'その他データ
    lblKekka.Caption = ""       '初期化実行表示部クリア
    
    frmKoumoku.Enabled = False
    frmMeisai.Enabled = False
    frmLog.Enabled = False
    chkMeisai(0).Enabled = False 'IC一件明細：DBデータ選択不可
    chkMeisai(1).Enabled = False 'IC一件明細：バックアップデータ選択不可
    chkMeisai(2).Enabled = False 'IC一件明細：再送データ選択不可
    chkLog(0).Enabled = False    'ログデータ：アプリケーションログ選択不可
    chkLog(1).Enabled = False    'ログデータ：保守プログラムログ選択不可
    chkLog(2).Enabled = False    'ログデータ：判定ICモジュールログ選択不可
    chkSonota.Enabled = False    'その他データ選択不可
            
    fraKoumoku.BorderStyle = 0
    OptShosai(0).Enabled = True  '初期化項目部：詳細釦押下可能
    OptShosai(0).Value = True    '初期化項目部：詳細釦押下
    For ii = 1 To 8
        OptShosai(ii).Enabled = False  '項目部：詳細釦押下可能
    Next
    
    OptKoumoku(2).Enabled = False
    frmDLL.Enabled = False
    chkDLL.Enabled = False       'DLLデータ選択不可
    'ログインユーザチェック
    If pbUserLevel = 1 Then
        OptKoumoku(2).Enabled = True
        frmDLL.Enabled = True
        chkDLL.Enabled = False
        chkDLL.Value = 1
    Else
        OptKoumoku(2).Enabled = False
        frmDLL.Enabled = False
        chkDLL.Enabled = False
        chkDLL.Value = 0
    End If
    '初期化実行フラグOFF
    bSysFormat = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
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
   'タイマ値設定
   tmrAplTimer.Interval = MN_MAIL_INTERVAL
   tmrAplTimer.Enabled = False
   'V1.5.0.1 ADD END

   'V1.20.0.1 ADD START
   'INIファイルよりログ起動タイマ値を取得
   lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
   '取得値が0の場合、デフォルト値を設定
   If lngLogMAX_Time = 0 Then
      lngLogMAX_Time = LOG_INTERVAL
   End If
   
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
'//  関数名称  : OptKoumoku_Click
'//  機能名称  : ラジオ釦押下時処理
'//  機能概要  : 初期化項目指定部：ラジオ釦押下時処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下ラジオ釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub OptKoumoku_Click(Index As Integer)
    Dim ii As Integer  'カウンター
    
    On Error Resume Next
   
    Select Case Index
        Case 1:        '項目選択時
            OptShosai(0).Enabled = False  '初期化項目指定：詳細釦選択不可
            OptShosai(1).Value = True     '項目指定：DBデータ詳細釦押下
            For ii = 1 To 7
                OptShosai(ii).Enabled = True  '項目指定：詳細釦選択可
            Next
            
            frmKoumoku.Enabled = True
            frmMeisai.Enabled = True
            frmLog.Enabled = True
            chkMeisai(0).Enabled = True  'IC一件明細：DBデータ選択可能
            chkMeisai(1).Enabled = True  'IC一件明細：バックアップデータ選択可能
            chkMeisai(2).Enabled = True  'IC一件明細：再送データ選択可能
            chkLog(0).Enabled = True     'ログデータ：アプリケーションログ選択可能
            chkLog(1).Enabled = True     'ログデータ：保守プログラムログ選択可能
            chkLog(2).Enabled = True     'ログデータ：判定ICモジュールログ選択可能
            chkSonota.Enabled = True     'その他データ選択可能

            'ログインユーザチェック
            If pbUserLevel = 1 Then
                frmDLL.Enabled = True        'DLLデータタ選択可能
                chkDLL.Enabled = True
                OptShosai(8).Enabled = True  'DLLデータ詳細釦押下可能
            End If
            '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：項目選択選択時」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_KOUMOKU, 0)
        Case Else:
            OptShosai(0).Enabled = True          '初期化項目指定：詳細釦選択可能
            OptShosai(0).Value = True            '初期化項目指定：詳細釦押下
            For ii = 1 To 7
                OptShosai(ii).Enabled = False    '項目部：詳細釦選択不可能
                OptShosai(ii).Value = False      '項目部：詳細釦選択未押下
            Next
            frmKoumoku.Enabled = False
            frmMeisai.Enabled = False
            frmLog.Enabled = False
            chkMeisai(0).Enabled = False         'IC一件明細：DBデータ選択不可
            chkMeisai(1).Enabled = False         'IC一件明細：バックアップデータ選択不可
            chkMeisai(2).Enabled = False         'IC一件明細：再送データ選択不可
            chkLog(0).Enabled = False            'ログデータ：アプリケーションログ選択不可
            chkLog(1).Enabled = False            'ログデータ：保守プログラムログ選択不可
            chkLog(2).Enabled = False            'ログデータ：判定ICモジュールログ選択不可
            chkSonota.Enabled = False            'その他データ選択不可

            'ログインユーザチェック
            If pbUserLevel = 1 Then
                frmDLL.Enabled = False           'DLLデータ選択不可
                chkDLL.Enabled = False
                OptShosai(8).Enabled = False     'DLLデータ詳細釦押下不可
            End If
            If Index = 0 Then
               '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：出荷時初期化選択時」ログ出力
               Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_SHUKKA, 0)
            Else
               '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：全て初期化選択時」ログ出力
               Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_ALL, 0)
            End If
        End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdZikko_Click
'//  機能名称  : 初期化実行釦押下処理
'//  機能概要  : 初期化を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　保存用設定ファイル作成処理追加
'//     REVISIONS  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加、確認ポップアップ修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdZikko_Click()
    Dim i As Integer
    Dim bRtn As Boolean
    Dim bSentaku As Boolean
    Dim iRet As Integer
    Dim sLine As String
    Dim lRetVal As Long
    Dim lExitCode As Long
    Dim sExecName As String
    Dim sDbInitCmd As String
    'ReDim bChk(8)                                'V1.5.0.1 DEL
    Dim bRtn1 As Boolean
    Dim bRtn2 As Boolean
    Dim uMail As MAIL_IDU_LDU_APLEND_CMD           'IDUアプリ終了要求
    Dim iRetApp         As Integer                 'IDUアプリ終了フラグ
    Dim iRetLog         As Integer                 'IDUログ終了フラグ
    Dim uIduEndMail As MAIL_IDU_LDU_LOGEND_CMD     'IDUログプロセス終了要求
    Dim lngErrCode As Long                      'エラーコード
    Dim iTargetDB As Integer                       '対象DB値
    Dim bDB_Code As Boolean
    'V1.5.0.1  ADD START
    Dim bIDUPCRet    As Boolean            'IDUアプリ処理結果
    Dim bIDULOGRet   As Boolean            'IDUログ処理結果
    
    bIDUPCRet = False
    bIDULOGRet = False
    'V1.5.0.1  ADD END
    On Error GoTo ERR_SPACE
    
    '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：初期化実行釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_START_BUTTOM, 0)

    '表示の初期化
    LstStatus.Clear
    lblKekka.Caption = ""

    '出荷時初期化選択時
    If OptKoumoku(0).Value = True Then
        For i = 2 To 8
            bChk(i) = True
        Next
        bChk(1) = False
    End If

    '項目選択選択時
    If OptKoumoku(1).Value = True Then
        bSentaku = False
        'ＩＣ一件明細
        'ＤＢデータ
        If chkMeisai(0).Value = 1 Then
            bSentaku = True
            bChk(2) = True
        Else
            bChk(2) = False
        End If
        'バックアップデータ
        If chkMeisai(1).Value = 1 Then
            bSentaku = True
            bChk(3) = True
        Else
            bChk(3) = False
        End If
        '再送データ
        If chkMeisai(2).Value = 1 Then
            bSentaku = True
            bChk(4) = True
        Else
            bChk(4) = False
        End If

        'ログデータ
        'アプリケーションログ
        If chkLog(0).Value = 1 Then
            bSentaku = True
            bChk(5) = True
        Else
            bChk(5) = False
        End If
        '保守プログラムログ
        If chkLog(1).Value = 1 Then
            bSentaku = True
            bChk(6) = True
        Else
            bChk(6) = False
        End If
        '判定ＩＣモジュールログ
        If chkLog(2).Value = 1 Then
            bSentaku = True
            bChk(7) = True
        Else
            bChk(7) = False
        End If

        'その他データ
        If chkSonota.Value = 1 Then
            bSentaku = True
            bChk(8) = True
        Else
            bChk(8) = False
        End If

        'ＤＬＬデータ
        If chkDLL.Value = 1 Then
            bSentaku = True
            bChk(1) = True
        Else
            bChk(1) = False
        End If

        If bSentaku = False Then
            '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
            MsgBox "初期化するデータが選択されていません", vbExclamation, "データ無警告"
            Exit Sub
        End If
    End If

    '全て初期化（ＤＬＬデータ含む）選択時
    If OptKoumoku(2).Value = True Then
        For i = 1 To 8
            bChk(i) = True
        Next
    End If
    
    iRet = MsgBox("初期化処理を行います。よろしいですか？", vbExclamation + vbOKCancel, "初期化確認")
    If iRet = vbOK Then
        OptKoumoku(0).Enabled = False
        OptKoumoku(1).Enabled = False
        cmdZikko.Enabled = False
        cmdCancel.Enabled = False
        'ログインユーザチェック
        If pbUserLevel = 1 Then
         OptKoumoku(2).Enabled = False
        End If
        
        On Error GoTo ERR_SPACE2
    
        '正常で初期化
        iRetApp = 1
        iRetLog = 1

        'アプリ起動チェック
        If CheckAppStart(PROCESS_IDU_PC) = 1 Then
          'V1.20.0.1 DEL START
'          iRet = MsgBox("ID中継ユニットアプリケーションを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認")
'          If iRet = vbOK Then
          'V1.20.0.1 DEL END
             'IDUアプリ終了要求をID制に送信する
              uMail.mlHeader.dwId = ML_ID_IDU_APLEND_CMD
              uMail.mlHeader.dwSize = MlSize.IDUAPLEND_REQ
              uMail.mlHeader.dwProid = RHOSHU_ID
              uMail.mlHeader.dwSubArea = 0
              uMail.dwEndType = ML_ENDTYPE_APLEND
              uMail.dwCMDLevel = ML_CMDLEVEL_TUJYO        'V1.5.0.1 ADD
              'V1.5.0.1 DEL START
              'bRtn = DssSendMail(MAIL_SLOT_IDSEI, Len(uMail), uMail.mlHeader)
              'If bRtn = 0 Then
              'V1.5.0.1 DEL END
              'V1.5.0.1 ADD START
              bIDUPCRet = DssSendMail(MAIL_SLOT_IDSEI, Len(uMail), uMail.mlHeader)
              If bIDUPCRet = 0 Then
              'V1.5.0.1 ADD END
                 '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：メール送信異常」ログ出力
                 lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                 Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lngErrCode)
                 GoTo ERR_SPACE2:
              Else
                 '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：メール送信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                'iRetApp = CheckAppEndComplete(PROCESS_IDU_PC, lExitCode)    'V1.5.0.1 DEL
              End If
     'V1.20.0.1 DEL START
'               'IDUログ終了要求CMD送信
'               'V1.5.0.1 DEL START
'                'bRtn = EndIDULog
'                'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'                bIDULOGRet = EndIDULog
'                If bIDULOGRet = False Then
'               'V1.5.0.1 ADD END
'                  '送信異常
'                  lblKekka.ForeColor = SYSFORMAT_ERROR
'                  lblKekka.Caption = "初期化に失敗しました"
'                  OptKoumoku(0).Enabled = True
'                  OptKoumoku(1).Enabled = True
'                  cmdZikko.Enabled = True
'                  cmdCancel.Enabled = True
'                  'ログインユーザチェック
'                  If pbUserLevel = 1 Then
'                     OptKoumoku(2).Enabled = True
'                  End If
'                  '処理を抜ける
'                  Exit Sub
'                End If
'
'               'IDUログプロセス終了確認
'               'iRetLog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)  'V1.5.0.1 DEL
'          Else
'             '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'              OptKoumoku(0).Enabled = True
'              OptKoumoku(1).Enabled = True
'              cmdZikko.Enabled = True
'              cmdCancel.Enabled = True
'              'ログインユーザチェック
'              If pbUserLevel = 1 Then
'                OptKoumoku(2).Enabled = True
'              End If
'              '処理を抜ける
'              Exit Sub
'          End If
     'V1.20.0.1 DEL END
       Else
       bIDUPCRet = True                                 'V1.5.0.1 ADD
       
         'ログプロセス起動チェック
          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
             'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
             'V1.20.0.1 DEL START
'             iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
'             If iRet = vbOK Then
             'V1.20.0.1 DEL END
                'IDUログ終了要求CMD送信
                'V1.5.0.1 DEL START
                'bRtn = EndIDULog
                'If bRtn = False Then
                'V1.5.0.1 DEL END
                'V1.5.0.1 ADD START
                bIDULOGRet = EndIDULog
                If bIDULOGRet = False Then
                'V1.5.0.1 ADD END
                  '送信異常
                  lblKekka.ForeColor = SYSFORMAT_ERROR
                  lblKekka.Caption = "初期化に失敗しました"
                  OptKoumoku(0).Enabled = True
                  OptKoumoku(1).Enabled = True
                  cmdZikko.Enabled = True
                  cmdCancel.Enabled = True
                  'ログインユーザチェック
                  If pbUserLevel = 1 Then
                     OptKoumoku(2).Enabled = True
                  End If
                  '処理を抜ける
                  Exit Sub
                End If
               
               'IDUログプロセス終了確認
               'iRetLog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)    V1.5.0.1 DEL
         'V1.20.0.1 DEL START
'             Else
'               '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'                OptKoumoku(0).Enabled = True
'                OptKoumoku(1).Enabled = True
'                cmdZikko.Enabled = True
'                cmdCancel.Enabled = True
'                'ログインユーザチェック
'                If pbUserLevel = 1 Then
'                   OptKoumoku(2).Enabled = True
'                End If
'                '処理を抜ける
'                Exit Sub
'             End If
'           'V1.5.0.1 ADD　START
        'V1.20.0.1 DEL END
           Else
            bIDULOGRet = True
           'V1.5.0.1 ADD END
           End If
        End If

       '初期化実行フラグON
        bSysFormat = True
'V1.5.0.1 ADD START
         'IDUアプリ、IDUログのメール送信処理が全て正常だった場合のみ、アプリ起動タイマを起動させ、
         'アプリ起動チェックによりアプリの起動/未起動を判断する。
'         If (bIDUPCRet = True) And (bIDULOGRet = True) Then            'V1.20.0.1 DEL
         If (bIDUPCRet = True) Then                                     'V1.20.0.1 ADD
            lngtime = 0
            lngtime = MN_MAIL_INTERVAL
            tmrAplTimer.Enabled = True
         Else
           'IDUアプリ、IDUログのメール送信にてひとつでも異常があった場合、初期化処理を異常終了とする。
           '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
           OptKoumoku(0).Enabled = True
           OptKoumoku(1).Enabled = True
           cmdZikko.Enabled = True
           cmdCancel.Enabled = True
           'ログインユーザチェック
           If pbUserLevel = 1 Then
              OptKoumoku(2).Enabled = True
           End If
           '処理を抜ける
           Exit Sub
         End If
   End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'       'アプリまたはログプロセスで終了処理に失敗した場合
'       If (iRetApp <> 1) Or (iRetLog <> 1) Then
'         '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
'         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'         lblKekka.ForeColor = SYSFORMAT_ERROR
'         lblKekka.Caption = "初期化に失敗しました"
'         OptKoumoku(0).Enabled = True
'         OptKoumoku(1).Enabled = True
'         cmdZikko.Enabled = True
'         cmdCancel.Enabled = True
'         'ログインユーザチェック
'         If pbUserLevel = 1 Then
'            OptKoumoku(2).Enabled = True
'         End If
'         '処理を抜ける
'         Exit Sub
'       End If
'      'V1.4.0.1 ADD START
'      '出荷時初期化選択時、全て初期化(DLLデータ含)選択時、その他データが初期化対象時
'      If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
'        If sCreateShokiFile = False Then
'           '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
'           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "初期化に失敗しました"
'           OptKoumoku(0).Enabled = True
'           OptKoumoku(1).Enabled = True
'           cmdZikko.Enabled = True
'           cmdCancel.Enabled = True
'           'ログインユーザチェック
'           If pbUserLevel = 1 Then
'              OptKoumoku(2).Enabled = True
'           End If
'           '処理を抜ける
'           Exit Sub
'        End If
'      End If
'      'V1.4.0.1 ADD END
'
'      'システムファイルの削除
'      If bChk(8) = True Then
'         bRtn1 = sSysFileDelete()
'         DoEvents
'      Else
'         bRtn1 = True
'      End If
'
'      'フォルダ、ファイルの削除
'      If bRtn1 = True Then
'
'         If sFileDelete() = True Then
'
'            bDB_Code = True
'
'            'DB初期化処理
'            'DBデータ：IC一件明細
'             If bChk(2) = True Then
'                iTargetDB = chkMeisai(0).DataField
'                Me.LstStatus.AddItem "DB初期化:" & chkMeisai(0).Caption
'                DoEvents
'                bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DBデータ：アプリログ
'             If bChk(5) = True And bDB_Code = True Then
'               iTargetDB = chkLog(0).DataField
'               Me.LstStatus.AddItem "DB初期化:" & chkLog(0).Caption
'               DoEvents
'               bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DBデータ：保守ログ
'             If bChk(6) = True And bDB_Code = True Then
'               iTargetDB = chkLog(1).DataField
'               Me.LstStatus.AddItem "DB初期化:" & chkLog(1).Caption
'               DoEvents
'               bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DBデータ：判定ICモジュールログ
'             If bChk(7) = True And bDB_Code = True Then
'                iTargetDB = chkLog(2).DataField
'                Me.LstStatus.AddItem "DB初期化:" & chkLog(2).Caption
'                DoEvents
'                bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DBデータ：ネガリスト(出荷時初期化のみ有効)
'             If OptKoumoku(1).Value = False And bDB_Code = True Then
'                iTargetDB = stsIDUNega
'                bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'             End If
'
'             '戻り値を見る
'              If bDB_Code = True Then
'                 '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理正常」ログ出力
'                 Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
'                 lblKekka.ForeColor = SYSFORMAT_OK
'                 lblKekka.Caption = "初期化は成功しました"
'              Else
'                 '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：DB初期化処理異常」ログ出力
'                 Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
'                 lblKekka.ForeColor = SYSFORMAT_ERROR
'                 lblKekka.Caption = "初期化に失敗しました"
'              End If
'        Else
'          '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
'          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'          lblKekka.ForeColor = SYSFORMAT_ERROR
'          lblKekka.Caption = "初期化に失敗しました"
'        End If
'    Else
'       '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
'       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'       lblKekka.ForeColor = SYSFORMAT_ERROR
'       lblKekka.Caption = "初期化に失敗しました"
'    End If
'
'    '初期化正常終了時の処理
'    OptKoumoku(0).Enabled = True
'    OptKoumoku(1).Enabled = True
'    cmdZikko.Enabled = True
'    cmdCancel.Enabled = True
'    'ログインユーザチェック
'    If pbUserLevel = 1 Then
'       OptKoumoku(2).Enabled = True
'    End If
' End If
'V1.5.0.1 DEL END
Exit Sub

ERR_SPACE2:
        'エラー発生時の処理
        OptKoumoku(0).Enabled = True    '初期化項目指定：出荷時初期化選択可能
        OptKoumoku(1).Enabled = True    '初期化項目指定：項目選択選択可能
        cmdZikko.Enabled = True         '「初期化実行」釦押下可能
        cmdCancel.Enabled = True        '「メニュー画面へ戻る」釦押下可能
        'ログインユーザチェック
        If pbUserLevel = 1 Then
            OptKoumoku(2).Enabled = True
        End If
        '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
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

   '「ID中継ユニットシステム初期化：消去」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_SYSFORMAT_GAMEN_END, 0)
   frmSysformatMenu.ZOrder
      
   Unload Me
End Sub

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
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　画面更新処理
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 リストボックスのスクロール処理追加
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sFileDelete()
    Dim iFileNo As Integer
    Dim sFileData As String
    Dim iMozi, iKbn As Integer
    Dim sShubetu, sRoot, sPass, sKomoku As String
    Dim bSyori As Boolean
    Dim fs As Object
    Dim MyName As String
    Dim i As Integer
    Dim sChkPass As String
    Dim iRet As Integer
    Dim lngErrCode As Long       'エラーコード

    sFileDelete = False

    On Error GoTo ERR_SPACE
    
    'ファイル有無チェック
    MyName = Dir(PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE, vbNormal)
    If MyName = "" Then
        GoTo ERR_SPACE
    End If
  
    '未使用のファイル番号を取得する。
    iFileNo = FreeFile
    'システム初期化設定ファイルを開く。
    Open PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE For Input As #iFileNo
    ' １行目は全体バージョンなので読飛ばす。
    Line Input #iFileNo, sFileData
    Do While Not EOF(iFileNo)
        '1行分読み込む。
        Line Input #iFileNo, sFileData
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
            Select Case sRoot
                Case 1  'アプリルート
                    sPass = PATH_IDU_APP & "\\" & sPass
                Case 2  'バックアップ
                    sPass = PATH_BUC & "\\" & sPass
                Case 3      '未使用
'                    sPass = PATH_DAT & sPass
                Case 4  'ログルート
                    sPass = PATH_IDU_LOG & "\\" & sPass
                'EG20 V2.0.1.1 ADD START
                Case 5  'ＤＢルート
                    sPass = PATH_IDU_DB & "\\" & sPass
                'EG20 V2.0.1.1 ADD START
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
                          DoEvents          'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                      'フォルダの削除／作成
                      Case 3:
                          fs.DeleteFolder (sPass), True
                          fs.CreateFolder (sPass)
                          LstStatus.AddItem "削除／作成したフォルダ - " & sPass
                          DoEvents          'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                  End Select
                'オブジェクト解放
                Set fs = Nothing
            Else
                '指定ＰＡＳＳナシ
                Select Case sShubetu
                   Case 2:
                       LstStatus.AddItem "指定ファイルなし - " & sPass
                       DoEvents          'V1.5.0.1 ADD
                       LstStatus.Selected(LstStatus.ListCount - 1) = True           'V1.12.0.1 ADD
                   Case 3:
                       Set fs = CreateObject("Scripting.FileSystemObject")
                       'ファイル有無チェック
                       For i = 0 To Len(sPass)
                           If Mid(sPass, Len(sPass) - i, 1) = "\" Then
                               sChkPass = Left(sPass, Len(sPass) - i - 1)
                               Exit For
                           End If
                       Next
                       MyName = Dir(sChkPass, vbDirectory)
                       If MyName = "" Then
                           LstStatus.AddItem "フォルダ作成失敗 - " & sPass
                           DoEvents          'V1.5.0.1 ADD
                           LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
                       Else
                           fs.CreateFolder (sPass)
                           LstStatus.AddItem "作成したフォルダ - " & sPass
                           DoEvents          'V1.5.0.1 ADD
                           LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
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
    '「IDUシステム初期化画面：ファイル・フォルダ初期化異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TARGET_FILE_FOLDER_DELETE_ERROR, lngErrCode)
   'オブジェクト解放
    Set fs = Nothing
End Function

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
    Dim iRet As Integer          '削除処理戻り値
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
    '解放
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
    '「IDUシステム初期化画面：システムファイル削除異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SYSFILE_DELETE_ERROR, lngErrCode)
    Set fs = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : OptShosai_Click
'//  機能名称  : 「詳細」釦押下時処理
'//  機能概要  : 各データに対する詳細釦押下時処理を行う。
'//
'//              型        名称        意味
'//   引数     :Integer　　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub OptShosai_Click(Index As Integer)
   
   '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：詳細釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYS_INFO_BUTTOM, 0)
  
    LblShosai.Caption = ShosaiMoji(Index)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ShosaiMongonGet
'//  機能名称  : 「詳細」釦押下表示文言取得処理
'//  機能概要  : 「詳細」釦押下にて表示する文言をファイルより取得する。
'//
'//              型        名称        意味
'//   引数     :なし
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub ShosaiMongonGet()
   Dim sWork As String                      '作業エリア
   Dim iKey As String                       'キー名
   Dim lSts As Long                         '戻り値
   Dim lngRet As Long          '関数の返り値
   Dim iGate As Integer        '自改INDEX
   Dim j As Integer            'ワークINDEX
   Dim cWork As Byte           'ワークエリア
   Dim sGateData As String * SYSMOJI_SIZE    '１行分ファイル内容取得用
   Dim iFCnt As Integer
   Dim iFLoop As Integer
   Dim iFLoop2 As Integer
   Dim MyName As String
   Dim i As Integer
    
   'ファイル有無チェック
   MyName = Dir(PATH_SYSFORMAT_SHOUSAI_FILE, vbNormal)
   If MyName = "" Then
       sWork = ""
       For i = 0 To 8
        ShosaiMoji(i) = sWork
       Next
       Exit Sub
   End If
    
   For iGate = CNT_MIN To 8
      ' SysFormatShousai.iniより文言を取得する。
       sGateData = ""
       iKey = SYS_KEY_NAME & iGate
       lSts = GetPrivateProfileString(SYS_IDU_SECTION_NAME, _
                                      iKey, _
                                      DEFAILT, _
                                      sGateData, _
                                      Len(sGateData), _
                                      PATH_SYSFORMAT_SHOUSAI_FILE)
      If lSts = 0 Or sGateData = "" Then
         '定義なければ空白
         ShosaiMoji(iGate) = sWork
      ElseIf Len(sGateData) <> 0 Then
         'データの取得
          ReDim sFData(6)
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
           
           For i = 0 To 5
             If i = 0 Then
                 ShosaiMoji(iGate) = sFData(i + 1)
                 ShosaiMoji(iGate) = ShosaiMoji(iGate) & vbCrLf
             Else
                 ShosaiMoji(iGate) = ShosaiMoji(iGate) & sFData(i + 1)
                 ShosaiMoji(iGate) = ShosaiMoji(iGate) & vbCrLf
             End If
           Next
       End If
  Next
End Sub

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
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmIDUSysformat.Caption, False
        pfFormActive (frmIDUSysformat.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sCreateShokiFile() As Boolean

   Dim NameChk As String        'ファイル有無チェック戻り値
   Dim lngErrCode As Long       'エラーコード
    
    sCreateShokiFile = False
    
    On Error GoTo ERR_SPACE
           
    '///////////////////////////////////////////////////////////
    'IDU縮退チェック＆IDUファイル関連の保存用ファイルを作成する。
    '///////////////////////////////////////////////////////////
    'ファイル有無チェック
    If pbIDUSts = 1 Then
       sCreateShokiFile = True
       '「IDUシステム初期化画面：保存用設定ファイル作成正常」ログ出力
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
    '「IDUシステム初期化画面：保存用設定ファイル作成正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
    
    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「IDUシステム初期化画面：保存用設定ファイル作成異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SHOKI_CREATE_ERROR, lngErrCode)
    sCreateShokiFile = False
End Function
'V1.4.0.1　ADD END

'V1.5.0.1　ADD START
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
  
  Dim bIDURet As Boolean  'IDUログフラグ 'V1.20.0.1 ADD

  On Error Resume Next
 
  '待ち時間がINI定義を超えたかどうかチェック
  If lngtime <= lngMAX_Time Then
    'アプリ起動チェックを行う。IDU(アプリ、ログ)が終了したときのみ、初期化処理を行う。
    'If CheckAppStart(PROCESS_IDU_PC) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 Then  'V1.20.0.1 DEL
    If CheckAppStart(PROCESS_IDU_PC) = 0 Then   'V1.20.0.1 ADD
      'アプリ起動チェックタイマを停止する。
      tmrAplTimer.Enabled = False
     'V1.20.0.1 DEL START
'      '初期化処理
'      DeleteFile_Folder
     'V1.20.0.1 DEL END
     'V1.20.0.1 ADD START
     If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
        bIDURet = EndIDULog 'IDUログ起動時はIDUログに対してログ終了要求CMD送信
     Else
        bIDURet = True
     End If
     
     If bIDURet = True Then
        lngtime = 0
        lngtime = MN_MAIL_INTERVAL
        tmrLogTimer.Enabled = True
     Else
       '「一括システム初期化画面：システム初期化処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "初期化に失敗しました"
        OptKoumoku(0).Enabled = True
        OptKoumoku(1).Enabled = True
        cmdZikko.Enabled = True
        cmdCancel.Enabled = True
        'ログインユーザチェック
        If pbUserLevel = 1 Then
           OptKoumoku(2).Enabled = True
        End If        'アプリ起動チェックタイマを停止する。
        Exit Sub
     End If
     'V1.20.0.1 ADD END
    Else
    '起動アプリ有りの場合、タイマを張り直す
      tmrAplTimer.Interval = MN_MAIL_INTERVAL
    '合計経過待ち時間をアップ
     lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI定義値を超えた場合、初期化処理異常とする。
    '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "初期化に失敗しました"
    OptKoumoku(0).Enabled = True
    OptKoumoku(1).Enabled = True
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
    'ログインユーザチェック
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True
     End If
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()
     
   On Error Resume Next

   '待ち時間がINI定義を超えたかどうかチェック
  If lngtime <= lngLogMAX_Time Then
    'ログ起動チェックを行う。全て終了したときのみ、初期化処理を行う。
    If CheckAppStart(PROCESS_IDU_LOG) = 0 Then
      'ログ起動チェックタイマを停止する。
      tmrLogTimer.Enabled = False
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
    '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "初期化に失敗しました"
    OptKoumoku(0).Enabled = True
    OptKoumoku(1).Enabled = True
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
    'ログインユーザチェック
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True
    End If
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub DeleteFile_Folder()
    
    Dim bRtn As Boolean
    Dim lExitCode As Long
    Dim bRtn1 As Boolean
    Dim lngErrCode As Long                      'エラーコード
    Dim iTargetDB As Integer                       '対象DB値
    Dim bDB_Code As Boolean
   
    On Error GoTo ERR_SPACE
   
    '出荷時初期化選択時、全て初期化(DLLデータ含)選択時、その他データが初期化対象時
    If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
       If sCreateShokiFile = False Then
          '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "初期化に失敗しました"
          OptKoumoku(0).Enabled = True
          OptKoumoku(1).Enabled = True
          cmdZikko.Enabled = True
          cmdCancel.Enabled = True
          'ログインユーザチェック
          If pbUserLevel = 1 Then
             OptKoumoku(2).Enabled = True
          End If
          '処理を抜ける
          Exit Sub
       End If
    End If
     
    'システムファイルの削除
    If bChk(8) = True Then
       bRtn1 = sSysFileDelete()
       DoEvents
    Else
       bRtn1 = True
    End If

    'フォルダ、ファイルの削除
    If bRtn1 = True Then

       If sFileDelete() = True Then

          bDB_Code = True

          'DB初期化処理
          'DBデータ：IC一件明細
          If bChk(2) = True Then
             iTargetDB = chkMeisai(0).DataField
             Me.LstStatus.AddItem "DB初期化:" & chkMeisai(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
         End If

' EG20 V2.0.1.1 ADD START
          'DBデータ：バックアップデータ
          If bChk(3) = True And bDB_Code = True Then
             iTargetDB = chkLog(0).DataField
             Me.LstStatus.AddItem "DB初期化:" & chkLog(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If

          'DBデータ：再送データ
          If bChk(4) = True And bDB_Code = True Then
             iTargetDB = chkLog(0).DataField
             Me.LstStatus.AddItem "DB初期化:" & chkLog(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If
' EG20 V2.0.1.1 ADD START

          'DBデータ：アプリログ
          If bChk(5) = True And bDB_Code = True Then
             iTargetDB = chkLog(0).DataField
             Me.LstStatus.AddItem "DB初期化:" & chkLog(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If
          
          'DBデータ：保守ログ
          If bChk(6) = True And bDB_Code = True Then
             iTargetDB = chkLog(1).DataField
             Me.LstStatus.AddItem "DB初期化:" & chkLog(1).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If

          'DBデータ：判定ICモジュールログ
          If bChk(7) = True And bDB_Code = True Then
             iTargetDB = chkLog(2).DataField
             Me.LstStatus.AddItem "DB初期化:" & chkLog(2).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If

          'DBデータ：ネガリスト(出荷時初期化のみ有効)
          If OptKoumoku(1).Value = False And bDB_Code = True Then
             iTargetDB = stsIDUNega
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             DoEvents
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If
          
          '戻り値を見る
          If bDB_Code = True Then
             '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理正常」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
             lblKekka.ForeColor = SYSFORMAT_OK
             lblKekka.Caption = "初期化は成功しました"
          Else
             '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：DB初期化処理異常」ログ出力
             Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
             lblKekka.ForeColor = SYSFORMAT_ERROR
             lblKekka.Caption = "初期化に失敗しました"
          End If
     Else
        '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "初期化に失敗しました"
     End If
  Else
    '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "初期化に失敗しました"
  End If

 '初期化終了時の処理
 OptKoumoku(0).Enabled = True
 OptKoumoku(1).Enabled = True
 cmdZikko.Enabled = True
 cmdCancel.Enabled = True
 'ログインユーザチェック
 If pbUserLevel = 1 Then
    OptKoumoku(2).Enabled = True
 End If
 
Exit Sub

ERR_SPACE2:
        'エラー発生時の処理
        OptKoumoku(0).Enabled = True    '初期化項目指定：出荷時初期化選択可能
        OptKoumoku(1).Enabled = True    '初期化項目指定：項目選択選択可能
        cmdZikko.Enabled = True         '「初期化実行」釦押下可能
        cmdCancel.Enabled = True        '「メニュー画面へ戻る」釦押下可能
        'ログインユーザチェック
        If pbUserLevel = 1 Then
            OptKoumoku(2).Enabled = True
        End If
        '「ID中継ﾕﾆｯﾄｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "初期化に失敗しました"

ERR_SPACE:
End Sub
'V1.5.0.1 ADD　END
