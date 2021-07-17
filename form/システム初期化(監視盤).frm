VERSION 5.00
Begin VB.Form frmKansiSysformat 
   BorderStyle     =   0  'なし
   Caption         =   "システム初期化機能"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chksolog 
      Caption         =   "操作卓ログデータ"
      Height          =   495
      Left            =   5280
      TabIndex        =   23
      Top             =   1800
      Value           =   1  'ﾁｪｯｸ
      Width           =   2535
   End
   Begin VB.Timer tmrLogTimer 
      Left            =   11400
      Top             =   6480
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   8640
      Top             =   7800
   End
   Begin VB.Timer tmrMail 
      Left            =   8640
      Top             =   6000
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
      TabIndex        =   9
      Top             =   5400
      Width           =   2415
   End
   Begin VB.ListBox LstStatus 
      Height          =   3210
      Left            =   120
      TabIndex        =   8
      Top             =   5400
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
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   11775
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   4
         Left            =   4200
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   7
         Left            =   7680
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   5
         Left            =   4200
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   27
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   3
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   2
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   25
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   24
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   6
         Left            =   7560
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "詳細"
         Height          =   375
         Index           =   0
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Frame frmKoumoku 
         Caption         =   "項目"
         Height          =   3615
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   11295
         Begin VB.Frame FraShosai 
            Caption         =   "項目詳細"
            Height          =   1725
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   11100
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
               Height          =   1380
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   10845
            End
         End
         Begin VB.CheckBox chkDLL 
            Caption         =   "プログラム判定データ"
            Height          =   375
            Left            =   8160
            TabIndex        =   5
            Top             =   360
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.CheckBox chkIC 
            Caption         =   "IC関連データ"
            Height          =   375
            Left            =   8280
            TabIndex        =   4
            Top             =   840
            Value           =   1  'ﾁｪｯｸ
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.CheckBox chkSonota 
            Caption         =   "その他データ"
            Height          =   375
            Left            =   5040
            TabIndex        =   3
            Top             =   840
            Value           =   1  'ﾁｪｯｸ
            Width           =   3000
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "統合監視盤ログデータ"
            Height          =   375
            Left            =   1200
            TabIndex        =   13
            Top             =   1320
            Value           =   1  'ﾁｪｯｸ
            Width           =   3000
         End
         Begin VB.CheckBox chkBackUp 
            Caption         =   "バックアップデータ  　ログ中継機転送データ"
            Height          =   495
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Value           =   1  'ﾁｪｯｸ
            Width           =   3000
         End
         Begin VB.CheckBox chkMeisai 
            Caption         =   "集計関連データ"
            Height          =   375
            Left            =   1200
            TabIndex        =   1
            Top             =   360
            Value           =   1  'ﾁｪｯｸ
            Width           =   3000
         End
      End
      Begin VB.Frame FraKomoku 
         Height          =   620
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   10455
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "出荷時初期化"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   225
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "項目選択"
            Height          =   300
            Index           =   1
            Left            =   2640
            TabIndex        =   19
            Top             =   225
            Width           =   1575
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "全て初期化（プログラム判定データ含む）"
            Height          =   300
            Index           =   2
            Left            =   5160
            TabIndex        =   18
            Top             =   225
            Visible         =   0   'False
            Width           =   4935
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
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "統合監視盤システム初期化"
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
      TabIndex        =   14
      Top             =   0
      Width           =   12015
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
      TabIndex        =   11
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Caption         =   "初期化結果"
      Height          =   255
      Left            =   8760
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
End
Attribute VB_Name = "frmKansiSysformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmKansiSysformat.frm
'//  パッケージ名：システム初期化(監視盤)画面
'/
'//  概要：システム初期化(監視盤)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　保存用設定ファイル処理追加
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
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'Private bChk() As Boolean           'V1.5.0.1 DEL

'初期化実行フラグ
Private bSysFormat As Boolean

Private ShosaiMoji(0 To 7) As String '詳細文言格納エリア
Private Const SYSMOJI_SIZE = 500
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000     'アプリ起動タイマデフォルト値
Dim lngMAX_Time As Long                    'INI取得設定値
Dim lngtime     As Long                    '現在タイマ値
Private bChk(8) As Boolean
'V1.5.0.1 ADD END
'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        'ログ起動タイマデフォルト値(30秒)
Dim lngLogMAX_Time As Long                'INI取得設定値(ログ）
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : システム初期化(監視盤)画面(アクティブ時)
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
'//  機能名称  : システム初期化(監視盤)画面(ディアクティブ時)
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
'//  機能名称  : システム初期化(監視盤)画面(ロード時)
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
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//     REVISIONS :(EG20 v2.0.1.1) 2011-11-24  REVISED BY [TCC] T.Koyama
'//                ＥＧ２０フェーズ２対応【残件№54】
'//                ・初期化項目追加（操作卓ログデータ）
'//                ・項目詳細削除(ＩＣ関連データ,プログラム・判定データ）
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer    'カウンター
   
    On Error Resume Next

    '「監視盤システム初期化画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SYSFORMAT_GAMEN_START, 0)

    '「詳細」釦押下文言取得処理
    ShosaiMongonGet

   '初期化
    OptShosai(0).Value = True   '初期化項目指定：詳細釦押下
    LstStatus.Clear             '削除ファイル表示部クリア
    OptKoumoku(0).Value = True  '初期化項目指定「出荷時初期化」指定有り選択
    chkMeisai.Value = 1         '集計関連データ：チェック有り
    chkMeisai.Enabled = False   '集計関連データ：選択不可
    chkBackUp.Value = 1         'バックアップデータ：チェック有り
    chkBackUp.Enabled = False   'バックアップデータ：選択不可
    chkLog.Value = 1            'ログデータ：チェック有り（統合監視盤ログデータ）
    chkLog.Enabled = False      'ログデータ：選択不可（統合監視盤ログデータ）
' EG20 V2.0.1.1【残件№54】ADD START
    chksolog.Value = 1          '操作卓ログデータ：チェック有り
    chksolog.Enabled = False    '操作卓ログデータ：選択不可
' EG20 V2.0.1.1【残件№54】ADD START
    chkSonota.Value = 1         'その他データ：チェック有り
    chkSonota.Enabled = False   'その他データ：選択不可
' EG20 V2.0.1.1【残件№54】DEL START
'    chkDLL.Value = 0            'プログラム判定データ：チェック無し
'    chkDLL.Enabled = False      'プログラム判定データ：選択不可
'    chkIC.Value = 1             'IC関連データ：チェック有り
'    chkIC.Enabled = False       'IC関連データ：選択不可
' EG20 V2.0.1.1【残件№54】DEL　END
    lblKekka.Caption = ""       '初期化実行表示部クリア
    frmKoumoku.Enabled = False  '項目部押下不可

    OptKoumoku(2).Enabled = False
    
    'ログインユーザチェック
    If pbUserLevel = 1 Then
        OptKoumoku(2).Enabled = True
        chkDLL.Value = 1
        chkDLL.Enabled = False   'プログラム判定データ：選択可
    Else
        OptKoumoku(2).Enabled = False
    End If
    
    '初期化実行フラグOFF
    bSysFormat = False
    
    OptShosai(0).Enabled = True '初期化項目部：詳細釦押下可能
    OptShosai(0).Value = True   '初期化項目部：詳細釦押下
    For i = 1 To 6
        OptShosai(i).Enabled = False '項目部：詳細釦押下不可
    Next

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
'//     REVISIONS :(Eg20 V2.0.1.1) 2011-11-24  REVISED BY [TCC] T.Koyama
'//                ＥＧ２０フェーズ２対応【残件№54】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub OptKoumoku_Click(Index As Integer)
    Dim i As Integer    'カウンター

    On Error Resume Next
     
     Select Case Index
          Case 0  '出荷時初期化選択時
            frmKoumoku.Enabled = False       '項目フレーム
            chkMeisai.Enabled = False        '集計関連データ
            chkBackUp.Enabled = False        'バックアップデータ
            chkLog.Enabled = False           'ログデータ
            chkSonota.Enabled = False        'その他データ
' EG20 V2.0.1.1【残件№54】ADD START
            chksolog.Enabled = False         '操作卓ログデータ
' EG20 V2.0.1.1【残件№54】ADD END
' EG20 V2.0.1.1【残件№54】DEL START
'            chkIC.Enabled = False            'IC関連データ
'
'            'ログインユーザチェック
'            If pbUserLevel = 1 Then
'                chkDLL.Enabled = False       'プログラム判定データ
'            End If
' EG20 V2.0.1.1【残件№54】DEL END
            
            OptShosai(0).Enabled = True      '初期化項目部：詳細釦押下可能
            OptShosai(0).Value = True        '初期化項目部：詳細釦押下
            For i = 1 To 6
                OptShosai(i).Enabled = False '初期化項目部：詳細釦押下不可
            Next
            '「監視盤ｼｽﾃﾑ初期化画面：出荷時初期化選択時」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_SHUKKA, 0)
        Case 1   '項目選択時
            frmKoumoku.Enabled = True        '項目フレーム
            chkMeisai.Enabled = True         '集計関連データ
            chkBackUp.Enabled = True         'バックアップデータ
            chkLog.Enabled = True            'ログデータ
            chkSonota.Enabled = True         'その他データ
' EG20 V2.0.1.1【残件№54】ADD START
            chksolog.Enabled = True          '操作卓ログデータ
' EG20 V2.0.1.1【残件№54】ADD END
' EG20 V2.0.1.1【残件№54】DEL START
'            chkIC.Enabled = True             'IC関連データ
'
'            'ログインユーザチェック
'            If pbUserLevel = 1 Then
'                chkDLL.Enabled = True        'プログラム判定データ
'                OptShosai(6).Enabled = True  'プログラム判定データ詳細釦押下可能
'            End If
' EG20 V2.0.1.1【残件№54】DEL END
            OptShosai(0).Enabled = False     '初期化項目指定：詳細釦選択不可
            OptShosai(1).Value = True        '初期化項目指定：詳細釦非不可
            For i = 1 To 5
                OptShosai(i).Enabled = True  '項目指定：詳細釦選択可能
            Next
            
            '「監視盤ｼｽﾃﾑ初期化画面：項目選択選択時」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_KOUMOKU, 0)
        Case Else:
            frmKoumoku.Enabled = False       '項目フレーム
            chkMeisai.Enabled = False        '集計関連データ
            chkBackUp.Enabled = False        'バックアップデータ
            chkLog.Enabled = False           'ログデータ
            chkSonota.Enabled = False        'その他データ
' EG20 V2.0.1.1【残件№54】ADD START
            chksolog.Enabled = True          '操作卓ログデータ
' EG20 V2.0.1.1【残件№54】ADD END
' EG20 V2.0.1.1【残件№54】DEL START
'            chkIC.Enabled = False            'IC関連データ
'
'            'ログインユーザチェック
'            If pbUserLevel = 1 Then
'                chkDLL.Enabled = False       'プログラム判定データ
'            End If
' EG20 V2.0.1.1【残件№54】DEL END
            OptShosai(0).Enabled = True      '初期化項目指定：詳細釦選択可能
            OptShosai(0).Value = True        '初期化項目指定：詳細釦押下
            For i = 1 To 6
                OptShosai(i).Enabled = False '項目指定：詳細釦選択不可
            Next
            '「監視盤ｼｽﾃﾑ初期化画面：全て初期化選択時」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_ALL, 0)
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
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                         保守総点検結果修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加、確認ポップアップ修正
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-240】
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
    Dim bRtn1 As Boolean
    Dim bRtn2 As Boolean
    Dim iRetApp         As Integer
    Dim iRetLog         As Integer
    Dim uMail As ML_KYOTU_INF           'メール
    'ReDim bChk(8)                      'V1.5.0.1 DEL
    Dim lngErrCode As Long              'エラーコード
    Dim iTargetDB As Integer            '対象DB値
    Dim bDB_Code As Boolean
    Dim iRetIDULog As Integer           'IDUログ起動フラグ
    Dim iRetLDULog As Integer           'IDUログ起動フラグ
    Dim bRet As Boolean
    'V1.5.0.1  ADD START
    Dim bKansiRet As Boolean            '監視盤アプリ処理結果
    Dim bIDURet   As Boolean            'IDUアプリ処理結果
    Dim bLDURet   As Boolean            'LDUアプリ処理結果
   
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1  ADD END
    On Error GoTo ERR_SPACE

    '「監視盤ｼｽﾃﾑ初期化画面：初期化実行釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_START_BUTTOM, 0)

    '表示の初期化
    LstStatus.Clear
    lblKekka.Caption = ""
    iRetIDULog = 0
    iRetLDULog = 0
 
    '出荷時初期化選択時
    If OptKoumoku(0).Value = True Then
        For i = 1 To 6
           bChk(i) = True
        Next
           'プログラム判定データはチェック無し
'           bChk(5) = False                     ' EG20 V3.3.0.1【結合TR-240】削除
           bChk(6) = False                      ' EG20 V3.3.0.1【結合TR-240】追加
    End If

    '項目選択選択時
    If OptKoumoku(1).Value = True Then
        bSentaku = False
        '集計関連データ
        If chkMeisai.Value = 1 Then
            bSentaku = True
            bChk(1) = True
        Else
            bChk(1) = False
        End If

        'バックアップデータ
        If chkBackUp.Value = 1 Then
            bSentaku = True
            bChk(2) = True
        Else
            bChk(2) = False
        End If

        'ログデータ
        If chkLog.Value = 1 Then
            bSentaku = True
            bChk(3) = True
        Else
            bChk(3) = False
        End If
' EG20 V 2.0.1.1【残件№54】DEL START
'        'その他データ
'        If chkSonota.Value = 1 Then
'           bSentaku = True
'           bChk(4) = True
'        Else
'           bChk(4) = False
'        End If
' EG20 V 2.0.1.1【残件№54】DEL END
' EG20 V 2.0.1.1【残件№54】ADD START
        '操作卓ログデータ
        If chksolog.Value = 1 Then
           bSentaku = True
           bChk(4) = True
        Else
           bChk(4) = False
        End If

        'その他データ
        If chkSonota.Value = 1 Then
           bSentaku = True
           bChk(5) = True
        Else
           bChk(5) = False
        End If
' EG20 V 2.0.1.1【残件№54】ADD END
' EG20 V 2.0.1.1【残件№54】DEL START
'        'ＤＬＬデータ
'        If chkDLL.Value = 1 Then
'            bSentaku = True
'            bChk(5) = True
'        Else
'            bChk(5) = False
'        End If
'
'        'IC関連データ
'        If chkIC.Value = 1 Then
'            bSentaku = True
'            bChk(6) = True
'        Else
'            bChk(6) = False
'        End If
' EG20 V 2.0.1.1【残件№54】DEL END
        bChk(6) = False                      ' EG20 V3.3.0.1【結合TR-240】追加

        If bSentaku = False Then
            '「監視盤ｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
            MsgBox "初期化するデータが選択されていません", vbExclamation, "データ無警告"
            Exit Sub
        End If
    End If

' EG20 V 2.0.1.1【残件№54】DEL START
'    '全て初期化（ＤＬＬデータ含む）選択時
'    If OptKoumoku(2).Value = True Then
'        For i = 1 To 6
'            '全て選択チェック有り
'            bChk(i) = True
'        Next
'    End If
' EG20 V 2.0.1.1【残件№54】DEL END
    
    iRet = MsgBox("初期化処理を行います。よろしいですか？", vbExclamation + vbOKCancel, "初期化確認")
    If iRet = vbOK Then
        '初期化正常終了時の処理
         OptKoumoku(0).Enabled = False      '「出荷時初期化」ラジオ釦選択不可
         OptKoumoku(1).Enabled = False      '「項目選択」ラジオ釦選択不可
' EG20 V 2.0.1.1【残件№54】DEL START
'        'ログインユーザチェック
'        If pbUserLevel = 1 Then
'           OptKoumoku(2).Enabled = False  '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
'        End If
' EG20 V 2.0.1.1【残件№54】DEL END
        cmdZikko.Enabled = False          '「初期化実行」釦押下不可
        cmdCancel.Enabled = False         '「メニュー画面へ戻る」釦押下不可
    
        On Error GoTo ERR_SPACE2

        '監視盤(管理プロセス)が起動しているかどうかチェックする。
        If CheckAppStart(PROC_KANRI) <> 0 Then
          'V1.20.0.1 DEL START
          ' iRet = MsgBox("監視盤アプリケーションを終了します。" & Chr(vbKeyReturn) & _
          '               "よろしいですか？", vbQuestion + vbOKCancel, "終了確認")
          'If iRet = vbOK Then
          'V1.20.0.1 DEL END
              'アプリ終了要求を管理に送信する
               uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
               uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
               uMail.udtlHeader.dwProid = RHOSHU_ID
               uMail.udtlHeader.dwSubArea = 0
               'V1.5.0.1 DEL START
               'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               'If bRtn = 0 Then
               'V1.5.0.1 DEL END
               'V1.5.0.1 ADD START
               bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               If bKansiRet = 0 Then
               'V1.5.0.1 ADD END
                 '「監視盤ｼｽﾃﾑ初期化画面：メール送信異常」ログ出力
                 lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                 Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lngErrCode)
                 GoTo ERR_SPACE2:
               Else
                 '「監視盤ｼｽﾃﾑ初期化画面：メール送信正常」ログ出力
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                 'アプリ終了確認
                  'iRetApp = CheckAppEndComplete(PROC_KANRI, lExitCode)   'V1.5.0.1 DEL
               End If
         'V1.20.0.1 DEL START
'              'IDUログプロセス起動チェック
'              If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'
'                 'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
'                 iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
'
'                 If iRet = vbOK Then
'                   'IDUログ終了要求CMD送信
'                   'V1.5.0.1 DEL START
'                   'bRet = EndIDULog
'                   'If bRtn = False Then
'                   'V1.5.0.1 DEL END
'                   'V1.5.0.1 ADD START
'                   bIDURet = EndIDULog
'                   If bIDURet = False Then
'                   'V1.5.0.1 ADD END
'                     '送信異常処理
'                     lblKekka.ForeColor = SYSFORMAT_ERROR
'                     lblKekka.Caption = "初期化に失敗しました"
'                     OptKoumoku(0).Enabled = True
'                     OptKoumoku(1).Enabled = True
'                     'ログインユーザチェック
'                     If pbUserLevel = 1 Then
'                        OptKoumoku(2).Enabled = True
'                     End If
'                     cmdZikko.Enabled = True
'                     cmdCancel.Enabled = True
'                     Exit Sub
'                  End If
'                  'IDUログプロセス終了確認
'                  'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)  'V1.5.0.1 DEL
'                'V1.7.0.1 ADD START
'                Else
'                 GoTo ERR_SPACE3
'                'V1.7.0.1 ADD END
'                End If
'              'V1.5.0.1 ADD START
'              Else
'              bIDURet = True
'              'V1.5.0.1 ADD END
'              End If
'              'LDUログプロセス起動チェック
'              'If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then 'V1.7.0.1 DEL
'              If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then  'V1.7.0.1 ADD
'
'                 'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
'                 iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
'
'                 If iRet = vbOK Then
'                   'IDUログ終了要求CMD送信
'                   'V1.5.0.1 DEL START
'                   'bRet = EndLDULog
'                   'If bRtn = False Then
'                   'V1.5.0.1 DEL END
'                   'V1.5.0.1 ADD START
'                   bLDURet = EndLDULog
'                   If bLDURet = False Then
'                   'V1.5.0.1 ADD END
'                     '送信異常処理
'                     lblKekka.ForeColor = SYSFORMAT_ERROR
'                     lblKekka.Caption = "初期化に失敗しました"
'                     OptKoumoku(0).Enabled = True
'                     OptKoumoku(1).Enabled = True
'                     'ログインユーザチェック
'                     If pbUserLevel = 1 Then
'                        OptKoumoku(2).Enabled = True
'                     End If
'                     cmdZikko.Enabled = True
'                     cmdCancel.Enabled = True
'                     Exit Sub
'                  End If
'                  'LDUログプロセス終了確認
'                  'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)  'V1.5.0.1 DEL
'                'V1.7.0.1 ADD START
'                Else
'                 GoTo ERR_SPACE3
'                'V1.7.0.1 ADD END
'                End If
'              'V1.5.0.1 ADD START
'              Else
'              bLDURet = True
'              'V1.5.0.1 ADD END
'              End If
'          Else
'             '「キャンセル釦押下」
'             '「監視盤ｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'              OptKoumoku(0).Enabled = True    '「出荷時初期化」ラジオ釦選択不可
'              OptKoumoku(1).Enabled = True    '「項目選択」ラジオ釦選択不可
'              'ログインユーザチェック
'              If pbUserLevel = 1 Then
'                OptKoumoku(2).Enabled = True  '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
'              End If
'              cmdZikko.Enabled = True         '「初期化実行」釦押下不可
'              cmdCancel.Enabled = True        '「メニュー画面へ戻る」釦押下不可
'              Exit Sub
'          End If
         'V1.20.0.1 DEL END
      Else
        iRetApp = 1
        bKansiRet = True    'V1.5.0.1 ADD
      'End If               'V1.5.0.1 DEL
      
       If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
          
          'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
        'V1.20.0.1 DEL START
'          iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
'
'           If iRet = vbOK Then
        'V1.20.0.1 DEL END
              'IDUログ終了要求CMD送信
               'V1.5.0.1 DEL START
               'bRet = EndIDULog
               'If bRtn = False Then
               'V1.5.0.1 DEL END
               'V1.5.0.1 ADD START
               bIDURet = EndIDULog
               If bIDURet = False Then
               'V1.5.0.1 ADD END
                 '送信異常処理
                  lblKekka.ForeColor = SYSFORMAT_ERROR
                  lblKekka.Caption = "初期化に失敗しました"
                  OptKoumoku(0).Enabled = True
                  OptKoumoku(1).Enabled = True
                  'ログインユーザチェック
                   If pbUserLevel = 1 Then
                      OptKoumoku(2).Enabled = True
                   End If
                   cmdZikko.Enabled = True
                   cmdCancel.Enabled = True
                   Exit Sub
               End If
               'IDUログプロセス終了確認
               'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)    'V1.5.0.1 DEL
           'V1.7.0.1 ADD START
        'V1.20.0.1 DEL START
'           Else
'              GoTo ERR_SPACE3
'           'V1.7.0.1 ADD END
'           End If
        'V1.20.0.1 DEL END
      Else
        iRetIDULog = 1
        bIDURet = True 'V1.5.0.1 ADD
      End If
       
      If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
         
         'iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "終了確認") 'V1.8.0.1 DEL
        'V1.20.0.1 DEL START
'         iRet = MsgBox("ログプロセスを終了します。よろしいですか？", vbQuestion + vbOKCancel, "ログ終了確認")  'V1.8.0.1 ADD
'
'         If iRet = vbOK Then
        'V1.20.0.1 DEL END
           'IDUログ終了要求CMD送信
            'V1.5.0.1 DEL START
            'bRet = EndLDULog
            'If bRtn = False Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 DEL START
             bLDURet = EndLDULog
             If bLDURet = False Then
            'V1.5.0.1 DEL END
                '送信異常処理
                 lblKekka.ForeColor = SYSFORMAT_ERROR
                 lblKekka.Caption = "初期化に失敗しました"
                 OptKoumoku(0).Enabled = True
                 OptKoumoku(1).Enabled = True
                 'ログインユーザチェック
                 If pbUserLevel = 1 Then
                    OptKoumoku(2).Enabled = True
                 End If
                 cmdZikko.Enabled = True
                 cmdCancel.Enabled = True
                 Exit Sub
              End If
            'LDUログプロセス終了確認
            'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)  'V1.5.0.1 DEL
         'V1.7.0.1 ADD START
       'V1.20.0.1 DEL START
'         Else
'            GoTo ERR_SPACE3
'         'V1.7.0.1 ADD END
'         End If
       'V1.20.0.1 DEL END
      Else
         iRetLDULog = 1
         bLDURet = True 'V1.5.0.1 ADD
      End If
     End If             'V1.5.0.1 ADD
'V1.5.0.1 ADD START
     '監視盤、IDU、LDUアプリのメール送信処理が全て正常だった場合のみ、アプリ起動タイマを起動させ、
     'アプリ起動チェックによりアプリの起動/未起動を判断する。
     'If (bKansiRet = True) And (bIDURet = True) And (bLDURet = True) Then         'V1.20.0.1 DEL
     If (bKansiRet = True) Then                                                    'V1.20.0.1 ADD
         lngtime = 0
         lngtime = MN_MAIL_INTERVAL
         tmrAplTimer.Enabled = True
     Else
         '監視盤、IDU、LDUアプリのメール送信にてひとつでも異常があった場合、初期化処理を異常終了とする。
         '「監視盤システム初期化画面：システム初期化処理異常」ログ出力
         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "初期化に失敗しました"
         '初期化正常終了時の処理
         OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
         OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
         'ログインユーザチェック
          If pbUserLevel = 1 Then
             OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
          End If
          cmdZikko.Enabled = True        '「初期化実行」釦押下不可
          cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
          '処理を抜ける
           Exit Sub
      End If
   End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'       'アプリまたはログプロセスで終了処理に失敗した場合
'       If (iRetApp <> 1) Or (iRetIDULog <> 1) Or (iRetLDULog <> 1) Then
'           '「一括システム初期化画面：システム初期化処理異常」ログ出力
'           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "初期化に失敗しました"
'           '初期化正常終了時の処理
'            OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
'            OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
'            'ログインユーザチェック
'            If pbUserLevel = 1 Then
'               OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
'            End If
'            cmdZikko.Enabled = True        '「初期化実行」釦押下不可
'            cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
'            '処理を抜ける
'            Exit Sub
'       End If
'
'      '初期化実行フラグON
'      bSysFormat = True
'
'      'V1.4.0.1 ADD START
'      '出荷時初期化選択時、全て初期化(DLLデータ含)選択時、その他データが初期化対象時
'      If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
'        If sCreateShokiFile = False Then
'          '「一括システム初期化画面：システム初期化処理異常」ログ出力
'          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'          lblKekka.ForeColor = SYSFORMAT_ERROR
'          lblKekka.Caption = "初期化に失敗しました"
'          '初期化正常終了時の処理
'           OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
'           OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
'           'ログインユーザチェック
'           If pbUserLevel = 1 Then
'              OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
'           End If
'           cmdZikko.Enabled = True        '「初期化実行」釦押下不可
'           cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
'           '処理を抜ける
'           Exit Sub
'        End If
'      End If
'      'V1.4.0.1 ADD END
'
'      'システムファイルの削除
'      If bChk(4) = True Then
'           bRtn1 = sSysFileDelete()
'           DoEvents
'      Else
'           bRtn1 = True
'      End If
'
'      'フォルダ、ファイルの削除
'      If bRtn1 = True Then
'
'        If sFileDelete() = True Then
'
'           bDB_Code = True
'
'          If bChk(1) = True Then
'             Me.LstStatus.AddItem "DB初期化:" & chkMeisai.Caption
'             DoEvents
'             Me.AutoRedraw = True
'
'             '監視盤：一件明細
'             iTargetDB = stsKansiMeisai
'             bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'             DoEvents
'             Me.AutoRedraw = True
'             If bDB_Code = True Then
'                '監視盤：別集札
'                iTargetDB = stsKansiBetu
'                'DB初期化処理
'                bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'                DoEvents
'                Me.AutoRedraw = True
'             End If
'          End If
'
'          If bDB_Code = True Then
'             '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理正常」ログ出力
'             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
'             lblKekka.ForeColor = SYSFORMAT_OK
'             lblKekka.Caption = "初期化は成功しました"
'          Else
'             '「監視盤ｼｽﾃﾑ初期化画面：DB初期化処理異常」ログ出力
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DBFORMAT_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "初期化に失敗しました"
'          End If
'        Else
'          '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
'          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'          lblKekka.ForeColor = SYSFORMAT_ERROR
'          lblKekka.Caption = "初期化に失敗しました"
'        End If
'    Else
'       '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
'       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'       lblKekka.ForeColor = SYSFORMAT_ERROR
'       lblKekka.Caption = "初期化に失敗しました"
'    End If
'End If
'    '初期化正常終了時の処理
'    OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
'    OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
'    'ログインユーザチェック
'    If pbUserLevel = 1 Then
'       OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
'    End If
'     cmdZikko.Enabled = True        '「初期化実行」釦押下不可
'     cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
'V1.5.0.1 DEL END
Exit Sub

'V1.7.0.1 ADD START
ERR_SPACE3:
'「キャンセル釦押下」
'「監視盤ｼｽﾃﾑ初期化画面：初期化処理未実行」ログ出力
Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
OptKoumoku(0).Enabled = True    '「出荷時初期化」ラジオ釦選択不可
OptKoumoku(1).Enabled = True    '「項目選択」ラジオ釦選択不可
'ログインユーザチェック
If pbUserLevel = 1 Then
   OptKoumoku(2).Enabled = True  '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
End If
cmdZikko.Enabled = True         '「初期化実行」釦押下不可
cmdCancel.Enabled = True        '「メニュー画面へ戻る」釦押下不可
Exit Sub
'V1.7.0.1 ADD END

ERR_SPACE2:
        'エラー発生時の処理
        OptKoumoku(0).Enabled = True    '「出荷時初期化」ラジオ釦選択不可
        OptKoumoku(1).Enabled = True    '「項目選択」ラジオ釦選択不可
        'ログインユーザチェック
        If pbUserLevel = 1 Then
           OptKoumoku(2).Enabled = True '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
        End If
        cmdZikko.Enabled = True         '「初期化実行」釦押下不可
        cmdCancel.Enabled = True        '「メニュー画面へ戻る」釦押下不可
        '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
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
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         フェーズ１不具合対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    On Error Resume Next

     '「監視盤システム初期化：消去」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SYSFORMAT_GAMEN_END, 0)
     'frmALLSysformat.ZOrder 'V1.5.0.1 DEL
     frmSysformatMenu.ZOrder 'V1.5.0.1 ADD
     Unload Me

End Sub

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
   '「監視盤システム初期化画面：システムファイル削除異常」ログ出力
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
Private Function sFileDelete()
    Dim iFileNo As Integer
    Dim sFileData As String
    Dim iMozi, iKbn As Integer
    Dim sShubetu As String
    Dim sRoot As String
    Dim sPass As String
    Dim sKomoku As String
    Dim bSyori As Boolean
    Dim fs As Object
    Dim MyName As String
    Dim i As Integer
    Dim sChkPass As String
    Dim iRet As Integer
    Dim lngErrCode As Long       'エラーコード
    Dim lBool As Boolean         ' EG20 V2.0.1.1【結合TR-240】追加

    sFileDelete = False

    On Error GoTo ERR_SPACE
        
    'ファイル有無チェック
    MyName = Dir(KANSI_SYSTEMFILE, vbNormal)
    If MyName = "" Then
        GoTo ERR_SPACE
    End If

' EG20 V3.3.0.1【結合TR-240】追加開始（位置移動）
    ' 保守ログファイルCLOSE
    lBool = dllCloseHoshuLogFile()
' EG20 V3.3.0.1【結合TR-240】追加終了（位置移動）

    iFileNo = FreeFile                                           '未使用のファイル番号を取得する。
    Open KANSI_SYSTEMFILE For Input As #iFileNo                  'システム初期化設定ファイルを開く。
    Line Input #iFileNo, sFileData                               ' １行目は全体バージョンなので読飛ばす。
    Do While Not EOF(iFileNo)
    Line Input #iFileNo, sFileData                               ' １行分読込む。
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
                          DoEvents  'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                      'フォルダの削除／作成
                      Case 3:
                          fs.DeleteFolder (sPass), True
                          fs.CreateFolder (sPass)
                          LstStatus.AddItem "削除／作成したフォルダ - " & sPass
                          DoEvents  'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                  End Select
                'オブジェクト解放
                Set fs = Nothing
            Else
                '指定ＰＡＳＳナシ
                Select Case sShubetu
                   Case 2:
                       LstStatus.AddItem "指定ファイルなし - " & sPass
                       DoEvents  'V1.5.0.1 ADD
                       LstStatus.Selected(LstStatus.ListCount - 1) = True           'V1.12.0.1 ADD
                   Case 3:
                       Set fs = CreateObject("Scripting.FileSystemObject")
                       'ファイル有無チェック
'                       For i = 0 To Len(sPass)         'EG20 V2.1.0.1 DEL 【統-313対応】
                       For i = 0 To Len(sPass) - 1      'EG20 V2.1.0.1 ADD 【統-313対応】
                           If Mid(sPass, Len(sPass) - i, 1) = "\" Then
                               sChkPass = Left(sPass, Len(sPass) - i - 1)
                               Exit For
                           End If
                       Next
                       MyName = Dir(sChkPass, vbDirectory)
                       If MyName = "" Then
                           LstStatus.AddItem "フォルダ作成失敗 - " & sPass
                           DoEvents  'V1.5.0.1 ADD
                           LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
                       Else
                           fs.CreateFolder (sPass)
                           LstStatus.AddItem "作成したフォルダ - " & sPass
                           DoEvents  'V1.5.0.1 ADD
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
   '「監視盤システム初期化画面：ファイル・フォルダ初期化異常」ログ出力
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TARGET_FILE_FOLDER_DELETE_ERROR, lngErrCode)
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
   
   '「監視盤ｼｽﾃﾑ初期化画面：詳細釦押下」ログ出力
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
       For i = 0 To 7
        ShosaiMoji(i) = sWork
       Next
       Exit Sub
   End If
    
   For iGate = CNT_MIN To 7
      ' SysFormatShousai.iniより文言を取得する。
       sGateData = ""
       iKey = SYS_KEY_NAME & iGate
       lSts = GetPrivateProfileString(SYS_KANSI_SECTION_NAME, _
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmKansiSysformat.Caption, False
        pfFormActive (frmKansiSysformat.hwnd)
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
    
    sCreateShokiFile = True
    '「監視盤システム初期化画面：保存用設定ファイル作成正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
    
    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「監視盤システム初期化画面：保存用設定ファイル作成異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SHOKI_CREATE_ERROR, lngErrCode)
    sCreateShokiFile = False
End Function
'V1.4.0.1　ADD END

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
    'If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then 'V1.20.0.1 DEL
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
         '初期化正常終了時の処理
          OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
          OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
          'ログインユーザチェック
          If pbUserLevel = 1 Then
             OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
          End If
          cmdZikko.Enabled = True        '「初期化実行」釦押下不可
          cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
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
    '「監視盤システム初期化画面：システム初期化処理異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "初期化に失敗しました"
    '初期化正常終了時の処理
    OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
    OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
    'ログインユーザチェック
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
    End If
    cmdZikko.Enabled = True        '「初期化実行」釦押下不可
    cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
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
    If CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
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
    '「一括システム初期化画面：システム初期化処理異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "初期化に失敗しました"
    '初期化正常終了時の処理
    OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
    OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
    'ログインユーザチェック
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
    End If
    cmdZikko.Enabled = True        '「初期化実行」釦押下不可
    cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
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
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-23  REVISED BY [TCC] T.Koyama
'//                ＥＧ２０フェーズ２対応【残件№54】
'//                ・保守ログファイルＣＬＯＳＥ処理追加
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-240】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub DeleteFile_Folder()

    Dim i As Integer
    Dim bRtn As Boolean
    Dim bSentaku As Boolean
    Dim iRet As Integer
    Dim lExitCode As Long
    Dim bRtn1 As Boolean
    Dim bRtn2 As Boolean
    Dim lngErrCode As Long              'エラーコード
    Dim iTargetDB As Integer            '対象DB値
    Dim bDB_Code As Boolean
    Dim iRetIDULog As Integer           'IDUログ起動フラグ
    Dim iRetLDULog As Integer           'IDUログ起動フラグ
    Dim bRet As Boolean
  
    Dim lBool As Boolean                ' EG20 V2.0.1.1【残件№54】ADD
 
    'EG20 V2.1.0.1 ADD START 【統-313対応】
    Dim intLoop As Integer
    Dim lSts As Long
    'EG20 V2.1.0.1 ADD END
    
    On Error GoTo ERR_SPACE
  
  '出荷時初期化選択時、全て初期化(DLLデータ含)選択時、その他データが初期化対象時
  If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
     
' EG20 V3.3.0.1【結合TR-240】削除開始（位置移動）
'     ' EG20 V2.0.1.1【残件№54】ADD START
'     If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkLog.Value = 1 Then
'
'        ' 保守ログファイルCLOSE
'         lBool = dllCloseHoshuLogFile()
'      End If
'     ' EG20 V2.0.1.1【残件№54】ADD START
' EG20 V3.3.0.1【結合TR-240】削除終了（位置移動）
      
     If sCreateShokiFile = False Then
        '「一括システム初期化画面：システム初期化処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "初期化に失敗しました"
        '初期化正常終了時の処理
        OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
        OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
        'ログインユーザチェック
        If pbUserLevel = 1 Then
           OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
        End If
        cmdZikko.Enabled = True        '「初期化実行」釦押下不可
        cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可
        '処理を抜ける
        Exit Sub
      End If
   End If

   'システムファイルの削除
'   If bChk(4) = True Then              ' EG20 V3.3.0.1【結合TR-240】削除
   If bChk(5) = True Then               ' EG20 V3.3.0.1【結合TR-240】追加
      bRtn1 = sSysFileDelete()
   Else
      bRtn1 = True
   End If

   'フォルダ、ファイルの削除
   If bRtn1 = True Then

      If sFileDelete() = True Then

         bDB_Code = True
         
         If bChk(1) = True Then
            Me.LstStatus.AddItem "DB初期化:" & chkMeisai.Caption
            DoEvents
            LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
            
            '監視盤：一件明細
            Me.LstStatus.AddItem "一件明細コーナ１　DB初期化開始"
            DoEvents
            iTargetDB = stsKansiMeisai
            bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
            Me.LstStatus.AddItem "一件明細コーナ１　DB初期化終了"
            DoEvents
            
            If bDB_Code = True Then
               '監視盤：一件明細（コーナ２）
               Me.LstStatus.AddItem "一件明細コーナ２　DB初期化開始"
               DoEvents
               iTargetDB = stsKansiMeisai2
               'DB初期化処理
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "一件明細コーナ２　DB初期化終了"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '監視盤：一件明細（コーナ３）
               Me.LstStatus.AddItem "一件明細コーナ３　DB初期化開始"
               DoEvents
               iTargetDB = stsKansiMeisai3
               'DB初期化処理
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "一件明細コーナ３　DB初期化終了"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '監視盤：一件明細（コーナ４）
               Me.LstStatus.AddItem "一件明細コーナ４　DB初期化開始"
               DoEvents
               iTargetDB = stsKansiMeisai4
               'DB初期化処理
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "一件明細コーナ４　DB初期化終了"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '監視盤：一件明細（コーナ５）
               Me.LstStatus.AddItem "一件明細コーナ５　DB初期化開始"
               DoEvents
               iTargetDB = stsKansiMeisai5
               'DB初期化処理
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "一件明細コーナ５　DB初期化終了"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '監視盤：一件明細（コーナ６）
               Me.LstStatus.AddItem "一件明細コーナ６　DB初期化開始"
               DoEvents
               iTargetDB = stsKansiMeisai6
               'DB初期化処理
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "一件明細コーナ６　DB初期化終了"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '監視盤：別集札
               Me.LstStatus.AddItem "別集札　DB初期化開始"
               DoEvents
               iTargetDB = stsKansiBetu
               'DB初期化処理
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
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
         End If

         If bDB_Code = True Then
            '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
            lblKekka.ForeColor = SYSFORMAT_OK
            lblKekka.Caption = "初期化は成功しました"
         Else
            '「監視盤ｼｽﾃﾑ初期化画面：DB初期化処理異常」ログ出力
             Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
             lblKekka.ForeColor = SYSFORMAT_ERROR
             lblKekka.Caption = "初期化に失敗しました"
         End If
      Else
        '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "初期化に失敗しました"
      End If
   Else
      '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
      lblKekka.ForeColor = SYSFORMAT_ERROR
      lblKekka.Caption = "初期化に失敗しました"
   End If
 
 '初期化正常終了時の処理
 OptKoumoku(0).Enabled = True      '「出荷時初期化」ラジオ釦選択不可
 OptKoumoku(1).Enabled = True      '「項目選択」ラジオ釦選択不可
 'ログインユーザチェック
 If pbUserLevel = 1 Then
    OptKoumoku(2).Enabled = True   '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
 End If
 cmdZikko.Enabled = True        '「初期化実行」釦押下不可
 cmdCancel.Enabled = True       '「メニュー画面へ戻る」釦押下不可

Exit Sub

ERR_SPACE2:
  'エラー発生時の処理
  OptKoumoku(0).Enabled = True    '「出荷時初期化」ラジオ釦選択不可
  OptKoumoku(1).Enabled = True    '「項目選択」ラジオ釦選択不可
  'ログインユーザチェック
  If pbUserLevel = 1 Then
     OptKoumoku(2).Enabled = True '「全て初期化(プログラム判定データ含む)」ラジオ釦選択不可
  End If
  cmdZikko.Enabled = True         '「初期化実行」釦押下不可
  cmdCancel.Enabled = True        '「メニュー画面へ戻る」釦押下不可
  '「監視盤ｼｽﾃﾑ初期化画面：システム初期化処理異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
  lblKekka.ForeColor = SYSFORMAT_ERROR
  lblKekka.Caption = "初期化に失敗しました"
ERR_SPACE:
End Sub
