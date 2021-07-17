VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEkiDataSubGate 
   BorderStyle     =   0  'なし
   Caption         =   "駅都度データ確認"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
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
   Begin VB.ComboBox CmbCornerName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8280
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton CmdMoveEkiInfoGamen 
      Caption         =   "駅情報画面へ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7080
      TabIndex        =   11
      Top             =   8400
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "駅設定出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "駅設定入力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   2400
      TabIndex        =   9
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "テキスト媒体出力(ｴﾝｺｰﾄﾞｺｰﾅ号機設定)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   4680
      TabIndex        =   8
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton CmdMenu 
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
      Height          =   555
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdMoveGateGamen 
      Caption         =   "改札機画面へ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7080
      TabIndex        =   6
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   8160
      Top             =   6000
   End
   Begin VB.ComboBox cmbGoki 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9480
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   10425
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "  機器情報設定    画面へ戻る"
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   5730
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10107
      _Version        =   393216
      Rows            =   33
      Cols            =   9
      FixedCols       =   3
      RowHeightMin    =   350
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "駅都度データ確認（エンコードコーナ号機情報定義）"
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
   Begin VB.Label LblEkiName 
      Caption         =   "駅名：○○○○○○○○○○○○○○○○○○○○"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   7815
   End
End
Attribute VB_Name = "frmEkiDataSubGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：駅都度データ確認（エンコード）画面.frm
'//  パッケージ名：駅都度データ確認（エンコード）画面のフォームモジュール
'//
'//  概要：駅都度データ確認（エンコード）画面.frm
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//                 EG-R阪急　流用
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応
'//                 【HKRK_Kansi07_003_01】 Sub_gate_kan.iniフォーマット見直し対応
'//                 【HKRK_Kansi07_008_01】 駅都度データの小分類を意識して表示、媒体出力を行う
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   'メイルタイマのインターバル値
'Private Const TITOL_EKI_NAME = "駅名　　　："           '駅名タイトル      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
Private Const TITOL_EKI_NAME = "駅名："                 '駅名タイトル       ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

'機器構成データ（エンコードコーナ号機情報定義）イメージファイル読取用の構造体
Private Type SUBGATE_IMAGE_FILE
    sType       As String                '種別
    sGoki       As String                '号機
    sNo         As String                '種別毎通番
    sCorner     As String                'コーナ        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    sKoumoku    As String                '項目
    sKubun      As String                '区分
    sSettei     As String                '設定値
    sSyosai     As String                '設定値詳細
End Type

'機器構成データ（エンコードコーナ号機情報定義）出力ファイル作成用の定数定義
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'Private Const ONE_PARAGRAPH_OUTPUT_KOUMOKU = 48  '1段落出力項目数
'Private Const ONE_PARAGRAPH_OUTPUT_ROW = 16      '1段落出力用配列の要素数
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'Private Const ONE_PARAGRAPH_OUTPUT_KOUMOKU = 24  '1段落出力項目数
'Private Const ONE_PARAGRAPH_OUTPUT_ROW = 3       '1段落出力用配列の要素数
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
Private Const ONE_PARAGRAPH_OUTPUT_KOUMOKU = 48  '1段落出力項目数
Private Const ONE_PARAGRAPH_OUTPUT_ROW = 0       '1段落出力用配列の要素数
'EG20 V30.1.0.1 ADD END
Private Const ONE_PARAGRAPH_OUTPUT_GOKI = 8      '1段落に出力する号機数
Private Const GOKI_JISHA_START = "1"             '小分類「自社」の最小項番
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'Private Const GOKI_JISHA_END = "6"               '小分類「自社」の最大項番
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'Private Const GOKI_JISHA_END = "3"               '小分類「自社」の最大項番
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'EG20 V30.1.0.1 DEL END
Private Const GOKI_JISHA_END = "6"               '小分類「自社」の最大項番

'機器構成データ（エンコードコーナ号機情報定義）出力ファイル作成用のテーブル定義
Private Type SUBGATE_OUT_DEF_TBL
    iRow            As Integer                   '行番号
    iNoStart        As Integer                   '小分類の最小項番
    iNoEnd          As Integer                   '小分類の最大項番
End Type

'機器構成データ（エンコードコーナ号機情報定義）出力ファイル作成用の構造体（1行）
Private Type SUBGATE_IMAGE_FILE_ONE_ROW
    sShakyoku       As String                    '社局
    sKubun          As String                    '区分
    sSettei         As String                    '設定値
End Type

'Private Const START_DATA_COL_INDEX = 2           '1行のデータ設定を開始するカラムインデックス  'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
Private Const START_DATA_COL_INDEX = 3           '1行のデータ設定を開始するカラムインデックス   'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
'Private Const MAX_DATA_COL_INDEX = 7             '1行の最大設定カラム数    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'Private Const MAX_DATA_COL_INDEX = 4             '1行の最大設定カラム数     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加 'EG20 V30.1.0.1 DEL
'Private Const MAX_DATA_COL_INDEX = 7             '1行の最大設定カラム数     ' EG20 V30.1.0.1 ADD 'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
Private Const MAX_DATA_COL_INDEX = 8             '1行の最大設定カラム数     ' EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD

Private gstrFileName        As String                       ' 出力ファイル名    ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmbCornerName_Click
'//  機能名称  : コーナ選択処理
'//  機能概要  : グリッドデータを再設定する
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度修正対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmbCornerName_Click()
    
    Dim iIndex          As Integer                  'インデックス
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    Call SetEnableFalse

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_GAMEN_CORNER_SELECT, 0)
    
    '画面表示処理
    Call sDisp

    '全ボタンを押下可とする。
    Call SetEnableTrue


End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 駅都度データ確認（エンコード）画面(アクティブ時：イベントプロシージャ)
'//  機能概要  : 最前前表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'エラールーチンを宣言
    On Error Resume Next
    
    '自画面最前面表示処理を行う。
    pfFormActive (hwnd)
    
    'タイマを起動する
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 駅都度データ確認（エンコード）画面(ディアクティブ時)
'//  機能概要  : メール受信用、タイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'タイマを停止する
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 駅都度データ確認（エンコード）画面(ロード時：イベントプロシージャ)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(V30.1.0.1) 2014-05-20 CODED BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード
    Dim iLoopCnt             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_SUBGATE_GAMEN_START, 0)
    
    '----------------------------------------------------
    '画面初期値設定
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイル作成
    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイル削除
        Kill EKI_TUDO_CHK_SUBGATE_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
    End If

    '号機コンボボックス初期値
    cmbGoki.Clear

    'For iLoopCnt = 0 To 15 'EG20 V30.1.0.1 DEL
    For iLoopCnt = 0 To 31  'EG20 V30.1.0.1 ADD
            cmbGoki.AddItem iLoopCnt + 1 & "号機"
    Next
    cmbGoki.ListIndex = 0
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    'コーナ設定コンボボックスの初期化処理
    Call InitCornerComboBox
    Call sDisp
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ処理（タイムアップ時：イベントプロシージャ）
'//  機能概要  : 汎用メイル受信処理を行う
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(V30.1.0.1) 2014-06-04 REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    Dim intStatus As Integer         '受信メールチェック結果
    Dim iResponse As Integer
    Dim iLoopCnt As Integer          ' ループ
    
    On Error Resume Next
    
    'メールを受信する。
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
    '受信メールがあれば、メールＩＤ毎の処理をする。
        Select Case udtReadMail.udtlHeader.dwId        'メールＩＤ
            Case ML_ID_PROEND_ORD
                '「プロセス終了指示」を受信した場合、
                '「プロセス終了指示受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
                'プロセスの終了処理を行う
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示」を受信した場合
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '表示元画面（保守データ収集画面）をアクティブ表示する。
                'AppActivate frmInputMstData.Caption, False      'EG20 V30.1.0.1 DEL
                'EG20 V30.1.0.1 ADD START
                AppActivate frmEkiDataSubGate.Caption, False
                pfFormActive (frmEkiDataSubGate.hwnd)
                'EG20 V30.1.0.1 ADD END
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '「保守操作卓プログラム送信要求」を受信した場合
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
                    Kill gstrFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
                    'プログレスバーを消去する
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
                    Call SetEnableTrue
                Else
                    Call pfuncInstallEkiSettei
                End If
            Case Else
                 'その他のメールを受信した場合
                 '「メールID不正」ログ出力
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_SUBGATE_GAMEN_END, 0)
    
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDisp
'//  機能名称  : 画面再描画処理
'//  機能概要  : 画面を再描画する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim strFileName          As String          'ファイル名
    Dim iLoopCnt             As Integer         'ループカウンタ
    Dim iLoopCnt2            As Integer         'ループカウンタ EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
    Dim bRet                 As Boolean         '関数戻り値
    Dim strKubun             As String          '区分
    Dim strIniData           As String          'INIファイル設定値
    Dim nCornerIndex         As Integer         ' コーナ選択状態

    'エラールーチンを宣言
    On Error Resume Next

' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    If CmbCornerName.ListIndex < 0 Then
        Exit Sub
    End If
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

    '初期値設定
    strFileName = ""                            'ファイル名
    cmbGoki.Enabled = False                     '号機コンボボックス選択不可設定
    CmbCornerName.Enabled = False               ' コーナ選択部選択不可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    LblEkiName.Caption = TITOL_EKI_NAME         '駅名ラベル初期化
    
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
    '画面レイアウト変更につき、コーナと号機のコンボボックスは不要になった。
    cmbGoki.Visible = False
    CmbCornerName.Visible = False
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
    
    '----------------------------------------------------
    'グリッドタイトル設定
    '----------------------------------------------------
    Call sDispGridTitol
    
    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear

        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '駅名ラベル更新
    '----------------------------------------------------
    LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo(NotEkiVer)
    
    '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイル
    strFileName = Dir(EKI_TUDO_CHK_SUBGATE_FILE)
    
    'ファイルが存在した場合
    If strFileName <> "" Then
    
        'グリッドデータ部設定
'        Call sDispDataSet(cmbGoki.ListIndex + 1)                                   ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'        nCornerIndex = CmbCornerName.ListIndex
'        Call sDispDataSet(cmbGoki.ListIndex + 1, nCornerIndex + 1)
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        '大分類:5の駅都度データはコーナ指定ではなくなったため、コーナは0固定とする。（駅都度のコーナは0で検索する）
        '1～32号機
        For iLoopCnt = 0 To 31
            '項目①～⑥
            For iLoopCnt2 = 0 To 5
                Call sDispDataSet(iLoopCnt + 1, 0, iLoopCnt2 + 1)
            Next
        Next
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_SUBGATE_IMAGE, 0)
        
        'グリッドデータ部クリア処理
        Call sDispDataClear
    End If


End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDispGridTitol
'//  機能名称  : グリッドタイトル部設定処理
'//  機能概要  : グリッドの初期値、タイトルを設定する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28 CODED BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19 CODED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispGridTitol()
    
    Dim ColCount                As Integer         ' カラムカウンタ
    Dim RowCount                As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    '設定中は非表示に設定
    GridIni.Visible = False
    
    'グリッドタイトル設定
    With GridIni
    
        '----------------------------------
        'グリッドの初期化
        '----------------------------------
        .Clear
        
        '----------------------------------
        'グリッドセル数設定
        '----------------------------------
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'        .Rows = 18
'        .Cols = 8
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'        .Rows = 5
'        .Cols = 5
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'EG20 V30.1.0.1 DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
''EG20 V30.1.0.1 ADD START
'        .Rows = 2
'        .Cols = 8
''EG20 V30.1.0.1 ADD END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        .Rows = 33
        .Cols = 9
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
        
        '----------------------------------
        'グリッド幅設定
        '----------------------------------
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】DEL START
'        .ColWidth(0) = 900
'        .ColWidth(1) = 700
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        .ColWidth(0) = 700
        .ColWidth(1) = 700
        .ColWidth(2) = 700
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'        For ColCount = 2 To (.Cols - 1)
'            .ColWidth(ColCount) = 1675
'        Next
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        For ColCount = 3 To (.Cols - 1)
            .ColWidth(ColCount) = 1550
        Next
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
        
        '----------------------------------
        'タイトル設定
        '----------------------------------
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        '号機設定
        .Col = 0
        .Row = 0: .Text = "号機"
        .CellAlignment = flexAlignCenterCenter
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END

        '区分設定
        '.Col = 1    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】DEL
        .Col = 2     'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】ADD
        .Row = 0: .Text = "区分"
        .CellAlignment = flexAlignCenterCenter
        For RowCount = 1 To (.Rows - 1)
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】ADD START
            '縦方向の固定表示部を設定する
            '号機設定（１～３２）
            .Col = 0
            .Row = RowCount: .Text = RowCount
            .CellAlignment = flexAlignCenterCenter
            
            '自社・他社設定（北陸新幹線の駅都度は自社のみ）
            .Col = 1
            .Row = RowCount: .Text = "自社"
            .CellAlignment = flexAlignCenterCenter
            
            '区分（北陸新幹線では区分は統合のみ）
            .Col = 2
            .Row = RowCount: .Text = "統合"
            .CellAlignment = flexAlignCenterCenter
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】ADD END
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】DEL START
'            If RowCount = 1 Then
'                '自社設定
'                .Col = 0
'                .Row = RowCount: .Text = "自社"
'                .CellAlignment = flexAlignLeftCenter
'
'            Else
'                '他社設定
'                .Col = 0
'                .Row = RowCount: .Text = "他社" & RowCount - 1
'                .CellAlignment = flexAlignLeftCenter
'            End If
'
'            .Col = 1
'            .Row = RowCount
''            .Text = "監視"                 ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'            .Text = "統合"                  ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'            .CellAlignment = flexAlignCenterCenter
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
        Next

        .RowHeight(0) = 500
    
    End With

    GridIni.Visible = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDispDataClear
'//  機能名称  : グリッドデータ部クリア処理
'//  機能概要  : グリッドデータ部をクリアする
'//
'//              型        名称      意味
'//  引数      : Integer   intStartRow  開始行位置
'//              Integer   intEndRow    終了行位置
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear()
    
    Dim iLoopCnt             As Integer         'ループカウンタ
    Dim ColCount             As Integer         'ループカウンタ
    
    'エラールーチンを宣言
    On Error Resume Next

    '設定中は非表示に設定
    GridIni.Visible = False
    
    'グリッド初期化
    With GridIni

        For iLoopCnt = 1 To (.Rows - 1)

            '項目設定
            'For ColCount = 2 To (.Rows - 1) 'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
            For ColCount = 3 To (.Rows - 1) 'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
                .Col = ColCount
                .Text = ""
                .CellAlignment = flexAlignLeftCenter
            Next

        Next

    End With

    GridIni.Visible = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDispDataSet
'//  機能名称  : グリッドデータ部設定処理
'//  機能概要  : グリッドデータ部を設定する
'//
'//              型        名称         意味
'//  引数      : Integer   iBunrui_Dai  大分類
'//            : Integer   iCorner      コーナ  ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応
'//                 【HKRK_Kansi07_003_01】、【HKRK_Kansi07_008_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sDispDataSet(iGoki As Integer)                             ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer)          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加 'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL
Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer, iKomoku As Integer)    'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
    
    Dim intFileNumber       As Integer                      ' ファイルポインタ
    Dim iLoopCnt            As Integer                      ' ループカウンタ
    Dim ColCount            As Integer                      ' カラムカウンタ
    Dim RowCount            As Integer                      ' 行カウンタ
    
    Dim strBunrui_Dai       As String                       ' 大分類
    Dim strBunrui_Tyu       As String                       ' 中分類
    Dim strBunrui_Sho       As String                       ' 小分類
    Dim strKomoku           As String                       ' 項目
    Dim strKubun            As String                       ' 区分
    Dim strData             As String                       ' 設定値
    Dim strSetShosai        As String                       ' 設定値詳細
    
    Dim strDispData         As String                       ' 表示データ
    Dim strCorner           As String                       ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim iCmpCorner          As Integer                      ' コーナ    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
    'エラールーチンを宣言
    On Error Resume Next

    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイルをオープンする。
    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
    
    GridIni.Visible = False             '設定中は非表示に設定

    ColCount = START_DATA_COL_INDEX     'データ設定のスタートカラムインデックス
    RowCount = 1                        'データ設定のスタート行インデックス
    Do While Not EOF(intFileNumber)
        '１ 行読み込み
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, _
'                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strCorner, _
                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

        '号機番号チェック
        If CStr(iGoki) = strBunrui_Tyu Then
            If iKomoku = CInt(strBunrui_Sho) Then       'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
                ' コーナ判定追加
                ' 選択したコーナのレコードを採用する
                iCmpCorner = CInt(strCorner)
                If (iCorner = iCmpCorner) Then
        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
                    'グリッド設定
                    With GridIni
                
                        'カラムインデックス設定
                        '.Col = ColCount                    'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】DEL
                        .Col = ColCount + (iKomoku - 1)     'EG20 V30.3.0.1 【HKRK_Kansi07_007_01】ADD
        
                       'タイトル設定
                        If (strKomoku <> "") Then
                            .Row = 0
                            .Text = strKomoku
                            .CellAlignment = flexAlignLeftCenter
                        End If
        
                        '項目設定
                        '.Row = RowCount        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL
                        .Row = iGoki            'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
                        .Text = Format(pfDispIniData(.Text, strData, strKubun), "000")
                        .CellAlignment = flexAlignLeftCenter
                        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD START
                        '駅都度データ1レコード分の設定値をセルにセットしたので、一旦終わらす。
                        Exit Do
                        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD END
                        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL START 上記Exit Doにロジック変更したため不要になった。
'                        ColCount = ColCount + 1
'                        If ColCount > MAX_DATA_COL_INDEX Then
'                         ColCount = START_DATA_COL_INDEX
'                         RowCount = RowCount + 1
'                        End If
                        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL END
        
                    End With
                
                End If          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            End If          'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】ADD
        End If
    
    Loop

    GridIni.Visible = True
    
    'ファイルをクローズする。
    Close #intFileNumber

    '号機コンボボックス選択可設定
    cmbGoki.Enabled = True
    CmbCornerName.Enabled = True               ' コーナ選択部選択不可      ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

    Exit Sub

'エラー処理
Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
    
    'グリッドタイトル設定
    Call sDispGridTitol
    
    'グリッドデータ部クリア処理
    Call sDispDataClear

    GridIni.Visible = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmbGoki_Click
'//  機能名称  : 号機選択処理
'//  機能概要  : グリッドデータを再設定する
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmbGoki_Click()
    
    Dim iIndex          As Integer                  'インデックス
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    Call SetEnableFalse

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_SUBGATE_GAMEN_GOKI_SELECT, 0)
    
    '画面表示処理
    Call sDisp

    '全ボタンを押下可とする。
    Call SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmdMenu_Click
'//  機能名称  : 「駅設定出力」「駅設定入力」「駅設定テキスト出力」
'//              「媒体取外」釦押下処理
'//  機能概要  : 各釦名称処理を行う。
'//
'//              型        名称         意味
'//  引数      : なし
'//
'//              型        値           意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmdMenu_Click(Index As Integer)
  
    Dim bUnlock             As Boolean          ' ロック解除フラグ      ' EG20 V3.0.0.2 追加
  
  'エラールーチンを宣言
  On Error Resume Next
    
  '全ボタンを押下不可とする。
  Call SetEnableFalse
    
' EG20 V3.0.0.2 追加開始
' 押下した釦に応じてロック解除を制限する
' ※メール受信を待つため
    bUnlock = True
' EG20 V3.0.0.2 追加終了
    
  Select Case Index
       Case 0                                  '駅設定出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_OUTPUT, 0)
            
            '駅設定出力処理
            Call sEkiSetteiOutPut
        
        Case 1                                  '駅設定入力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_INPUT, 0)
            
            '駅設定入力処理
            Call sInstolEkiSettei
        
            bUnlock = False                     ' EG20 V3.0.0.2 追加
        
        Case 2                                  '駅設定テキスト出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_DISP_TEXT, 0)
            
            '駅設定テキスト出力処理
            Call sDispTextEkiDataNow
        
        Case 3                                  '媒体取外
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
            
            '媒体取外処理
            Call pfRemove(Me)
 End Select

  '全ボタンを押下可とする。
' EG20 V3.0.0.2 追加開始
    If bUnlock = True Then
        Call SetEnableTrue
    End If
' EG20 V3.0.0.2 追加終了
'    Call SetEnableTrue                 ' EG20 V3.0.0.2 削除

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sEkiSetteiOutPut
'//  機能名称  : 「駅設定出力」釦押下時処理
'//  機能概要  : 現在駅設定ファイルを外部媒体に出力する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sEkiSetteiOutPut()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
    Dim iResponse            As Integer         'MsgBox戻り値

    'エラールーチンを宣言
    On Error Resume Next
    iResponse = MsgBox("選択されている駅の現在の駅都度データ１駅分を出力します。" & Chr(vbKeyReturn) & _
                        "よろしいですか？", _
                        vbOKCancel + vbQuestion, _
                        "駅設定出力確認")

    If iResponse = vbCancel Then Exit Sub

    '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '異常終了
        MsgBox "媒体出力するデータがありません。", _
                vbOKOnly + vbExclamation, _
                 "データ無警告"
        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '媒体出力処理
    '----------------------------------------------------
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
'        FileCopy EKI_SETTI_FILE, sWriteDir & Dir(EKI_SETTI_FILE)                   ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
        FileCopy EKI_SETTI_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(EKI_SETTI_FILE)    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
       '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定出力結果")
    
    End If
    
  Exit Sub
 
COPY_ERROR:

    Select Case Err.Number
        Case 61 ' 媒体出力空き容量不足
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case 71 ' 媒体なし
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_NOT_DISK, 0)
        Case Else
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

    iResponse = MsgBox("異常終了しました", vbOKOnly + vbCritical, "駅設定出力結果")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sInstolEkiSettei
'//  機能名称  : 「駅設定入力」釦押下時処理
'//  機能概要  : 外部媒体から現在駅設定ファイルインストールする
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.76修正対応】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sInstolEkiSettei()

    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bRet                As Boolean          '関数戻り値
    Dim strFileName         As String           '媒体ファイル名

    Dim objFso As New FileSystemObject          'ファイルシステムオブジェクト

    Dim lResult             As Long             ' 処理結果

    'エラールーチンを宣言
    On Error Resume Next
    iResponse = MsgBox("駅都度データ１駅分をインストールします。" & Chr(vbKeyReturn) & _
                        "よろしいですか？", _
                        vbOKCancel + vbQuestion, _
                        "駅設定入力確認")
    If iResponse = vbCancel Then
        Set objFso = Nothing
        Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
        Exit Sub
    End If
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
    ' 拡張子を設定
    CommonDialog1.Filter = "ＣＳＶ（カンマ区切り）(*.csv)|*.csv|"
    ' ファイル選択画面を開く
    CommonDialog1.ShowOpen
    ' 選択したファイル名を取得
    strFileName = CommonDialog1.FileName
    
    Call ChDrive("D")  'V2.5.0.1 ADD

    'ファイル存在チェック
    If strFileName <> "" Then
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL Start
'        ' 出力先ファイル名を保存
'        gstrFileName = strFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL End
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
        ' 出力先ファイル名を保存
        gstrFileName = PATH_HOSHUWRK_EKI_INFO
        '一時保存フォルダにデータをコピーし読取専用を解除する
        If pfChangeAttrNormal(strFileName, PATH_HOSHUTMP_EKI_INFO, PATH_HOSHUWRK_EKI_INFO) = False Then
            Kill gstrFileName
            '一時保存フォルダを削除する
            psDeleteFolder PATH_HOSHUTMP
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' 異常終了
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
            Call SetEnableTrue
            Exit Sub
        End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End

        ' //////////////////////////////////////////////
        ' // 操作卓プログラム処理
        ' //////////////////////////////////////////////
        lResult = pubfuncTakuProgramData(2, gstrFileName)
        If lResult = 0 Then
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
            Kill gstrFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' 異常終了
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
            Call SetEnableTrue
            Exit Sub
        ElseIf lResult = 1 Then
            ' メール送信中
            ' ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
            Exit Sub
        End If


        ' //////////////////////////////////////////////
        ' // 統合監視盤非動作中のためメール応答を待たずに
        ' // 即時更新
        ' //////////////////////////////////////////////
        bRet = pfuncInstallEkiSettei

    End If
    Call SetEnableTrue                      ' EG20 V5.0.2.1【結合TR-No.76修正対応】追加
End Sub

' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]削除開始（全体見直し）
'Private Sub sInstolEkiSettei()
'
'    Dim iResponse           As Integer          'MsgBox戻り値
'    Dim bRet                As Boolean          '関数戻り値
'    Dim lErrCode            As Long             'エラーコード
'    Dim strFileName         As String           '媒体ファイル名
'
'    Dim iRet                    As Integer      'メッセージボックス戻り値
'    Dim lSekuta                 As Long         'セクタ（クラスタ当り）
'    Dim lByte                   As Long         'バイト数（セクタ当り）
'    Dim lKurasuta               As Long         'フリークラスタ数
'    Dim lDrive                  As Long         'ドライブのクラスタ数（合計）
'    Dim strDrive                As String       'ドライブ
'    Dim bSysChange              As Boolean      'システム設定処理戻り値
'    Dim bUpData                 As Boolean      '画面更新処理戻り値
'    Dim iLoopCnt                As Integer      'ループカウンタ
'
'    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
'
'    Dim lResult             As Long             ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加
'
'    'エラールーチンを宣言
'    On Error Resume Next
'    iResponse = MsgBox("駅都度データ１駅分をインストールします。" & Chr(vbKeyReturn) & _
'                        "よろしいですか？", _
'                        vbOKCancel + vbQuestion, _
'                        "駅設定入力確認")
'    If iResponse = vbCancel Then
'        Set objFso = Nothing
'        Exit Sub
'    End If
'    '取得ファイル名を初期化
'    CommonDialog1.FileName = ""
'    '初期ディレクトリを設定
'    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'フォルダ選択画面デフォルトパス１が存在するか
'        '存在するため、デフォルトパス１（H:）を設定
'        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
'    Else
'        '存在しないため、デフォルトパス２（C:）を設定
'        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
'    End If
'    Set objFso = Nothing
'    '拡張子を設定
'    CommonDialog1.Filter = "ＣＳＶ（カンマ区切り）(*.csv)|*.csv|"
'    'ファイル選択画面を開く
'    CommonDialog1.ShowOpen
'    '選択したファイル名を取得
'    strFileName = CommonDialog1.FileName
'
'    Call ChDrive("D")
'
'    'ファイル存在チェック
'    If strFileName <> "" Then
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'        'プログレスバーを表示する
'        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'        '現在駅設定データインストール処理
'        bRet = dllInstolEkiDataNow(strFileName, EKI_SETTI_FILE, lErrCode)
'
'        If bRet = False Then
'
'            '異常ログ出力
'            Call pfOutPutErrLog(lErrCode)
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'            'プログレスバーを消去する
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'            '異常終了
'            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
'
'        Else
'            '----------------------------------------------------
'            'コンピュータ名、ネットワーク変更処理
'            '----------------------------------------------------
'            bSysChange = True
'            bUpData = True
'            bSysChange = pfNetWorkChng(Me)
'             'ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'           '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイル作成
'            bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
'            If bRet = False Then
'                '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイル削除
'                Kill EKI_TUDO_CHK_SUBGATE_FILE
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'                'プログレスバーを消去する
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'                '異常ログ出力
'                Call pfOutPutErrLog(lErrCode)
'                bUpData = False
'            End If
'
'' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加開始
'            ' //////////////////////////////////////////////
'            ' // 操作卓プログラム処理
'            ' //////////////////////////////////////////////
'             lResult = pubfuncTakuProgramData(2)
'             If lResult = 0 Then
'                'プログレスバーを消去する
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'                ' 異常終了
'                iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "設定反映結果")
'                Call SetEnableTrue
'                Exit Sub
'             ElseIf lResult = 1 Then
'                ' メール送信中
'                ' ログ出力
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'
'                Exit Sub
'             End If
'' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加終了
'
'            '号機コンボボックス初期値
'            cmbGoki.Clear
'            For iLoopCnt = 0 To 15
'                    cmbGoki.AddItem iLoopCnt + 1 & "号機"
'            Next
'            cmbGoki.ListIndex = 0
'
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'            'コーナ設定コンボボックスの初期化処理
'            Call InitCornerComboBox
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'            'プログレスバーを消去する
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'            If bSysChange = True And bUpData = True Then
'            '正常終了
'            iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定入力結果")
'            End If
'        End If
'    End If
'
'End Sub
' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]削除終了（全体見直し）

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDispTextEkiDataNow
'//  機能名称  : 「駅設定テキスト出力」釦押下時処理
'//  機能概要  : 現在駅設定ファイルをテキスト表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi07_003_01】、【HKRK_Kansi07_008_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispTextEkiDataNow()

    Dim strFileName          As String          'ファイル名
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim lRetVal              As Long            '戻り値
    Dim sCommand             As String          'コマンド文字列
    Dim sWriteDir            As String          '書き込み先フォルダ名
    Dim intFileNumber        As Integer         'ファイルポインタ
    Dim ColCount             As Integer         'カラムカウンタ
    Dim RowCount             As Integer         'ループカウンタ
    Dim TypeCount            As Integer         'ループカウンタ
    Dim sData                As String          '入力用文字列
    Dim strData_Kansi()      As String          '監視盤情報保存配列
    Dim strData_Ldu()        As String          'LDU情報保存配列
    Dim iLength              As Integer         '改行コード検索用（長さ）
    Dim iLeft                As Integer         '改行コード検索用（先頭）
    Dim iRight               As Integer         '改行コード検索用（終端）
    
    Dim ReadFileSettei()     As SUBGATE_IMAGE_FILE          'ファイル読込用構造体
    Dim OutFileData1()       As SUBGATE_IMAGE_FILE_ONE_ROW  'ファイル出力用構造体
    Dim OutFileData2()       As SUBGATE_IMAGE_FILE_ONE_ROW  'ファイル出力用構造体
    Dim strOutDefTbl()       As SUBGATE_OUT_DEF_TBL         '出力情報テーブル
    Dim i                    As Integer             'ループカウンタ１
    Dim j                    As Integer             'ループカウンタ２
    Dim k                    As Integer             'ループカウンタ３
    Dim strLineCount         As String              '行数カウンタ
    Dim fso                  As New FileSystemObject        'ファイルシステムオブジェクト
    Dim FsoTS                As TextStream
    Dim strSaveFileName      As String          ' 保存ファイル名        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim szCornerName         As String          ' コーナ名称            ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    Dim nNullIndex           As Integer         ' 文字数ワーク          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'エラールーチンを宣言
    On Error Resume Next

    '書き込み先ファイル選択
    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    If sWriteDir = "" Then
       'フォルダ選択画面「取消」釦押下時は処理終了
       Exit Sub
    End If

    '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then

        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)

        '異常終了
        MsgBox "テキスト表示するデータがありません。", _
                vbOKOnly + vbExclamation, _
                 "データ無警告"
        Exit Sub

    End If

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了

    On Error GoTo OUTPUT_ERROR
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】、【HKRK_Kansi07_008_01】ADD START
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    '現在駅設定ファイルをオープンする
    Open PATH_WORK & EKI_SETTI_SUBGATE_FILE For Output As #intFileNumber
    
    'タイトル表示
    Print #intFileNumber, "設置駅　　：" & Trim(pfGetEkiNameInfo(NotEkiVer))
    Print #intFileNumber, "【エンコードコーナ号機情報定義】"
    
    'グリッドタイトル設定
    With GridIni
    
        '行数分ループさせる
        For RowCount = 0 To .Rows - 1
            'sData初期化
            sData = ""
            '各項目表示
            If RowCount = 0 Then
                For ColCount = 0 To .Cols - 1
                    .Col = ColCount
                    .Row = RowCount
                    
                    If ColCount <> .Cols - 1 Then
                        sData = sData & Replace(.Text, " ", "") & ","
                    Else
                        sData = sData & Replace(.Text, " ", "")
                    End If
                Next
                Print #intFileNumber, sData
            Else
                .Row = RowCount
                .Col = 0
                
                '再定義
                
                '項目分ループする
                For ColCount = 0 To .Cols - 1
                    .Col = ColCount
                    .Row = RowCount
                   
                    '設定値取得
                    If ColCount <> .Cols - 1 Then
                        sData = sData & .Text & ","
                    Else
                        sData = sData & .Text
                    End If
                Next
                Print #intFileNumber, sData
            End If
        Next
    End With
    
    'ファイルをクローズする
    Close #intFileNumber
    
    ' コーナ名称の付加
    'strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & szCornerName & "_" & EKI_SETTI_SUBGATE_FILE
    strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & EKI_SETTI_SUBGATE_FILE
    '一時ファイルを媒体にコピーする
    Call FileCopy(PATH_WORK & EKI_SETTI_SUBGATE_FILE, sWriteDir & strSaveFileName)
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】、【HKRK_Kansi07_008_01】DEL END
'EG20 V30.3.0.1 媒体出力フォーマット大幅見直しに付き削除【HKRK_Kansi07_003_01】、【HKRK_Kansi07_008_01】DEL START
'    '///////////////////////////////////////////////////////////////////////////
'    '機器構成データ（エンコードコーナ号機情報定義）イメージファイル情報読み込み
'    '///////////////////////////////////////////////////////////////////////////
'    'ファイル番号取得
'    intFileNumber = FreeFile
'
'    'CSVファイルオープン
'    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
'
'    'CSVファイル行数カウント（ファイル終端までループを繰り返す）
''    Do While Not EOF(1)                                    ' EG20 V3.3.0.1削除
'    Do While Not EOF(intFileNumber)                         ' EG20 V3.3.0.1追加
'        Line Input #intFileNumber, strLineCount
'        j = j + 1
'    Loop
'
'    'CSVファイルクローズ
'    Close #intFileNumber
'
'    'ファイル番号取得
'    intFileNumber = FreeFile
'
'    '再設定
'    ReDim ReadFileSettei(j) As SUBGATE_IMAGE_FILE        'ファイル読込用エリア
'
'    'CSVファイルオープン
'    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
'
'    'リスト表示分読み込み（ファイル終端までループを繰り返す）
'    For i = 0 To j - 1
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
''        Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, _
''         ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'        Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, ReadFileSettei(i).sCorner, _
'         ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'    Next
'
'    'CSVファイルクローズ
'    Close #intFileNumber
'
'    '再設定
'    ReDim OutFileData1(ONE_PARAGRAPH_OUTPUT_ROW) As SUBGATE_IMAGE_FILE_ONE_ROW         'ファイル出力用構造体
'    ReDim OutFileData2(ONE_PARAGRAPH_OUTPUT_ROW) As SUBGATE_IMAGE_FILE_ONE_ROW         'ファイル出力用構造体
'
'    '各項目設定値を出力用構造体に変換
'
'    ReDim strOutDefTbl(ONE_PARAGRAPH_OUTPUT_ROW) As SUBGATE_OUT_DEF_TBL                '出力情報テーブル
'    strOutDefTbl(0).iRow = 0
'    strOutDefTbl(0).iNoStart = GOKI_JISHA_START
'    strOutDefTbl(0).iNoEnd = GOKI_JISHA_END
'
'    For RowCount = 1 To ONE_PARAGRAPH_OUTPUT_ROW
'        strOutDefTbl(RowCount).iRow = RowCount
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
''        strOutDefTbl(RowCount).iNoStart = strOutDefTbl(RowCount - 1).iNoStart + 6
''        strOutDefTbl(RowCount).iNoEnd = strOutDefTbl(RowCount - 1).iNoEnd + 6
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'' EG20 V30.1.0.1 DEL START
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
''        strOutDefTbl(RowCount).iNoStart = strOutDefTbl(RowCount - 1).iNoStart + 3
''        strOutDefTbl(RowCount).iNoEnd = strOutDefTbl(RowCount - 1).iNoEnd + 3
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
''EG20 V30.1.0.1 DEL END
''EG20 V30.1.0.1 ADD START
'        strOutDefTbl(RowCount).iNoStart = strOutDefTbl(RowCount - 1).iNoStart + 6
'        strOutDefTbl(RowCount).iNoEnd = strOutDefTbl(RowCount - 1).iNoEnd + 6
''EG20 V30.1.0.1 ADD END
'
'    Next
'
'    '1号機～8号機
'    For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'        '社局を設定
'        If RowCount = 0 Then
'            OutFileData1(RowCount).sShakyoku = "自社"
'        Else
'            OutFileData1(RowCount).sShakyoku = "他社" & StrConv(CStr(RowCount), vbWide)
'        End If
'
'        '区分を設定
''        OutFileData1(RowCount).sKubun = "監視"         ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'        OutFileData1(RowCount).sKubun = "統合"          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'
'        '設定値を設定
'        If RowCount = strOutDefTbl(RowCount).iRow Then
'            'i = 0  EG20 V30.0.1.1 DEL
'            i = CmbCornerName.ListIndex * MAX_GATE_NO * GOKI_JISHA_END  ' コーナ毎に出力する  EG20V30.0.1.1 ADD
'
'            Do While (i < j)
'             If (CInt(ReadFileSettei(i).sGoki) <= ONE_PARAGRAPH_OUTPUT_GOKI) And _
'                (CInt(ReadFileSettei(i).sNo) >= strOutDefTbl(RowCount).iNoStart) And _
'                (CInt(ReadFileSettei(i).sNo) <= strOutDefTbl(RowCount).iNoEnd) Then
'
'                 If (CInt(ReadFileSettei(i).sGoki) = ONE_PARAGRAPH_OUTPUT_GOKI) And _
'                    (CInt(ReadFileSettei(i).sNo) = strOutDefTbl(RowCount).iNoEnd) Then
'
'                     OutFileData1(RowCount).sSettei = OutFileData1(RowCount).sSettei + _
'                                                      Format(ReadFileSettei(i).sSettei, "000") & vbCrLf
'                    Exit Do
'
'                 Else
'
'                     OutFileData1(RowCount).sSettei = OutFileData1(RowCount).sSettei + _
'                                                       Format(ReadFileSettei(i).sSettei, "000") & ","
'                 End If
'
'             End If
'
'             i = i + 1
'
'            Loop
'
'        End If
'
'   Next
'
'    '9号機～16号機
'   For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'       '社局を設定
'       If RowCount = 0 Then
'           OutFileData2(RowCount).sShakyoku = "自社"
'       Else
'           OutFileData2(RowCount).sShakyoku = "他社" & StrConv(CStr(RowCount), vbWide)
'       End If
'
'       '区分を設定
''       OutFileData2(RowCount).sKubun = "監視"          ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
'       OutFileData2(RowCount).sKubun = "統合"           ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
'
'       '設定値を設定
'       If RowCount = strOutDefTbl(RowCount).iRow Then
'           'i = 0       'EG20 V30.0.1.1 DEL
'           i = (CmbCornerName.ListIndex * MAX_GATE_NO * GOKI_JISHA_END) + 8  ' コーナ毎に出力する  EG20V30.0.1.1 ADD
'
'           Do While (i < j)
'            If (CInt(ReadFileSettei(i).sGoki) > ONE_PARAGRAPH_OUTPUT_GOKI) And _
'               (CInt(ReadFileSettei(i).sGoki) <= (ONE_PARAGRAPH_OUTPUT_GOKI * 2)) And _
'               (CInt(ReadFileSettei(i).sNo) >= strOutDefTbl(RowCount).iNoStart) And _
'               (CInt(ReadFileSettei(i).sNo) <= strOutDefTbl(RowCount).iNoEnd) Then
'
'                If (CInt(ReadFileSettei(i).sGoki) = (ONE_PARAGRAPH_OUTPUT_GOKI * 2)) And _
'                   (CInt(ReadFileSettei(i).sNo) = strOutDefTbl(RowCount).iNoEnd) Then
'
'                    OutFileData2(RowCount).sSettei = OutFileData2(RowCount).sSettei + _
'                                                     Format(ReadFileSettei(i).sSettei, "000") & vbCrLf
'                   Exit Do
'
'                Else
'
'                    OutFileData2(RowCount).sSettei = OutFileData2(RowCount).sSettei + _
'                                                     Format(ReadFileSettei(i).sSettei, "000") & ","
'                End If
'
'            End If
'
'                i = i + 1
'
'           Loop
'
'       End If
'
'    Next
'
'
'    '///////////////////////////////////////////////////////////////////////////
'    '機器構成データ（エンコードコーナ号機情報定義）ファイル出力処理
'    '///////////////////////////////////////////////////////////////////////////
'    '一時ファイルを作る
'    Set FsoTS = fso.CreateTextFile(PATH_WORK & EKI_SETTI_SUBGATE_FILE, True)
'
'    'タイトル出力
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'    ' コーナ名称の付加
'    nNullIndex = InStr(gstrCornerName(CmbCornerName.ListIndex), Chr(0))
'    If nNullIndex <> 0 Then
'        szCornerName = Left(gstrCornerName(CmbCornerName.ListIndex), nNullIndex - 1)
'    Else
''        szCornerName = ""                                          ' EG20 V3.3.0.1削除
'        szCornerName = gstrCornerName(CmbCornerName.ListIndex)      ' EG20 V3.3.0.1追加
'    End If
'
'    FsoTS.Write ("設置駅　　：" & Trim(pfGetEkiNameInfo(NotEkiVer)) & vbCrLf)
'    FsoTS.Write ("設置コーナ：" & szCornerName & vbCrLf)
'    FsoTS.Write (vbCrLf)
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'    FsoTS.Write ("【エンコードコーナ号機情報定義】" & vbCrLf)
'
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
''    '項目タイトル出力
''    FsoTS.Write ("項目,区分,1号機,,,,,,2号機,,,,,,3号機,,,,,,4号機,,,,,,5号機,,,,,,6号機,,,,,,7号機,,,,,,8号機" & vbCrLf)
''
''    '項目タイトル出力
''    FsoTS.Write (",,")
''    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
''        FsoTS.Write ("①,②,③,④,⑤,⑥,")     '1～7号機
''    Next
''    FsoTS.Write ("①,②,③,④,⑤,⑥" & vbCrLf) '8号機
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'    '項目タイトル出力
'    'FsoTS.Write ("項目,区分,1号機,,,2号機,,,3号機,,,4号機,,,5号機,,,6号機,,,7号機,,,8号機" & vbCrLf)   'EG20 V30.1.0.1 DEL
'    FsoTS.Write ("項目,区分,1号機,,,,,,2号機,,,,,,3号機,,,,,,4号機,,,,,,5号機,,,,,,6号機,,,,,,7号機,,,,,,8号機" & vbCrLf)   'EG20 V30.1.0.1 ADD
'
'    '項目タイトル出力
'    FsoTS.Write (",,")
'    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
'        'FsoTS.Write ("①,②,③,")     '1～7号機     'EG20 V30.1.0.1 DEL
'        FsoTS.Write ("①,②,③,④,⑤,⑥,")     '1～7号機     'EG20 V30.1.0.1 ADD
'    Next
'    'FsoTS.Write ("①,②,③" & vbCrLf) '8号機   'EG20 V30.1.0.1 DEL
'    FsoTS.Write ("①,②,③,④,⑤,⑥" & vbCrLf) '8号機   'EG20 V30.1.0.1 ADD
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'
'    '各項目設定値出力
'    '1号機～8号機
'    For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'        FsoTS.Write (OutFileData1(RowCount).sShakyoku & ",")
'        FsoTS.Write (OutFileData1(RowCount).sKubun & ",")
'        FsoTS.Write (OutFileData1(RowCount).sSettei)
'    Next
'
'    '空行出力
'    FsoTS.Write (vbCrLf)
'
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
''    '項目タイトル出力
''    FsoTS.Write ("項目,区分,9号機,,,,,,10号機,,,,,,11号機,,,,,,12号機,,,,,,13号機,,,,,,14号機,,,,,,15号機,,,,,,16号機" & vbCrLf)
''
''    '項目タイトル出力
''    FsoTS.Write (",,")
''    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
''        FsoTS.Write ("①,②,③,④,⑤,⑥,")     '1～7号機
''    Next
''    FsoTS.Write ("①,②,③,④,⑤,⑥" & vbCrLf) '8号機
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'    '項目タイトル出力
'    'FsoTS.Write ("項目,区分,9号機,,,10号機,,,11号機,,,12号機,,,13号機,,,14号機,,,15号機,,,16号機" & vbCrLf)    'EG20　V30.1.0.1 DEL
'    FsoTS.Write ("項目,区分,9号機,,,,,,10号機,,,,,,11号機,,,,,,12号機,,,,,,13号機,,,,,,14号機,,,,,,15号機,,,,,,16号機" & vbCrLf)    'EG20 V30.1.0.1 ADD
'
'    '項目タイトル出力
'    FsoTS.Write (",,")
'    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
'        'FsoTS.Write ("①,②,③,")     '1～7号機    'EG20 V30.1.0.1 DEL
'        FsoTS.Write ("①,②,③,④,⑤,⑥,")     '1～7号機    'EG20 V30.1.0.1 ADD
'    Next
'    'FsoTS.Write ("①,②,③" & vbCrLf) '8号機   'EG20 V30.1.0.1 DEL
'    FsoTS.Write ("①,②,③,④,⑤,⑥" & vbCrLf) '8号機   'EG20 V30.1.0.1 ADD
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'
'    '9号機～16号機
'    For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'        FsoTS.Write (OutFileData2(RowCount).sShakyoku & ",")
'        FsoTS.Write (OutFileData2(RowCount).sKubun & ",")
'        FsoTS.Write (OutFileData2(RowCount).sSettei)
'    Next
'
'    'ファイルをクローズする。
'    FsoTS.Close
'
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
''    '一時ファイルを媒体にコピーする
''    Call FileCopy(PATH_WORK & EKI_SETTI_SUBGATE_FILE, sWriteDir & EKI_SETTI_SUBGATE_FILE)
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
'    ' コーナ名称の付加
'    strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & szCornerName & "_" & EKI_SETTI_SUBGATE_FILE
'    '一時ファイルを媒体にコピーする
'    Call FileCopy(PATH_WORK & EKI_SETTI_SUBGATE_FILE, sWriteDir & strSaveFileName)
'' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】、【HKRK_Kansi07_008_01】DEL END

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了

    '正常終了
    iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定テキスト出力結果")
    
    Exit Sub

OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing

    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, CREATE_FILE_ERROR, 0)
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
    '異常終了
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定テキスト出力結果")

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pfStartUpProc
'//  機能名称  : ファイル選択画面処理
'//  機能概要  : ファイル選択画面を表示し、選択されたファイル名を返す。
'//
'//              型        名称      意味
'//  引数      : String　　sDrive　　[IN]初期表示ドライブ名
'//  　　      : String　　sPattern　[IN]選択対象ファイル拡張子
'//  　　      : String　　sTitle　　[IN]画面表示ラベル
'//
'//              型        値        意味
'//  戻り値    :String　　　　　　　 [OUT]戻り値
'//                                      選択されたファイルパス:正常　""：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function pfFileSelection(sDrive As String, _
                                sPattern As String, _
                                sTitle As String) As String
                                
    Dim sWorkDrive As String                    'ワーク用初期表示ドライブ名

    'ドライブ異常処理を定義する。
    On Error GoTo Drive_Error
    
    sWorkDrive = sDrive                         '初期表示ドライブ名をワーク用にセットする。
    frmFil.filSelection.Pattern = sPattern      '選択対象拡張子をセットする。
    frmFil.lblFileSelection = sTitle            'サブタイトルをセットする。

Retry:
    frmFil.drvSelection.Drive = sWorkDrive      'ドライブをセットする。
    frmFil.dirSelection.Path = sWorkDrive & "\" 'ディレクトリをセットする。
    
    'ファイル選択画面を表示する。
    frmFil.Show 1
    
    '選択されたファイル名を返す。
    pfFileSelection = gstrMyPath
    
    Exit Function

'**ドライブ指定異常処理**
Drive_Error:

    If Left$(sWorkDrive, 1) = "H" Then
        'a:ドライブが異常なら、カレントドライブを表示させる。
        sWorkDrive = Left$(App.Path, 2)
        GoTo Retry
    End If
    
    'その他のドライブなら、ファイル選択なしで戻る。
    pfFileSelection = ""

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmdMoveGateGamen_Click
'//  機能名称  : 「自改画面へ」釦押下処理
'//  機能概要  : 駅都度データ確認(自改)画面を表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    :なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveGateGamen_Click()
   
    '全ボタンを押下不可とする。
    Call SetEnableFalse
    DoEvents
   
   '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)
    Unload Me
    Load frmEkiDataGate
    frmEkiDataGate.Show 1
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  関数名称  : SetEnableFalse
'//  機能名称  : 画面ロック処理
'//  機能概要  : 画面をロックする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    cmbGoki.Enabled = False
    CmdMenu(0).Enabled = False
    CmdMenu(1).Enabled = False
    CmdMenu(2).Enabled = False
    CmdMenu(3).Enabled = False
    CmdMoveGateGamen.Enabled = False
    CmdMoveEkiInfoGamen.Enabled = False
    cmdCancel.Enabled = False
    
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    CmbCornerName.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
    
    DoEvents
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  関数名称  : SetEnableTrue
'//  機能名称  : 画面ロック解除処理
'//  機能概要  : 画面のロックを解除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()

    Dim strFileName         As String           'ファイル名

    '初期値設定
    strFileName = ""                            'ファイル名
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下可とする。
    CmdMenu(0).Enabled = True
    CmdMenu(1).Enabled = True
    CmdMenu(2).Enabled = True
    CmdMenu(3).Enabled = True
    CmdMoveGateGamen.Enabled = True
    CmdMoveEkiInfoGamen.Enabled = True
    cmdCancel.Enabled = True

    'コンボボックスは条件によっては元々押下不可のため、画面表示用ファイルの有無で判定を行う
    strFileName = Dir(EKI_TUDO_CHK_SUBGATE_FILE)
    'ファイルが存在する場合
    If strFileName <> "" Then
        cmbGoki.Enabled = True
    End If

' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
    CmbCornerName.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了

    DoEvents
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : CmdMoveGateGamen_Click
'//  機能名称  : エンコードコーナ設定画面切替
'//  機能概要  : 駅都度データ確認（駅情報）画面を表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EGR HK1.1.0.1) 2011-05-11  CODED   BY [TCC] M.Kuroki
'//                 EG-R阪急　新規開発
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveEkiInfoGamen_Click()
    
    '全ボタンを押下不可とする。
    Call SetEnableFalse
    DoEvents
   
   '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIINFO_GAMEN_GO_BUTTOM, 0)

    '表示中画面アンロード
    Unload Me
                
    Load frmEkiData
    frmEkiData.Show 1

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : InitCornerComboBox
'//  機能名称  : コーナ設定コンボボックスの初期化処理
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub InitCornerComboBox()

    Dim intLoop   As Integer            ' ループカウンタ
    Dim strCorner As String             ' 文字列格納エリア
    
    On Error Resume Next
    
    ' /////////////////////////////////////////////////////
    ' // 初期化処理
    ' /////////////////////////////////////////////////////
    ' コーナ名称設定処理
    Call gsGetCornerName
    
    CmbCornerName.Clear
    For intLoop = 0 To 5
    
        '設定ありのコーナを活性にする
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
            strCorner = gstrCornerName(intLoop)
            CmbCornerName.AddItem strCorner
        End If
    Next intLoop
    CmbCornerName.ListIndex = 0

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pfuncInstallEkiSettei
'//  機能名称  : 駅設定インストール処理
'//  機能概要  :
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfuncInstallEkiSettei() As Boolean

    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bRet                As Boolean          '関数戻り値
    Dim lErrCode            As Long             'エラーコード

    Dim bSysChange              As Boolean      'システム設定処理戻り値　’V1.8.0.1　ADD
    Dim bUpData                 As Boolean      '画面更新処理戻り値　　　'V1.8.0.1　ADD
    Dim iLoopCnt                As Integer      'ループカウンタ

    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    Call SetEnableFalse

    pfuncInstallEkiSettei = True

    '現在駅設定データインストール処理
    bRet = dllInstolEkiDataNow(gstrFileName, EKI_SETTI_FILE, lErrCode)
    
    If bRet = False Then
            
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
            
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        pfuncInstallEkiSettei = False
        '異常終了
        iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
            
    Else
        
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        '----------------------------------------------------
        'コンピュータ名、ネットワーク変更処理
        '----------------------------------------------------
        bUpData = True
        bSysChange = True
        bSysChange = pfNetWorkChng(Me)
         'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
            
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
            
        '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイル作成
         bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
        If bRet = False Then
            '駅都度データ確認（エンコードコーナ号機情報定義）イメージファイル削除
            Kill EKI_TUDO_CHK_SUBGATE_FILE
               
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
            bUpData = False
            pfuncInstallEkiSettei = False
        End If

        '号機コンボボックス初期値
        cmbGoki.Clear
        'For iLoopCnt = 0 To 15 'EG20 V30.1.0.1 DEL
        For iLoopCnt = 0 To 31  'EG20 V30.1.0.1 ADD
                cmbGoki.AddItem iLoopCnt + 1 & "号機"
        Next
        cmbGoki.ListIndex = 0

        'コーナ設定コンボボックスの初期化処理
        Call InitCornerComboBox
            
        '画面表示処理
        Call sDisp
            
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
        If bSysChange = True And bUpData = True Then
            
            '正常終了
            iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定入力結果")
        End If
    End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    Kill gstrFileName
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
    gstrFileName = ""
    Call SetEnableTrue
End Function


