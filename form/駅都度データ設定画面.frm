VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEkisettei 
   BorderStyle     =   0  'なし
   Caption         =   "駅都度データ設定"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
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
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   11400
      Top             =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "  機器情報設定    画面へ戻る"
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
      Left            =   9480
      TabIndex        =   17
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdOut 
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
      Height          =   1095
      Left            =   6600
      TabIndex        =   15
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox txtDummy 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   10000
      Width           =   975
   End
   Begin VB.CommandButton cmdDataHanei 
      Caption         =   "設置駅変更"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "駅都度データ媒体 インストール"
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
      Left            =   360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "▲"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   2
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdPageUp 
      Caption         =   "▲ ▲"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPageDown 
      Caption         =   "▼  ▼"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   5
      Top             =   6180
      Width           =   1215
   End
   Begin VB.ListBox LstStation 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2225
      Width           =   9135
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "駅都度データ設定"
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
      TabIndex        =   16
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblNo 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "No."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label lblStation 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "駅名"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   1860
      Width           =   6135
   End
   Begin VB.Label lblVer 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   " バージョン"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   1860
      Width           =   2295
   End
   Begin VB.Label lblZen 
      Caption         =   "Z9.Z9.Z9.Z9"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblZenTop 
      Caption         =   "全体バージョン"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblNow 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   9225
   End
   Begin VB.Label lblNowTop 
      Caption         =   "現在の設置駅"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmEkisettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：駅都度データ設定画面.frm
'//  パッケージ名：駅都度データ設定画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//                 ディスク情報取得位置変更
'//                 画面ロック／画面ロック解除処理追加
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                画面再前面表示修正(不具合修正)
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000       'メイルタイマのインターバル値

'パターン番号定義
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]変更開始
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除開始
'Private Const PtnZenVersion = "000000"      '全体バージョン
'Private Const PtnEkiName = "000001"         '駅名
'Private Const PtnEkiVersion = "000002"      '駅バージョン
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加開始
' 小分類を3桁に修正
Private Const PtnZenVersion = "0000000"      '全体バージョン
Private Const PtnEkiName = "0000001"         '駅名
Private Const PtnEkiVersion = "0000002"      '駅バージョン
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加終了
' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]変更終了
Private gstrFileName        As String                       ' 出力ファイル名    ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 駅都度データ設定画面(アクティブ時：イベントプロシージャ)
'//  機能概要  : 最前前表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
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
'//  機能名称  : 駅都度データ設定画面（エンコード）画面(ディアクティブ時)
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
'EG20 V2.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 駅都度データ設定画面(ロード時：イベントプロシージャ)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim bRet As Boolean
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_GAMEN_START, 0)
    
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

    '初期画面表示
    bRet = sDisp
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdUp_Click
'//  機能名称  : 「▲」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdUp_Click()
    
    'エラールーチンを宣言
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex <= 0 Then
            LstStation.ListIndex = 1
            LstStation.ListIndex = 0
        Else
            LstStation.ListIndex = LstStation.ListIndex - 1
        End If
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdPageUp_Click
'//  機能名称  : 「▲▲」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdPageUp_Click()

    'エラールーチンを宣言
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex <= 18 Then
            LstStation.ListIndex = 1
            LstStation.ListIndex = 0
        Else
            LstStation.ListIndex = LstStation.ListIndex - 18
        End If
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdPageDown_Click
'//  機能名称  : 「▼▼」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdPageDown_Click()

    Dim iCnt As Integer
    
    'エラールーチンを宣言
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex >= LstStation.ListCount - 19 Or LstStation.ListIndex = -1 Then
            LstStation.ListIndex = LstStation.ListCount - 2
            LstStation.ListIndex = LstStation.ListCount - 1
        Else
            iCnt = LstStation.ListIndex
            LstStation.ListIndex = LstStation.ListCount - 1
            LstStation.ListIndex = iCnt + 18
        End If
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdDown_Click
'//  機能名称  : 「▼」釦押下時処理
'//  機能概要  : リストボックスのインデックスを動かす。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdDown_Click()

    'エラールーチンを宣言
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex < LstStation.ListCount - 1 Then
            If LstStation.ListIndex = -1 Then
                LstStation.ListIndex = LstStation.ListCount - 2
                LstStation.ListIndex = LstStation.ListCount - 1
            Else
                LstStation.ListIndex = LstStation.ListIndex + 1
            End If
        Else
            LstStation.ListIndex = LstStation.ListCount - 2
            LstStation.ListIndex = LstStation.ListCount - 1
        End If
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdInstall_Click
'//  機能名称  : 「駅都度データ媒体インストール」釦押下時処理
'//  機能概要  : 駅都度データ媒体より、インストールし駅情報を表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//                 ディスク情報取得位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdInstall_Click()

    Dim strFileName             As String       '媒体ファイル名
    Dim strZenVersion           As String       '全体バージョン
    Dim iRet                    As Integer      'メッセージボックス戻り値
    Dim lSekuta                 As Long         'セクタ（クラスタ当り）
    Dim lByte                   As Long         'バイト数（セクタ当り）
    Dim lKurasuta               As Long         'フリークラスタ数
    Dim lDrive                  As Long         'ドライブのクラスタ数（合計）
    Dim strDrive                As String       'ドライブ
    Dim bFrmShow                As Boolean
    Dim bRet                    As Boolean
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト  'V1.20.0.1 ADD
    
    'エラールーチンを宣言
    On Error Resume Next

'V1.12.0.1 ADD START
    '全ボタンを押下不可とする。
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_INSTOL, 0)
    
    'V1.20.0.1 DEL START
    ''ディスク情報を取得
''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD
    '
    'If lDrive = 0 Then
    '    strDrive = "d:"
    'Else
''        strDrive = "a:"    'V1.12.0.1 DEL
    '    strDrive = "H:"     'V1.12.0.1 ADD
    'End If
    '
    ''媒体ファイル名取得
    'strFileName = pfFileSelection(strDrive, "*.csv", "駅都度ﾌｧｲﾙ選択")
    'V1.20.0.1 DEL END
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
    CommonDialog1.Filter = "ＣＳＶ（カンマ区切り）(*.csv)|*.csv|"
    'ファイル選択画面を開く
    CommonDialog1.ShowOpen
    '選択したファイル名を取得
    strFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
    
    Call ChDrive("D")  'V2.5.0.1 ADD

    'ファイル存在チェック
    If strFileName <> "" Then
        
        '内部ファイルエラーのトラップ
        On Error GoTo Err_LOG

' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL Start
        'ワークフォルダに媒体ファイルをコピーする
'        Call FileCopy(strFileName, PATH_WORK_EKI_DATA_FILE)
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL End
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
        '一時保存フォルダにデータをコピーし読取専用を解除する
        If pfChangeAttrNormal(strFileName, PATH_HOSHUTMP_EKI_DATA, PATH_WORK_EKI_DATA_FILE) = False Then
            Goto Err_LOG
        End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
        '全体バージョン取得（ワーク）
        strZenVersion = sGetZenVersion
    
        iRet = MsgBox("全体バージョンＶｅｒ" & strZenVersion & "の統合駅都度データを" & vbCrLf & _
                      "インストールしますがよろしいですか？", _
                      vbOKCancel + vbQuestion, _
                      "媒体インストール確認")
        
        If iRet = vbOK Then
        
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを表示する
            Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
        
            'エラールーチンを宣言
            On Error Resume Next
            
            '処理番号格納（媒体インストール中）
            glShoriNo = SHORI_NO.NO_INSTOL
        
            '媒体インストール中ポップアップ画面表示
            Load frmSyorityu
            frmSyorityu.lblLogMessage.Caption = "媒体インストール中"
            frmSyorityu.Caption = "媒体インストール中"
            frmSyorityu.Show vbModal
        
            '統合駅都度データインストール結果
            If gTgEkiData = False Then GoTo Err_LOG
            
            '初期画面表示
            bRet = sDisp
            If bRet = False Then
                '統合駅都度データファイルを元に戻す
                Kill EKI_DATA_FILE
                Name EKI_DATA_RENAME_FILE As EKI_DATA_FILE
                bRet = sDisp
                GoTo Err_LOG
            End If
    
            '統合駅都度データバックアップファイル削除
            Kill EKI_DATA_RENAME_FILE
            
            'ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
    
            '媒体インストール結果ポップアップ画面表示
            iRet = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体インストール結果")
            
        End If
                
        'ワークフォルダ内の統合駅都度データファイルを削除
        iRet = DeleteFile(PATH_WORK_EKI_DATA_FILE)
        
    End If
    
'V1.12.0.1 ADD START
    '全ボタンを押下可とする。
    Call SetEnableTrue
'V1.12.0.1 ADD END
    
    Exit Sub

Err_LOG:

    'ワークフォルダ内の統合駅都度データファイルを削除
    iRet = DeleteFile(PATH_WORK_EKI_DATA_FILE)

' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    '一時保存フォルダを削除する
    psDeleteFolder PATH_HOSHUTMP
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End

    'ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_NG, 0)

' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了

    '媒体インストール結果ポップアップ画面表示
    'iRet = MsgBox("異常終了しました。", vbOKOnly + vbExclamation, "媒体インストール結果")  'V1.8.0.1 DEL
     iRet = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "媒体インストール結果") 'V1.8.0.1 ADD

'V1.12.0.1 ADD START
    '全ボタンを押下可とする。
    Call SetEnableTrue
'V1.12.0.1 ADD END

    Set objFso = Nothing    'V1.20.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdDataHanei_Click
'//  機能名称  : 「設置駅データ反映」釦押下時処理
'//  機能概要  : 指定駅情報をINIファイルに反映する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdDataHanei_Click()

    Dim iRet As Integer         'メッセージボックス戻り値
    Dim bRet As Boolean         '関数戻り値
    Dim lErrCode As Long        'エラーコード
    Dim bSysChange As Boolean   'システム設定処理戻り値　'V1.8.0.1 ADD
    Dim bInstolType As Boolean  '統合駅タイプ都度データインストール処理済みフラグ
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    Dim lResult             As Long             ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加
    Dim iResponse           As Integer          ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加

    '初期化
    bInstolType = False  '統合駅タイプ都度データインストール処理済みフラグ

    'エラールーチンを宣言
    On Error Resume Next
       
    '全ボタンを押下不可とする。
    Call SetEnableFalse
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_UPDATE, 0)
    
    bRet = False
    
    iRet = MsgBox("設置駅名を変更します。よろしいですか？" & vbCrLf & "反映は再起動後になります。", _
                  vbOKCancel + vbInformation, _
                  "統合駅都度データ反映")
             
    If iRet = vbCancel Then
        Call SetEnableTrue
        Set objFso = Nothing  'V2.1.0.1 ADD
        Exit Sub
    End If

' EG20 V3.3.0.1 追加開始
    ' リストに１件もデータがない場合は異常終了
    If LstStation.ListCount = 0 Then
        GoTo ErrorHandler
    End If
' EG20 V3.3.0.1 追加終了

    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
        
    ' /////////////////////////////////////////////////////////
    ' // 統合駅都度データから選択した駅データを切りだし
    ' // [IN]統合駅都度データファイル名
    ' // [IN]駅データファイルの保存ファイル名
    ' // [IN]選択された統合駅都度データファイルのインデックス
    ' // [out] エラーコード
    bRet = dllCreateFile_ChooseEkiData(EKI_DATA_FILE, EKI_DATA_CHOOSE_FILE, LstStation.ListIndex, lErrCode)

    '処理結果
    If bRet = False Then
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, lErrCode)
        GoTo ErrorHandler               ' EG20 V3.3.0.1追加
    Else
        gstrFileName = EKI_DATA_CHOOSE_FILE
        ' //////////////////////////////////////////////
        ' // 操作卓プログラム処理
        ' //////////////////////////////////////////////
        lResult = pubfuncTakuProgramData(2, gstrFileName)
        If lResult = 0 Then
            GoTo ErrorHandler           ' EG20 V3.3.0.1追加
' EG20 V3.3.0.1 削除開始
'           'プログレスバーを消去する
'           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'           ' 異常終了
'           iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "統合駅都度データ反映")
'           Set objFso = Nothing  'V2.1.0.1 ADD
'           Call SetEnableTrue
'           Exit Sub
' EG20 V3.3.0.1 削除終了
        ElseIf lResult = 1 Then
           ' メール送信中
           ' ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
           Set objFso = Nothing  'V2.1.0.1 ADD
            
           Exit Sub
        End If
    
        ' //////////////////////////////////////////////
        ' // 統合監視盤非動作中のためメール応答を待たずに
        ' // 即時更新
        ' //////////////////////////////////////////////
        bRet = pfuncInstallEkiSettei
    
    End If

    Exit Sub                            ' EG20 V3.3.0.1 追加
' EG20 V3.3.0.1 追加開始（エラー処理をまとめる）
ErrorHandler:
    Set objFso = Nothing
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    MsgBox "統合駅都度データ反映処理が異常終了しました。", _
            vbOKOnly + vbCritical, _
             "統合駅都度データ反映結果"
    Call SetEnableTrue
    Exit Sub
' EG20 V3.3.0.1 追加終了（エラー処理をまとめる）
End Sub

' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]削除開始（全体見直し）
'Private Sub cmdDataHanei_Click()
'
'    Dim iRet As Integer         'メッセージボックス戻り値
'    Dim bRet As Boolean         '関数戻り値
'    Dim lErrCode As Long        'エラーコード
'    Dim bSysChange As Boolean   'システム設定処理戻り値　'V1.8.0.1 ADD
''V2.1.0.1 ADD START
'    Dim bInstolType As Boolean  '統合駅タイプ都度データインストール処理済みフラグ
'    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
'    Dim lResult             As Long             ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加
'    Dim iResponse           As Integer          ' 処理結果     ' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加
'
'    '初期化
'    bInstolType = False  '統合駅タイプ都度データインストール処理済みフラグ
''V2.1.0.1 ADD END
'
'    'エラールーチンを宣言
'    On Error Resume Next
'
''V1.12.0.1 ADD START
'    '全ボタンを押下不可とする。
'    Call SetEnableFalse
''V1.12.0.1 ADD END
'
'    '画面操作ログ出力
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_UPDATE, 0)
'
'    bRet = False
'
'    iRet = MsgBox("設置駅名を変更します。よろしいですか？" & vbCrLf & "反映は再起動後になります。", _
'                  vbOKCancel + vbInformation, _
'                  "統合駅都度データ反映")
'
''    If iRet = vbCancel Then Exit Sub   'V1.12.0.1 DEL
''V1.12.0.1 ADD START
'    If iRet = vbCancel Then
'        Call SetEnableTrue
'        Set objFso = Nothing  'V2.1.0.1 ADD
'        Exit Sub
'    End If
''V1.12.0.1 ADD END
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'    'プログレスバーを表示する
'    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'    '統合駅都度データインストール処理
'    bRet = dllInstolEkiData(EKI_DATA_FILE, EKI_NAME_FILE, EKI_SETTI_FILE, LstStation.ListIndex, lErrCode)
'
'    '処理結果
'    If bRet = False Then
'        '異常ログ出力
''       Call pfOutPutErrLog(lErrCode)    'V2.1.0.1 DEL
'        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, lErrCode)    'V2.1.0.1 ADD
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'        'プログレスバーを消去する
'        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'        '異常終了
'        'V1.8.0.1 DEL START
'        'MsgBox "異常終了しました。", _
'        '        vbOKOnly + vbExclamation, _
'        '         "統合駅都度データ反映結果"
'        'V1.8.0.1 DEL END
''V2.1.0.1 DEL START
''       'V1.8.0.1 ADD START
''       MsgBox "異常終了しました。", _
''               vbOKOnly + vbCritical, _
''                "統合駅都度データ反映結果"
''       'V1.8.0.1 ADD END
''V2.1.0.1 DEL END
''V2.1.0.1 ADD START
'        MsgBox "統合駅都度データ反映処理が異常終了しました。", _
'                vbOKOnly + vbCritical, _
'                 "統合駅都度データ反映結果"
''V2.1.0.1 ADD END
'    Else
''V2.1.0.1 ADD START
'        '統合駅タイプ都度データファイルが存在する？
'        If (objFso.FileExists(EKI_TYPE_DATA_FILE) = True) Then
'
'            '統合駅タイプ都度データインストール処理関数
'            bRet = dllInstolEkiTypeData(EKI_TYPE_DATA_FILE, lErrCode)
'            '統合駅タイプ都度データインストール処理済みフラグを立てる
'            bInstolType = True
'
'        End If
''V2.1.0.1 ADD END
'        '----------------------------------------------------
'        '現在の設置駅ラベル更新
'        '----------------------------------------------------
'        Call sDispNowEkiLabel
'
'        'V1.8.0.1 START ADD
'        '----------------------------------------------------
'        'コンピュータ名、ネットワーク変更処理
'        '----------------------------------------------------
'        'Call pfNetWorkChng(Me)
'        bSysChange = pfNetWorkChng(Me)
'        'V1.8.0.1 START END
'
''V2.1.0.1 DEL START
''       'ログ出力
''       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
''V2.1.0.1 DEL END
'
'        If bSysChange = True Then 'V1.8.0.1 ADD
'        'V2.1.0.1 DEL START
'        ''正常終了
'        'MsgBox "正常終了しました。", _
'        '        vbOKOnly + vbInformation, _
'        '         "統合駅都度データ反映結果"
'        ''V2.1.0.1 DEL END
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
'                Set objFso = Nothing  'V2.1.0.1 ADD
'                Call SetEnableTrue
'                Exit Sub
'             ElseIf lResult = 1 Then
'                ' メール送信中
'                ' ログ出力
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'                Set objFso = Nothing  'V2.1.0.1 ADD
'
'                Exit Sub
'             End If
'' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]追加終了
'
'            'V2.1.0.1 ADD START
'            'ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET, 0)
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'            'プログレスバーを消去する
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'            '正常終了
'            MsgBox "統合駅都度データ反映処理が正常終了しました。", _
'                    vbOKOnly + vbInformation, _
'                     "統合駅都度データ反映結果"
'            'V2.1.0.1 ADD END
''V1.12.0.1 ADD START
'        Else
'            'V2.1.0.1 DEL START
'            ''異常終了
'            'MsgBox "異常終了しました。", _
'            '        vbOKOnly + vbCritical, _
'            '         "統合駅都度データ反映結果"
'            'V2.1.0.1 DEL END
'            'V2.1.0.1 ADD START
'            '異常ログ出力
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, 0)
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'            'プログレスバーを消去する
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'            '異常終了
'            MsgBox "統合駅都度データ反映処理が異常終了しました。", _
'                    vbOKOnly + vbCritical, _
'                     "統合駅都度データ反映結果"
'            'V2.1.0.1 ADD END
''V1.12.0.1 ADD END
'       End If                    'V1.8.0.1 ADD
''V2.1.0.1 ADD START
'
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加開始
'        'プログレスバーを消去する
'        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 プログレスバー対応]追加終了
'
'        '統合駅タイプ都度データインストール処理を実行したか？
'        If (True = bInstolType) Then
'            '統合駅タイプ都度データインストール処理結果
'            If ((False = bRet) And (ERR_EKITYPE_NO_TYPE = lErrCode)) Then
'                '該当駅タイプ都度データなし終了
'                'ログ出力
'                Call sLogTraceReq(LTYP_WARNING, L3AN_ETC, EKITUDODATASET_NO_EKITYPE_DATA, 0)
'                MsgBox "該当する駅タイプ都度データが存在しませんでした。", _
'                        vbOKOnly + vbExclamation, _
'                        "統合駅タイプ都度データ反映結果"
'            ElseIf (False = bRet) Then
'                '異常終了
'                '異常ログ出力
'                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKITYPE_TUDO_SET_NG, lErrCode)
'                MsgBox "統合駅タイプ都度データ反映処理が異常終了しました。", _
'                        vbOKOnly + vbCritical, _
'                        "統合駅タイプ都度データ反映結果"
'            Else
'                '正常終了
'                'ログ出力
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_EKITYPE_TUDO_SET, 0)
'                MsgBox "統合駅タイプ都度データ反映処理が正常終了しました。", _
'                        vbOKOnly + vbInformation, _
'                        "統合駅タイプ都度データ反映結果"
'            End If
'        Else
'            '統合駅タイプ都度データインストール処理未実行
'            'ログ出力
'            Call sLogTraceReq(LTYP_WARNING, L3AN_FILE, EKITUDODATASET_NO_EKITYPE_FILE, 0)
'        End If
''V2.1.0.1 ADD END
'    End If
'
'    Set objFso = Nothing  'V2.1.0.1 ADD
''V1.12.0.1 ADD START
'    '全ボタンを押下可とする。
'    Call SetEnableTrue
''V1.12.0.1 ADD END
'
'End Sub
' EG20 V3.0.0.2[Mainte_03_01 駅都度対応]削除終了（全体見直し）

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdOut_Click
'//  機能名称  : 「媒体取り外し」釦押下時処理
'//  機能概要  : USB取り外し準備を行う
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOut_Click()

    'エラールーチンを宣言
    On Error Resume Next
    
'V1.12.0.1 ADD START
    '全ボタンを押下不可とする。
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
    
    '媒体取外処理
    Call pfRemove(Me)

'V1.12.0.1 ADD START
    '全ボタンを押下可とする。
    Call SetEnableTrue
'V1.12.0.1 ADD END
        
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()

    'エラールーチンを宣言
    On Error Resume Next

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_GAMEN_END, 0)
    
    '自画面消去
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                画面再前面表示修正(不具合修正)
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    Dim intStatus As Integer         '受信メールチェック結果
    Dim iResponse As Integer
    
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
'                AppActivate frmInputMstData.Caption, False ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
                AppActivate frmEkisettei.Caption, False     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
                pfFormActive (frmEkisettei.hwnd)            ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '「保守操作卓プログラム送信要求」を受信した場合
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
                    'プログレスバーを消去する
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                    MsgBox "統合駅都度データ反映処理が異常終了しました。", _
                           vbOKOnly + vbCritical, _
                           "統合駅都度データ反映結果"
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
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sDisp() As Boolean

    'エラールーチンを宣言
    On Error Resume Next
    
    Dim bRet                 As Boolean         '関数戻り値
    
    sDisp = False
    
    '----------------------------------------------------
    '釦初期値設定
    '----------------------------------------------------
    cmdUp.Enabled = False                       '▼釦
    cmdPageUp.Enabled = False                   '▼▼釦
    cmdPageDown.Enabled = False                 '▲釦
    cmdDown.Enabled = False                     '▲▲釦
    cmdInstall.Enabled = True                   '駅都度データ媒体 インストール釦
    cmdDataHanei.Enabled = False                '設置駅データ反映釦
    cmdOut.Enabled = True                       '媒体取外釦
    
    '----------------------------------------------------
    '初期値設定
    '----------------------------------------------------
    lblNow.Caption = ""
    lblZen.Caption = ""
    LstStation.Clear

    '----------------------------------------------------
    '現在の設置駅ラベル更新
    '----------------------------------------------------
    Call sDispNowEkiLabel
    
    '----------------------------------------------------
    '駅情報更新
    '----------------------------------------------------
    bRet = sDispEkiData
    
    '釦押下設定
    If bRet = True Then
        cmdDataHanei.Enabled = True             '設置駅データ反映釦
        cmdUp.Enabled = True                    '▼釦
        cmdPageUp.Enabled = True                '▼▼釦
        cmdPageDown.Enabled = True              '▲釦
        cmdDown.Enabled = True                  '▲▲釦
        
        '駅名コンボボックスのインデックス設定
        LstStation.ListIndex = 0
        
        sDisp = True
    
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sDispNowEkiLabel
'//  機能名称  : 現在の設置駅ラベル更新処理
'//  機能概要  : 現在の設置駅ラベルに駅名、駅バージョンを設定する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispNowEkiLabel()

    Dim strFileName          As String          'ファイル名
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在した場合
    If strFileName <> "" Then
    
        '駅バージョン取得
'       lblNow.Caption = pfGetEkiNameInfo               'V2.1.0.1 DEL
        lblNow.Caption = pfGetEkiNameInfo(SetEkiVer)    'V2.1.0.1 ADD
    
    Else
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
    
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sDispEkiData
'//  機能名称  : 駅情報更新処理
'//  機能概要  : 全体バージョンラベルをバージョンを設定し、
'//              駅名コンボボックスに駅情報を設定する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　 TURE      正常終了
'//                        FALSE     異常終了（コンボボックスデータデータが存在しない場合）
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sDispEkiData() As Boolean

    Dim LOG_Event            As String          'ログのイベント名
    Dim LOG_FukaData         As String          '付加データ名
    Dim ECOD3                As Integer         '小分類
    
    Dim intLoopCount         As Integer         'ループカウンタ
    Dim intFileNumber        As Integer
    Dim strFileName          As String          'ファイル名
    
    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード
    Dim strData              As String          'ファイル読込データ
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '初期値設定
    sDispEkiData = False

    '----------------------------------------------------
    '統合駅都度データファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_DATA_FILE)
    
    'ファイルが存在した場合
    If strFileName <> "" Then
    
        ' 統合駅都度データ駅名ファイル作成
        bRet = dllCreateEkiNameFile(EKI_DATA_FILE, EKI_NAME_FILE, lErrCode)
        If bRet = False Then
            '統合駅都度データ駅名ファイル削除
            Kill EKI_NAME_FILE
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
            Exit Function
        End If
        
        '駅名ファイル検索
        strFileName = Dir(EKI_NAME_FILE)
        
        'ファイルが存在した場合
        If strFileName <> "" Then
        
            '内部ファイルエラーのトラップ
            On Error GoTo Err_LOG
        
            '未使用のファイル番号取得
            intFileNumber = FreeFile
            
            '現在駅設定ファイルをオープンする。
            Open EKI_NAME_FILE For Input As #intFileNumber
            
            intLoopCount = 0
            Do While Not EOF(intFileNumber)
                '１ 行読み込み
                Input #intFileNumber, strData
                
                '先頭一行目は全体バージョンを取得し、ラベルを更新する
                If intLoopCount = 0 Then
                  '  lblZen.Caption = "V" & strData 'V1.8.0.1 DEL
                     lblZen.Caption = strData       'V1.8.0.1 ADD
                    intLoopCount = 1
                Else
                    'リストボックスに駅名情報を追加
                    LstStation.AddItem strData
                End If
            Loop
            
            'ファイルをクローズする。
            Close #intFileNumber
        
        Else
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKINAME, 0)
            
            '統合駅都度データ駅名ファイルなし
            sDispEkiData = False
        End If
        
    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_TG_EKITUDO, 0)
        
        '統合駅都度データファイルなし
        sDispEkiData = False
    End If
    
    sDispEkiData = True
    
    Exit Function
    
'エラー処理
Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfFileSelection
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
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

'    If Left$(sWorkDrive, 1) = "a" Then     'V1.12.0.1 DEL
    If Left$(sWorkDrive, 1) = "H" Then      'V1.12.0.1 ADD
        'a:ドライブが異常なら、カレントドライブを表示させる。
        sWorkDrive = Left$(App.Path, 2)
        GoTo Retry
    End If
    
    'その他のドライブなら、ファイル選択なしで戻る。
    pfFileSelection = ""

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sGetZenVersion
'//  機能名称  : 全体バージョン取得
'//  機能概要  : ワークフォルダ内の統合駅都度データファイルから全体バージョンを取得する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】駅都度対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sGetZenVersion() As String

    Dim LOG_Event            As String          'ログのイベント名
    Dim LOG_FukaData         As String          '付加データ名
    Dim ECOD3                As Integer         '小分類
    
    Dim intFileNumber        As Integer
    Dim strFileName          As String          'ファイル名
    
    Dim intBunrui_Dai        As Integer         '大分類
    Dim intBunrui_Tyu        As Integer         '中分類
    Dim intBunrui_Sho        As Integer         '小分類
    Dim strData              As String          '設定値
    
    Dim strPtnNo             As String          'パターン番号
    Dim strZenVersion        As String          '全体バージョン
    
    Dim iGetDataCount        As Integer         'データ取得カウンタ
    Dim intBunrui_Corner     As Integer         '小分類
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '初期値設定
    sGetZenVersion = ""
    strZenVersion = ""
    iGetDataCount = 0

    '----------------------------------------------------
    '統合駅都度データファイル検索
    '----------------------------------------------------
    strFileName = Dir(PATH_WORK_EKI_DATA_FILE)

    'ファイルが存在した場合
    If strFileName <> "" Then
    
        '未使用のファイル番号取得
        intFileNumber = FreeFile
    
        '内部ファイルエラーのトラップ
        On Error GoTo Err_LOG
        
        '現在駅設定ファイルをオープンする。
        Open PATH_WORK_EKI_DATA_FILE For Input As #intFileNumber
    
        Do While Not EOF(intFileNumber)
            '１ 行づつ変数読み込み
'            Input #intFileNumber, intBunrui_Dai, intBunrui_Tyu, intBunrui_Sho, strData                     ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
            Input #intFileNumber, intBunrui_Dai, intBunrui_Tyu, intBunrui_Sho, intBunrui_Corner, strData    ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    
            'パターン番号取得
'            strPtnNo = Format(intBunrui_Dai, "00") & Format(intBunrui_Tyu, "00") & Format(intBunrui_Sho, "00") ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]削除
            strPtnNo = Format(intBunrui_Dai, "00") & Format(intBunrui_Tyu, "00") & Format(intBunrui_Sho, "000") ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
            
            Select Case strPtnNo
                
                '全体バージョン取得
                Case PtnZenVersion
                    strZenVersion = strData
                    iGetDataCount = iGetDataCount + 1
                
                Case Else
                    '処理なし
            End Select
            
            '全体バージョンを取得したらループを抜ける
            If iGetDataCount > 0 Then Exit Do

        Loop
        
        'ファイルをクローズする。
        Close #intFileNumber
        
        '戻り値設定（全体バージョン）
        sGetZenVersion = strZenVersion
    
    Else
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_TG_EKITUDO, 0)
    End If
    
    Exit Function
    
'エラー処理
Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
    
End Function

'V1.12.0.1 ADD START
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
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    cmdInstall.Enabled = False
    cmdDataHanei.Enabled = False
    cmdOut.Enabled = False
    cmdCancel.Enabled = False
    cmdUp.Enabled = False
    cmdPageUp.Enabled = False
    cmdPageDown.Enabled = False
    cmdDown.Enabled = False
    
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
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下可とする。
    cmdInstall.Enabled = True
    cmdDataHanei.Enabled = True
    cmdOut.Enabled = True
    cmdCancel.Enabled = True
    
    'リストボックスに項目がない場合、「▲」「▲▲」「▼」「▼▼」はFalseのままとする。
    If LstStation.ListCount <> 0 Then
        cmdUp.Enabled = True
        cmdPageUp.Enabled = True
        cmdPageDown.Enabled = True
        cmdDown.Enabled = True
    End If
    
End Sub
'V1.12.0.1 ADD END

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfuncInstallEkiSettei() As Boolean

    Dim bRet As Boolean                  ' 関数戻り値
    Dim lErrCode As Long                 ' エラーコード
    Dim bSysChange As Boolean            ' システム設定処理戻り値
    Dim objFso As New FileSystemObject   ' ファイルシステムオブジェクト

    'エラールーチンを宣言
    On Error Resume Next

    '全ボタンを押下不可とする。
    Call SetEnableFalse

    '現在駅設定データインストール処理
    bRet = dllInstolEkiDataNow(gstrFileName, EKI_SETTI_FILE, lErrCode)

    If bRet = False Then

        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)

        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)

        pfuncInstallEkiSettei = False

        '異常終了
        MsgBox "統合駅都度データ反映処理が異常終了しました。", _
                vbOKOnly + vbCritical, _
                "統合駅都度データ反映結果"

    Else

        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)

        '----------------------------------------------------
        '現在の設置駅ラベル更新
        '----------------------------------------------------
        Call sDispNowEkiLabel
        
        '----------------------------------------------------
        'コンピュータ名、ネットワーク変更処理
        '----------------------------------------------------
        bSysChange = pfNetWorkChng(Me)

        If bSysChange = True Then

            'ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET, 0)

            '正常終了
            MsgBox "統合駅都度データ反映処理が正常終了しました。", _
                    vbOKOnly + vbInformation, _
                     "統合駅都度データ反映結果"
        Else

            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, 0)

            '異常終了
            MsgBox "統合駅都度データ反映処理が異常終了しました。", _
                    vbOKOnly + vbCritical, _
                     "統合駅都度データ反映結果"
        End If

    End If

    gstrFileName = ""
    Call SetEnableTrue

End Function


