VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLduVer 
   BorderStyle     =   0  'なし
   Caption         =   "                                                               ＬＤユーティリティバージョン管理"
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
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAplTimer 
      Left            =   8640
      Top             =   4080
   End
   Begin VB.Timer tmrLogTimer 
      Left            =   8760
      Top             =   3360
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " 媒体 → ワーク コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   18
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyWork_Jikko 
      Caption         =   " ワーク → 実行 コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   17
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   旧 → 実行     コピー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   16
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   8760
      Top             =   5040
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "　バージョン管理　画面へ戻る"
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
      Left            =   9360
      TabIndex        =   15
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdRefresh 
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
      Height          =   615
      Left            =   9360
      TabIndex        =   14
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdOutPut 
      Caption         =   " バージョン情報 媒体出力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   13
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdRemove 
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
      Height          =   615
      Left            =   9360
      TabIndex        =   12
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox txtDummy 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   12480
      Width           =   2535
   End
   Begin VB.Frame frmFolder 
      Height          =   1815
      Left            =   9360
      TabIndex        =   6
      Top             =   480
      Width           =   2055
      Begin VB.CheckBox chkFolder 
         Caption         =   "O 旧"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "N 実行"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "W ワーク"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.ListBox LstFile 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   240
      MultiSelect     =   2  '拡張
      TabIndex        =   1
      Top             =   2520
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0000C000&
      Caption         =   "LDUアプリケーションバージョン管理"
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
      TabIndex        =   11
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label lblZenVer 
      Caption         =   "全体バージョン"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label lblVer 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "バージョン番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblTime 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "更新日付"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lblFolder 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ﾌｫﾙﾀﾞ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblFile 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ファイル名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "frmLduVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmLduVer.frm
'//  パッケージ名：バージョン管理(LDU)画面
'//
'//  概要：バージョン管理(LDU)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・LDユーティリティより、バージョン管理(LDU)画面流用。
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.129】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.123】
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.273修正対応】
'//                 EG20フェーズ２対応【03統合TR-No.22修正対応】
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.59修正対応】
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ2,3統合対応
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【アプリ切替改善対応】
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-23  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応（媒体取外しエラー対応）
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'全体バージョン情報保存管理
Private sMainVer As String

'フォルダ種別部
Public mlngChkFolderType        As Long
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値  'V1.3.0.1 ADD

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
Private Const HEADERTITLE_WRK = "LDUアプリケーションバージョン（ワーク）："
Private Const HEADERTITLE_NOW = "　　　　　　　　　　　　　　 （実行）　："
Private Const HEADERTITLE_OLD = "　　　　　　　　　　　　　　 （旧）　　："
Private Const HEADERVERSION_NON = "--.--.--.--"
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

' EG20 V3.3.0.1【結合TR-No.123】 追加開始
Private Const APL_INTERVAL = 390000         ' アプリ起動タイマデフォルト値
Private Const LOG_INTERVAL = 30000          ' ログ起動タイマデフォルト値(30秒)
Dim lngAplMAX_Time As Long                  ' INI取得設定値（ＡＰＬ）
Dim lngLogMAX_Time As Long                  ' INI取得設定値（ログ）
Dim lngtime        As Long                  ' 現在タイマ値
Dim lngChangeKind  As Long                  ' バージョン切替種別
' EG20 V3.3.0.1【結合TR-No.123】 追加終了

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdCopyBaitai_Work_Click
'//  機能名称    ：媒体→ワークコピー
'//  機能概要    ：
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()

    '「媒体→ワークコピー」ボタンの場合。
    '「バージョン管理画面：媒体→ワークコピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_USB_COPY_WRK_BUTTOM, 0)
    sCmdBtnEnabled False                        ' 画面操作不可
    'インストール媒体をワークフォルダ内にコピーする
    Call sFDInstall
    sCmdBtnEnabled True                         ' 画面操作可
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sFDInstall
'//  機能名称  : 媒体インストール処理
'//  機能概要  : インストール媒体ファイルを、ワークフォルダにコピーする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.273修正対応】
'//  REVISIONS   ：(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ：(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ2,3統合対応
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ： (EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  量産対応【アプリ切替改善対応】
'//  REVISIONS   ： (EG20 V30.3.0.1)2014-10-23  CODED BY  [TCC] T.Nakajima
'//                  北陸新幹線フェーズ２対応（媒体取外しエラー対応）
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：改札機バージョン管理画面のsFDInstall流用
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall()
    Dim MyName As String            'ファイルフルパス名
    Dim iResponse As Integer        'MsgBoxボタンコード
    Dim sInputPass As String        'インストール元ディレクトリ名(STD)orファイル名(LZH)
    Dim lngErrCode As Long          'エラーコード
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    Dim objFi As File                    'ファイルオブジェクト
    
    Dim lngProcId As Long                ' プロセスID
    Dim hProc As Variant                 ' プロセスハンドル
    Dim FileName As String               ' 抽出ファイル名                ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim FileKaku As String               ' 拡張子                        ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim ExecCommand As String            ' 実行文字列                    ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim CurrentDirectory As String       ' カレントディレクトリ          ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim ExecDirectory As String          ' 実行ファイルディレクトリ      ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    
    On Error GoTo ErrorHandler      'エラーハンドルの登録

    '圧縮ファイル指定の時:
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
    '拡張子を設定
    CommonDialog1.Filter = "実行ファイル（*.EXE）|*.exe|" & _
                            "バッチファイル（*.BAT）|*.bat|" & _
                            "スクリプトファイル（*.VBS）|*.vbs|"
    'ファイル選択画面を開く
    CommonDialog1.ShowOpen
    '選択したファイル名を取得
    sInputPass = CommonDialog1.FileName
    If sInputPass = "" Then 'ファイル未選択
        Set objFso = Nothing
        Set objFi = Nothing
        Exit Sub    'ファイルが選択されなければ処理中断
    End If
        
    '「ワークコピー確認」ポップアップ画面表示
    iResponse = MsgBox("選択されたインストール部材の内容をＬＤＵアプリケーションの" _
                       & Chr(vbKeyReturn) _
                       & "切替領域に展開します。よろしいですか？", _
                       (vbOKCancel + vbExclamation), _
                       "媒体→ワーク　コピー")
        
    If iResponse = vbCancel Then
    '[いいえ] ボタンを選択:何もしない。
        'V1.20.0.1 ADD START
        Set objFso = Nothing
        Set objFi = Nothing
        'V1.20.0.1 ADD END
        Exit Sub
    End If
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
'    lngProcId = Shell(sInputPass, vbNormalFocus)       ' EG20 V3.6.0.1【統合TR-No.273修正対応】削除
' EG20 V3.6.0.1【統合TR-No.273修正対応】追加開始
    ' カレントディレクトリ取得
    CurrentDirectory = CurDir$()
    Call psFolderPathGet(sInputPass, ExecDirectory)
    Call ChDir(ExecDirectory)
    ' ファイル名前取得
    psFileNameGet sInputPass, FileName, FileKaku
    If UCase(FileKaku) = "VBS" Then
        ExecCommand = "wscript.exe " & sInputPass
    Else
        ExecCommand = sInputPass
    End If
    lngProcId = Shell(ExecCommand, vbNormalFocus)
' EG20 V3.6.0.1【統合TR-No.273修正対応】追加終了
    
    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, lngProcId)   ' プロセスハンドルを取得します。
    If hProc > 0 Then                                           ' プロセスハンドルを取得できた場合
        Call dllWaitForSingleObject(hProc)                      ' プロセスがシグナル状態になるまで待ちます。
        CloseHandle hProc                                       ' プロセスハンドルを解放します。
    End If

    'EG20 V30.0.3.1 ADD START
    'ChDirではCommonDialogの場合、Hドライブが選択されたまま変更されず、媒体取外しができなくなるため、ChDriveに変更
    ChDrive "C"
    'EG20 V30.0.3.1 ADD END
    Call ChDir(CurrentDirectory)                        ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    
    Set objFso = Nothing
    Set objFi = Nothing
    
    '監視盤のバージョン番号を表示する｡
    psVersionDisp
    
' EG20 V5.8.0.1削除開始
'    ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
    ' 運改状態更新
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_LDU, BOOTINFO_UNKAI_ARI)
    Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMELDU, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1追加終了
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD START
    ' 切替実行コピーツールパラメータ更新処理
    Call funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_CLEAR_LDU)
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD END

' EG20 V5.8.0.1 ADD START
    '読み取り外しの関数を実行
    dllChangeAttributeContents (PATH_LDU_APLNEW)
' EG20 V5.8.0.1 ADD END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    Exit Sub    '処理を終了する

ErrorHandler:   ' エラー処理。
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing

' EG20 V5.8.0.1 ADD START
    '読み取り外しの関数を実行
    dllChangeAttributeContents (PATH_LDU_APLNEW)
' EG20 V5.8.0.1 ADD END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    MsgBox "インストール媒体からのコピーエラーが発生しました。" _
            & Chr(vbKeyReturn) & "エラーコード＝" _
            & str$(Err.Number), _
            vbOKOnly + vbExclamation, _
            "媒体→ワーク　コピー"
    
    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_USB_COPY_WRK_ERROR, lngErrCode)
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdCopyOld_Jikko_Click
'//  機能名称    ：旧→実行コピー
'//  機能概要    ：
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.123】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.22修正対応】
'//                 EG20フェーズ２対応【統合TR-No.372修正対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【アプリ切替改善対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    
    Dim udtSendData As ML_KANEND_REQ_CMD  ' 共通エリア
    Dim lngSendSize As Long               ' 送信するメールサイズ
    Dim lngErrCode  As Long               ' エラーコード
    Dim bRet        As Boolean            ' メール送信処理戻り値
    Dim iResponse   As Integer            ' MsgBoxボタンコード
    Dim iAplChk     As Integer            ' アプリ起動チェック戻り値    'EG20 V3.6.0.1【03統合TR-No.22修正対応】追加
    
    On Error Resume Next
    
    '「バージョン管理画面：旧→実行コピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OLD_COPY_NOW_BUTTOM, 0)

    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("「旧」フォルダの内容を、" _
            & Chr(vbKeyReturn) & "「実行」フォルダに戻すことにより、" _
            & Chr(vbKeyReturn) & "ＬＤＵの一世代前バージョンを､実行バージョンとします｡" _
            & Chr(vbKeyReturn) & "よろしいですか？", _
           vbOKCancel + vbExclamation, _
           "旧→実行　コピー")
    If iResponse = vbCancel Then
        Exit Sub
    End If

'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加開始
    ' 旧バージョンフォルダに代表バージョンファイルが存在しない場合は異常とする。
    ' 旧バージョン・EW4500JR・LDU・
    bRet = dllCheckAplVersion(4, PATH_LDU_APP, 1)
    If bRet = False Then
        MsgBox "異常終了しました。", vbCritical, "旧→実行　コピー"
        Exit Sub
    End If
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加終了
        
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD START
    ' 切替実行コピーツールパラメータ更新処理
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONDOWN_LDU)
    If bRet = False Then
        MsgBox "異常終了しました。", vbCritical, "旧→実行　コピー"
        Exit Sub
    End If

    ' 終了確認
    iResponse = MsgBox("実行コピーを適用するために統合監視盤を" & Chr(vbKeyReturn) _
                        & "再起動しますか？", _
                        vbOKCancel + vbExclamation, _
                        "旧→実行　コピー")
    If iResponse = vbCancel Then
        Exit Sub
    End If
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD END
        
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】DEL START
''EG20 V3.6.0.1【03統合TR-No.22修正対応】追加開始
'    ' 統合監視盤が起動中の場合にメッセージボックスを表示する。
'    iAplChk = CheckAppStart(PROC_KANRI)
'    If iAplChk <> 0 Then
''EG20 V3.6.0.1【03統合TR-No.22修正対応】追加終了
'        '確認ポップアップウィンドウを表示する。
'        iResponse = MsgBox("統合監視盤､ＩＤＵ､ＬＤＵアプリケーションを" _
'                & Chr(vbKeyReturn) & "終了します。よろしいですか？", _
'               vbOKCancel + vbExclamation, _
'               "終了確認")
'
'        If iResponse = vbCancel Then
'            Exit Sub
'        End If
'    End If  'EG20 V3.6.0.1【03統合TR-No.22修正対応】追加
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】DEL END

' EG20 V2.1.0.1[Mainte_03_01]削除開始
' AplVersionChangeProcにモジュール化
'    ' メールの送信内容を編集する
'    udtSendData.udtlHeader.dwId = ML_ID_KANEND_REQ      ' メールＩＤ　=”"監視装置終了要求"
'    udtSendData.udtlHeader.dwSize = MlSize.KANEND_REQ   ' メールサイズ=”"監視装置終了要求"
'    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' 送信元プロセスＩＤ=”保守”
'    udtSendData.udtlHeader.dwSubArea = 0                ' 補助情報　=　0
'
'    udtSendData.dwStartProc = ML_DT_VERSIONDOWN         ' 起動プロセス種別 = バージョンダウン
'    ' 送信サイズを設定する。
'    lngSendSize = udtSendData.udtlHeader.dwSize
'
'    ' 監マに対して、設定情報要求メールを送信する。
'    bRet = DssSendMail(MAIL_SLOT_KANRI, lngSendSize, udtSendData.udtlHeader)
'    ' メールを正常に送信した時のログ
'    If bRet = False Then
'        '「設定情報要求メール送信異常」ログ出力
'        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, lngErrCode)
'    Else
'        '「設定情報要求メール送信正常」ログ出力
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, 0)
'    End If
' EG20 V2.1.0.1[Mainte_03_01]削除終了
' EG20 V3.3.0.1【結合TR-No.123】削除開始
'' EG20 V2.1.0.1[Mainte_03_01]追加開始
'    ' アプリケーションバージョン切替実行処理
'    If (AplVersionChangeProc(ML_DT_VERSIONDOWN_LDU) = False) Then
'        ' // 保守を終了する。
'        Call psEndHoshuProc
'        '保守プロセス終了
'        End
'    End If
'' EG20 V2.1.0.1[Mainte_03_01]追加終了
' EG20 V3.3.0.1【結合TR-No.123】削除終了
' EG20 V3.3.0.1【結合TR-No.123】追加開始

    sCmdBtnEnabled False                            ' 画面操作不可
    ' 統合監視盤へアプリ終了要求の送信
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
               vbOKOnly + vbExclamation, _
               "ＬＤＵバージョン管理"
        sCmdBtnEnabled True                         ' 画面操作可
    Else

        lngtime = MN_MAIL_INTERVAL                  ' 現在タイマ値初期化
        tmrAplTimer.Enabled = True                  ' 現在タイマ起動
    
        lngChangeKind = ML_DT_VERSIONDOWN_LDU       ' 切替種別を設定
    End If
' EG20 V3.3.0.1【結合TR-No.123】追加終了

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdCopyWork_Jikko_Click
'//  機能名称    ：ワーク→実行コピー
'//  機能概要    ：
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.123】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.22修正対応】
'//                 EG20フェーズ２対応【統合TR-No.372修正対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【アプリ切替改善対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
    
    Dim udtSendData As ML_KANEND_REQ_CMD  ' 共通エリア
    Dim lngSendSize As Long               ' 送信するメールサイズ
    Dim lngErrCode  As Long               ' エラーコード
    Dim bRet        As Boolean            ' メール送信処理戻り値
    Dim iResponse   As Integer            ' MsgBoxボタンコード
    Dim iAplChk     As Integer            ' アプリ起動チェック戻り値    'EG20 V3.6.0.1【03統合TR-No.22修正対応】追加
    
    On Error Resume Next
    
    '「バージョン管理画面：ワーク→実行コピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_COPY_NOW_BUTTOM, 0)


    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("「ワーク」フォルダの内容を、" _
            & Chr(vbKeyReturn) & "「実行」フォルダに登録することにより、" _
            & Chr(vbKeyReturn) & " ＬＤＵの最新バージョンを、実行バージョンとします。" _
            & Chr(vbKeyReturn) & "よろしいですか？", _
           vbOKCancel + vbExclamation, _
           "ワーク→実行 コピー")
    If iResponse = vbCancel Then
        Exit Sub
    End If
        
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加開始
    ' ワークバージョンフォルダに代表バージョンファイルが存在しない場合は異常とする。
    ' ワークバージョン・EW4500JR・LDU・
    bRet = dllCheckAplVersion(1, PATH_LDU_APP, 1)
    If bRet = False Then
        MsgBox "異常終了しました。", vbCritical, "ワーク→実行 コピー"
        Exit Sub
    End If
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加終了

' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD START
    ' 切替実行コピーツールパラメータ更新処理
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONUP_LDU)
    If bRet = False Then
        MsgBox "異常終了しました。", vbCritical, "ワーク→実行 コピー"
        Exit Sub
    End If

    ' 終了確認
    iResponse = MsgBox("実行コピーを適用するために統合監視盤を" & Chr(vbKeyReturn) _
                        & "再起動しますか？", _
                        vbOKCancel + vbExclamation, _
                        "ワーク→実行 コピー")
    If iResponse = vbCancel Then
        Exit Sub
    End If
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD END
        
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】DEL START
'EG20 V3.6.0.1【03統合TR-No.22修正対応】追加開始
'    ' 統合監視盤が起動中の場合にメッセージボックスを表示する。
'    iAplChk = CheckAppStart(PROC_KANRI)
'    If iAplChk <> 0 Then
''EG20 V3.6.0.1【03統合TR-No.22修正対応】追加終了
'        '確認ポップアップウィンドウを表示する。
'        iResponse = MsgBox("統合監視盤､ＩＤＵ､ＬＤＵアプリケーションを" _
'                & Chr(vbKeyReturn) & "終了します。よろしいですか？", _
'               vbOKCancel + vbExclamation, _
'               "終了確認")
'
'        If iResponse = vbCancel Then
'            Exit Sub
'        End If
'    End If  'EG20 V3.6.0.1【03統合TR-No.22修正対応】追加
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】DEL END

' EG20 V2.1.0.1[Mainte_03_01]削除開始
' AplVersionChangeProcにモジュール化
'    ' メールの送信内容を編集する
'    udtSendData.udtlHeader.dwId = ML_ID_KANEND_REQ      ' メールＩＤ　=”"監視装置終了要求"
'    udtSendData.udtlHeader.dwSize = MlSize.KANEND_REQ   ' メールサイズ=”"監視装置終了要求"
'    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' 送信元プロセスＩＤ=”保守”
'    udtSendData.udtlHeader.dwSubArea = 0                ' 補助情報　=　0
'
'    udtSendData.dwStartProc = ML_DT_VERSIONUP           ' 起動プロセス種別 = バージョンアップ
'    ' 送信サイズを設定する。
'    lngSendSize = udtSendData.udtlHeader.dwSize
'
'    ' 監マに対して、設定情報要求メールを送信する。
'    bRet = DssSendMail(MAIL_SLOT_KANRI, lngSendSize, udtSendData.udtlHeader)
'    ' メールを正常に送信した時のログ
'    If bRet = False Then
'        '「設定情報要求メール送信異常」ログ出力
'        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, lngErrCode)
'    Else
'        '「設定情報要求メール送信正常」ログ出力
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, 0)
'    End If
' EG20 V2.1.0.1[Mainte_03_01]削除終了
' EG20 V3.3.0.1【結合TR-No.123】削除開始
'' EG20 V2.1.0.1[Mainte_03_01]追加開始
'    ' アプリケーションバージョン切替実行処理
'    If (AplVersionChangeProc(ML_DT_VERSIONUP_LDU) = False) Then
'        ' // 保守を終了する。
'        Call psEndHoshuProc
'        '保守プロセス終了
'        End
'    End If
'' EG20 V2.1.0.1[Mainte_03_01]追加終了
' EG20 V3.3.0.1【結合TR-No.123】削除終了
' EG20 V3.3.0.1【結合TR-No.123】追加開始

    sCmdBtnEnabled False                            ' 画面操作不可
    ' 統合監視盤へアプリ終了要求の送信
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
               vbOKOnly + vbExclamation, _
               "ＬＤＵバージョン管理"
        sCmdBtnEnabled True                         ' 画面操作可
    Else

        lngtime = MN_MAIL_INTERVAL                  ' 現在タイマ値初期化
        tmrAplTimer.Enabled = True                  ' 現在タイマ起動
    
        lngChangeKind = ML_DT_VERSIONUP_LDU         ' 切替種別を設定
    End If
' EG20 V3.3.0.1【結合TR-No.123】追加終了

End Sub

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
'//  機能名称  : バージョン管理(LDU)画面(アクティブ時)
'//  機能概要  : 画面最前面表示
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
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    'V1.3.0.1 ADD START
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
    'V1.3.0.1 ADD END

End Sub

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : バージョン管理(LDU)画面(ディアクティブ時)
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
'//  機能名称  : バージョン管理(LDU)画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.123】
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    '「LDﾕｰﾃｨﾘﾃｨﾊﾞｰｼﾞｮﾝ管理画面：表示」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_VERASION_KANRI_GAMEN_START, 0)

    gStrCurrentForm = sFormName_LDUVer

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   
   '初期化
    LstFile.Clear
    lblZenVer.Caption = ""
    mlngChkFolderType = 0

    'フォルダ選択部：選択有り
    chkFolder(0).Value = 1
    chkFolder(1).Value = 1
    chkFolder(2).Value = 1
        
    mlngChkFolderType = 7       ' EG20 V2.1.0.1[Mainte_03_01]追加
        
    'バージョン情報出力処理
    Call psVersionDisp
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   
   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END

' EG20 V3.3.0.1【結合TR-No.123】 追加開始
    ' INIファイルよりアプリ起動タイマ値を取得
    lngAplMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                       APL_INTERVAL, HOSHU_FILE)
    ' 取得値が0の場合、デフォルト値を設定
    If lngAplMAX_Time = 0 Then
       lngAplMAX_Time = APL_INTERVAL
    End If

    ' タイマ値設定
    tmrAplTimer.Interval = MN_MAIL_INTERVAL
    tmrAplTimer.Enabled = False

    ' INIファイルよりログ起動タイマ値を取得
    lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
    ' 取得値が0の場合、デフォルト値を設定
    If lngLogMAX_Time = 0 Then
       lngLogMAX_Time = LOG_INTERVAL
    End If

    ' タイマ値設定
    tmrLogTimer.Interval = MN_MAIL_INTERVAL
    tmrLogTimer.Enabled = False

    ' 切替種別を初期化
    lngChangeKind = 0
' EG20 V3.3.0.1【結合TR-No.123】 追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : chkFolder_Click
'//  機能名称  : 「フォルダチェック」チェック押下処理
'//  機能概要  : フォルダ選択部チェックを行う。
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
Private Sub chkFolder_Click(Index As Integer)
    
    Dim ValueCnt                As Integer

    '種類によって増減値を変更する
    ValueCnt = 0
    'ワーク
    If Index = 0 Then
        ValueCnt = 1
    '実行
    ElseIf Index = 1 Then
        ValueCnt = 2
    '旧
    ElseIf Index = 2 Then
        ValueCnt = 4
    End If

    'チェックがはずされた時
    If chkFolder(Index).Value = 0 Then
        mlngChkFolderType = mlngChkFolderType - ValueCnt
    'チェックされた時
    ElseIf chkFolder(Index).Value = 1 Then
        mlngChkFolderType = mlngChkFolderType + ValueCnt
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdRefresh_Click
'//  機能名称  : 「表示更新」釦押下時処理
'//  機能概要  : 最新のバージョン情報を表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdRefresh_Click()
    Dim i As Integer        'カウンター
    Dim bFlag As Boolean    '表示フォルダ選択チェック(TRUE：チェック有。FALSE：チェック無)
  
    On Error Resume Next
  
  '「LDユーティリティバージョン管理画面：表示更新釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
   
   '表示フォルダ選択チェックをチェック無しで初期化する。
   bFlag = False
   '表示フォルダ選択チェック有無をチェックする。
   For i = 0 To 2
     If chkFolder(i).Value = CHECKBOX_ON Then
       '１つでもチェック有りの場合、表示フォルダ選択チェックを、チェック有にする。
        bFlag = True
         Exit For
      End If
   Next
   
   '表示フォルダ選択のチェックがない場合は、「表示フォルダ指定なし」ポップアップ表示
   If bFlag = False Then
      MsgBox "表示ﾌｫﾙﾀﾞ指定がひとつも選択されていません。", _
              vbOKOnly + vbExclamation, _
               "ＬＤＵバージョン管理"
       Exit Sub
    End If

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    'バージョン情報出力処理
    Call psVersionDisp

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdOutPut_Click
'//  機能名称  : 「バージョン情報媒体出力」釦押下時処理
'//  機能概要  : 画面上に表示されたバージョン情報を、媒体に出力する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.129】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()

'*******************************
'VBエラー処理
On Error GoTo Error_cmdOutPut_Click
'*******************************

    Dim strVerFile  As String               'LDユーティリティファイルパス
    Dim strCopySaki As String               '出力ファイルパス
    Dim strWriteDir As String               '出力先フォルダ
    Dim fso         As New FileSystemObject 'ファイルシステムオブジェクト
    Dim lngErrCode  As Long                 'エラーコード
  
    Dim strStationName       As String          ' 駅名名                ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim szCornerName         As String          ' コーナ名称            ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim nNullIndex           As Integer         ' 文字数ワーク          ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim strRecord            As String          ' ワーク
    Dim strFileName         As String           ' ファイル名
    Dim bRet                As Boolean          ' 戻り値

    Set fso = CreateObject("Scripting.FileSystemObject")

  
   '「LDユーティリティバージョン管理画面：バージョン情報媒体出力釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OUTPUT, 0)

' EG20 V3.3.0.1 【結合TR-No.129】追加開始
    ' リストに１件もデータがない場合は異常終了
    If LstFile.ListCount = 0 Then
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Set fso = Nothing           ' EG20 V3.3.0.1追加
        Exit Sub
    End If
' EG20 V3.3.0.1 【結合TR-No.129】追加終了

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'フォルダ選択ポップアップ画面表示
    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)

    '指定フォルダなし
    If Len(strWriteDir) = 0 Then
        Set fso = Nothing
        Exit Sub
    End If

    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了


' EG20 V2.1.0.1[Mainte_03_01]追加開始
    strStationName = gsGetStationEkiName
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ 保守専用関数:IDUバージョンファイル（帳表用）作成
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, LDUVERLIST_REPORTFILE, PATH_LDU_APP, _
                                    VERLISTKIND_REPORT, 1)

    ' バージョンファイル成功
    If bRet Then
        '「バージョン情報ファイル作成正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    ' バージョンファイル失敗
    Else
        '「バージョン情報ファイル作成異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Set fso = Nothing           ' EG20 V3.3.0.1追加
       Exit Sub
    End If
    
    'ファイルの有無確認
    If fso.FileExists(LDUVERLIST_REPORTFILE) = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Set fso = Nothing           ' EG20 V3.3.0.1追加
        Exit Sub
    End If
    strFileName = Dir(LDUVERLIST_REPORTFILE)
' EG20 V2.1.0.1[Mainte_03_01]追加終了

' EG20 V2.1.0.1[Mainte_03_01]削除開始
'    'LDユーティリティバージョンファイル
'    strVerFile = PATH_LDU_APP & PATH_LDU_WORK & LDU_VER_FILE
'
'    'ファイルの有無確認
'    If fso.FileExists(strVerFile) = False Then
'        'ファイル無し異常ポップアップ画面表示
'        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
'        Exit Sub
'    End If
' EG20 V2.1.0.1[Mainte_03_01]削除終了

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除開始
'    'フォルダ選択ポップアップ画面表示
''    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")                       'V1.12.0.1 DEL
'    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.12.0.1 ADD
'
'    '指定フォルダなし
'    If Len(strWriteDir) = 0 Then
'        Set fso = Nothing           ' EG20 V3.3.0.1追加
'        Exit Sub
'    End If
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除終了

' EG20 V2.1.0.1[Mainte_03_01]追加開始
    'コピー先フォルダの有無確認
    If fso.FolderExists(strWriteDir) = False Then
        'コピー先フォルダ作成
        fso.CreateFolder (strWriteDir)
    End If
' EG20 V2.1.0.1[Mainte_03_01]追加終了

' EG20 V2.1.0.1[Mainte_03_01]削除開始
'    'コピー先フォルダパス作成(指定フォルダ￥LDU_VER)
'    strWriteDir = strWriteDir & "\" & LDU_VER
'
'    'コピー先フォルダの有無確認
'    If fso.FolderExists(strWriteDir) = False Then
'
'        'コピー先フォルダ作成
'        fso.CreateFolder (strWriteDir)
'
'    End If
' EG20 V2.1.0.1[Mainte_03_01]削除終了

    'コピー先ファイル名作成
' EG20 V2.1.0.1[Mainte_03_01]追加開始
    'コピー先ファイル名作成
    strCopySaki = strWriteDir & "\" & strStationName & "_" & strFileName

    'ファイルコピー（既に存在した場合は上書きするする）
    fso.CopyFile LDUVERLIST_REPORTFILE, strCopySaki, True
' EG20 V2.1.0.1[Mainte_03_01]追加終了

    'ファイルコピー（既に存在した場合は上書きするする）
'    fso.CopyFile strVerFile, strCopySaki, True
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    MsgBox "正常終了しました。", vbInformation + vbOKOnly, "出力結果"
   '「LDユーティリティバージョン管理画面：バージョン情報媒体出力処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_INFO_OUTPUT_OK, 0)

    Set fso = Nothing

    Exit Sub

'*******************************
'VBエラー処理
Error_cmdOutPut_Click:
   '「LDユーティリティバージョン管理画面：バージョン情報媒体出力処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OUTPUT_ERROR, lngErrCode)
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    MsgBox "異常終了しました。", vbCritical, "出力結果"
    Set fso = Nothing

'*******************************

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
   
   '「LDユーティリティバージョン管理画面：消去」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_VERASION_KANRI_GAMEN_END, 0)
   frmVersion.ZOrder
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psVersionDisp
'//  機能名称  : バージョン情報表示処理
'//  機能概要  : バージョン情報表示部の表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psVersionDisp()

    Dim strFilePath     As String   'バージョンファイルパス
    Dim bRet            As Boolean  '戻り値
    Dim intFileNo       As Integer  'ファイル番号
    Dim strWork         As String   '作業エリア
    Dim strVerData      As String   '全体バージョン
    Dim intCnt          As Integer  'カウンター
    Dim lngErrCode      As Long     'エラーコード

'*******************************
'VBエラー処理
On Error GoTo Error_psVersionDisp
'*******************************

    '媒体出力釦押下不可
    cmdOutPut.Enabled = False

    'リスト初期化
    LstFile.Clear

    '全体バージョン初期化
' EG20 V2.1.0.1[Mainte_03_01]削除開始
'    lblZenVer.Caption = "全体バージョン（ワーク）:--.--.--.--" & vbCrLf & _
'                        "　　　　　　　（実行）　:--.--.--.--" & vbCrLf & _
'                        "　　　　　　　（旧）    :--.--.--.--"
' EG20 V2.1.0.1[Mainte_03_01]削除終了
' EG20 V2.1.0.1[Mainte_03_01]追加開始
    lblZenVer.Caption = HEADERTITLE_WRK & HEADERVERSION_NON & vbCrLf & _
                        HEADERTITLE_NOW & HEADERVERSION_NON & vbCrLf & _
                        HEADERTITLE_OLD & HEADERVERSION_NON
' EG20 V2.1.0.1[Mainte_03_01]追加終了

    '作業エリア初期化
    strWork = ""

    '全体バージョン初期化
    strVerData = ""

    'LDユーティリティ画面表示用バージョンファイルパス作成
    strFilePath = PATH_LDU_APP & PATH_LDU_WORK & LDU_VER_FILE
    
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ 共通DA:LDユーティリティ画面表示用バージョンファイル作成
    '///////////////////////////////////////////////////////////////////////////////////////////
'    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, strFilePath, PATH_LDU_APP)                       ' EG20 V2.1.0.1[Mainte_03_01]削除
    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, strFilePath, PATH_LDU_APP, VERLISTKIND_DISP, 1)   ' EG20 V2.1.0.1[Mainte_03_01]追加

    'LDユーティリティ画面表示用バージョンファイル作成成功
    If bRet Then
       '「LDユーティリティバージョン管理画面：バージョン情報ファイル作成正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    'LDユーティリティ画面表示用バージョンファイル作成失敗
    Else
       '「LDユーティリティバージョン管理画面：バージョン情報ファイル作成異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
       Exit Sub
    End If

    'LDユーティリティ画面表示用バージョンファイルの有無確認
    If Len(Trim(Dir(strFilePath))) = 0 Then
        Exit Sub
    End If

    'LDユーティリティ画面表示用バージョンファイルのファイル番号を取得する。
    intFileNo = FreeFile

    'LDユーティリティ画面表示用バージョンファイルオープン
    Open strFilePath For Input As #intFileNo


        'ワーク
        Line Input #intFileNo, strWork

        If (Trim(strWork) = "") Then
'            strVerData = "全体バージョン（ワーク）：--.--.--.--" & vbCrLf                              ' EG20 V2.1.0.1[Mainte_03_01]削除
            strVerData = HEADERTITLE_WRK & HEADERVERSION_NON & vbCrLf                                   ' EG20 V2.1.0.1[Mainte_03_01]追加
        Else
            '全体バージョン文字列作成
'            strVerData = strVerData & strWork & vbCrLf                                                 ' EG20 V2.1.0.1[Mainte_03_01]削除
            strVerData = strWork & vbCrLf                                                               ' EG20 V2.1.0.1[Mainte_03_01]追加
        End If

        '実行
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "　　　　　　　（実行）　：--.--.--.--" & vbCrLf                 ' EG20 V2.1.0.1[Mainte_03_01]削除
            strVerData = strVerData & HEADERTITLE_NOW & HEADERVERSION_NON & vbCrLf                      ' EG20 V2.1.0.1[Mainte_03_01]追加
        Else
            '全体バージョン文字列作成
            strVerData = strVerData & strWork & vbCrLf
        End If

        '旧
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "　　　　　　　（旧）    ：--.--.--.--" & vbCrLf                 ' EG20 V2.1.0.1[Mainte_03_01]削除
            strVerData = strVerData & HEADERTITLE_OLD & HEADERVERSION_NON & vbCrLf                      ' EG20 V2.1.0.1[Mainte_03_01]追加
        Else
            '全体バージョン文字列作成
            strVerData = strVerData & strWork & vbCrLf
        End If

        '全体バージョン出力
        lblZenVer.Caption = strVerData

        strWork = ""

        'リスト表示分読み込み（ファイル終端までループを繰り返す）
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1削除
        Do While Not EOF(intFileNo)                         ' EG20 V3.3.0.1追加

            Line Input #intFileNo, strWork

            '改行コードのみは読みとばす
            If Trim(strWork) <> "" Then

                'リストに出力
                LstFile.AddItem (strWork)

            End If
        Loop

    'ファイルクローズ
    Close #intFileNo

    '媒体出力釦押下可
    cmdOutPut.Enabled = True

    Exit Sub

'*******************************
'VBエラー処理
Error_psVersionDisp:
   '「LDユーティリティバージョン管理画面：バージョン情報ファイル作成異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'    ファイルクローズ
    Close #intFileNo
'*******************************
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
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応（監視盤バージョンアップ対応）
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then          ' EG20 V3.0.0.2削除
    If pfVersionDispMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then   ' EG20 V3.0.0.2追加
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmLduVer.Caption, False
        pfFormActive (frmLduVer.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'/
'/  関数名称     : sCmdBtnEnabled
'/  機能名称     : コマンドボタン押下可・不可処理
'/  機能概要     : コマンドボタンを引数に基いて押下可・不可処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                EG20フェーズ２対応
'//                EG20統合監視盤USDM対応番号【Mainte_02_06】
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)
    Dim iLoopCnt    As Integer

    'フォルダ選択部：選択有り
    chkFolder(0).Enabled = blnFlg
    chkFolder(1).Enabled = blnFlg
    chkFolder(2).Enabled = blnFlg

    cmdRefresh.Enabled = blnFlg                     ' 表示更新
    cmdCopyBaitai_Work.Enabled = blnFlg             ' 媒体→ワークコピー
    cmdCopyWork_Jikko.Enabled = blnFlg              ' ワーク→実行コピー
    cmdCopyOld_Jikko.Enabled = blnFlg               ' 旧→実行コピー
    cmdOutPut.Enabled = blnFlg                      ' バージョン情報媒体出力
    cmdRemove.Enabled = blnFlg                      ' 媒体取外
    cmdCancel.Enabled = blnFlg                      ' バージョン管理画面へ戻る

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
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
'//  ORIGINAL  : (EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応【結合TR-No.123】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplTimer_Timer()

    Dim bIDURet As Boolean
    Dim bLDURet As Boolean

    On Error Resume Next

    If CheckAppStart(PROC_KANRI) <> 0 Then
        If lngtime >= lngAplMAX_Time Then
            tmrAplTimer.Enabled = False
            '管理、IDUログ、LDUログが終了していなければ、終了処理異常
            '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)

' EG20 V5.2.0.1削除開始
'            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
'                    vbOKOnly + vbExclamation, _
'                    "LDユーティリティバージョン管理"
' EG20 V5.2.0.1削除終了
' EG20 V5.2.0.1追加開始
            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
                    vbOKOnly + vbExclamation, _
                    "ＬＤＵバージョン管理"
' EG20 V5.2.0.1追加終了
            sCmdBtnEnabled True                         ' 画面操作可
        Else
            'タイマ張り直し
            tmrAplTimer.Interval = MN_MAIL_INTERVAL
            lngtime = lngtime + MN_MAIL_INTERVAL
        End If
    Else
        tmrAplTimer.Enabled = False
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
            lngtime = MN_MAIL_INTERVAL
            tmrLogTimer.Enabled = True
        Else
            '管理、IDUログ、LDUログが終了していなければ、終了処理異常
            '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
' EG20 V5.2.0.1削除開始
'            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
'                    vbOKOnly + vbExclamation, _
'                    "LDユーティリティバージョン管理"
' EG20 V5.2.0.1削除終了
' EG20 V5.2.0.1追加開始
            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
                    vbOKOnly + vbExclamation, _
                    "ＬＤＵバージョン管理"
' EG20 V5.2.0.1追加終了
            sCmdBtnEnabled True                         ' 画面操作可
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : tmrLogTimer_Timer
'//  機能名称  : ログ起動チェックタイマ処理
'//  機能概要  : ログ起動チェックタイマ処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//  ORIGINAL  : (EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応【結合TR-No.123】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()

    On Error Resume Next

    If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
        Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then

        If lngtime >= lngLogMAX_Time Then
            'ログ起動チェックタイマを停止する。
            tmrLogTimer.Enabled = False
            '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
' EG20 V5.2.0.1削除開始
'            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
'                    vbOKOnly + vbExclamation, _
'                    "LDユーティリティバージョン管理"
' EG20 V5.2.0.1削除終了
' EG20 V5.2.0.1追加開始
            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
                    vbOKOnly + vbExclamation, _
                    "ＬＤＵバージョン管理"
' EG20 V5.2.0.1追加終了
            sCmdBtnEnabled True                         ' 画面操作可
        Else
            'タイマ張り直し
            tmrLogTimer.Interval = MN_MAIL_INTERVAL
            lngtime = lngtime + MN_MAIL_INTERVAL
        End If
    Else
        tmrLogTimer.Enabled = False
        '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)

        '切替ツール起動
        Call AplVersionChangeProc(lngChangeKind)

        '終了処理
        psEndHoshuProc
        '保守プロセス終了
        End
    End If
End Sub

