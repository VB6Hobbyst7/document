VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKVer 
   BorderStyle     =   0  'なし
   Caption         =   "バージョン管理（監視盤）"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -60
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer tmrLogTimer 
      Left            =   11520
      Top             =   1560
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   11520
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   19
      Top             =   4560
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
      TabIndex        =   18
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ワーククリア"
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
      TabIndex        =   17
      Top             =   5280
      Width           =   2415
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
      TabIndex        =   16
      Top             =   3120
      Width           =   2415
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
      Height          =   615
      Left            =   9360
      TabIndex        =   14
      Top             =   6960
      Width           =   2415
   End
   Begin VB.ListBox lstKan 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   8655
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
      TabIndex        =   5
      Top             =   6240
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
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame fraVersion 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   9360
      TabIndex        =   7
      Top             =   480
      Width           =   2055
      Begin VB.CheckBox chkFolder 
         Caption         =   "W ワーク"
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
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "N 実行"
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
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "O 旧"
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
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " バージョン管理   画面へ戻る"
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
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   9000
      Top             =   7200
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "統合監視盤バージョン管理"
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
      TabIndex        =   15
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblKansibanVersion 
      Caption         =   "全体バージョン"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label lblKan 
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
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ﾌｫﾙﾀﾞ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ｻｲｽﾞ(ﾊﾞｲﾄ)"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblKan 
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
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblKan 
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
      Index           =   4
      Left            =   7080
      TabIndex        =   8
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmKVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmKVer.frm
'//  パッケージ名：バージョン管理(監視盤)画面
'//
'//  概要：バージョン管理(監視盤)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・EG10保守より、バージョン管理(監視盤)画面(frmKVer.frm)流用。
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.100】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.100】【結合TR-No.184】
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.273修正対応】
'//                 EG20フェーズ２対応【03統合TR-No.22修正対応】
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                 【残件:保守運改の切替結果通知対応】
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ2,3統合対応
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【アプリ切替改善対応】
'//     REVISIONS :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS ：(EG20 V30.3.0.1)2014-10-23  CODED BY  [TCC] T.Nakajima
'//                  北陸新幹線フェーズ２対応（媒体取外しエラー対応）
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
Dim uVersion() As MN_VERSION_LIST       'バージョン情報格納エリア

'フォルダ種別部
Public mlngChkFolderType        As Long

Private Const VERSION_STA = 28
Private Const VERSION_SIZE = 12
Private Const VERMOJI_STA = 1
Private Const FOLDER_STS = 27
Private Const HIDUKE_STA = 40
Private Const VERSION_END = 30

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
Private Const HEADERTITLE_WRK = "統合監視盤バージョン（ワーク）："
Private Const HEADERTITLE_NOW = "　　　　　　　　　　（実行）　："
Private Const HEADERTITLE_OLD = "　　　　　　　　　　（旧）　　："
Private Const HEADERVERSION_NON = "--.--.--.--"
' EG20 V2.1.0.1[Mainte_03_01] 追加終了

' EG20 V3.3.0.1【結合TR-No.184】 追加開始
Private Const APL_INTERVAL = 390000         ' アプリ起動タイマデフォルト値
Private Const LOG_INTERVAL = 30000          ' ログ起動タイマデフォルト値(30秒)
Dim lngAplMAX_Time As Long                  ' INI取得設定値（ＡＰＬ）
Dim lngLogMAX_Time As Long                  ' INI取得設定値（ログ）
Dim lngtime        As Long                  ' 現在タイマ値
Dim lngChangeKind  As Long                  ' バージョン切替種別
' EG20 V3.3.0.1【結合TR-No.184】 追加終了

' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
Private Const DESHU_ID = 242                              'デ集1コーナID
Private Const WAIT_TIME_OUT = 180000                      'タイムアウト値（３分）
Private Const DESHU_CONNECT = 1                           'デ集接続設定
Private Const GATE_CONNECT = 2                            '改札機接続設定
Private Const ERROR_TUSHIN_DISP = 1                       '通信切断異常メッセージ表示
Private Const ERROR_MISOU_DISP = 2                        '未送データ出力失敗メッセージ表示
Private Const ERROR_END_DISP = 3                          '異常終了メッセージ表示
Private udtMail          As MAIL_CONECT_CMD               '通信設定要求CMD
Public miCornerNo        As Integer                       'コーナー番号
Public mbMisouResult     As Boolean                       '未送データ作成結果　TRUE：正常　FALSE：異常
Public miErrorSts        As Integer                       '異常時通信種別
Public miErrorDisp       As Integer                       '異常時表示文言
Private byDeshuCnctSet(CONECT_CORNER_MAXINDEX)  As Byte   'デ集切離設定
Private byGateCnctSet(CONECT_JIKAI_CHK_MAX)   As Byte     '自改切離設定
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdClear_Click
'//  機能名称    ：ワーククリア
'//  機能概要    ：
'//
'//                   型          名称            意味
'//  引数        ：
'//  戻り値      ：
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ： (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 【運改表示改善対応】
'//  REVISIONS   ： (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ： (EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  量産対応【アプリ切替改善対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()
   
    Dim iResponse As Integer         ' MsgBoxボタンコード
    Dim bResult As Boolean           ' 処理結果
    
    On Error Resume Next

    '「バージョン管理画面：ワーククリア釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_CREA_BUTTOM, 0)
    
    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("「ワーク」フォルダ内のファイルを、" _
           & Chr(vbKeyReturn) & "全て削除します。    よろしいですか？", _
           vbOKCancel + vbExclamation, _
           "ワーク クリア")
    
    If iResponse <> vbCancel Then
        sCmdBtnEnabled False                        ' 画面操作不可
        '[はい] ボタンを選択した場合
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        'ワークフォルダ内のファイルを削除する
        bResult = sWrkFolderRemove
        sCmdBtnEnabled True                         ' 画面操作可
        If bResult = True Then
            ' 監視盤のバージョン情報を表示する｡
            Call psVersionDisp
            
' EG20 V5.8.0.1削除開始
'            ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'            Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
            ' 運改状態更新
'            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_KANSI, BOOTINFO_UNKAI_NASHI)      ' EG20 V5.11.0.1削除
            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_KANSI, BOOTINFO_UNKAI_CLEAR)       ' EG20 V5.11.0.1追加
            Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMEKANSI, BOOTINFO_UNKAI_NASHI)
' EG20 V5.8.0.1追加終了

' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD START
            ' 切替実行コピーツールパラメータ更新処理
            Call funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_CLEAR)
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD END

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        End If
    End If

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1追加終了

End Sub

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
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()

    On Error Resume Next
    '「媒体→ワークコピー」ボタンの場合。
    '「バージョン管理画面：媒体→ワークコピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_USB_COPY_WRK_BUTTOM, 0)
    
    sCmdBtnEnabled False                        ' 画面操作不可
    'インストール媒体をワークフォルダ内にコピーする
    Call sFDInstall
    sCmdBtnEnabled True                         ' 画面操作可
    Call psVersionDisp

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

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
'//                 EG20フェーズ２対応【結合TR-No.184】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.22修正対応】
'//                 EG20フェーズ２対応【統合TR-No.372修正対応】
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
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

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1追加終了

    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("「旧」フォルダの内容を、" _
            & Chr(vbKeyReturn) & "「実行」フォルダに戻すことにより、" _
            & Chr(vbKeyReturn) & "統合監視盤の一世代前バージョンを､実行バージョンとします｡" _
            & Chr(vbKeyReturn) & "よろしいですか？", _
           (vbOKCancel + vbExclamation), _
           "旧→実行　コピー")
    If iResponse = vbCancel Then
        Exit Sub
    End If
        
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加開始
    ' 旧バージョンフォルダに代表バージョンファイルが存在しない場合は異常とする。
    ' 旧バージョン・KANSI・統合監視盤・
    bRet = dllCheckAplVersion(4, PATH_KANSI, 2)
    If bRet = False Then
        MsgBox "異常終了しました。", vbCritical, "旧→実行　コピー"
        Exit Sub
    End If
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加終了
        
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD START
    ' 切替実行コピーツールパラメータ更新処理
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONDOWN)
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
'EG20 V3.6.0.1【03統合TR-No.22修正対応】追加開始
'    ' 統合監視盤が起動中の場合にメッセージボックスを表示する。
'    iAplChk = CheckAppStart(PROC_KANRI)
'    If iAplChk <> 0 Then
''EG20 V3.6.0.1【03統合TR-No.22修正対応】追加終了
'        '確認ポップアップウィンドウを表示する。
'        iResponse = MsgBox("統合監視盤､ＩＤＵ､ＬＤＵアプリケーションを" _
'                & Chr(vbKeyReturn) & "終了します。よろしいですか？", _
'               (vbOKCancel + vbExclamation), _
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
' EG20 V3.3.0.1【結合TR-No.184】削除開始
'' EG20 V2.1.0.1[Mainte_03_01]追加開始
'    ' アプリケーションバージョン切替実行処理
'    If (AplVersionChangeProc(ML_DT_VERSIONDOWN) = False) Then
'        ' // 保守を終了する。
'        Call psEndHoshuProc
'        '保守プロセス終了
'        End
'    End If
'' EG20 V2.1.0.1[Mainte_03_01]追加終了
' EG20 V3.3.0.1【結合TR-No.184】削除終了
' EG20 V3.3.0.1【結合TR-No.184】追加開始

    sCmdBtnEnabled False                            ' 画面操作不可
    ' 統合監視盤へアプリ終了要求の送信
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
               vbOKOnly + vbExclamation, _
               "監視盤バージョン管理"
        sCmdBtnEnabled True                         ' 画面操作可
    Else

        lngtime = MN_MAIL_INTERVAL                  ' 現在タイマ値初期化
        tmrAplTimer.Enabled = True                  ' 現在タイマ起動
    
        lngChangeKind = ML_DT_VERSIONDOWN           ' 切替種別を設定
    End If
' EG20 V3.3.0.1【結合TR-No.184】追加終了


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
'//                 EG20フェーズ２対応【結合TR-No.184】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.22修正対応】
'//                 EG20フェーズ２対応【統合TR-No.372修正対応】
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 量産対応【アプリ切替改善対応】
'//  REVISIONS    :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
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

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1追加終了

    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("「ワーク」フォルダの内容を、" _
            & Chr(vbKeyReturn) & "「実行」フォルダに登録することにより、" _
            & Chr(vbKeyReturn) & " 統合監視盤の最新バージョンを、実行バージョンとします。" _
            & Chr(vbKeyReturn) & "よろしいですか？", _
           vbOKCancel + vbExclamation, _
           "ワーク→実行 コピー")
    If iResponse = vbCancel Then
        Exit Sub
    End If
        
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加開始
    ' ワークバージョンフォルダに代表バージョンファイルが存在しない場合は異常とする。
    ' ワークバージョン・KANSI・統合監視盤
    bRet = dllCheckAplVersion(1, PATH_KANSI, 2)
    If bRet = False Then
        MsgBox "異常終了しました。", vbCritical, "ワーク→実行 コピー"
        Exit Sub
    End If
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加終了

' EG20 V7.3.0.1【EG20_KANSI03_01】ADD START
        'ワーク→実行コピー前処理
    bRet = fWorktoNow_Before1
    If bRet = False Then
        Exit Sub
    End If
' EG20 V7.3.0.1【EG20_KANSI03_01】ADD END

' EG20 V7.3.0.1【EG20_KANSI03_01】DEL START
'' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD START
'    ' 切替実行コピーツールパラメータ更新処理
'    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONUP)
'    If bRet = False Then
'        MsgBox "異常終了しました。", vbCritical, "ワーク→実行 コピー"
'        Exit Sub
'    End If
'
'    ' 終了確認
'    iResponse = MsgBox("実行コピーを適用するために統合監視盤を" & Chr(vbKeyReturn) _
'                        & "再起動しますか？", _
'                        vbOKCancel + vbExclamation, _
'                        "ワーク→実行 コピー")
'    If iResponse = vbCancel Then
'        Exit Sub
'    End If
'' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD END
'
'' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】DEL START
'''EG20 V3.6.0.1【03統合TR-No.22修正対応】追加開始
''    ' 統合監視盤が起動中の場合にメッセージボックスを表示する。
''    iAplChk = CheckAppStart(PROC_KANRI)
''    If iAplChk <> 0 Then
'''EG20 V3.6.0.1【03統合TR-No.22修正対応】追加終了
''        '確認ポップアップウィンドウを表示する。
''        iResponse = MsgBox("統合監視盤､ＩＤＵ､ＬＤＵアプリケーションを" _
''                & Chr(vbKeyReturn) & "終了します。よろしいですか？", _
''               vbOKCancel + vbExclamation, _
''               "終了確認")
''
''        If iResponse = vbCancel Then
''            Exit Sub
''        End If
''    End If  'EG20 V3.6.0.1【03統合TR-No.22修正対応】追加
'' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】DEL END
'
'' EG20 V2.1.0.1[Mainte_03_01]削除開始
'' AplVersionChangeProcにモジュール化
''    ' メールの送信内容を編集する
''    udtSendData.udtlHeader.dwId = ML_ID_KANEND_REQ      ' メールＩＤ　=”"監視装置終了要求"
''    udtSendData.udtlHeader.dwSize = MlSize.KANEND_REQ   ' メールサイズ=”"監視装置終了要求"
''    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' 送信元プロセスＩＤ=”保守”
''    udtSendData.udtlHeader.dwSubArea = 0                ' 補助情報　=　0
''
''    udtSendData.dwStartProc = ML_DT_VERSIONUP           ' 起動プロセス種別 = バージョンアップ
''    ' 送信サイズを設定する。
''    lngSendSize = udtSendData.udtlHeader.dwSize
''
''    ' 監マに対して、設定情報要求メールを送信する。
''    bRet = DssSendMail(MAIL_SLOT_KANRI, lngSendSize, udtSendData.udtlHeader)
''    ' メールを正常に送信した時のログ
''    If bRet = False Then
''        '「設定情報要求メール送信異常」ログ出力
''        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
''        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, lngErrCode)
''    Else
''        '「設定情報要求メール送信正常」ログ出力
''        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, 0)
''    End If
'' EG20 V2.1.0.1[Mainte_03_01]削除終了
'' EG20 V3.3.0.1【結合TR-No.184】削除開始
''' EG20 V2.1.0.1[Mainte_03_01]追加開始
''    ' アプリケーションバージョン切替実行処理
''    If (AplVersionChangeProc(ML_DT_VERSIONUP) = False) Then
''        ' // 保守を終了する。
''        Call psEndHoshuProc
''        '保守プロセス終了
''        End
''    End If
''' EG20 V2.1.0.1[Mainte_03_01]追加終了
'' EG20 V3.3.0.1【結合TR-No.184】削除終了
'' EG20 V3.3.0.1【結合TR-No.184】追加開始
'
'    sCmdBtnEnabled False                            ' 画面操作不可
'    ' 統合監視盤へアプリ終了要求の送信
'    bRet = pubFuncAplEndRequest()
'    If bRet = False Then
'        MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
'               vbOKOnly + vbExclamation, _
'               "監視盤バージョン管理"
'        sCmdBtnEnabled True                         ' 画面操作可
'    Else
'
'        lngtime = MN_MAIL_INTERVAL                  ' 現在タイマ値初期化
'        tmrAplTimer.Enabled = True                  ' 現在タイマ起動
'
'        lngChangeKind = ML_DT_VERSIONUP             ' 切替種別を設定
'    End If
'' EG20 V3.3.0.1【結合TR-No.184】追加終了
' EG20 V7.3.0.1【EG20_KANSI03_01】DEL END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン管理(監視盤)画面(アクティブ時)
'//  機能概要  : メール受信用のタイマ起動
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
'//  機能名称  : バージョン管理(監視盤)画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
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
'//  機能名称  : バージョン管理(監視盤)画面(ロード時)
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
'//                 EG20フェーズ２対応【結合TR-No.184】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
   
   '「監視盤バージョン管理画面：表示」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_KANRI_GAMEN_START, 0)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    '初期化
    lstKan.Clear
    mlngChkFolderType = 0

    'フォルダ選択部：選択有り
    chkFolder(0).Value = 1
    chkFolder(1).Value = 1
    chkFolder(2).Value = 1
    
    mlngChkFolderType = 7
    
' EG20 V2.1.0.1[Mainte_03_01]削除開始
''    監視盤のバージョン番号を表示する｡
'    sKansibanVersionSet
'
''   バージョン情報のリストボックスを作成する
'    fMakeListbox
' EG20 V2.1.0.1[Mainte_03_01]削除終了
' EG20 V2.1.0.1[Mainte_03_01]追加開始
    ' 統合監視盤のバージョン情報を表示する｡
    Call psVersionDisp
' EG20 V2.1.0.1[Mainte_03_01]追加終了
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   
'   メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

' EG20 V3.3.0.1【結合TR-No.184】 追加開始
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
' EG20 V3.3.0.1【結合TR-No.184】 追加終了

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
'//  機能名称  : 「表示更新」釦押下処理
'//  機能概要  : 最新の状態を表示する。
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
    
    '「監視盤バージョン管理画面：表示更新釦押下」
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
                "監視盤バージョン管理"
        Exit Sub
    End If
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   
   'バージョン情報のリストボックスを作成する
'    fMakeListbox           ' EG20 V2.1.0.1[Mainte_03_01]削除
    Call psVersionDisp      ' EG20 V2.1.0.1[Mainte_03_01]追加

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdOutPut_Click
'//  機能名称  : 「バージョン情報媒体出力」釦押下処理
'//  機能概要  : 表示されたバージョン情報を媒体に出力する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 フォルダ選択ポップアップ画面の初期フォルダ変更
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.100】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()
'*******************************
'VBエラー処理
On Error GoTo Error_cmdOutPut_Click
'*******************************
    Dim iRet        As Integer                '戻り値
    Dim strCopySaki As String                 '出力先ファイルパス
    Dim strWriteDir As String                 '出力先フォルダ
    Dim fso         As New FileSystemObject   'ファイルシステムオブジェクト
    Dim iFileNumber As Integer                'ファイル番号
    Dim iMaxLine As Integer                   'リストボックスの行数
    Dim iLine As Integer                      'リストボックスの行カウンタ
    Dim sCopymoto As String                   '出力元ファイルパス
    Dim lngErrCode  As Long              'エラーコード
    
    Dim strStationName       As String          ' 駅名名                ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim szCornerName         As String          ' コーナ名称            ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim nNullIndex           As Integer         ' 文字数ワーク          ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim strWork              As String          ' ワーク                ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim strFileName         As String           ' ファイル名            ' EG20 V2.1.0.1[Mainte_03_01]追加
    Dim bRet                As Boolean          ' 戻り値                ' EG20 V2.1.0.1[Mainte_03_01]追加

   '「監視盤バージョン管理画面：バージョン情報媒体出力釦押下」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OUTPUT, 0)

' EG20 V3.3.0.1 【結合TR-No.100】追加開始
    ' リストに１件もデータがない場合は異常終了
    If lstKan.ListCount = 0 Then
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Exit Sub
    End If
' EG20 V3.3.0.1 【結合TR-No.100】追加終了

' EG20 V2.1.0.1[Mainte_03_01]追加開始
    strStationName = gsGetStationEkiName
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'フォルダ選択ポップアップ画面表示
    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)

    '指定フォルダなし
    If Len(strWriteDir) = 0 Then
        Exit Sub
    End If

    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ 保守専用関数:操作卓バージョンファイル（画面表示用）作成
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllKansiCreateVerFile(mlngChkFolderType, MN_VERSI_FILE, VERLISTKIND_REPORT)
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
       Exit Sub
    End If

    'ファイルの有無確認
    If fso.FileExists(MN_VERSI_FILE) = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Exit Sub
    End If
    strFileName = Dir(MN_VERSI_FILE)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除開始（処理移動）
'    'フォルダ選択ポップアップ画面表示
'    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
'
'    '指定フォルダなし
'    If Len(strWriteDir) = 0 Then
'        Exit Sub
'    End If
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除終了

    'コピー先フォルダの有無確認
    If fso.FolderExists(strWriteDir) = False Then
        'コピー先フォルダ作成
        fso.CreateFolder (strWriteDir)
    End If

    'コピー先ファイル名作成
    strCopySaki = strWriteDir & "\" & strStationName & "_" & strFileName

    'ファイルコピー（既に存在した場合は上書きするする）
    fso.CopyFile MN_VERSI_FILE, strCopySaki, True
' EG20 V2.1.0.1[Mainte_03_01]追加終了
' EG20 V2.1.0.1[Mainte_03_01]削除開始
''V1.7.0.1 DEL START
''    'フォルダ選択ポップアップ画面表示
''    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")
''
''    '指定フォルダなし
''    If Len(strWriteDir) = 0 Then
''        Exit Sub
''    End If
''V1.7.0.1 DEL END
'    iFileNumber = FreeFile              '未使用のファイル番号を取得する
'
'    sCopymoto = PATH_WORK + VER_TXT_NAME
'
'    'バージョンテキストファイルをオープンする。ファイルがなければ新規に作成する。
'    Open sCopymoto For Output Access Write As #iFileNumber
'
'    iMaxLine = lstKan.ListCount
'    For iLine = 0 To lstKan.ListCount - 1
'        'リストボックス１行分ををバージョンテキストファイルに書き込む。
'        Print #iFileNumber, lstKan.List(iLine) & Chr(vbKeyReturn)
'    Next
'    'バージョンテキストファイルをクローズする。
'    Close #iFileNumber
'
''V1.7.0.1 DEL START
''    'コピー先フォルダの有無確認
''    If fso.FolderExists(strWriteDir) = False Then
''        'コピー先フォルダ作成
''        fso.CreateFolder (strWriteDir)
''    End If
''V1.7.0.1 DEL END
'
'   'ファイルの有無確認
'    If fso.FileExists(sCopymoto) = False Then
'        'ファイル無し異常ポップアップ画面表示
'        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
'        Exit Sub
'    End If
'
''V1.7.0.1 ADD  START
'    'フォルダ選択ポップアップ画面表示
''    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")                       'V1.12.0.1 DEL
'    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.12.0.1 ADD
'
'    '指定フォルダなし
'    If Len(strWriteDir) = 0 Then
'        Exit Sub
'    End If
'
'    'コピー先フォルダの有無確認
'    If fso.FolderExists(strWriteDir) = False Then
'        'コピー先フォルダ作成
'        fso.CreateFolder (strWriteDir)
'    End If
''V1.7.0.1 ADD END
'
'    'コピー先ファイル名作成
'    strCopySaki = strWriteDir & "\" & VER_TXT_NAME
'
'    'ファイルコピー（既に存在した場合は上書きするする）
'    fso.CopyFile sCopymoto, strCopySaki, True
' EG20 V2.1.0.1[Mainte_03_01]削除終了

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
  
    '出力結果ポップアップ(正常)表示
    MsgBox "正常終了しました。", vbInformation + vbOKOnly, "出力結果"
    '「監視盤バージョン管理画面：バージョン情報媒体出力処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_INFO_OUTPUT_OK, 0)
    
    Set fso = Nothing
    
    Exit Sub
'*******************************
'VBエラー処理
Error_cmdOutPut_Click:
' EG20 V2.1.0.1[Mainte_03_01]削除開始
'        'V1.21.0.1 ADD  START
'        If iFileNumber > 0 Then
'           Close #iFileNumber
'        End If
'        'V1.21.0.1 ADD  END
' EG20 V2.1.0.1[Mainte_03_01]削除終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '処理異常の場合、出力結果ポップアップ(異常)表示
        MsgBox "異常終了しました。", vbCritical, "出力結果"
        '「監視盤バージョン管理画面：バージョン情報媒体出力処理異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OUTPUT_ERROR, lngErrCode)
        Set fso = Nothing
'*******************************
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdRemove_Click
'//  機能名称  : 「媒体取外」釦押下処理
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
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '「監視盤バージョン管理画面：消去」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_KANRI_GAMEN_END, 0)
 
 ' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1追加終了

    'バージョン管理（監視盤）画面を閉じる
    Unload Me
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：psVersionDisp
'//  機能名称    ：バージョン情報作成処理
'//  機能概要    ：バージョン情報ファイル作成/画面表示を行う。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub psVersionDisp()
    Dim bRet            As Boolean  '戻り値
    Dim intFileNo       As Integer  'ファイル番号
    Dim strWork         As String   '作業エリア
    Dim strVerData      As String   '全体バージョン
    Dim lngErrCode      As Long     'エラーコード
    Dim strList         As String
    Dim strVer          As String
    Dim strWork1        As String
    Dim strWork2        As String
    Dim strWork3        As String
    Dim sFileName       As String


'*******************************
'VBエラー処理
On Error GoTo Error_psVersionDisp
'*******************************

    '媒体出力釦押下不可
    cmdOutPut.Enabled = False

    'リスト初期化
    lstKan.Clear
    
    '作業エリア初期化
    strWork = ""

    '全体バージョン初期化
    strVerData = ""

    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ 保守専用関数:操作卓バージョンファイル（画面表示用）作成
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllKansiCreateVerFile(mlngChkFolderType, KANSI_VERSION_CSVFILE, VERLISTKIND_DISP)

    ' バージョンファイル成功
    If bRet Then
       '「バージョン情報ファイル作成正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    ' バージョンファイル失敗
    Else
       '「バージョン情報ファイル作成異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
       Exit Sub
    End If
    
    ' バージョンファイルの有無確認
    If Len(Trim(Dir(KANSI_VERSION_CSVFILE))) = 0 Then
        Exit Sub
    End If

    ' バージョンファイルのファイル番号を取得する。
    intFileNo = FreeFile

    ' バージョンファイルオープン
    Open KANSI_VERSION_CSVFILE For Input As #intFileNo
    
        'ワーク
        Line Input #intFileNo, strWork
        
        If (Trim(strWork) = "") Then
            strVerData = HEADERTITLE_WRK & HEADERVERSION_NON & vbCrLf
        Else
            '全体バージョン文字列作成
            strVerData = strWork & vbCrLf
        End If

        '実行
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
            strVerData = strVerData & HEADERTITLE_NOW & HEADERVERSION_NON & vbCrLf
        Else
            strVerData = strVerData & strWork & vbCrLf
        End If

        '旧
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
            strVerData = strVerData & HEADERTITLE_OLD & HEADERVERSION_NON & vbCrLf
        Else
            strVerData = strVerData & strWork & vbCrLf
        End If

        '全体バージョン出力
        lblKansibanVersion.Caption = strVerData

        strWork = ""

        'リスト表示分読み込み（ファイル終端までループを繰り返す）
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1削除
        Do While Not EOF(intFileNo)                         ' EG20 V3.3.0.1追加

            Line Input #intFileNo, strWork

            '改行コードのみは読みとばす
            If Trim(strWork) <> "" Then

                strWork1 = Right(strWork, 42)
                strWork2 = Mid(strWork1, 1, 12)   'サイズのみ抽出
                strWork3 = Mid(strWork1, 13, 30)
                strVer = Format(strWork2, "#,##0")
                strVer = Format(strVer, "@@@@@@@@@@@@")
                sFileName = StrConv(MidB(StrConv(Mid(strWork, 1, 27) & Space(20), vbFromUnicode), 1, 27), vbUnicode)
                strList = sFileName & strVer & strWork3
                'リストに出力
                lstKan.AddItem (strList)

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
    'バージョン情報ファイル作成異常ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
    'ファイルクローズ
    Close #intFileNo
'*******************************

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sKansibanVersionSet
'//  機能名称  : 監視盤のバージョン取得表示する。
'//  機能概要  : KansiVersion.iniより、バージョンを取得・表示する。
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sKansibanVersionSet()
    Dim lSts As Long                            '関数戻り値
    Dim strKansiVersion As String * 128         '監視盤全体バージョン
    Dim strKansiVersionNow As String            ' 監視盤全体バージョン（現行）  EG20 V2.1.0.1[Mainte_03_01]追加
    Dim strKansiVersionOld As String            ' 監視盤全体バージョン（旧）    EG20 V2.1.0.1[Mainte_03_01]追加
    Dim strKansiVersionWrk As String            ' 監視盤全体バージョン（ワーク）EG20 V2.1.0.1[Mainte_03_01]追加
    
    On Error Resume Next
        
' EG20 V2.1.0.1[Mainte_03_01] コメント追加開始
    ' /////////////////////////////////////////////////////
    ' // 実行バージョン
' EG20 V2.1.0.1[Mainte_03_01] コメント追加終了
    
    strKansiVersion = ""
    
    ' KansiVersion.iniから監視盤の全体バージョンを取得し表示する
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSION_FILE)
     If lSts > 0 Then
        '取得したバージョン番号を表示
'        lblKansibanVersion.Caption = "全体バージョン：" & Left$(strKansiVersion, lSts)     ' EG20 V2.1.0.1[Mainte_03_01] 削除
        strKansiVersionNow = HEADERTITLE_NOW & Left$(strKansiVersion, lSts)                 ' EG20 V2.1.0.1[Mainte_03_01] 追加
    Else
        'バージョン番号の取得異常の場合、「--,--,--,--」を表示
'        lblKansibanVersion.Caption = "全体バージョン：--.--.--.-- "                        ' EG20 V2.1.0.1[Mainte_03_01] 削除
        strKansiVersionNow = HEADERTITLE_NOW & HEADERVERSION_NON                            ' EG20 V2.1.0.1[Mainte_03_01] 追加
    End If

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
    ' /////////////////////////////////////////////////////
    ' // 旧バージョン
    strKansiVersion = ""
    
    ' KansiVersion.iniから監視盤の全体バージョンを取得し表示する
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSIONOLD_FILE)
     If lSts > 0 Then
        '取得したバージョン番号を表示
        strKansiVersionOld = HEADERTITLE_OLD & Left$(strKansiVersion, lSts)
    Else
        'バージョン番号の取得異常の場合、「--,--,--,--」を表示
        strKansiVersionOld = HEADERTITLE_OLD & HEADERVERSION_NON
    End If
    
    ' /////////////////////////////////////////////////////
    ' // ワークバージョン
    strKansiVersion = ""
    
    ' KansiVersion.iniから監視盤の全体バージョンを取得し表示する
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSIONWRK_FILE)
     If lSts > 0 Then
        '取得したバージョン番号を表示
        strKansiVersionWrk = HEADERTITLE_WRK & Left$(strKansiVersion, lSts)
    Else
        'バージョン番号の取得異常の場合、「--,--,--,--」を表示
        strKansiVersionWrk = HEADERTITLE_WRK & HEADERVERSION_NON
    End If

    
    ' /////////////////////////////////////////////////////
    ' // 表示内容の合体
    lblKansibanVersion.Caption = strKansiVersionWrk & vbCrLf & _
                                strKansiVersionNow & vbCrLf & _
                                strKansiVersionOld
' EG20 V2.1.0.1[Mainte_03_01] 追加終了


End Sub

' EG20 V2.1.0.1[Mainte_03_01]削除開始
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : fMakeListbox
''//  機能名称  : バージョン表示対象よりバージョンを取得表示する。
''//  機能概要  : 旧、実行、ワーク、INI内にある、
''//              *.exe、*.dll、*.OCX、*.INIのバージョンを取得する。
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
''//                 フェーズ３　結合検査　不具合修正
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Function fMakeListbox() As Boolean
'    Dim strFilePath     As String   'バージョンファイルパス
'    Dim bRet            As Boolean  '戻り値
'    Dim intFileNo       As Integer  'ファイル番号
'    Dim strWork         As String   '作業エリア
'    Dim strVerData      As String   '全体バージョン
'    Dim intCnt          As Integer  'カウンター
'    Dim lngErrCode      As Long     'エラーコード
'    Dim strVerformat As String
'    Dim strList As String
'    Dim strVer As String
''V1.8.0.1 ADD START
'    Dim strWork1 As String
'    Dim strWork2 As String
'    Dim strWork3 As String
'    Dim strWork4 As String
'    Dim sFileName As String
''V1.8.0.1 ADD END
''*******************************
''VBエラー処理
'On Error GoTo Error_psVersionDisp
''*******************************
'
'    fMakeListbox = False
'
''    媒体出力釦押下不可
'    cmdOutPut.Enabled = False
'
''    リスト初期化
'    lstKan.Clear
'
''    作業エリア初期化
'    strWork = ""
'
''    監視盤画面表示用バージョンファイルパス作成
'    strFilePath = KANSI_VERSION_CSVFILE
'
'    bRet = True
''    ///////////////////////////////////////////////////////////////////////////////////////////
''    / 共通DA:LDユーティリティ画面表示用バージョンファイル作成
''    ///////////////////////////////////////////////////////////////////////////////////////////
'    bRet = dllKansiCreateVerFile(mlngChkFolderType, strFilePath)
'
''    監視盤画面表示用バージョンファイル作成成功
'    If bRet Then
''       「監視盤バージョン管理画面：バージョン情報ファイル作成正常」ログ出力
'       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
''    監視盤画面表示用バージョンファイル作成失敗
'    Else
''       「監視盤バージョン管理画面：バージョン情報ファイル作成異常」ログ出力
'       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'       Exit Function
'    End If
'
''    監視盤画面表示用バージョンファイルの有無確認
'    If Len(Trim(Dir(strFilePath))) = 0 Then
'        Exit Function
'    End If
'
''    監視盤画面表示用バージョンファイルのファイル番号を取得する｡
'    intFileNo = FreeFile
'
''    監視盤画面表示用バージョンファイルオープン
'    Open strFilePath For Input As #intFileNo
'
'    strWork = ""
'
''    リスト表示分読み込み (ファイル終端までループを繰り返す)
'    Do While Not EOF(1)
'
'        Line Input #intFileNo, strWork
'
''        改行コードのみは読みとばす
'        If Trim(strWork) <> "" Then
'            'バージョンファイル内のバージョン値を「zzz,zzz,zzz」フォーマットに変換する処理
'            'V1.8.0.1 DEL START
''            strVer = Mid(strWork, VERSION_STA, VERSION_SIZE)
''            strVerformat = Format(strVer, "#,##0")
''            strVerformat = Format(strVerformat, "@@@@@@@@@@@@")
''            strList = Mid(strWork, VERMOJI_STA, FOLDER_STS)
''            strList = strList & strVerformat
''            strList = strList & Mid(strWork, HIDUKE_STA, VERSION_END)
'            'V1.8.0.1 DEL END
'            'V1.8.0.1 ADD START
'            strWork1 = Right(strWork, 42)
'            strWork2 = Mid(strWork1, 1, 12)   'サイズのみ抽出
'            strWork3 = Mid(strWork1, 13, 30)
'            strVer = Format(strWork2, "#,##0")
'            strVer = Format(strVer, "@@@@@@@@@@@@")
'            sFileName = StrConv(MidB(StrConv(Mid(strWork, 1, 27) & Space(20), vbFromUnicode), 1, 27), vbUnicode)
'            strList = sFileName & strVer & strWork3
'            'V1.8.0.1 ADD END
''           リストに出力
'            lstKan.AddItem (strList)
'        End If
'    Loop
'
''    ファイルクローズ
'    Close #intFileNo
'
'    fMakeListbox = True
'
''    媒体出力釦押下可
'    cmdOutPut.Enabled = True
'
'    Exit Function
'
''*******************************
''VBエラー処理
'Error_psVersionDisp:
''   「監視盤バージョン管理画面：バージョン情報ファイル作成異常」ログ出力
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
''    ファイルクローズ
'   Close #intFileNo
''*******************************
'End Function
' EG20 V2.1.0.1[Mainte_03_01]削除終了

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
'//     REVISIONS :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    '汎用メール受信処理を行う
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then          ' EG20 V3.0.0.2削除
'    If pfVersionDispMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then   ' EG20 V3.0.0.2追加 ' EG20 V7.3.0.1【EG20_KANSI03_01】DEL
    If pfMailRecieve_KansiVerDisp = ML_ID_HOSHU_ACTIVE_REQ Then  ' EG20 V7.3.0.1【EG20_KANSI03_01】ADD
        AppActivate frmKVer.Caption, False
        pfFormActive (frmKVer.hwnd)                              ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称    ： sWrkFolderRemove
'//  機能名称    ： ワークフォルダ内ファイル削除処理
'//  機能概要    ： ワークフォルダ内のファイルを削除する。
'//
'//                 型        名称      意味
'//  引数        ： なし
'//
'//                 型        値        意味
'//  戻り値      ： なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：改札機バージョン管理画面のsWrkFolderRemove流用
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim stringWorkFolder As String      ' フォルダ名
    Dim lngErrCode As Long              'エラーコード
    
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    
    On Error GoTo ErrorHandler          'エラーハンドルの登録

    '初期値設定
    sWrkFolderRemove = True
   
    '//////////////////////////////////////////////////////////////////////////
    '// ワークフォルダ内の操作卓フォルダを消去
    ' ワークフォルダ内のディレクトリの名前を表示します。
    stringWorkFolder = PATH_KANSI_APLNEW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    Set objFso = Nothing

'    '「正常終了」ポップアップ画面表示
'    MsgBox "正常終了しました。", _
'           vbOKOnly + vbInformation, _
'           "実行結果"

    Exit Function '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '「ワーククリア異常終了」ポップアップ画面表示
     MsgBox "異常終了しました。", _
           vbOKOnly + vbCritical, _
           "実行結果"
           
   '「自改ﾊﾞｰｼﾞｮﾝ：ﾜｰｸﾌｫﾙﾀﾞﾌｧｲﾙ削除異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    Set objFso = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ2,3統合対応
'//  REVISIONS   ： (EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  量産対応【アプリ切替改善対応】
'//  REVISIONS   ： (EG20 V30.3.0.1)2014-10-23  CODED BY  [TCC] T.Nakajima
'//                  北陸新幹線フェーズ２対応（媒体取外しエラー対応）
'//  備考        ：改札機バージョン管理画面のsFDInstall流用
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall()
    Dim MyName As String            'ファイルフルパス名
    Dim iResponse As Integer        'MsgBoxボタンコード
    Dim sInputPass As String        'インストール元ディレクトリ名(STD)orファイル名(LZH)
    Dim lngErrCode As Long          'エラーコード
    
    Dim lngProcId As Long                ' プロセスID
    Dim hProc As Variant                 ' プロセスハンドル
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    Dim objFi As File                    'ファイルオブジェクト
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
    iResponse = MsgBox("選択されたインストール部材の内容を統合監視盤アプリケーションの" _
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
    
    Set objFso = Nothing
    Set objFi = Nothing
    
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
    
' EG20 V5.8.0.1削除開始
'    ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
    ' 運改状態更新
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_KANSI, BOOTINFO_UNKAI_ARI)
    Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMEKANSI, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1追加終了
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD START
    ' 切替実行コピーツールパラメータ更新処理
    Call funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_CLEAR)
' EG20 V6.9.0.1【量産対応：アプリ切替改善対応】ADD END

' EG20 V5.8.0.1 ADD START
    '読み取り外しの関数を実行
    dllChangeAttributeContents (PATH_KANSI_APLNEW)
' EG20 V5.8.0.1 ADD END
    
    Exit Sub    '処理を終了する

ErrorHandler:   ' エラー処理。
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing

' EG20 V5.8.0.1 ADD START
    '読み取り外しの関数を実行
    dllChangeAttributeContents (PATH_KANSI_APLNEW)
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
    cmdClear.Enabled = blnFlg                       ' ワーククリア
    cmdCopyBaitai_Work.Enabled = blnFlg             ' 媒体→ワークコピー
    cmdCopyWork_Jikko.Enabled = blnFlg              ' ワーク→実行コピー
    cmdCopyOld_Jikko.Enabled = blnFlg               ' 旧→実行コピー
    cmdOutPut.Enabled = blnFlg                      ' バージョン情報媒体出力
    cmdRemove.Enabled = blnFlg                      ' 媒体取外
    cmdReturn.Enabled = blnFlg                      ' バージョン管理画面へ戻る

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
'//               EG20フェーズ２対応【結合TR-No.184】
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

            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
                    vbOKOnly + vbExclamation, _
                    "監視盤バージョン管理"
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
            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
                    vbOKOnly + vbExclamation, _
                    "監視盤バージョン管理"
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
'//               EG20フェーズ２対応【結合TR-No.184】
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
            MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
                    vbOKOnly + vbExclamation, _
                    "監視盤バージョン管理"
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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : fWorktoNow_Before1
'//  機能名称  : ワーク→実行コピー前処理1
'//  機能概要  : ワーク→実行コピー前に下記処理を行う
'//　　　　　　　・監視盤アプリ未起動チェック
'//　　　　　　　・締切未送データ存在チェック
'//　　　　　　　・デ集通信切断
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　 TRUE      正常終了
'//　　　　　　　　　　　　FALSE     異常終了
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fWorktoNow_Before1() As Boolean

    Dim iCnt As Integer

    fWorktoNow_Before1 = False

    '-------------------------------------------------------------------------------------------
    '監視盤アプリ未起動時は、本シーケンスを実施しない
    '-------------------------------------------------------------------------------------------
    If CheckAppStart(PROC_KANRI) = 0 Then
        MsgBox "保守単独起動のため、ワーク→実行コピーが行えません。", _
                vbOKOnly + vbCritical, _
                "ワーク→実行 コピー"
        Exit Function
    End If

    '-------------------------------------------------------------------------------------------
    '締切未送データが存在する場合、本シーケンスを実施しない
    '-------------------------------------------------------------------------------------------
    If fChkSimekiriMisouUmu = False Then
        MsgBox "締切未送データがあるため、ワーク→実行コピーが行えません。", _
                vbOKOnly + vbCritical, _
                "ワーク→実行 コピー"
        Exit Function
    End If

    sCmdBtnEnabled False                            ' 画面操作不可

    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
    
    Erase byDeshuCnctSet  'デ集切離設定初期化
    Erase byGateCnctSet   '自改切離設定初期化
    miErrorSts = 0        '異常時通信種別初期化
    miErrorDisp = 0       '異常時異常時表示文言初期化
    
    'デ集の接続／切断設定を取得
    For iCnt = CNT_MIN To CONECT_CORNER_MAXINDEX
        If gblnCornerSet(iCnt) = True Then
            If (0 = pfGetJyouiKikiConectSet(DESHU_ID + iCnt)) Then
                byDeshuCnctSet(iCnt) = 1
            End If
        End If
    Next
    
    '自改の接続／切断設定を取得
    For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
        If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
            If (0 = pfGetGateConectSet(iCnt + 1)) Then
                byGateCnctSet(iCnt) = 1
            End If
        End If
    Next

    '-------------------------------------------------------------------------------------------
    'デ集通信切断
    '-------------------------------------------------------------------------------------------
    If False = pfKill_TusinConect(ML_DT_DESHU) Then
        '通信切断異常処理
        Call ConnctErrorProc(DESHU_CONNECT, ERROR_TUSHIN_DISP)
        Exit Function
        
    End If

    fWorktoNow_Before1 = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : fWorktoNow_Before2
'//  機能名称  : ワーク→実行コピー前処理2
'//  機能概要  : 通信設定要求RES（デ集、切断）受信後、下記処理を行う
'//　　　　　　　・自改通信切断
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　 TRUE      正常終了
'//　　　　　　　　　　　　FALSE     異常終了
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fWorktoNow_Before2() As Boolean

    fWorktoNow_Before2 = False

    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
    
    '-------------------------------------------------------------------------------------------
    '自改通信切断
    '-------------------------------------------------------------------------------------------
    If False = pfKill_TusinConect(ML_DT_JIKAI) Then
        '通信切断異常処理
        Call ConnctErrorProc(GATE_CONNECT, ERROR_TUSHIN_DISP)
        Exit Function
    End If
    
    fWorktoNow_Before2 = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : fWorktoNow_Start
'//  機能名称  : ワーク→実行コピー処理
'//  機能概要  : 未送締切データ作成後、ワーク→実行コピー処理を行う
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　 TRUE      正常終了
'//　　　　　　　　　　　　FALSE     異常終了
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY ]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fWorktoNow_Start() As Boolean

    Dim bRet        As Boolean            ' メール送信処理戻り値
    Dim iResponse   As Integer            ' MsgBoxボタンコード
    
    On Error Resume Next
    
    fWorktoNow_Start = True
    
    ' 切替実行コピーツールパラメータ更新処理
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONUP)
    If bRet = False Then
        'デ集＆自改通信接続
        Call ConnctErrorProc(GATE_CONNECT, ERROR_END_DISP)
        Exit Function
    End If

    'ワーク→実行コピー切断設定パラメータ更新処理
    bRet = funcUpdateConnectSetParam(byDeshuCnctSet, byGateCnctSet)
    If bRet = False Then
        'デ集＆自改通信接続
        Call ConnctErrorProc(GATE_CONNECT, ERROR_END_DISP)
        Exit Function
    End If

    ' 終了確認
    iResponse = MsgBox("実行コピーを適用するために統合監視盤を" & Chr(vbKeyReturn) _
                        & "再起動しますか？", _
                        vbOKCancel + vbExclamation, _
                        "ワーク→実行 コピー")
    If iResponse = vbCancel Then
        fWorktoNow_Start = False
        Exit Function
    End If
    
    sCmdBtnEnabled False                            ' 画面操作不可
    
    ' 統合監視盤へアプリ終了要求の送信
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "アプリケーションの終了処理中に異常が発生しました。", _
               vbOKOnly + vbExclamation, _
               "監視盤バージョン管理"
        sCmdBtnEnabled True                         ' 画面操作可
    Else

        lngtime = MN_MAIL_INTERVAL                  ' 現在タイマ値初期化
        tmrAplTimer.Enabled = True                  ' 現在タイマ起動
    
        lngChangeKind = ML_DT_VERSIONUP             ' 切替種別を設定
        
    End If
        
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : fChkSimekiriMisouUmu
'//  機能名称  : 締切未送データ存在チェック
'//  機能概要  : 締切未送データが存在する場合、本シーケンスを実施しない
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　 TRUE      締切未送データなし
'//　　　　　　　　　　　　FALSE     締切未送データあり
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fChkSimekiriMisouUmu() As Boolean

    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim nLoop As Integer                                ' ループ
    Dim bEnable As Boolean                              ' ボタン状態
    Dim szFileName As String

    On Error GoTo ErrorHandler                          ' エラーハンドルの登録
    
    fChkSimekiriMisouUmu = True

    For nLoop = 0 To UBound(gblnCornerSet)

        bEnable = False
        If gblnCornerSet(nLoop) = True Then
            ' /////////////////////////////////////////////////////////////////////////
            ' // 締切出力データは存在するか？（D:\KANSI\SHUKEI\OUT_DATA\CORNER##\SIME##.DAT）
            szFileName = Replace(PATH_SHUKEI_SHIMEDAT, "##", Format(nLoop + 1, "0#"))
            If objFso.FileExists(szFileName) = True Then
                fChkSimekiriMisouUmu = False
                Exit Function
            End If
        End If
        
    Next nLoop
    
    Set objFso = Nothing
    
    Exit Function

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:
    Set objFso = Nothing
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfKill_TusinConect
'//  機能名称  : 通信回線切断処理
'//  機能概要  : 指定した外部機器の通信回線を切断する
'//
'//              型        名称      意味
'//  引数      : Long      dwKiki    外部機器要求種別
'//
'//              型        値        意味
'//  戻り値    : Boolean　 TRUE      メッセージ送信正常
'//　　　　　　　　　　　　FALSE     メッセージ送信異常
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfKill_TusinConect(dwKiki As Long) As Boolean

    Dim bRet As Boolean                 'メール送信戻り値
    Dim iCnt As Integer                 'カウンター
    Dim lngErrCode As Long              'エラーコード
    
    pfKill_TusinConect = False

    '-------------------------------------------------------------------------------------------
    '通信設定要求CMDメッセージ作成
    '-------------------------------------------------------------------------------------------
    'ヘッダ部共通作成処理
    Call SendMailHeader(ML_ID_CONECT_CMD, MlSize.CONECT_CMD)
    
    'データ部設定
    udtMail.dwRequestKIKI = dwKiki
    udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
    For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
        udtMail.dwGouki(iCnt) = ML_TARGET_OFF
    Next
    
    '外部機器要求種別が自改？
    If dwKiki = ML_DT_JIKAI Then
        '外部機器要求種別が自改の場合
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
            '改札機が実装されているか？
            If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
                udtMail.dwGouki(iCnt) = ML_TARGET_ON
            End If
        Next
    Else
        '外部機器要求種別が自改以外の場合
        For iCnt = 0 To UBound(gblnCornerSet)
            'コーナ接続されているか？
            If gblnCornerSet(iCnt) = True Then
                udtMail.dwGouki(iCnt) = ML_TARGET_ON
            End If
        Next
    End If
    
    '-------------------------------------------------------------------------------------------
    '通信設定要求CMD(対象ID)を監マプロセスに送信する
    '-------------------------------------------------------------------------------------------
    bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
    If False = bRet Then
        '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
        Exit Function
    Else
        '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
    End If

    pfKill_TusinConect = True

End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfConnect_TusinConect
'//  機能名称  : 通信回線接続処理
'//  機能概要  : 指定した外部機器の通信回線を接続する
'//
'//              型        名称      意味
'//  引数      : Long      dwKiki    外部機器要求種別
'//
'//              型        値        意味
'//  戻り値    : Boolean　 TRUE      メッセージ送信正常
'//　　　　　　　　　　　　FALSE     メッセージ送信異常
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfConnect_TusinConect(dwKiki As Long) As Boolean

    Dim bRet As Boolean                 'メール送信戻り値
    Dim iCnt As Integer                 'カウンター
    Dim lngErrCode As Long              'エラーコード
    
    pfConnect_TusinConect = False
    
    '-------------------------------------------------------------------------------------------
    '通信設定要求CMDメッセージ作成
    '-------------------------------------------------------------------------------------------
    'ヘッダ部共通作成処理
    Call SendMailHeader(ML_ID_CONECT_CMD, MlSize.CONECT_CMD)
    
    'データ部設定
    udtMail.dwRequestKIKI = dwKiki
    udtMail.dwRequestConectType = ML_REQUEST_CONECT
    For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
        udtMail.dwGouki(iCnt) = ML_TARGET_OFF
    Next
    
    '外部機器要求種別が自改？
    If dwKiki = ML_DT_JIKAI Then
        '外部機器要求種別が自改の場合
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
            '改札機が実装されているか？
            If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
                udtMail.dwGouki(iCnt) = byGateCnctSet(iCnt)
            End If
        Next
    Else
        '外部機器要求種別が自改以外の場合
        For iCnt = 0 To UBound(gblnCornerSet)
            'コーナ接続されているか？
            If gblnCornerSet(iCnt) = True Then
                udtMail.dwGouki(iCnt) = byDeshuCnctSet(iCnt)
            End If
        Next
    End If
    
    '-------------------------------------------------------------------------------------------
    '通信設定要求CMD(対象ID)を監マプロセスに送信する
    '-------------------------------------------------------------------------------------------
    bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
    If False = bRet Then
        '「通信接続・切断画面：通信設定要求CMD送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
        Exit Function
    Else
        '「通信接続・切断画面：通信設定要求CMD送信正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
    End If

    pfConnect_TusinConect = True

End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : SendMailHeader
'//  機能名称  : 送信メール作成処理
'//  機能概要  : 送信メール(ヘッダ部)作成を行う
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SendMailHeader(dwId As Long, dwSize As Long)

    Dim bytWork()   As Byte
    Dim i           As Integer
    
    Erase bytWork
    
      udtMail.mlHeader.dwId = dwId
      udtMail.mlHeader.dwSize = dwSize
      udtMail.mlHeader.dwProid = RHOSHU_ID
      udtMail.mlHeader.dwSubArea = 0
      
      bytWork = StrConv(MAIL_SLOT_HOSHU, vbFromUnicode)
      '動的配列の内容をログパラメータ構造体の静的配列に格納する。
      For i = 0 To UBound(bytWork)
        'Null値になったら処理を抜ける。
         If bytWork(i) = vbVEmpty Then Exit For
               
            udtMail.byMailName(i) = bytWork(i)
                
            '動的配列の最大要素になったら処理を抜ける
             If i = UBound(bytWork) Then Exit For
      Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfMailRecieve_KansiVerDisp
'//  機能名称  : 汎用メール受信処理（監視盤バージョン管理）
'//  機能概要  : 保守メールスロットから、メールを受信。
'//              ※プロセス終了指示時は強制終了
'//              →プロセス終了指示を受信した場合に応答を通知する。
'//
'//              型        名称      意味
'//   引数     :なし
'//
'//              型        値        意味
'//  戻り値    :Long　　　　　　　　[OUT]戻り値
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function pfMailRecieve_KansiVerDisp() As Long

    Dim lLen As Long                    'メールサイズ
    Dim uMail As ML_KYOTU_INF           '汎用メールフォーマット
    Dim lngErrCode As Long              'エラーコード
    Dim bRet As Boolean                 '戻り値
    Dim iCnt As Integer                 'カウンタ
   
    On Error Resume Next
    
    pfMailRecieve_KansiVerDisp = 0      '戻り値を正常とする

    '保守メール･スロットからメールを取り出す
    lLen = DssMailRead(plMSlot_MN, uMail)
    If lLen > 0 Then                            '受信?
    
        '------------------------------------------------------------------------------
        'プロセス終了指示の場合
        '------------------------------------------------------------------------------
        If uMail.udtlHeader.dwId = ML_ID_PROEND_ORD Then
           
           '「プロセス終了指示受信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            
            'プロセス終了通知を送信する
            uMail.udtlHeader.dwId = ML_ID_PROEND_INF
            uMail.udtlHeader.dwSize = MlSize.PROEND_INF
            uMail.udtlHeader.dwProid = RHOSHU_ID
            uMail.udtlHeader.dwSubArea = 0
            bRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.PROEND_INF, uMail.udtlHeader)
            If bRet = True Then
               '「プロセス終了通知送信：正常」ログ出力
               Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, PROCESS_END_REQ_SEND, 0)
            Else
               '「プロセス終了通知送信：異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, PROCESS_END_REQ_SEND, lngErrCode)
            End If
            
            '強制終了処理を行う
            pfAbortProc
            Exit Function       '処理を終了する
            
        '------------------------------------------------------------------------------
        '保守画面アクティブ表示の場合
        '------------------------------------------------------------------------------
        ElseIf uMail.udtlHeader.dwId = ML_ID_HOSHU_ACTIVE_REQ Then
           
           '「保守画面アクティブ表示要求受信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
            
            pfMailRecieve_KansiVerDisp = ML_ID_HOSHU_ACTIVE_REQ
        
        '------------------------------------------------------------------------------
        '通信設定要求RESの場合
        '------------------------------------------------------------------------------
        ElseIf uMail.udtlHeader.dwId = ML_ID_CONECT_RES Then
           '「通信設定要求RES受信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, CONECT_CONECTSETTEI_CMD_RECV, 0)
            
            'エラー時通信種別が0（正常）でない場合
            If (miErrorSts <> 0) Then
                '外部機器要求種別が自改？
                If uMail.lngData(0) = ML_DT_JIKAI Then
                    Call ErrorProc
                 
                '外部機器要求種別が自改でない場合
                Else
                    'エラー種別デ集の場合
                    If (miErrorSts = DESHU_CONNECT) Then
                        Call ErrorProc

                    'エラー種別デ集でない場合
                    Else
                        '自改と通信接続
                        pfConnect_TusinConect (ML_DT_JIKAI)
                    End If
                End If
                Exit Function       '処理を終了する
            End If
                    
            '外部機器要求種別が自改？
            If uMail.lngData(0) = ML_DT_JIKAI Then
                If uMail.lngData(1) = ML_CONECT_ERROR Then
                    Call ConnctErrorProc(GATE_CONNECT, ERROR_TUSHIN_DISP) '通信切断異常処理
                    Exit Function                                   '処理を終了する
                End If

                sCmdBtnEnabled False                                ' 画面操作不可

                '自改通信切断待機処理
                bRet = pfCheakGateConectSts
                If bRet = False Then
                    Call ConnctErrorProc(GATE_CONNECT, ERROR_TUSHIN_DISP) '通信切断異常処理
                    Exit Function                                   '処理を終了する
                End If

                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                
                '締切未送データ作成（接続コーナー数分繰り返す）
                For iCnt = 0 To UBound(gblnCornerSet)
                    'コーナ接続されているか？
                    If gblnCornerSet(iCnt) = True Then
                        miCornerNo = iCnt
                        '締切データ出力中画面表示
                        frmShimekiriOutPut2.Show vbModal
                        If (mbMisouResult = False) Then
                            tmrMail.Enabled = True
                            Call ConnctErrorProc(GATE_CONNECT, ERROR_MISOU_DISP) '通信切断異常処理
                            Exit Function
                        End If
                    End If
                Next
                tmrMail.Enabled = True
               
                'ワーク→実行コピー処理
                bRet = fWorktoNow_Start
                If bRet = False Then
                    sCmdBtnEnabled True                             ' 画面操作可
                End If
            
            Else
                If uMail.lngData(1) = ML_CONECT_ERROR Then
                    Call ConnctErrorProc(DESHU_CONNECT, ERROR_TUSHIN_DISP) ' 通信切断異常処理
                    Exit Function                                   ' 処理を終了する
                End If
                
                '上位機器通信切断待機処理
                bRet = pfCheakJyouiKikiConectSts
                If (bRet = False) Then
                    Call ConnctErrorProc(DESHU_CONNECT, ERROR_TUSHIN_DISP) ' 通信切断異常処理
                    Exit Function                                   ' 処理を終了する
                End If
                
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
                '外部機器要求種別がデ集の場合、自改通信切断処理を実施
                bRet = fWorktoNow_Before2
                If bRet = False Then
                    sCmdBtnEnabled True                     ' 画面操作可
                End If
            End If
            
        '------------------------------------------------------------------------------
        '上記以外
        '------------------------------------------------------------------------------
        Else
        
           '「メールID不正」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        
        End If
        
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfCheakJyouiKikiConectSts
'//  機能名称  : 上位機器通信切断待機処理
'//  機能概要  : 上位機器通信状態が全コーナー通信異常になるまで繰り返す
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean   True      全コーナー通信切断
'//　　　　　　　　　　　　False　　 通信切断待ちタイムアウト発生
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfCheakJyouiKikiConectSts() As Boolean

    Dim iAreId(0 To 5) As Integer       '上位機器通信状態エリアID
    Dim bCheakAns As Boolean            'チェック結果
    Dim iConectSts As Integer           '通信状態
    Dim iCnt As Integer                 'カウンタ
    Dim LngSleepTotal As Long           'スリープカウント合計値
    
    On Error Resume Next
    
    pfCheakJyouiKikiConectSts = False
    bCheakAns = True
    LngSleepTotal = 0
    
    iAreId(0) = 1                       '上位機器通信状態エリアID：データ集計機(コーナ１)
    iAreId(1) = 9                       '上位機器通信状態エリアID：データ集計機(コーナ２)
    iAreId(2) = 10                      '上位機器通信状態エリアID：データ集計機(コーナ３)
    iAreId(3) = 11                      '上位機器通信状態エリアID：データ集計機(コーナ４)
    iAreId(4) = 12                      '上位機器通信状態エリアID：データ集計機(コーナ５)
    iAreId(5) = 13                      '上位機器通信状態エリアID：データ集計機(コーナ６)
    
    '------------------------------------------------------------------------------------
    '上位機器通信状態チェック
    '------------------------------------------------------------------------------------
    '全コーナー通信切断されるまで数分繰り返す
    '※全コーナーの通信が切断されるまで、３分以上経過した場合異常終了とする。
    Do While LngSleepTotal < WAIT_TIME_OUT
    
        '接続コーナー数分繰り返す
        For iCnt = 0 To UBound(gblnCornerSet)
            
            bCheakAns = True
            
            'コーナ接続されているか？
            If gblnCornerSet(iCnt) = True Then
                
                '上位機器通信状態取得
                iConectSts = pfGetJyouiKikiConectSts(iAreId(iCnt))
                '通信状態が「0:立上中」or「1:通信正常」か？
                If iConectSts = 1 Then
                    bCheakAns = False
                    Exit For
                End If
                
            End If
        
        Next
        
        '全コーナー通信切断されたか？
        If bCheakAns = True Then
            Exit Do
        End If
        
        Sleep (100)
        LngSleepTotal = LngSleepTotal + 100
    
    Loop
    
    '３分以上経過したか？
    If LngSleepTotal >= WAIT_TIME_OUT Then
        Exit Function
    End If
    
    pfCheakJyouiKikiConectSts = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfCheakGateConectSts
'//  機能名称  : 自改通信切断待機処理
'//  機能概要  : 自改通信状態が全コーナー通信異常になるまで繰り返す
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean   True      全号機通信切断
'//　　　　　　　　　　　　False　　 通信切断待ちタイムアウト発生
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfCheakGateConectSts() As Boolean

    Dim bCheakAns As Boolean            'チェック結果
    Dim iConectSts As Integer           '通信状態
    Dim iCnt As Integer                 'カウンタ
    Dim LngSleepTotal As Long           'スリープカウント合計値
    
    On Error Resume Next
    
    pfCheakGateConectSts = False
    bCheakAns = True
    LngSleepTotal = 0
    
    '------------------------------------------------------------------------------------
    '自改通信状態チェック
    '------------------------------------------------------------------------------------
    '全号機通信切断されるまで数分繰り返す
    '※自改の通信が全号機切断されるまで、３分以上経過した場合異常終了とする。
    Do While LngSleepTotal < WAIT_TIME_OUT
    
        '号機数分繰り返す
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
            
            bCheakAns = True
            
            '改札機が実装されているか？
            If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
                
                '自改機器通信状態取得
                iConectSts = pfGetGateConectSts(iCnt + 1)
                '通信状態が「0:立上中」or「1:通信正常」か？
                If iConectSts = 0 Or iConectSts = 1 Then
                    bCheakAns = False
                    Exit For
                End If
                
            End If
        
        Next
        
        '全号機通信切断されたか？
        If bCheakAns = True Then
            Exit Do
        End If
        
        Sleep (100)
        LngSleepTotal = LngSleepTotal + 100
    
    Loop
    
    '３分以上経過したか？
    If LngSleepTotal >= WAIT_TIME_OUT Then
        Exit Function
    End If
    
    pfCheakGateConectSts = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfGetJyouiKikiConectSts
'//  機能名称  : 上位機器通信状態取得処理
'//  機能概要  : 上位機器の通信状態を取得する
'//
'//              型        名称      意味
'//  引数      : Integer　iAreId  　[IN]上位機器通信状態エリアID
'//
'//              型        値        意味
'//  戻り値    : Integer　　　　　　　上位機器通信状態
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetJyouiKikiConectSts(iAreId As Integer) As Integer
    
    On Error Resume Next
    
    pfGetJyouiKikiConectSts = -1
    
    'ＩＤ別情報操作クラスの生成
    Set Idinf_Jyoui = New IdInfProc
   '参照(上位機器通信状態)エリア名を設定
    Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui
    Idinf_Jyoui.IdOpen
    If Idinf_Jyoui.Errsts <> 0 Then
      Set Idinf_Jyoui = Nothing
      Exit Function
    End If
    
    '参照(上位機器通信状態)エリアをＬＯＣＫする。
    Idinf_Jyoui.IdLock
    If Idinf_Jyoui.Errsts <> 0 Then
      Idinf_Jyoui.IdFree
      Set Idinf_Jyoui = Nothing
      Exit Function
    End If
    
     'エリアの内容を読み込む。
    Idinf_Jyoui.id = iAreId
    Idinf_Jyoui.GetInf (CONECT)
    If Idinf_Jyoui.Errsts <> 0 Then
       Idinf_Jyoui.IdFree
       Set Idinf_Jyoui = Nothing
       Exit Function
    End If
    
    '上位機器通信状態を取得
    pfGetJyouiKikiConectSts = CInt(Idinf_Jyoui.DataArea(0))
    
    Idinf_Jyoui.IdFree
    Set Idinf_Jyoui = Nothing
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfGetGateConectSts
'//  機能名称  : 自改通信状態取得処理
'//  機能概要  : 自改の通信状態を取得する
'//
'//              型        名称      意味
'//  引数      : Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : Integer　　　　　　　自改通信状態
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetGateConectSts(iGouki As Integer) As Integer
    
    On Error Resume Next
    
    pfGetGateConectSts = -1
    
    Set Idinf_JikaiTuushin = New IdInfProc
    '参照(自改通信状態)エリア名を設定
    Idinf_JikaiTuushin.ProcMode = DATA_ID.Data_Id_JikaiTuushinJyotai
    Idinf_JikaiTuushin.IdOpen
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       Set Idinf_JikaiTuushin = Nothing
       Exit Function
    End If
     
    '参照(自改通信状態)エリアをＬＯＣＫする。
    Idinf_JikaiTuushin.IdLock
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       Idinf_JikaiTuushin.IdFree
       Set Idinf_JikaiTuushin = Nothing
       Exit Function
    End If
    
    'エリアの内容を読み込む。
    Idinf_JikaiTuushin.id = IdGateComSts.GATE_COM
    Idinf_JikaiTuushin.GetJikai_Tuusin iGouki - 1
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       Idinf_JikaiTuushin.IdFree
       Set Idinf_JikaiTuushin = Nothing
       Exit Function
    End If
        
    pfGetGateConectSts = CInt(Idinf_JikaiTuushin.DataArea(iGouki - 1))
    
    Idinf_JikaiTuushin.IdFree
    Set Idinf_JikaiTuushin = Nothing
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfGetJyouiKikiConectSet
'//  機能名称  : 上位機器通信接続／切断設定取得処理
'//  機能概要  : 上位機器の通信接続／切断設定を取得する
'//
'//              型        名称      意味
'//  引数      : Integer　iKansiId  　[IN]エリアID
'//
'//              型        値        意味
'//  戻り値    : Integer　　1　　　　　切離設定
'//　　　　　　　　　　　　 0          接続設定
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetJyouiKikiConectSet(iKansiId As Integer) As Integer
    
    On Error Resume Next
    
    pfGetJyouiKikiConectSet = -1
    
    'ＩＤ別情報操作クラスの生成
    Set Idinf_KansiSettei = New IdInfProc
    '共有エリアオープン
    Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
    Idinf_KansiSettei.IdOpen
    If Idinf_KansiSettei.Errsts <> 0 Then
        Set Idinf_KansiSettei = Nothing
        Exit Function
    End If
       
    '監視設定エリアをＬＯＣＫする。
    Idinf_KansiSettei.IdLock
    If Idinf_KansiSettei.Errsts <> 0 Then
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing
        Exit Function
    End If
    
    '監視設定エリアIDを設定
    Idinf_KansiSettei.id = iKansiId
    Idinf_KansiSettei.IdGet
    If Idinf_KansiSettei.Errsts <> 0 Then
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing
        Exit Function
    End If
    
    pfGetJyouiKikiConectSet = Idinf_KansiSettei.DataArea(0)   '設定内容
    
    Idinf_KansiSettei.IdFree
    Set Idinf_KansiSettei = Nothing

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfGetGateConectSet
'//  機能名称  : 自改通信接続／切断設定取得処理
'//  機能概要  : 自改の通信接続／切断設定を取得する
'//
'//              型        名称      意味
'//  引数      : Integer　iGouki  　　[IN]号機番号
'//
'//              型        値        意味
'//  戻り値    : Integer　　1　　　　　切離設定
'//　　　　　　　　　　　　 0          接続設定
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetGateConectSet(iGouki As Integer) As Integer
    
    On Error Resume Next
    
    pfGetGateConectSet = -1
    
    Set Idinf_JikaiSettei = New IdInfProc
    '自改設定エリアをオープンする。
    Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
    Idinf_JikaiSettei.IdOpen
    If Idinf_JikaiSettei.Errsts <> 0 Then
        Set Idinf_JikaiSettei = Nothing
        Exit Function
    End If
    
    '自改設定エリアをＬＯＣＫする。
    Idinf_JikaiSettei.IdLock
    If Idinf_JikaiSettei.Errsts <> 0 Then
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing
        Exit Function
    End If
    
    'エリアの内容を読み込む。
    Idinf_JikaiSettei.id = IdGate.JIKAI_CONECT_SETTEI
    Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
    If Idinf_JikaiSettei.Errsts <> 0 Then
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing
        Exit Function
    End If
    
    '設定内容を取得
    pfGetGateConectSet = Idinf_JikaiSettei.DataArea(iGouki - 1)
    
    '状態：正常
    Idinf_JikaiSettei.IdFree
    Set Idinf_JikaiSettei = Nothing
    
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : ConnctErrorProc
'//  機能名称  : 通信切断異常処理
'//  機能概要  : デ集または自改の通信切断で異常発生時の処理を行う
'//
'//              型        名称      意味
'//  引数      : Integer　iGouki  　　[IN]号機番号
'//
'//              型        値        意味
'//  戻り値    : Long　　  1　　　　　切離設定
'//　　　　　　　Long　　　0          接続設定
'//　　　　　　　Long　　　0          接続設定
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub ConnctErrorProc(iTusinSts As Integer, iErrorDisp As Integer)

    On Error Resume Next

    '異常時通信種別設定
    miErrorSts = iTusinSts

    '異常時表文言設定
    miErrorDisp = iErrorDisp
    
    'デ集接続
    pfConnect_TusinConect (ML_DT_DESHU)

End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : ErrorProc
'//  機能名称  : 異常処理
'//  機能概要  :
'//
'//              型        名称      意味
'//  引数      : Integer　iGouki  　　[IN]号機番号
'//
'//              型        値        意味
'//  戻り値    : Long　　  1　　　　　切離設定
'//　　　　　　　Long　　　0          接続設定
'//　　　　　　　Long　　　0          接続設定
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013年度施策 遠隔対応【EG20_KANSI03_01】
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub ErrorProc()

    On Error Resume Next

                                        
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                                                                                
    '異常文言表示
    If (miErrorDisp = ERROR_TUSHIN_DISP) Then
        '通信切断失敗メッセージ表示
        MsgBox "通信切断処理中に異常が発生しました。ワーク→実行コピーが行えません。", _
                        vbOKOnly + vbCritical, _
                        "ワーク→実行 コピー"
    ElseIf (miErrorDisp = ERROR_MISOU_DISP) Then
        '未送データ作成失敗メッセージ表示
        MsgBox "未送データ作成処理中に異常が発生しました。ワーク→実行コピーが行えません。", _
                        vbOKOnly + vbCritical, _
                        "ワーク→実行 コピー"
    Else
        '異常終了メッセージ表示
        MsgBox "異常終了しました。", vbCritical, "ワーク→実行 コピー"
    End If
    
    ' 画面操作可
    sCmdBtnEnabled True

End Sub

