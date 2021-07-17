VERSION 5.00
Begin VB.Form frmSousaTakuVerKanri 
   BorderStyle     =   0  'なし
   Caption         =   "バージョン管理（操作卓）"
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
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   18
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " 媒体 → ワーク　コピー"
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
      Top             =   3360
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
      TabIndex        =   16
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   旧 → 実行   コピー"
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
      TabIndex        =   15
      Top             =   4800
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
      TabIndex        =   13
      Top             =   6960
      Width           =   2415
   End
   Begin VB.ListBox lstTaku 
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
      TabIndex        =   6
      Top             =   2520
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
      Top             =   1920
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
      Height          =   1335
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
         Top             =   240
         Value           =   1  'ﾁｪｯｸ
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
         Top             =   600
         Value           =   1  'ﾁｪｯｸ
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
         Top             =   960
         Value           =   1  'ﾁｪｯｸ
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " バージョン管理  画面へ戻る"
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
   Begin VB.Label lblZenVer 
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
      TabIndex        =   19
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "操作卓バージョン管理"
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
      Top             =   2160
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
      Top             =   2160
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
      Top             =   2160
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
      Top             =   2160
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
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmSousaTakuVerKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmSousaTakuVerKanri.frm
'//  パッケージ名：バージョン管理(監視盤)画面
'//
'//  概要：バージョン管理(監視盤)画面
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.100関連】
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.273修正対応】
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                 【指摘事項No.02修正対応】
'//                 【残件:保守運改の切替結果通知対応】
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ2,3統合対応
'//     REVISIONS :(EG20 V5.9.0.1) 2012-05-02   REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'フォルダ種別部
Public mlngChkFolderType        As Long

'Dim uVersion() As MN_VERSION_LIST       'バージョン情報格納エリア

Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

Private Const HEADERTITLE_WRK = "操作卓バージョン（ワーク）："
Private Const HEADERTITLE_NOW = "　　　　　　　　（実行）　："
Private Const HEADERTITLE_OLD = "　　　　　　　　（旧）　　："
Private Const HEADERVERSION_NON = "--.--.--.--"


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
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-11-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                 【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ： (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 【運改表示改善対応】
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
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
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '[はい] ボタンを選択した場合
        'ワークフォルダ内のファイルを削除する
        bResult = sWrkFolderRemove
        sCmdBtnEnabled True                         ' 画面操作可
        If bResult = True Then
            ' 操作卓のバージョン情報を表示する｡
            Call psVersionDisp
        
' EG20 V5.8.0.1削除開始
'            ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'            Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
            ' 運改状態更新
'            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_NASHI)    ' EG20 V5.11.0.1削除
            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_CLEAR)     ' EG20 V5.11.0.1追加
            Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_NASHI)
' EG20 V5.8.0.1追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
            'プログレスバーを消去する
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            '「正常終了」ポップアップ画面表示
            MsgBox "正常終了しました。", _
                   vbOKOnly + vbInformation, _
                   "実行結果"
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

        End If
    End If

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

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
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()
    Dim iResponse As Integer         'MsgBoxボタンコード

    On Error Resume Next

    '「媒体→ワークコピー」ボタンの場合。
    '「バージョン管理画面：媒体→ワークコピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_USB_COPY_WRK_BUTTOM, 0)

    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("インストール媒体をワークフォルダに" _
           & Chr(vbKeyReturn) & "コピーします。よろしいですか？", _
           vbOKCancel + vbExclamation, _
           "媒体→ワークコピー")
    If iResponse <> vbCancel Then
        '[はい] ボタンを選択した場合
        sCmdBtnEnabled False                        ' 画面操作不可
        'インストール媒体をワークフォルダ内にコピーする
        Call sFDInstall
        sCmdBtnEnabled True                         ' 画面操作可

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除開始
'        ' 操作卓のバージョン情報を表示する｡
'        Call psVersionDisp
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除終了
    End If

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

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
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.372修正対応】
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY[TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    Dim iResponse As Integer         'MsgBoxボタンコード
    Dim bRet As Boolean              ' 処理結果

    On Error Resume Next

    '「バージョン管理画面：旧→実行コピー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OLD_COPY_NOW_BUTTOM, 0)

    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("実行フォルダをクリアし旧フォルダの" _
           & Chr(vbKeyReturn) & "ファイルをコピーしますがよろしいですか？", _
           vbOKCancel + vbExclamation, _
           "旧→実行コピー")
    If iResponse <> vbCancel Then
        
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加開始
        ' 旧バージョンフォルダに代表バージョンファイルが存在しない場合は異常とする。
        ' 旧バージョン・OPERATE・操作卓・
        bRet = dllCheckAplVersion(4, PATH_OPERATE_APL, 3)
        If bRet = False Then
'            MsgBox "異常終了しました。", vbCritical, "旧→実行　コピー"        ' EG20 V5.8.0.1削除
            MsgBox "異常終了しました。", vbCritical, "実行結果"                 ' EG20 V5.8.0.1追加
' EG20 V5.6.0.1追加開始
            pubSubCreateFolder (PATH_OPERATE_APL)
            pubSubCreateFolder (PATH_OPERATE_APLNEW)
            pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V6.9.0.1ADD START
            pubSubCreateFolder (FLD_OPERATEPROGNOW)
            pubSubCreateFolder (FLD_OPERATEPROGWRK)
            pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END
            Exit Sub
        End If
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加終了
        
        '[はい] ボタンを選択した場合
        sCmdBtnEnabled False                        ' 画面操作不可
        'インストール媒体をワークフォルダ内にコピーする
        Call sVersionRollBack
        sCmdBtnEnabled True                         ' 画面操作可
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除開始
'        ' 操作卓のバージョン情報を表示する｡
'        Call psVersionDisp
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除終了
    End If

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

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
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.372修正対応】
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY[TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
    Dim iResponse As Integer         'MsgBoxボタンコード
    Dim bRet As Boolean              ' 処理結果

    On Error Resume Next

    '「バージョン管理画面：旧→実行コピー釦押下」ログ出力
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OLD_COPY_NOW_BUTTOM, 0)
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_COPY_NOW_BUTTOM, 0)

    '確認ポップアップウィンドウを表示する。
    '確認ポップアップウィンドウを表示する。
    iResponse = MsgBox("実行フォルダをクリアしワークフォルダの" _
            & Chr(vbKeyReturn) & "ファイルをコピーしますがよろしいですか？", _
           vbOKCancel + vbExclamation, _
           "ワーク→実行コピー")
    If iResponse <> vbCancel Then
        
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加開始
        ' 旧バージョンフォルダに代表バージョンファイルが存在しない場合は異常とする。
        ' 旧バージョン・OPERATE・操作卓・
        bRet = dllCheckAplVersion(1, PATH_OPERATE_APL, 3)
        If bRet = False Then
'            MsgBox "異常終了しました。", vbCritical, "ワーク→実行 コピー"     ' EG20 V5.8.0.1削除
            MsgBox "異常終了しました。", vbCritical, "実行結果"                 ' EG20 V5.8.0.1追加
' EG20 V5.6.0.1追加開始
            pubSubCreateFolder (PATH_OPERATE_APL)
            pubSubCreateFolder (PATH_OPERATE_APLNEW)
            pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V6.9.0.1ADD START
            pubSubCreateFolder (FLD_OPERATEPROGNOW)
            pubSubCreateFolder (FLD_OPERATEPROGWRK)
            pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END
            Exit Sub
        End If
'EG20 V3.6.0.1【03統合TR-No.372修正対応】追加終了
        
        '[はい] ボタンを選択した場合
        sCmdBtnEnabled False                        ' 画面操作不可
        'インストール媒体をワークフォルダ内にコピーする
        Call sVersionUpdate
        sCmdBtnEnabled True                         ' 画面操作可

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除開始
'        ' 操作卓のバージョン情報を表示する｡
'        Call psVersionDisp
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除終了
    End If

' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

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

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：Form_Load
'//  機能名称    ：バージョン管理(監視盤)画面(ロード時)
'//  機能概要    ：初期処理を行う。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
   
   '「操作卓バージョン管理画面：表示」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_TAKU_GAMEN_START, 0)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    '初期化
    lstTaku.Clear
    mlngChkFolderType = 0

    'フォルダ選択部：選択有り
    chkFolder(0).Value = 1
    chkFolder(1).Value = 1
    chkFolder(2).Value = 1
    
    mlngChkFolderType = 7
    
    ' 操作卓のバージョン情報を表示する｡
    Call psVersionDisp
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

'   メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：chkFolder_Click
'//  機能名称    ：「フォルダチェック」チェック押下処理
'//  機能概要    ：フォルダ選択部チェックを行う。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
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

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdRefresh_Click
'//  機能名称    ：「表示更新」釦押下処理
'//  機能概要    ：最新の状態を表示する。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdRefresh_Click()
    Dim i As Integer        'カウンター
    Dim bFlag As Boolean    '表示フォルダ選択チェック(TRUE：チェック有。FALSE：チェック無)
   
    On Error Resume Next
    
    '「操作卓バージョン管理画面：表示更新釦押下」
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
                "操作卓バージョン管理"
        Exit Sub
    End If
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   
    ' 操作卓のバージョン情報を表示する｡
    Call psVersionDisp

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
End Sub


'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdOutPut_Click
'//  機能名称    ：「バージョン情報媒体出力」釦押下処理
'//  機能概要    ：表示されたバージョン情報を媒体に出力する。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.100関連】
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()
'*******************************
'VBエラー処理
On Error GoTo Error_cmdOutPut_Click
'*******************************
    Dim strCopySaki    As String        ' 出力先ファイルパス
    Dim strWriteDir    As String        ' 出力先フォルダ
    Dim fso            As New FileSystemObject   'ファイルシステムオブジェクト
    Dim lngErrCode     As Long          'エラーコード
    
    Dim strStationName As String        ' 駅名名
    Dim szCornerName   As String        ' コーナ名称
    Dim nNullIndex     As Integer       ' 文字数ワーク
    Dim strWork        As String        ' ワーク
    Dim strFileName    As String        ' ファイル名
    Dim bRet           As Boolean  '戻り値

   '「監視盤バージョン管理画面：バージョン情報媒体出力釦押下」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OUTPUT, 0)
    
' EG20 V3.3.0.1 【結合TR-No.100関連】追加開始
    ' リストに１件もデータがない場合は異常終了
    If lstTaku.ListCount = 0 Then
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Exit Sub
    End If
' EG20 V3.3.0.1 【結合TR-No.100関連】追加終了
    
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
    
    strStationName = gsGetStationEkiName
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ 保守専用関数:操作卓バージョンファイル（画面表示用）作成
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateTakuVersionFile(mlngChkFolderType, TAKUVERLIST_REPORTFILE, VERLISTKIND_REPORT)
    
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
    If fso.FileExists(TAKUVERLIST_REPORTFILE) = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        'ファイル無し異常ポップアップ画面表示
        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
        Exit Sub
    End If
    strFileName = Dir(TAKUVERLIST_REPORTFILE)
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除開始
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
    fso.CopyFile TAKUVERLIST_REPORTFILE, strCopySaki, True
  
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '出力結果ポップアップ(正常)表示
    MsgBox "正常終了しました。", vbInformation + vbOKOnly, "出力結果"
    '「操作卓バージョン管理画面：バージョン情報媒体出力処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_INFO_OUTPUT_OK, 0)
    
    Set fso = Nothing
    
    Exit Sub
'*******************************
'VBエラー処理
Error_cmdOutPut_Click:
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    '処理異常の場合、出力結果ポップアップ(異常)表示
    MsgBox "異常終了しました。", vbCritical, "出力結果"
    '「操作卓バージョン管理画面：バージョン情報媒体出力処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OUTPUT_ERROR, lngErrCode)
    Set fso = Nothing
'*******************************
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdRemove_Click
'//  機能名称    ：「媒体取外」釦押下処理
'//  機能概要    ：媒体の取り外しを行う。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
   
   On Error Resume Next
       
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：cmdReturn_Click
'//  機能名称    ：「メニュー画面へ戻る」釦押下処理
'//  機能概要    ：自画面を消去する。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '「操作卓バージョン管理画面：消去」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_TAKU_GAMEN_END, 0)
 
 ' EG20 V5.6.0.1追加開始
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1追加終了
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

    'バージョン管理（操作卓）画面を閉じる
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
    lstTaku.Clear
    
    '作業エリア初期化
    strWork = ""

    '全体バージョン初期化
    strVerData = ""

    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ 保守専用関数:操作卓バージョンファイル（画面表示用）作成
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateTakuVersionFile(mlngChkFolderType, TAKUVERLIST_DISPFILE, VERLISTKIND_DISP)

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
    If Len(Trim(Dir(TAKUVERLIST_DISPFILE))) = 0 Then
        Exit Sub
    End If

    ' バージョンファイルのファイル番号を取得する。
    intFileNo = FreeFile

    ' バージョンファイルオープン
    Open TAKUVERLIST_DISPFILE For Input As #intFileNo
    
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
        lblZenVer.Caption = strVerData

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
                lstTaku.AddItem (strList)

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

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：tmrMail_Timer
'//  機能名称    ：メール受信タイマ、タイムアップ処理
'//  機能概要    ：メールを受信する。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    '汎用メール受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSousaTakuVerKanri.Caption, False
        pfFormActive (frmSousaTakuVerKanri.hwnd)            ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：sWrkFolderRemove
'//  機能名称    ：ワークフォルダ内ファイル削除処理
'//  機能概要    ：ワークフォルダ内のファイルを削除する。
'//
'//                   型          名称            意味
'//  引数        ：なし
'//  戻り値      ：なし
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'/////////////////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim stringWorkFolder As String      ' フォルダ名
    Dim MyName As String                'ファイル名
    Dim lngErrCode As Long              'エラーコード
    
    Dim objFso As New FileSystemObject     'ファイルシステムオブジェクト
    Dim objFi As File                     'ファイルオブジェクト
    Dim objFolder As Folder               'フォルダオブジェクト         ' EG20 V6.9.0.1 ADD
    
    On Error GoTo ErrorHandler          'エラーハンドルの登録

    '初期値設定
    sWrkFolderRemove = True
   
    '//////////////////////////////////////////////////////////////////////////
    '// 監視盤フォルダ内の操作卓ワークフォルダを消去
    ' ワークフォルダ内のディレクトリの名前を表示します。
    stringWorkFolder = FLD_OPERATEPROGWRK & "\"
    For Each objFi In objFso.GetFolder(stringWorkFolder).files  ' ループを開始
        If objFso.FileExists(objFi.Path) = True Then            ' ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            'ファイルを削除する
            Kill stringWorkFolder & MyName
        End If
    Next

' EG20 V6.9.0.1 ADD START
    For Each objFolder In objFso.GetFolder(stringWorkFolder).SubFolders  ' ループを開始
        If objFso.FolderExists(objFolder.Path) = True Then               ' ファイル名の取得チェック
            'ディレクトリを削除
            Call objFso.DeleteFolder(objFolder.Path)
        End If
    Next
' EG20 V6.9.0.1 ADD END

    '//////////////////////////////////////////////////////////////////////////
    '// ワークフォルダ内の操作卓フォルダを消去
    ' ワークフォルダ内のディレクトリの名前を表示します。
    stringWorkFolder = PATH_OPERATE_APLNEW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing         ' EG20 V6.9.0.1 ADD

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除開始
'    '「正常終了」ポップアップ画面表示
'    MsgBox "正常終了しました。", _
'           vbOKOnly + vbInformation, _
'           "実行結果"
'
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】削除終了
    
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
    Set objFi = Nothing
    Set objFolder = Nothing         ' EG20 V6.9.0.1 ADD
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
'//                 【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ：(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 フェーズ2,3統合対応
'//  REVISIONS   ：(EG20 V5.9.0.1) 2012-05-02  REVISED BY [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：改札機バージョン管理画面のsFDInstall流用
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall()

    Dim sInputPass As String                ' インストール元ディレクトリ名
    Dim sInputFolder As String              ' インストール元フォルダ名
    Dim objFso As New FileSystemObject      ' ファイルシステムオブジェクト
    Dim objFi As File                       ' ファイルオブジェクト
    Dim MyName As String                    ' ファイルフルパス名
    Dim sSrcFileName As String              ' コピー元ファイル名
    Dim sDstFileName As String              ' コピー先ファイル名
    Dim lngErrCode     As Long              ' エラーコード
    Dim lngProcId As Long                   ' プロセスID
    Dim hProc As Variant                    ' プロセスハンドル
    Dim objFolder As Folder                 ' フォルダオブジェクト          ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim FileName As String                  ' 抽出ファイル名                ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim FileKaku As String                  ' 拡張子                        ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim ExecCommand As String               ' 実行文字列                    ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim CurrentDirectory As String          ' カレントディレクトリ          ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加
    Dim ExecDirectory As String             ' 実行ファイルディレクトリ      ' EG20 V5.9.0.1追加

    On Error GoTo ErrorHandler              ' エラーハンドルの登録

    ' /////////////////////////////////////////////////////////////////////////
    ' // インストール部材のコピー
    sInputPass = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    '指定フォルダなし
    If Len(sInputPass) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
        Exit Sub
    End If
 
    sInputFolder = sInputPass

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

' EG20 V3.6.0.1【統合TR-No.273修正対応】削除開始
'    For Each objFi In objFso.GetFolder(sInputFolder).files      'ループを開始
'        If objFso.FileExists(objFi.Path) = True Then            'ファイル名の取得チェック
'            'ディレクトリ名を取得
'            MyName = objFi.Name
'            '媒体内ファイル名を作成
'            sSrcFileName = sInputFolder & "\" & MyName
'            ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
'            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
'                'ワークフォルダ内ファイル名を作成する
'                sDstFileName = FLD_OPERATEPROGWRK & "\" & MyName
'                'ファイルコピー（既に存在した場合は上書きするする）
'                objFso.CopyFile sSrcFileName, sDstFileName, True
'            End If
'        End If
'    Next
' EG20 V3.6.0.1【統合TR-No.273修正対応】削除終了
' EG20 V3.6.0.1【統合TR-No.273修正対応】追加開始
    'フォルダオブジェクトを取得
    Set objFolder = objFso.GetFolder(sInputFolder)

    '//////////////////////////////////////////////////////
    '// ワークフォルダを消去
    If objFso.FolderExists(FLD_OPERATEPROGWRK) Then
        Call objFso.DeleteFolder(FLD_OPERATEPROGWRK)
    End If

    objFolder.Copy FLD_OPERATEPROGWRK
' EG20 V3.6.0.1【統合TR-No.273修正対応】追加終了

    
    ' /////////////////////////////////////////////////////////////////////////
    ' // インストール部材の中から指定ファイルの実行
    sSrcFileName = pfGetFileNameTakuExec
    ' ファイル設定なし
    If Len(sSrcFileName) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        Call psVersionDisp
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        MsgBox "異常終了しました。", vbCritical, "出力結果"
        Exit Sub
    End If
   
    sDstFileName = FLD_OPERATEPROGWRK & "\" & sSrcFileName
    If objFso.FileExists(sDstFileName) = False Then
        Set objFso = Nothing
        Set objFi = Nothing
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        Call psVersionDisp
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        MsgBox "異常終了しました。", vbCritical, "出力結果"
        Exit Sub
    End If
    
    Set objFso = Nothing
    Set objFi = Nothing
    
'    lngProcId = Shell(sDstFileName, vbNormalFocus)             ' EG20 V3.6.0.1【統合TR-No.273修正対応】削除
' EG20 V3.6.0.1【統合TR-No.273修正対応】追加開始
    ' カレントディレクトリ取得
    CurrentDirectory = CurDir$()
'    Call ChDir(FLD_OPERATEPROGWRK)                             ' EG20 V5.9.0.1削除
' EG20 V5.9.0.1追加開始
    Call psFolderPathGet(sDstFileName, ExecDirectory)
    Call ChDrive("D")
    Call ChDir(ExecDirectory)
' EG20 V5.9.0.1追加終了

    ' ファイル名前取得
    psFileNameGet sDstFileName, FileName, FileKaku
    If UCase(FileKaku) = "VBS" Then
        ExecCommand = "wscript.exe " & sDstFileName
    Else
        ExecCommand = sDstFileName
    End If
    lngProcId = Shell(ExecCommand, vbNormalFocus)
' EG20 V3.6.0.1【統合TR-No.273修正対応】追加終了

    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, lngProcId)   ' プロセスハンドルを取得します。
    If hProc > 0 Then                                           ' プロセスハンドルを取得できた場合
        Call dllWaitForSingleObject(hProc)                      ' プロセスがシグナル状態になるまで待ちます。
        CloseHandle hProc                                       ' プロセスハンドルを解放します。
    End If
    
'    Call ChDir(CurrentDirectory)                ' EG20 V3.6.0.1【統合TR-No.273修正対応】追加   ' EG20 V5.9.0.1削除
    Call ChDir("D:\")                                           ' EG20 V5.9.0.1追加
    
' EG20 V5.8.0.1削除開始
'    ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
    ' 運改状態更新
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_ARI)
    Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1追加終了
' EG20 V5.8.0.1 ADD START
    '読み取り外しの関数を実行
    dllChangeAttributeContents (PATH_OPERATE_APLNEW)
    '読み取り外しの関数を実行
    dllChangeAttributeContents (FLD_OPERATEPROGWRK)
' EG20 V5.8.0.1 ADD END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    Call psVersionDisp
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    '「正常終了」ポップアップ画面表示
    MsgBox "正常終了しました。", _
           vbOKOnly + vbInformation, _
           "実行結果"
    
    Exit Sub    '処理を終了する

ErrorHandler:   ' エラー処理。
    Set objFso = Nothing
    Set objFi = Nothing

' EG20 V5.8.0.1 ADD START
    '読み取り外しの関数を実行
    dllChangeAttributeContents (PATH_OPERATE_APLNEW)
    '読み取り外しの関数を実行
    dllChangeAttributeContents (FLD_OPERATEPROGWRK)
' EG20 V5.8.0.1 ADD END

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    Call psVersionDisp
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    MsgBox "異常終了しました。", vbCritical, "出力結果"
    '「自改ﾊﾞｰｼﾞｮﾝ：媒体→ﾜｰｸｺﾋﾟｰ処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_USB_COPY_WRK_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sVersionRollBack
'//  機能名称  : バージョン戻し処理
'//  機能概要  : 現行バージョンを旧バージョンへ戻す
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : String    ファイル名
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.273修正対応】
'//  REVISIONS   ：(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                【指摘事項No.02修正対応】
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Sub sVersionRollBack()

    Dim objFso As New FileSystemObject      ' ファイルシステムオブジェクト
    Dim objFi As File                       ' ファイルオブジェクト
    Dim objFolder As Folder                 ' フォルダオブジェクト
    Dim stringWorkFolder As String          ' フォルダ名
    Dim MyName As String                    ' ファイル名
    Dim lngErrCode     As Long              ' エラーコード
    Dim strSrcFile As String                ' コピー元
    Dim strDstFile As String                ' コピー先
    Dim bResult As Boolean                  ' 処理結果      ' EG20 V3.6.0.1追加
    Dim sSrcFileName As String              ' コピー元ファイル名    ' EG20 V5.8.0.1追加
    Dim sDstFileName As String              ' コピー先ファイル名    ' EG20 V5.8.0.1追加

    On Error GoTo ErrorHandler          'エラーハンドルの登録

' EG20 V5.8.0.1追加開始
    ' /////////////////////////////////////////////////////////////////////////
    ' // ワークフォルダのファイル存在チェック
    stringWorkFolder = FLD_OPERATEPROGOLD
    If objFso.FolderExists(stringWorkFolder) <> True Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' フォルダがそもそも存在しない
        MsgBox "異常終了しました。", vbCritical, "実行結果"
        Exit Sub                        ' 処理終了
    End If
    
    strSrcFile = pfGetFileNameTakuExec
    ' ファイル設定なし
    If Len(strSrcFile) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
            MsgBox "異常終了しました。", vbCritical, "実行結果"
        Exit Sub
    End If

    strDstFile = stringWorkFolder & "\" & strSrcFile
    If objFso.FileExists(strDstFile) = False Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' ファイルが存在しない
        MsgBox "異常終了しました。", vbCritical, "実行結果"
        Exit Sub                        ' 処理終了
    End If
    
' EG20 V5.8.0.1追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

    ' /////////////////////////////////////////////////////////////////////////
    ' // インストール部材のコピー
    
    '//////////////////////////////////////////////////////
    '// 監視盤フォルダ内の操作卓実行フォルダを消去
    stringWorkFolder = FLD_OPERATEPROGNOW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If

    '//////////////////////////////////////////////////////
    '// 旧→実行コピー
    strSrcFile = FLD_OPERATEPROGOLD
    strDstFile = FLD_OPERATEPROGNOW

    'フォルダオブジェクトを取得
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If

    ' /////////////////////////////////////////////////////////////////////////
    ' // 全体をコピー
    
    '//////////////////////////////////////////////////////
    '// 操作卓フォルダを消去
    stringWorkFolder = PATH_OPERATE_APL
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    strSrcFile = PATH_OPERATE_APLOLD
    strDstFile = PATH_OPERATE_APL

    'フォルダオブジェクトを取得
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If

    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing

' EG20 V3.6.0.1追加開始
    ' 操作卓プログラムデータ作成処理
    bResult = pfTakuProgramVersionCreateProc
' EG20 V3.6.0.1追加終了

    If bResult = True Then
' EG20 V5.8.0.1追加開始
        ' 運改状態更新
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_KIRIKAE)
        Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        ' 操作卓のバージョン情報を表示する｡
        Call psVersionDisp
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '「正常終了」ポップアップ画面表示
        MsgBox "正常終了しました。", _
               vbOKOnly + vbInformation, _
               "実行結果"
    Else
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        ' 操作卓のバージョン情報を表示する｡
        Call psVersionDisp
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        '「異常終了」ポップアップ画面表示
        MsgBox "異常終了しました。", _
               vbOKOnly + vbCritical, _
               "実行結果"
    End If
    
    Exit Sub '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    ' 操作卓のバージョン情報を表示する｡
    Call psVersionDisp
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    '「ワーククリア異常終了」ポップアップ画面表示
     MsgBox "異常終了しました。", _
           vbOKOnly + vbCritical, _
           "実行結果"
           
   '「バージョン戻し処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OLD_COPY_NOW_ERROR, lngErrCode)
           
    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sVersionUpdate
'//  機能名称  : バージョンアップ処理
'//  機能概要  : ワークバージョンを現行バージョンへ更新する。
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
'//                【指摘事項No.02修正対応】
'//                【残件:保守運改の切替結果通知対応】
'//  REVISIONS   ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY[TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS   ：(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 量産対応
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Sub sVersionUpdate()

    Dim objFso As New FileSystemObject      ' ファイルシステムオブジェクト
    Dim objFi As File                       ' ファイルオブジェクト
    Dim objFolder As Folder                 ' フォルダオブジェクト
    Dim stringWorkFolder As String          ' フォルダ名
    
    Dim lngErrCode     As Long              ' エラーコード
    Dim strSrcFile As String                ' コピー元
    Dim strDstFile As String                ' コピー先
    Dim bResult As Boolean                  ' 処理結果      ' EG20 V3.6.0.1追加

    On Error GoTo ErrorHandler          'エラーハンドルの登録

    ' /////////////////////////////////////////////////////////////////////////
    ' // ワークフォルダのファイル存在チェック
    stringWorkFolder = FLD_OPERATEPROGWRK
    If objFso.FolderExists(stringWorkFolder) <> True Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' フォルダがそもそも存在しない
' EG20 V5.8.0.1削除開始
'        MsgBox "ワークフォルダ内に、" _
'               & Chr(vbKeyReturn) & "ファイルが存在しません。", _
'               vbOKOnly + vbExclamation, _
'               "ワーク→実行コピー"
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
        MsgBox "異常終了しました。", vbCritical, "実行結果"
' EG20 V5.8.0.1追加終了
        Exit Sub                        ' 処理終了
    End If

    strSrcFile = pfGetFileNameTakuExec
    ' ファイル設定なし
    If Len(strSrcFile) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        MsgBox "異常終了しました。", vbCritical, "出力結果"
        Exit Sub
    End If

    strDstFile = stringWorkFolder & "\" & strSrcFile
    If objFso.FileExists(strDstFile) = False Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' ファイルが存在しない
' EG20 V5.8.0.1削除開始
'        MsgBox "ワークフォルダ内に、" _
'               & Chr(vbKeyReturn) & "ファイルが存在しません。", _
'               vbOKOnly + vbExclamation, _
'               "ワーク→実行コピー"
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
        MsgBox "異常終了しました。", vbCritical, "実行結果"
' EG20 V5.8.0.1追加終了
        Exit Sub                        ' 処理終了
    End If
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 旧フォルダの削除
    '//////////////////////////////////////////////////////
    '// 監視盤フォルダ内の操作卓旧フォルダを消去
    stringWorkFolder = FLD_OPERATEPROGOLD
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    '//////////////////////////////////////////////////////
    '// 操作卓フォルダを消去
    stringWorkFolder = PATH_OPERATE_APLOLD
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
   
    ' /////////////////////////////////////////////////////////////////////////
    ' // ワーク→実行コピー

    '//////////////////////////////////////////////////////
    '// ワーク→実行コピー
    strSrcFile = FLD_OPERATEPROGNOW
    strDstFile = FLD_OPERATEPROGOLD

    'フォルダオブジェクトを取得
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If

    ' /////////////////////////////////////////////////////
    ' // 全体をコピー
    strSrcFile = PATH_OPERATE_APL
    strDstFile = PATH_OPERATE_APLOLD

    'フォルダオブジェクトを取得
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If
   
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 実行フォルダの削除
    '//////////////////////////////////////////////////////
    '// 監視盤フォルダ内の操作卓実行フォルダを消去
    stringWorkFolder = FLD_OPERATEPROGNOW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    '//////////////////////////////////////////////////////
    '// 操作卓フォルダを消去
    stringWorkFolder = PATH_OPERATE_APL
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
   
    ' /////////////////////////////////////////////////////////////////////////
    ' // ワーク→実行コピー

    '//////////////////////////////////////////////////////
    '// ワーク→実行コピー
    strSrcFile = FLD_OPERATEPROGWRK
    strDstFile = FLD_OPERATEPROGNOW

    'フォルダオブジェクトを取得
    Set objFolder = objFso.GetFolder(strSrcFile)
    objFolder.Copy strDstFile

    ' /////////////////////////////////////////////////////
    ' // 全体をコピー
    strSrcFile = PATH_OPERATE_APLNEW
    strDstFile = PATH_OPERATE_APL

    'フォルダオブジェクトを取得
    Set objFolder = objFso.GetFolder(strSrcFile)
    objFolder.Copy strDstFile

    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing

' EG20 V3.6.0.1追加開始
    ' 操作卓プログラムデータ作成処理
    bResult = pfTakuProgramVersionCreateProc
' EG20 V3.6.0.1追加終了
    
    If bResult = True Then
' EG20 V5.8.0.1削除開始
'        ' 運改状態更新                                              ' EG20 V5.5.0.1追加
'        Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1追加
' EG20 V5.8.0.1削除終了
' EG20 V5.8.0.1追加開始
        ' 運改状態更新
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_KIRIKAE)
        Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        ' 操作卓のバージョン情報を表示する｡
        Call psVersionDisp
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '「正常終了」ポップアップ画面表示
        MsgBox "正常終了しました。", _
               vbOKOnly + vbInformation, _
               "実行結果"
    Else
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        ' 操作卓のバージョン情報を表示する｡
        Call psVersionDisp
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
        '「正常終了」ポップアップ画面表示
        MsgBox "異常終了しました。", _
               vbOKOnly + vbInformation, _
               "実行結果"
    End If
    
    Exit Sub '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    ' 操作卓のバージョン情報を表示する｡
    Call psVersionDisp
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    '「ワーククリア異常終了」ポップアップ画面表示
     MsgBox "異常終了しました。", _
           vbOKOnly + vbCritical, _
           "実行結果"
           
   '「バージョン戻し処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_WRK_COPY_NOW_BUTTOM, lngErrCode)
           
    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : pfGetFileNameTakuExec
'//  機能名称  : インストール実行ファイル名取得処理
'//  機能概要  : インストール実行ファイル名を取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : String    ファイル名
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Private Function pfGetFileNameTakuExec() As String

    Const lngBufSize = MAX_PATH         ' 取得文字列の文字数：ID、データ用
    Dim strRet As String * MAX_PATH     ' 取得文字列
    Dim lngRet As Long                  ' 戻り値
    Dim szFileName As String            ' ファイル名称
    Dim nNullIndex As Integer           ' 文字数ワーク
    
    pfGetFileNameTakuExec = ""
        
    'Iniファイルから実行ファイル名を取得
    lngRet = GetPrivateProfileString(HOSHUINI_SECTION_OPERATE, HOSHUINI_OPERATEKEY_INSTEXEC, _
                                        "", strRet, lngBufSize, HOSHU_FILE)
    
    nNullIndex = InStr(strRet, Chr(0))
    If nNullIndex <> 0 Then
        szFileName = Left(strRet, nNullIndex - 1)
    Else
        szFileName = ""                 ' EG20 V3.3.0.1削除
        szFileName = strRet             ' EG20 V3.3.0.1追加
    End If
    pfGetFileNameTakuExec = szFileName
    
End Function

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
    CmdRemove.Enabled = blnFlg                      ' 媒体取外
    cmdReturn.Enabled = blnFlg                      ' バージョン管理画面へ戻る

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'/
'/  関数名称     : pfTakuProgramVersionCreateProc
'/  機能名称     : 操作卓プログラムデータ作成処理
'/  機能概要     : 操作卓の実行バージョンを圧縮して操作卓データを作成する。
'/
'/                 型          名称            意味
'/  引数         : なし
'/  戻り値       : Boolean     False           異常終了
'/               :             True            正常終了
'/
'//  ORIGINAL    :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.273修正対応】
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考         :
'/////////////////////////////////////////////////////////////////////////////
Private Function pfTakuProgramVersionCreateProc() As Boolean

    Dim sInputFolder As String                  ' インストール元フォルダ名
    Dim objFso As New FileSystemObject          ' ファイルシステムオブジェクト
    Dim objFi As File                           ' ファイルオブジェクト
    Dim MyName As String                        ' ファイルフルパス名
    Dim sSrcFileName As String                  ' コピー元ファイル名
    Dim strCabTarget As String                  ' 圧縮対象ファイル
    Dim lngRetZip As Long                       ' 圧縮結果
    Dim objFolder As Folder                 ' フォルダオブジェクト

    Dim bResult As Long                     ' 処理結果

    On Error GoTo ErrorHandler                  ' エラーハンドルの登録

    pfTakuProgramVersionCreateProc = True
    sInputFolder = FLD_OPERATEPROGNOW
    strCabTarget = ""
    For Each objFi In objFso.GetFolder(sInputFolder).files      'ループを開始
        If objFso.FileExists(objFi.Path) = True Then            'ファイル名の取得チェック
            'ディレクトリ名を取得
            MyName = objFi.Name
            '媒体内ファイル名を作成
            sSrcFileName = sInputFolder & "\" & MyName
            strCabTarget = strCabTarget & sSrcFileName & " "
        End If
    Next

   ' すべてのディレクトリを列挙する
    For Each objFolder In objFso.GetFolder(sInputFolder).SubFolders
        MyName = objFolder.Path
        strCabTarget = strCabTarget & MyName & " "
    Next


    lngRetZip = gsubCabZip(MELTED_TAKUVERSION, strCabTarget)
    
    If (lngRetZip <> 0) Then   '圧縮結果が正常(0)以外
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LZH_ERROR, 0)
        Set objFso = Nothing
        Set objFi = Nothing
        pfTakuProgramVersionCreateProc = False
        Exit Function
    End If

    Set objFso = Nothing
    Set objFi = Nothing

    ' /////////////////////////////////////////////////////
    ' // 操作卓プログラムデータの作成
    bResult = dllCreateFile_TakuProgramData(1, MELTED_TAKUVERSION)
    If bResult = False Then
       pfTakuProgramVersionCreateProc = False
       Exit Function
    End If
    Exit Function

ErrorHandler:   ' エラー処理。
    Set objFso = Nothing
    Set objFi = Nothing
    pfTakuProgramVersionCreateProc = False

End Function


