VERSION 5.00
Begin VB.Form frmKikiSettei 
   BorderStyle     =   0  'なし
   Caption         =   "機器情報設定"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "駅務機器ID確認"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   6360
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8280
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "係員設定 保存／復元"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "LANカード設定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "駅都度データ確認"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6360
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "機器構成設定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "駅都度データ設定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " メンテナンス   画面へ戻る"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "機器情報設定"
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
      TabIndex        =   6
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：機器情報設定メニュー画面.frm
'//  パッケージ名：機器情報設定メニューのフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　駅務機器ID確認画面、LANカード上位下位設定画面追加
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//                 画面ロック処理／画面ロック解除処理追加
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                画面再前面表示修正(不具合修正)
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000       'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 機器情報設定メニュー画面(アクティブ時：イベントプロシージャ)
'//  機能概要  : メール受信タイマ起動
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
    
    'タイマを起動する
    tmrMail.Enabled = True

End Sub

'EG20 V2.1.0.1 ADD START 【フェーズ２対応】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 機器情報設定メニュー画面(ディアクティブ時)
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
'//  機能名称  : 機器情報設定メニュー画面(ロード時：イベントプロシージャ)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'V1.4.0.1　ADD START
    'IDU縮退チェック
    psIDUCheck
    
    If pbIDUSts = 1 Then
     '駅務機器ID確認非表示
      cmdFixedExe(5).Visible = False
    End If
    'V1.4.0.1　ADD END
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称     　　　意味
'//  引数      : Integer　 Index          選択釦のインデックス
'//
'//              型        値        　　 意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応
'//                 ・駅都度データ確認メニュー画面→駅都度データ確認(駅情報)画面
'//                 ・機器構成設定メニュー画面→機器構成設定(駅情報)画面
'//                 ・LANカード上位下位設定釦押下時、画面表示処理追加
'//                 ・駅務機器ID確認画面追加
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)

    'エラールーチンを宣言
    On Error Resume Next
    
'V1.12.0.1 ADD START
    '全ボタンを押下不可とする。
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    Select Case Index
        
        Case 0                                 '駅都度データ設定
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKITUDO_DATA_SETTEI, 0)
            
            '画面表示
            Load frmEkisettei
            frmEkisettei.Show 1

        Case 1                                 '駅都度データ確認
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKITUDO_DATA_KAKUNIN, 0)
            'V1.4.0.1 DEL START
            '画面表示
            'Load frmEkiDataGateMenu
            'frmEkiDataGateMenu.Show 1
            'V1.4.0.1 DEL END
            'V1.4.0.1 ADD START
            '画面表示
            Load frmEkiData
            frmEkiData.Show 1
            'V1.4.0.1 ADD END
   
        Case 2                                 '機器構成設定
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_KIKI_KOUSEI_SETTEI, 0)
            'V1.4.0.1 DEL START
            '画面表示
            'Load frmKikiDataMenu
            'frmKikiDataMenu.Show 1
            'V1.4.0.1 DEL END
            'V1.4.0.1 ADD START
            '画面表示
            Load frmKikiData
            frmKikiData.Show 1
            'V1.4.0.1 ADD END
            
        Case 3                                 'LANカード上位下位設定
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_LAN_CARD_SETTEI, 0)

'V1.4.0.1 ADD START
            '画面表示
            Load frmLanSettei
            frmLanSettei.Show 1
'V1.4.0.1 ADD END

        Case 4                                 '係員設定保持復元
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_KAKARI_SAVE_RESTORE, 0)
            
            '画面表示
            Load frmRenewData
            frmRenewData.Show 1
'V1.4.0.1 ADD START
        Case 5                                 '駅務機器ID
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKIMUKIKI_ID, 0)
            '画面表示
            Load frmEkimKikiId
            frmEkimKikiId.Show 1
'V1.4.0.1 ADD END

        Case Else
            '処理なし
            
    End Select

'V1.12.0.1 ADD START
    '全ボタンを押下可とする。
    Call SetEnableTrue
'V1.12.0.1 ADD END
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_END, 0)
    
    '自画面消去
    Unload Me

End Sub

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       'V1.17.0.1 DEL START
'        AppActivate frmRenewData.Caption, False
'        pfFormActive (frmRenewData.hwnd)
       'V1.17.0.1 DEL START
       'V1.17.0.1 ADD START
        AppActivate frmKikiSettei.Caption, False
        pfFormActive (frmKikiSettei.hwnd)
       'V1.17.0.1 ADD END
    End If

End Sub

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
    cmdFixedExe(0).Enabled = False
    cmdFixedExe(1).Enabled = False
    cmdFixedExe(2).Enabled = False
    cmdFixedExe(3).Enabled = False
    cmdFixedExe(4).Enabled = False
    cmdFixedExe(5).Enabled = False
    cmdReturn.Enabled = False
    
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
    cmdFixedExe(0).Enabled = True
    cmdFixedExe(1).Enabled = True
    cmdFixedExe(2).Enabled = True
    cmdFixedExe(3).Enabled = True
    cmdFixedExe(4).Enabled = True
    cmdFixedExe(5).Enabled = True
    cmdReturn.Enabled = True
    
End Sub
'V1.12.0.1 ADD END

