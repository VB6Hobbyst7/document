VERSION 5.00
Begin VB.Form frmEkiDataGateMenu 
   BorderStyle     =   0  'なし
   Caption         =   "駅都度データ確認"
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
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   360
      Top             =   8280
   End
   Begin VB.CommandButton cmdFixedExe 
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
      Height          =   855
      Index           =   5
      Left            =   2040
      TabIndex        =   7
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "駅設定テキスト出力"
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
      TabIndex        =   6
      Top             =   5040
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
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
      Height          =   855
      Index           =   3
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
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
      Height          =   855
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "自改"
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
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "駅情報"
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "    メニュー     画面へ戻る"
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
   Begin VB.Label Label7 
      Caption         =   "選択されている駅の現在の駅都度データ１駅分をテキスト表示する。"
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
      Left            =   6360
      TabIndex        =   13
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Label Label6 
      Caption         =   "・・・"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "・・・"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "駅都度データ１駅分をインストールする。"
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
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "・・・"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "選択されている駅の現在の駅都度データ１駅分を出力する。"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   2880
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "駅設定確認"
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
End
Attribute VB_Name = "frmEkiDataGateMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：駅都度データ確認メニュー画面.frm
'//  パッケージ名：駅都度データ確認メニュー画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 駅設定ファイル書込み先ディレクトリ位置変更
'//                 ディスク情報取得位置変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000       'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 駅都度データ確認メニュー画面(アクティブ時：イベントプロシージャ)
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
'//  機能名称  : 駅都度データ確認メニュー画面(ディアクティブ時)
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
'//  機能名称  : 駅都度データ確認メニュー画面(ロード時：イベントプロシージャ)
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

    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_START, 0)
    
    '----------------------------------------------------
    '画面初期値設定
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmRenewData.Caption, False
        pfFormActive (frmRenewData.hwnd)
    End If

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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)

    'エラールーチンを宣言
    On Error Resume Next
    
    Select Case Index
        Case 0                                 '駅情報
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKIINFO, 0)
            
            '画面表示
            Load frmEkiData
            frmEkiData.Show 1
        Case 1                                 '自改
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_GATE, 0)
            
            '画面表示
            Load frmEkiDataGate
            frmEkiDataGate.Show 1
        Case 2                                  '駅設定出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_OUTPUT, 0)
            
            '駅設定出力処理
            Call sEkiSetteiOutPut
        
        Case 3                                  '駅設定入力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_INPUT, 0)
            
            '駅設定入力処理
            Call sInstolEkiSettei
        
        Case 4                                  '駅設定テキスト出力
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_DISP_TEXT, 0)
            
            '駅設定テキスト出力処理
            Call sDispTextEkiDataNow
        
        Case 5                                  '媒体取外
            '画面操作ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
            
            '媒体取外処理
            Call pfRemove(Me)
        
    End Select
    
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_END, 0)
    
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 駅設定ファイル書込み先ディレクトリ位置変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sEkiSetteiOutPut()

    Dim strFileName          As String          'ファイル名
    Dim sWriteDir            As String          'フォルダ名
    Dim iResponse            As Integer         'MsgBox戻り値

    'エラールーチンを宣言
    On Error Resume Next
    
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
'    sWriteDir = pfDirSelection("a:", "駅設定ファイル書込み先のディレクトリ選択")   'V1.12.0.1 DEL
    sWriteDir = pfDirSelection("H:", "駅設定ファイル書込み先のディレクトリ選択")    'V1.12.0.1 ADD
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        'ファイルコピー
        FileCopy EKI_SETTI_FILE, sWriteDir & Dir(EKI_SETTI_FILE)
        
        'ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
       '正常終了
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "媒体出力結果")
    
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

    iResponse = MsgBox("異常終了しました", vbOKOnly, "媒体出力結果")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 ディスク情報取得位置変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sInstolEkiSettei()

    Dim iResponse           As Integer          'MsgBox戻り値
    Dim bRet                As Boolean          '関数戻り値
    Dim lErrCode            As Long             'エラーコード
    Dim strFileName         As String           '媒体ファイル名
    
    Dim iRet                    As Integer      'メッセージボックス戻り値
    Dim lSekuta                 As Long         'セクタ（クラスタ当り）
    Dim lByte                   As Long         'バイト数（セクタ当り）
    Dim lKurasuta               As Long         'フリークラスタ数
    Dim lDrive                  As Long         'ドライブのクラスタ数（合計）
    Dim strDrive                As String       'ドライブ
    
    'エラールーチンを宣言
    On Error Resume Next
    
    iResponse = MsgBox("駅都度データ１駅分をインストールします。" & Chr(vbKeyReturn) & _
                        "よろしいですか？", _
                        vbYesNo + vbExclamation, _
                        "駅設定入力確認")
    
    If iResponse = vbNo Then Exit Sub
    
    'ディスク情報を取得
'    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
    iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD

    If lDrive = 0 Then
        strDrive = "d:"
    Else
'        strDrive = "a:"        'V1.12.0.1 DEL
        strDrive = "H:"         'V1.12.0.1 ADD
    End If

    '媒体ファイル名取得
    strFileName = pfFileSelection(strDrive, "*.csv", "駅設定ﾌｧｲﾙ選択")
    
    'ファイル存在チェック
    If strFileName <> "" Then

        '現在駅設定データインストール処理
        bRet = dllInstolEkiDataNow(strFileName, EKI_SETTI_FILE, lErrCode)
    
        If bRet = False Then
            
            '異常ログ出力
            Call pfOutPutErrLog(lErrCode)
            
            '異常終了
            iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
            
        Else
        
             'ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
            '正常終了
            iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "駅設定入力結果")
            
        End If
    
    Else
        'ファイルなし
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, ERROR_MEDIUM_OTHER_ERR, 0)
        
        '異常終了
        iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定入力結果")
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDispTextEkiDataNow()

    Dim strFileName          As String          'ファイル名
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim lRetVal              As Long            '戻り値
    Dim sCommand             As String          'コマンド文字列

    'エラールーチンを宣言
    On Error Resume Next
    
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
    
    sCommand = MN_EXE_MEMO & EKI_SETTI_FILE         'メモ帳実行コマンドを作成する
    lRetVal = Shell(sCommand, vbMaximizedFocus)     'ノートパッドを起動する
    AppActivate lRetVal, True                       'アクティブ（前面表示）にする
    SendKeys "{LEFT}", True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
