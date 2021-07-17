VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   0  'なし
   Caption         =   "保守員パスワード入力"
   ClientHeight    =   9000
   ClientLeft      =   2700
   ClientTop       =   2520
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   Picture         =   "パスワード入力画面.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdKakutei 
      BackColor       =   &H00C0C0C0&
      Caption         =   "確  定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Frame fraPass 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5655
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ｃ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   11
         Left            =   2880
         TabIndex        =   12
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BS"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   1680
         TabIndex        =   11
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   9
         Left            =   2880
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   1680
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   2880
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   1680
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   480
         TabIndex        =   5
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   2880
         TabIndex        =   4
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   480
         Width           =   3375
      End
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
      Height          =   1455
      Left            =   8400
      TabIndex        =   14
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Timer tmrMail 
      Left            =   600
      Top             =   720
   End
   Begin VB.Label lblGuide 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H0000FFFF&
      Caption         =   "  この画面は、保守員専用です！       保守員以外の方は、 "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   0
      Left            =   2520
      TabIndex        =   16
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label lblGuide 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   20
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label lblGuide 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H0000FFFF&
      Caption         =   "  画面右下の「監視画面へ戻る」ボタンを押し戻って下さい。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   19
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "保守員パスワード入力"
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
      TabIndex        =   18
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  '透明
      Caption         =   "パスワードを入力して下さい。"
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
      Left            =   2400
      TabIndex        =   17
      Top             =   1560
      Width           =   4095
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmPass.frm
'//  パッケージ名：パスワード入力画面
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・EG10保守より、パスワード入力画面(frmPass.frm)を流用
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                 ・レジストリ取得異常時処理変更
'//                 ・フォームアンロード処理追加
'//     REVISIONS :(1.6.0.1) 2009-06-23   REVISED BY [TCC] S.Terao
'//                 ・画面表示/消去タイミング修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yoshimori
'//                 パスワード不一致の画面遷移変更
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(EG20 V30.3.0.1)  2014-09-18  REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応
'//                 【HKRK_Kansi07_007_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : パスワード入力画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
    
    Dim iKansiAplChk As Integer     'アプリ起動チェック戻り値用　' EG20 V2.1.0.1[Mainte_03_01]追加
    
    '最大化表示する。
    Me.WindowState = 2
    'メール受信タイマを起動する。
    tmrMail.Enabled = True

' EG20 V2.1.0.1[Mainte_03_01]追加開始
    '監視盤起動/未起動チェックを行う。チェック状態により処理分岐を行う。
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '監視盤起動時：戻る釦の文言「監視画面へ戻る」
        cmdReturn.Caption = "監視画面へ戻る"
    Else
        '監視未起動時：戻る釦の文言「Windowsに戻る」
        cmdReturn.Caption = "Windowsに戻る"
    End If
' EG20 V2.1.0.1[Mainte_03_01]追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : パスワード入力画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ停止
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
'//  機能名称  : パスワード入力画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.2.0.1) 2009-02-26   REVISED BY [TCC] S.Terao
'//             立ち上げ処理、レジストリ取得異常時起動させない修正
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//             レジストリのデフォルト値取得に伴う修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-09   REVISED BY [TCC] M.Matsumoto
'//             【フェーズ２対応】ログ出力フォルダ変更
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim bRet As Boolean     '戻り値
    Dim lngErrCode As Long  '関数エラーコード
    Dim slogPath As String  '保守ログ
    
    On Error Resume Next
   
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '保守ログクラスを生成
'    slogPath = PATH_LOG & HOSHULOG_FILE            'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    slogPath = PATH_HOSHULOG & HOSHULOG_FILE        'EG20 V2.1.0.1 ADD 【フェーズ２対応】
    bRet = dllHoshulogClass(slogPath, lngErrCode)
    
    '初回起動
    iStaFlag = 0
    '「保守起動」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_PROCESS_START, 0)
    '起動済
    iStaFlag = 1
    
     '「パスワード入力画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_START, 0)
    
    'パス設定
    ChDrive App.Path
    ChDir App.Path
     
    '立ち上げ処理を行う
    bRet = pfStartUpProc
    If bRet = False Then
    '立ち上げ処理異常時は強制終了する。
        'pbAbortFlag = True 'V1.2.0.1 DEL
        'V1.2.0.1 ADD START
        pfAbortProc
        End
        'V1.2.0.1 ADD END
    End If
   'パスワードファイル初期処理を行う。
    sPassFileInitialize
    
    'IDU/LDUのパスをレジストリより取得する。
    bRet = sGetRegIDU_LDU_Path
' V1.3.0.1 DEL START
'   If False = bRet Then
'       'Exit Sub 'V1.2.0.1 DEL
'       'V1.2.0.1 ADD START
'       pfAbortProc
'       End
'       'V1.2.0.1 ADD END
'    End If
' V1.3.0.1 DEL START
    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

    pfFormActive (frmPass.hwnd)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdKakutei_Click
'//  機能名称  : 「確定」釦押下時処理
'//  機能概要  : 入力パスワードをチェックし、
'//              エラー：「パスワードエラー」ポップアップを表示。
'//              正　常：保守画面を表示 。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                ・フォームアンロード処理追加
'//     REVISIONS :(1.6.0.1) 2009-06-23   REVISED BY [TCC] S.Terao
'//                 ・画面表示/消去タイミング修正
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yoshimori
'//                 パスワード不一致の画面遷移変更
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(EG20 V30.3.0.1)  2014-09-18  REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応
'//                 【HKRK_Kansi07_007_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdKakutei_Click()
    Dim iResponse As Integer      'メッセージボックス表示戻り値
    Dim bFlag As Boolean          '入力パスワードチェックフラグ(True：一致。False:不一致)
    Dim intPassFileNo As Integer  'パスワードファイルのファイル番号
    Dim sPassword As String       'パスワードファイルの１行分のデータ
    Dim sPassData As String       'システム月日。
    Dim sHoshuPass As String      '連結パスワード=ファイル定義パス+システム月日
    Dim sPass As String           'ファイル定義パス
    
    '「パスワード入力画面：確定釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKUTEI_BUTTOM, 0)
 
    '入力パスワードチェックフラグを不一致で初期化する。
    bFlag = False
    
    '未入力での「確定」釦押下時
    If txtPass = "" Or IsNull(txtPass) Then
       '「パスワード入力画面：パスワード未入力異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_NOT_KEY, 0)
        'V1.20.0.1 DEL START
        ''メール受信タイマを停止する
        'tmrMail.Enabled = False
        ''「パスワード入力画面：パスワード入力画面消去」ログ出力
        ' Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
        ''終了処理を行う
        'psEndProc
        ' 'パスワード入力画面を閉じる。
        'Unload Me
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
        'エラーメッセージを表示する。
        iResponse = MsgBox("パスワードが違います。", _
                            vbOKOnly, _
                            "パスワードエラー")
        'V1.20.0.1 ADD END
        Exit Sub
    End If
    
    'パスワード桁数チェックを行う。
    If (Len(txtPass) <= INPUT_PASSWORD_MAX) And (Len(txtPass) >= INPUT_PASSWORD_MIN) Then
       On Error GoTo FileError
       '未使用のファイル番号を取得する。
       intPassFileNo = FreeFile
       'パスワードファイルを開く。
       Open PASSWORD_FILE_FULLPASS For Input As #intPassFileNo
       'システム月日を「MMDD」で取得する。
       sPassData = Format(Date, "mmdd")
       'ファイルの終端まで繰り返す。
'       Do While Not EOF(1)                                 ' EG20 V3.3.0.1削除
       Do While Not EOF(intPassFileNo)                      ' EG20 V3.3.0.1追加
         'パスワードファイルの先頭から１行ずつ読込む。
         Line Input #intPassFileNo, sPassword
         'パスワードの定義設定があるかどうかチェックする。
         If sPassword <> "" Then
           'パスワードファイルにある定義と取得月日を連結させる。
           sPass = Mid(sPassword, 3, 8)
           If sPass <> "" Then
              sHoshuPass = sPass + sPassData
              '連結パスワードと、入力値を比較する。
              If sHoshuPass = txtPass Then
                 '一致した場合、パスワードチェックフラグを一致にする。
                 bFlag = True
                 '一致したパスワード部のユーザレベルを取得する。
                 pbUserLevel = CInt(Left$(sPassword, 1))
                 'Exit Do       'EG20 V30.3.0.1 DEL 【HKRK_Kansi07_007_01】
                 'EG20 V30.3.0.1 ADD START 【HKRK_Kansi07_007_01】
                 'ユーザレベルが「一般」「特権」の場合のみ一致扱いとし、ループを抜ける。
                 If pbUserLevel = 0 Or pbUserLevel = 1 Then
                    Exit Do
                 Else
                    'それ以外はループを抜けずに処理を続ける。
                    'パスワードエラー扱いなのでpbUserLevelはリセットする。
                    'そうしないと、pbUserLevelに2がセットされていた場合、戻る釦押下時にプロセス終了要求の補助エリアに2がセットされてしまい、
                    '業務系保守画面が表示されてしまう。
                    pbUserLevel = 0
                    bFlag = False
                 End If
                 'EG20 V30.3.0.1 DEL END 【HKRK_Kansi07_007_01】
              End If
           End If
         End If
       Loop
     'パスワードファイルを閉じる。
     Close #intPassFileNo
    End If
    
    'パスワードエラーの場合
    If bFlag = False Then
        'エラーメッセージを表示する。
        iResponse = MsgBox("パスワードが違います。", _
                            vbOKOnly, _
                            "パスワードエラー")
        '「パスワード入力画面：入力パスワード異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_KEY_ERROR, 0)
        'V1.20.0.1 DEL START
        ''メール受信タイマを停止する。
        'tmrMail.Enabled = False
        ''終了処理を行う。
        'psEndProc
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
        '入力パスワードを初期化する
        txtPass = ""
        'パスワード入力画面を閉じず、処理を終了する
        Exit Sub
        'V1.20.0.1 ADD END
        
    'パスワード正常の場合
    Else
        '機器構成データクラス生成処理
        Call MakeInitKikiClas
        
        'パスワード入力ログを保守操作ログに出力する。
        If pbUserLevel = 0 Then  'ユーザレベル
          '「パスワード入力画面：入力パスワード正常」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_OK, 0)
        ElseIf pbUserLevel = 1 Then
         '「パスワード入力画面：入力パスワード正常」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_OK, 0)
        Else
            'EG20 V30.3.0.1 ADD START【HKRK_Kansi07_007_01】
            'ユーザレベルが0、１以外はbFlagをFalseにしてしまっているので、ここが実行されることはないが念のため、
            '業務系保守が起動しないように、ここでもパスワード不一致扱いとする
            iResponse = MsgBox("パスワードが違います。", _
                                vbOKOnly, _
                                "パスワードエラー")
            '「パスワード入力画面：入力パスワード異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_KEY_ERROR, 0)
            '入力パスワードを初期化する
            txtPass = ""
            'パスワード入力画面を閉じず、処理を終了する
            Exit Sub
            'EG20 V30.3.0.1 ADD END【HKRK_Kansi07_007_01】
            'EG20 V30.3.0.1 DEL START 【HKRK_Kansi07_007_01】
'         '「パスワード入力画面：入力パスワード正常」ログ出力
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_OK, 0)
'           '特殊ユーザの場合は業務終了し、
'           '管理にプロセス終了要求ﾒｰﾙ(補助情報＝ﾕｰｻﾞﾚﾍﾞﾙを設定)を送信する。
'            tmrMail.Enabled = False
'            '「パスワード入力画面：パスワード入力画面消去」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
'            Call psEndProc
'            Unload Me
'            Exit Sub
            'EG20 V30.3.0.1 DEL END 【HKRK_Kansi07_007_01】
        End If
        'V1.6.0.1 ADD START
        'メンテナンスメニュー画面を表示する｡
        frmHoshu.Show  'V1.6.0.1 DEL
        'V1.6.0.1 ADD END
        'パスワード入力画面を非アクティブ表示にする｡
        '「パスワード入力画面：パスワード入力画面消去」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
        Me.Hide
        'メンテナンスメニュー画面を表示する｡
        'frmHoshu.Show  'V1.6.0.1 DEL
    End If
' V1.3.0.1 ADD START
        'パスワード入力画面を閉じる。
        Unload Me
' V1.3.0.1 ADD END
  Exit Sub

FileError:
   '「パスワード入力画面：パスワード入力画面消去」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
    'パスワード入力画面を閉じる。
   Unload Me
 End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「監視画面へ戻る」釦押下時処理
'//  機能概要  : プロセス終了処理をし、画面を閉じる。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
  On Error Resume Next
  
  'メール受信タイマを停止する
  tmrMail.Enabled = False
  '「パスワード入力画面：パスワード入力画面消去」ログ出力
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
  
' EG20 V2.1.0.1[Mainte_03_01]追加開始
    If CheckAppStart(PROC_KANRI) = 0 Then
        '管理プロセスが起動していない場合
        psEndHoshuProc
    Else
' EG20 V2.1.0.1[Mainte_03_01]追加終了
        '終了処理を行う
        psEndProc
    End If          ' EG20 V2.1.0.1[Mainte_03_01]追加
  'パスワード入力画面を閉じる。
  Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdNumber_Click
'//  機能名称  : テンキー各釦押下時処理
'//  機能概要  : 押下されたキーに従って、パスワード入力欄を更新する。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下キーの種別
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdNumber_Click(Index As Integer)
    
    'ＢＳキー押下時
    If Index = 10 Then
        '入力値有無チェックを行う。
        If (txtPass <> "") Then
            '入力値が有った場合、末尾入力値を１文字削除する。
            txtPass = Left(txtPass, Len(txtPass) - 1)
        End If
        '処理終了。
        Exit Sub
    End If
    
    'Ｃキー押下時
    If Index = 11 Then
        '入力値を全て削除する。
        txtPass = ""
        '処理終了。
        Exit Sub
    End If
    
    '０～９までの数字キーが押下時、入力済み文字列の末尾に追加する。
    txtPass = txtPass & Index
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ処理
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
On Error Resume Next
    
    '汎用メール受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmPass.Caption, False
        pfFormActive (frmPass.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sPassFileInitialize
'//  機能名称  : パスワードファイル初期処理
'//  機能概要  : パスワードファイルが存在しなければ作成し、
'//              デフォルトのパスワードを格納する。
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
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sPassFileInitialize()
    Dim intPassFileNo As Integer  'パスワードファイルのファイル番号
    Dim iLine As Integer          'パスワードファイルの行カウンタ
    Dim sPassword As String       'パスワードファイルの１行分のデータ
    
    '行カウンタを0にて初期化する。
    iLine = 0
    On Error GoTo FileError
    '未使用のファイル番号を取得する。
    intPassFileNo = FreeFile
    'パスワードファイルを開く。
    Open PASSWORD_FILE_FULLPASS For Input As #intPassFileNo
    'ファイルの終端まで繰り返す。
'    Do While Not EOF(1)                                    ' EG20 V3.3.0.1削除
    Do While Not EOF(intPassFileNo)                         ' EG20 V3.3.0.1追加
      'パスワードファイルに有効な定義設定が有るかチェックするため、1行分読み込む。
      Line Input #intPassFileNo, sPassword
        '定義設定がある場合、行カウンターをカウントアップする。
        If sPassword <> "" Then
            iLine = iLine + 1
            Exit Do
        End If
    Loop

    'パスワードファイルを閉じる。
    Close #intPassFileNo
'Exit Sub  'V1.7.0.1 DEL

FileError:
    '「パスワード入力画面：デフォルトパスワードファイル作成」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_KEY_FILE_CREATE, 0)
    '行カウンターが0の場合(=定義設定無し)
    If iLine = 0 Then
        'パスワードファイルを開く。
        Open PASSWORD_FILE_FULLPASS For Output As #intPassFileNo
        'デフォルトのパスワード「 '特権ユーザ用："１"」を書き込む。
        Print #intPassFileNo, "1,1"
        'パスワードファイルを閉じる。
        Close #intPassFileNo
    End If
End Sub
 
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sGetRegIDU_LDU_Path
'//  機能名称  : IDU/LDUのパスをレジストリより取得する。
'//  機能概要  : IDU/LDUのパスをレジストリより取得を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                 ・レジストリのデフォルト値設定及び処理変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sGetRegIDU_LDU_Path()
    
    On Error Resume Next
    
    sGetRegIDU_LDU_Path = False
    
    'IDU：アプリパス取得
    PATH_IDU_APP = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "AplRoot")
    If PATH_IDU_APP = "" Then
       '「レジストリ情報取得異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       'レジストリのデフォルト値を取得
       PATH_IDU_APP = REG_IDU_APLROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If

    'IDU：DBパス取得
    PATH_IDU_DB = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "DataRoot")
    If PATH_IDU_DB = "" Then
       '「レジストリ情報取得異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       'レジストリのデフォルト値を取得
       PATH_IDU_DB = REG_IDU_DBROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If
    
    'IDU：バックアップパス取得
    PATH_BUC = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "BackupRoot")
    If PATH_BUC = "" Then
       '「レジストリ情報取得異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       'レジストリのデフォルト値を取得
       PATH_BUC = REG_IDU_BACKUPROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If

    'IDU：ログパス取得
    PATH_IDU_LOG = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "LogRoot")
    If PATH_IDU_LOG = "" Then
       '「レジストリ情報取得異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       'レジストリのデフォルト値を取得
       PATH_IDU_LOG = REG_IDU_LOGROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If

'    LDUアプリパス取得
    PATH_LDU_APP = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\LD_Utility", "AplRoot")
    If PATH_LDU_APP = "" Then
      '「レジストリ情報取得異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       'レジストリのデフォルト値を取得
       PATH_LDU_APP = REG_LDU_APLROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If
    
'    LDUログパス取得
    PATH_LDU_LOG = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\LD_Utility", "LogRoot")
    If PATH_LDU_LOG = "" Then
      '「レジストリ情報取得異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       'レジストリのデフォルト値を取得
       PATH_LDU_LOG = REG_LDU_LOGROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If
    
    sGetRegIDU_LDU_Path = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : MakeInitKikiClas
'//  機能名称  : 機器構成データクラス生成処理
'//  機能概要  : 機器構成データクラスを生成する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.1.0.1) 2010-05-28   REVISED BY [TCC] S.Yoshimori
'//                 １ラッチ共同使用駅対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub MakeInitKikiClas()

    Dim lErrCode As Long          'エラーコード
    Dim bRet As Boolean           '関数戻り値
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト 'V2.1.0.1 ADD

    '---------------------------------------------
    '機器構成データクラス生成
    '---------------------------------------------
    bRet = dllInitKikiClass(lErrCode)
    
    'ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_CREATE_KIKICLASS, lErrCode)
    
    '---------------------------------------------
    '駅都度データ紐付けファイルメモリ展開
    '---------------------------------------------
    bRet = dllMemEkiDataChange(EKI_DATA_CHANGE_FILE, lErrCode)
    
    If bRet = False Then
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_CREATE_EKITUDOMAP_ERR, lErrCode)
    Else
        '正常ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_CREATE_EKITUDOMAP, 0)
    End If
    
    'V2.1.0.1 ADD START
    '---------------------------------------------
    '駅タイプ都度データ紐付けファイルメモリ展開
    '---------------------------------------------
    '駅タイプ都度データ紐付けファイルが存在するなら処理を行なう
    If (objFso.FileExists(EKI_TYPE_DATA_CHANGE_FILE) = True) Then

        '駅タイプ都度データ紐付けファイルメモリ展開関数
        bRet = dllMemEkiTypeDataChange(EKI_TYPE_DATA_CHANGE_FILE, lErrCode)
        
        If (False = bRet) Then
            '異常ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_CREATE_EKITYPE_TUDOMAP_ERR, lErrCode)
        Else
            '正常ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_CREATE_EKITYPE_TUDOMAP, 0)
        End If
    Else
        '駅タイプ都度データ紐付けファイルなし
        'ログ出力
        Call sLogTraceReq(LTYP_WARNING, L3AN_FILE, PASS_NOT_EKITYPE_TUDOFILE, 0)
    End If
    
    Set objFso = Nothing
    'V2.1.0.1 ADD END
End Sub

