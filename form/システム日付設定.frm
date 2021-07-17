VERSION 5.00
Begin VB.Form frmSystemDateSettei 
   BorderStyle     =   0  'なし
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   7200
      Top             =   6840
   End
   Begin Hoshu.ctlDateSetting ctlDateSetting1 
      Height          =   7000
      Left            =   720
      TabIndex        =   13
      Top             =   1000
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12356
   End
   Begin VB.Timer tmrKakunin 
      Left            =   6960
      Top             =   8400
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "確  定"
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
      Left            =   9870
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   12
      Top             =   6270
      Width           =   1725
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   9000
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   11
      Top             =   6270
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   9000
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   10
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   9870
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   10740
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   8
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   9000
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   7
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   9870
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   6
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   10740
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   5
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   9000
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   4
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   9870
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   3
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
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
      Left            =   10740
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   2
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton cmdModoru_Menu 
      Caption         =   "     システム設定     画面へ戻る"
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
      Left            =   8760
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   1
      Top             =   7800
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "システム日付設定"
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
      TabIndex        =   0
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmSystemDateSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  ファイル名     : frmSystemDateSettei
'//  パッケージ名   : システム日付設定画面
'//  概要           : システム日付設定画面の処理を定義する。
'//
'//  ORIGINAL       :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                   EG20フェーズ２対応
'//                   EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS      :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                   2014年度施策 【EG20_KANSI05_01】
'//  REVISIONS      :(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考           :
'//////////////////////////////////////////////////////////////////////////////

Option Explicit

Private intPos As Integer       '(0:未選択・1:年・2:月・3:日・4:時・5:分・6:秒)日時入力項目のカレント位置
Private gintCtrlIndex As Integer

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

Private Sub ctlDateSetting1_BtnOuka(intIndex As Integer)
    
    On Error Resume Next
    gintCtrlIndex = intIndex

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  関数名称     : Form_Load
'/  機能名称     : Form_Load時処理
'/  機能概要     : Form_Load時処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20フェーズ２対応
'/             EG20統合監視盤USDM対応番号【Mainte_03_01】
'/ REVISIONS :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next

    Dim strDateTime As String       '現在日時設定用

    '「システム日付設定画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

' EG20 V6.8.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
' EG20 V6.8.0.1 ADD END

    ' コントロールの保存インデックスを初期化
    gintCtrlIndex = -1
    
    ctlDateSetting1.psInitialize
    
    ' 日時設定釦コントロールに初期値を設定する。
    ctlDateSetting1.TotalArea = Format$(Now, "yyyymmddhhmmss")
    
    ' 設定した内容をコントロール上に表示する。
    ctlDateSetting1.DisplaySetUp
    ' コントロールを表示する。
    ctlDateSetting1.Enable = 0
    ctlDateSetting1.Visible = True
    
    ' 確定ボタンを押下不可
    cmdKakutei.Enabled = False
    
End Sub


'/////////////////////////////////////////////////////////////////////////////
'/   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  関数名称     : cmdTenkey_Click
'/  機能名称     : テンキー押下時処理
'/  機能概要     : テンキーが押下された処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20フェーズ２対応
'/             EG20統合監視盤USDM対応番号【Mainte_03_01】
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdTenkey_Click(Index As Integer)

    On Error Resume Next
    
    '「システム日付設定画面：テンキー釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_TENKEY_BUTTON, 0)

    
    ' 月とか時とかのボタンが押下されていない場合は何もしない    ' rev 02.15
    If (gintCtrlIndex = -1) Then
        Exit Sub
    End If
    
    '確認ボタンをＥｎａｂｌｅにする。
    cmdKakutei.Enabled = True
    
    'テンキーコントロールより、パラメータとして受け取った入力値を、
    '日時設定コントロールの個別入力値エリアプロパティへ設定する。
    ctlDateSetting1.InputArea = CStr(Index)
       
    '日時設定ボタンコントロールの入力値表示処理メソッドを行う。
    ctlDateSetting1.DisplayInput

End Sub
'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  関数名称     : cmdKakutei_Click
'/  機能名称     : 確定ボタン押下時処理
'/  機能概要     : 確定ボタンが押下された処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20フェーズ２対応
'/             EG20統合監視盤USDM対応番号【Mainte_03_01】
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdKakutei_Click()
    
    Dim i As Integer
    Dim udtSendData         As ML_KYOTU_INF     ' 共通エリア
    Dim lngSendSize         As Long             ' 送信するメールサイズ
    Dim lngErrCode          As Long             ' エラーコード
    Dim bRet                As Boolean          ' メール送信処理戻り値
    Dim iResponse           As Integer          ' メッセージボックス戻り値
    
    Dim strDate             As String
    
    '「システム日付設定画面：確定釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_KAKUTEI_BUTTON, 0)
    
    ' 日付チェック
    ctlDateSetting1.InputCheck
    
    ' 正しい場合
    If ctlDateSetting1.TotalArea <> -1 Then
        '確認ボタン押下用タイマを作動させる
        tmrKakunin.Interval = MN_MAIL_INTERVAL       'ボタン押下用タイマ時間設定
        tmrKakunin.Enabled = True
        
        strDate = ctlDateSetting1.TotalArea
        ctlDateSetting1.Enable = 1
        For i = 0 To cmdTenkey.Count - 1
            cmdTenkey(i).Enabled = False
        Next
        cmdKakutei.Enabled = False
        cmdModoru_Menu.Enabled = False

        ' 日時設定コントロールのトータル入力値エリアの値でシステム時刻を更新する。
        Date = Mid(strDate, 1, 4) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2)
        Time = Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2)

        ' 監視盤が動作していない場合はメール送信を行わない
        If CheckAppStart(PROC_KANRI) <> 0 Then

            ' メールの送信内容を編集する
            udtSendData.udtlHeader.dwId = ML_ID_DATE_SET_ORD       'メールＩＤ　=”"日時設定指示"
            udtSendData.udtlHeader.dwSize = MlSize.DATE_SET_ORD    'メールサイズ=”"日時設定指示"
            udtSendData.udtlHeader.dwProid = RHOSHU_ID             '送信元プロセスＩＤ=”保守”
            udtSendData.udtlHeader.dwSubArea = 0                   '補助情報　=　0

            ' 送信サイズを設定する。
            lngSendSize = udtSendData.udtlHeader.dwSize
                
            ' 監マに対して、日時設定指示メールを送信する。
            bRet = DssSendMail(MAIL_SLOT_KANMA, lngSendSize, udtSendData.udtlHeader)
            ' メールを正常に送信した時のログ
            If bRet = False Then
                '「画面表示要求メール送信異常」ログ出力
                lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, DATESETORDER_REQ_SEND, lngErrCode)
            Else
                '「画面表示要求メール送信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, DATESETORDER_REQ_SEND, 0)
            End If
        End If

        ' 保存エリアを再度初期化                ' rev 02.15
        gintCtrlIndex = -1

    ' 不正な場合
    Else

        iResponse = MsgBox("入力した値は不正です。", _
                           (vbOKOnly + vbExclamation), _
                           "入力異常")
    End If

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  関数名称     : cmdModoru_Menu_Click
'/  機能名称     : メニューに戻るボタン押下処理
'/  機能概要     : メニューに戻るボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20フェーズ２対応
'/             EG20統合監視盤USDM対応番号【Mainte_03_01】
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Menu_Click()

    On Error Resume Next
    
    '「システム設定メニュー画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_GAMEN_END, 0)
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 初期化メニュー画面(アクティブ時)
'//  機能概要  : 画面再表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//  ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//              EG20フェーズ２対応
'//              EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    tmrMail.Enabled = True         ' EG20 V6.8.0.1 ADD
    
    pfFormActive (hwnd)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 初期化メニュー画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//  ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//              EG20フェーズ２対応
'//              EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    'タイマを停止する。
    tmrKakunin.Enabled = False

    tmrMail.Enabled = False         ' EG20 V6.8.0.1 ADD
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  関数名称     : TmrKakunin_Timer
'/  機能名称     : 確認ボタン押下用タイマイベント時処理
'/  機能概要     : 確認ボタン押下用タイマイベント発生時の処理を行う。
'/                 確認ボタン、その他ボタンの色を押下色から元の色に戻す
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                 EG20フェーズ２対応
'/                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'/ REVISIONS     :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrKakunin_Timer()

    Dim i As Integer
    Dim blnRet As Boolean
    Dim intCount As Integer
    
    On Error Resume Next
    
    '確認ボタン押下用タイマを停止する
    tmrKakunin.Enabled = False                   '確認ボタン押下用タイマ停止
    tmrKakunin.Interval = 0                      '確認ボタン押下用時間初期化
        
    ctlDateSetting1.Enable = 0
        
    'テンキーを押下可にする
    For i = 0 To cmdTenkey.Count - 1
        cmdTenkey(i).Enabled = True
    Next
        
    '戻るボタンを押下可にする
    cmdModoru_Menu.Enabled = True
    
    '確定ボタンを押下可にする
    cmdKakutei.Enabled = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】

'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
'        AppActivate frmLogMenu.Caption, False          ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
'        pfFormActive (frmLogMenu.hwnd)                 ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
        AppActivate frmSystemDateSettei.Caption, False  ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
        pfFormActive (frmSystemDateSettei.hwnd)         ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If
End Sub

