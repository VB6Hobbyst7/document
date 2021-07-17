VERSION 5.00
Begin VB.Form frmOriTest 
   BorderStyle     =   0  'なし
   Caption         =   "折り返しテスト画面"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   14.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   8760
      Top             =   7800
   End
   Begin VB.Frame fraResource 
      Caption         =   "テスト先指定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9600
      TabIndex        =   3
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton optSyubetu 
         Caption         =   "交調"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "締切"
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
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdTestStart 
      Caption         =   "テスト開始"
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
      Left            =   9600
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "     メニュー        画面へ戻る"
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
      Left            =   9600
      TabIndex        =   2
      Top             =   7320
      Width           =   2175
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
      Height          =   7260
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "折り返しテスト"
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
      Height          =   403
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "ステータス"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "サーバー名"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "日時"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmOriTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'*
'*   ﾓｼﾞｭｰﾙ概要  : 折返しテスト画面のフォームモジュール
'*               :（集計および交調に対する状況確認画面）
'*
'*     ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'*              ・KK対応
'*     REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Option Explicit
'リソース定数

'選択中リソース種別 =0：集計、=1:交調
Dim iSelResource As Integer


Private Const MN_MAIL_INTERVAL = 1000 'メイルタイマのインターバル値

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  概要     : 折返しテスト結果表示
'  説明     : 折返しテストの結果を表示する文言を作成する。
'  ﾊﾟﾗﾒｰﾀ   : strMsg, I ,string, ：テスト結果表示文言
'           :  戻り値,O ,string, ：リストボックス表示文言
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              ・KK対応
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Function fMakeListbox(strMsg As String) As String
    Dim strRet As String
    Dim strServer As String
    
    On Error Resume Next
    
    strRet = vbNullString
    
    'システム時刻取得
    strRet = Format(Now, "YYYY/MM/DD HH:MM:SS")
    
    'サーバータイプ取得
    If (iSelResource = 0) Then
        strServer = "締切"
    Else
        strServer = "交調"
    End If

    'リストボックスに(日時 サーバー名 異常終了)を記載する
    strRet = strRet & Space(5) & strServer & Space(13) & strMsg

    fMakeListbox = strRet

End Function

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2001 All Right Reserved
'
'  概要     : 「保守画面に戻る」ボタン押下時のイベントプロシージャ
'  説明     : 折返しテスト結果表示画面を閉じる。
'  ﾊﾟﾗﾒｰﾀ   :
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              ・KK対応
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdReturn_Click()

    On Error Resume Next

    'ログを記載し、現画面を消去する
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ORI_TEST_GAMEN_END, 0)
    frmOriTest.ZOrder
    Unload Me

End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2001 All Right Reserved
'
'   概要    : テスト開始ボタン押下時のイベントプロシージャ
'   説明    : 折返しテスト開始要求を監マに送信する。
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              ・KK対応
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdTestStart_Click()
    
    Dim udtMail As MAIL_ORI_TEST        '送信メール
    Dim strServer As String             'サーバー名格納
    
    Dim bRet As Boolean
    Dim strRet As String
    
    On Error Resume Next
    
    'テスト開始釦押下ログ記載
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ORI_TEST_TEST_START_BUTTOM, 0)
    
    'ラジオ釦からサーバタイプ取得
    If optSyubetu(0) = True Then
        '締切
        iSelResource = 0
    Else
        '交調
        iSelResource = 1
    End If
    
    'テスト開始をリストに表示する。
    lstKan.AddItem fMakeListbox("開始")
    
    '監マに対して折返しテスト開始要求を送信する。
    udtMail.mlHeader.dwId = ML_ID_ORI_TEST_REQ
    udtMail.mlHeader.dwSize = MlSize.ORI_TEST_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwSvrType = iSelResource
    bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMail), udtMail.mlHeader)
    
    'メール送信処理チェック
    If bRet = False Then
        '送信失敗時
        
        'メール送信失敗ログ記載
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, ORI_TEST_TEST_MAIL_SEND_ERR, 0)

        '表示文章作成
        strRet = fMakeListbox("異常終了")
        
        '文章表示
        lstKan.AddItem strRet
    Else
        '送信成功時
        
        '釦制御を行う。
        cmdTestStart.Enabled = False
        optSyubetu(0).Enabled = False
        optSyubetu(1).Enabled = False
        cmdReturn.Enabled = False

    End If
       
    Exit Sub
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  概要      : 折返しテスト結果表示画面がアクティブになった時のイベントプロシージャ
'  説明      : メイル受信用のタイマを起動する。
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              ・KK対応
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Activate()
    'メール受信用のタイマを起動する。
    tmrMail.Enabled = True
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  概要     : 折返しテスト結果表示画面がﾃﾞｨｱｸﾃｨﾌﾞになった時のｲﾍﾞﾝﾄﾌﾟﾛｼｰｼﾞｬ
'  説明     : メール受信用のタイマを止める。
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              ・KK対応
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Deactivate()
    'メール受信用のタイマを止める。
    tmrMail.Enabled = False
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  概要     : 折返しテスト結果表示画面面がロードされた時のｲﾍﾞﾝﾄﾌﾟﾛｼｰｼﾞｬ
'  説明     : 初期処理を行う
'  ﾊﾟﾗﾒｰﾀ   :
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              ・KK対応
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Load()
    
    Dim iKansiAplChk As Integer
    
    On Error Resume Next

    '折り返しテスト画面表示ログ記載
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ORI_TEST_GAMEN_START, 0)

    'テスト結果表示用のリストボックスをクリアする。
    lstKan.Clear
        
    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

    '画面サイズ
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '監視盤起動/未起動チェックを行う。
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '監視盤起動時
        '処理の釦を押せるようにする
        cmdTestStart.Enabled = True
        optSyubetu(0).Enabled = True
        optSyubetu(1).Enabled = True
    Else
        '監視未起動時
        '処理の釦を押せなくする
        cmdTestStart.Enabled = False
        optSyubetu(0).Enabled = False
        optSyubetu(1).Enabled = False
    End If

End Sub


'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2001 All Right Reserved
'
'  概要     : メール受信用タイマがタイムアップした時のイベントプロシージャ
'  説明     : 受信メールの内容に基づき処理をする。
'  ﾊﾟﾗﾒｰﾀ   :
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              ・KK対応
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub tmrMail_Timer()
    Dim lLen As Long                    'メイルサイズ
    Dim bRet As Boolean                 '戻り値
    Dim udtReadMail As ML_KYOTU_INF

    On Error Resume Next

    'メールが届いているか確認する。
    lLen = DssMailRead(plMSlot_MN, udtReadMail)
    
    '受信したメールがサイズ0じゃなければ解析する
    If lLen <> 0 Then
        
        Select Case udtReadMail.udtlHeader.dwId   'メールＩＤ
        
        '「プロセス終了指示」を受信した場合
        Case ML_ID_PROEND_ORD
                    
            '強制終了処理を行う
            pfAbortProc

        '「折返しテスト完了通知」を受信した場合
        Case ML_ID_ORI_TEST_INF
            
            '結果内容に基づき処理を行う。
            Select Case udtReadMail.lngData(1)
                Case 0
                    'テスト正常終了を表示する。
                    lstKan.AddItem fMakeListbox("正常終了")
                Case 1
                    'テスト異常終了を表示する。
                    lstKan.AddItem fMakeListbox("異常終了")
                Case Else
                    'テスト実行不可能を表示する。
                    lstKan.AddItem fMakeListbox("実行不可能")
            End Select
            
            ' ボタンを押下可能にする。
            cmdTestStart.Enabled = True
            ' ラジオボタンを押下可能にする。
            optSyubetu(0).Enabled = True
            optSyubetu(1).Enabled = True
            ' 保守画面へ戻る釦を押下不可にする。
            cmdReturn.Enabled = True
            

        '保守画面アクティブ表示の場合
        Case ML_ID_HOSHU_ACTIVE_REQ

            '折返しテスト画面をアクティブにする。
            AppActivate frmOriTest.Caption, False

        '(メールＩＤ不正）
        Case Else
        End Select
    End If
    
End Sub

