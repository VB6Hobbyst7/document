VERSION 5.00
Begin VB.Form frmRebootTimeSettei 
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail2 
      Left            =   10440
      Top             =   7080
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   2
      Left            =   7680
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   1
      Left            =   9600
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "操作卓"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   11175
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FFFF&
         Caption         =   "切"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   9360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FFFF&
         Caption         =   "切"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   7560
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FFFF&
         Caption         =   "切"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   5760
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FF80&
         Caption         =   "入"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   4080
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   9
         Top             =   960
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FF80&
         Caption         =   "入"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2280
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   7
         Top             =   960
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FF80&
         Caption         =   "入"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   480
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   5
         Top             =   960
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   9120
         TabIndex        =   16
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   7320
         TabIndex        =   14
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5520
         TabIndex        =   12
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "統合監視盤"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   11175
      Begin VB.CommandButton cmdKakutei 
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   0
         Left            =   9240
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox ChkSet 
         BackColor       =   &H0080FF80&
         Caption         =   "入"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   4
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  システム設定    画面へ戻る"
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
   Begin VB.Timer tmrMail 
      Left            =   11400
      Top             =   7080
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "リブート設定"
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
      TabIndex        =   1
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmRebootTimeSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmRebootTimeSettei.frm
'//  パッケージ名：リブート設定画面
'//  概要        ：リブート設定画面
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-15  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS   ：(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-264】
'//  REVISIONS   ：(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   ：(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//  REVISIONS   ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考        ：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

Private Const REBOOTSW_ON_MESSAGE = "入"    ' 釦メッセージ：入状態
Private Const REBOOTSW_OFF_MESSAGE = "切"   ' 釦メッセージ：切状態
Private Const REBOOTSW_ON_COLOR = &H80FF80  ' 釦色：入状態
Private Const REBOOTSW_OFF_COLOR = &H80FFFF ' 釦色：切状態
Private Const REBOOTSW_ON_VALUE = 1         ' 釦状態：入状態
Private Const REBOOTSW_OFF_VALUE = 0        ' 釦状態：切状態

' DA設定内容
Private Const ID_KANSI_SET_RBOOT_SET = &H14  ' 監視装置設定ＩＤ：リブート設定
Private Const REBOOT_ON_DASTATUS = 1         ' 釦状態：入状態
Private Const REBOOT_OFF_DASTATUS = 0        ' 釦状態：切状態

Private Const HUTEI = 0                      ' 値不定


'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  関数名称     : Form_Load
'/  機能名称     : Form_Load時処理
'/  機能概要     : Form_Load時処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                 EG20フェーズ２対応
'/                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'/  REVISIONS    :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'/                 EG20フェーズ２対応【結合TR-264】
'/  REVISIONS    :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    Dim intLoop         As Integer          ' ループカウンタ
    Dim intStatus       As Integer          ' リブート設定状態
    Dim strSecName(5)   As String
    Dim strDefault      As String
    Dim lngRet          As Long
    Dim strRet          As String * 32
    Const lngBufSize = 32
    
    Dim strCorner1      As String           ' 文字列格納エリア1
    Dim strCorner2      As String           ' 文字列格納エリア2
    
    On Error Resume Next
    
    '「リブート設定画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
        
' EG20 V6.8.0.1 ADD START
    tmrMail2.Interval = MN_MAIL_INTERVAL
    tmrMail2.Enabled = False
' EG20 V6.8.0.1 ADD END
        
    ' /////////////////////////////////////////////////////////////////////////
    ' // 統合監視盤設定
    ' /////////////////////////////////////////////////////////////////////////
    
    ' /////////////////////////////////////////////////
    ' // 統合監視盤釦
    
    ' 現在の設定状態を取得
    intStatus = pfGetKansiArea_Sts(ID_KANSI_SET_RBOOT_SET)
    
    If intStatus = REBOOT_ON_DASTATUS Then
        ChkSet.Caption = REBOOTSW_ON_MESSAGE
        ChkSet.BackColor = REBOOTSW_ON_COLOR
        ChkSet.Value = REBOOTSW_ON_VALUE
    Else
        ChkSet.Caption = REBOOTSW_OFF_MESSAGE
        ChkSet.BackColor = REBOOTSW_OFF_COLOR
        ChkSet.Value = REBOOTSW_OFF_VALUE
    End If
    ChkSet.Visible = True
    ChkSet.Enabled = True
    
   
    ' /////////////////////////////////////////////////////////////////////////
    ' // 操作卓設定
    ' /////////////////////////////////////////////////////////////////////////
' EG20 V3.3.0.1【結合TR-264】削除開始
'    strDefault = ""
'    strSecName(0) = strAppName_station
'    strSecName(1) = strAppName_station2
'    strSecName(2) = strAppName_station3
'    strSecName(3) = strAppName_station4
'    strSecName(4) = strAppName_station5
'    strSecName(5) = strAppName_station6
' EG20 V3.3.0.1【結合TR-264】削除終了
    
    ' コーナ名称設定処理
    Call gsGetCornerName
    
    For intLoop = 0 To UBound(strSecName)
    
        '設定ありのコーナを活性にする
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
' EG20 V3.3.0.1【結合TR-264】削除開始
'            'Iniファイルからコーナー名を取得
'            lngRet = GetPrivateProfileString(strSecName(intLoop), IDU_PROFILE_KEY_NAME_CONER, _
'                                                strDefault, strRet, lngBufSize, KANSI_STATION_INI_FILE)
'            ' /////////////////////////////////////////////////
'            ' // ラベル（コーナー名称表示）
'            lblCornerName(intLoop).Caption = Trim(strRet)
' EG20 V3.3.0.1【結合TR-264】削除終了
' EG20 V3.3.0.1【結合TR-264】追加開始
            ' /////////////////////////////////////////////////
            ' // ラベル（コーナー名称表示）
            strCorner1 = MidB(gstrCornerName(intLoop), 1, 12)
            strCorner2 = MidB(gstrCornerName(intLoop), 13, 24)
            lblCornerName(intLoop).Caption = strCorner1 & vbCrLf & strCorner2
' EG20 V3.3.0.1【結合TR-264】追加終了
            lblCornerName(intLoop).Visible = True
        
            ' /////////////////////////////////////////////////
            ' // 釦
            
            ' 現在の設定状態を取得
            Call pfGetJikaiSts(intStatus, intLoop + 1)
            
            If intStatus = REBOOT_ON_DASTATUS Then
                ChkSetTaku(intLoop).Caption = REBOOTSW_ON_MESSAGE
                ChkSetTaku(intLoop).BackColor = REBOOTSW_ON_COLOR
                ChkSetTaku(intLoop).Value = REBOOTSW_ON_VALUE
            Else
                ChkSetTaku(intLoop).Caption = REBOOTSW_OFF_MESSAGE
                ChkSetTaku(intLoop).BackColor = REBOOTSW_OFF_COLOR
                ChkSetTaku(intLoop).Value = REBOOTSW_OFF_VALUE
            End If
            ChkSetTaku(intLoop).Visible = True
            ChkSetTaku(intLoop).Enabled = True

        Else
            lblCornerName(intLoop).Caption = Trim(strRet)
            lblCornerName(intLoop).Visible = False
            ChkSetTaku(intLoop).Enabled = False
            ChkSetTaku(intLoop).Visible = False
            ChkSetTaku(intLoop).Value = REBOOTSW_OFF_VALUE
        End If
    
    Next intLoop
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // その他コントロール設定
    ' /////////////////////////////////////////////////////////////////////////
    
    cmdKakutei(0).Enabled = False       ' 統合監視盤「確定」
    cmdKakutei(1).Enabled = False       ' 操作卓「確定」
    cmdKakutei(2).Enabled = False       ' 「確定」        ：押下不可 EG20 V5.13.0.1 ADD
    cmdReturn.Enabled = True            ' 「戻る」
    
End Sub


'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : cmdKakutei_Click
'// 機能名称    : 確定ボタン押下時処理
'// 機能概要    : 確定ボタンが押下された処理を行う
'//
'//                   型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応
'//               EG20統合監視盤USDM対応番号【Mainte_03_01】
'// REVISIONS :(EG20 V5.13.0.1) 2012-06-07  CODED BY  [TCC] H.Sano
'/             EG20確認釦1個対応
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdKakutei_Click(Index As Integer)
    
    Dim intLoop     As Integer          ' ループカウンタ
    Dim udtSendData As ML_REBOOT_REQ    ' リブート設定情報要求
    Dim lngSendSize As Long             ' 送信するメールサイズ
    Dim lngErrCode  As Long             ' エラーコード
    Dim bRet        As Boolean          ' メール送信処理戻り値
    Dim intLoopMail As Integer          ' ループカウンタ2 EG20 V5.13.0.1 ADD
    
    On Error Resume Next
    
    '「リブート設定画面：確定釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_KAKUTEI_BUTTOM, 0)

    ' コントロール設定
    cmdKakutei(0).Enabled = False       ' 統合監視盤「確定」    ：押下不可
    cmdKakutei(1).Enabled = False       ' 操作卓「確定」        ：押下不可
    cmdKakutei(2).Enabled = False       ' 「確定」              ：押下不可 EG20 V5.13.0.1 ADD
    cmdReturn.Enabled = False           ' 「戻る」              ：押下不可
    
    ChkSet.Enabled = False              ' 統合監視盤「入／切」  ：押下不可

    For intLoop = 0 To CONECT_CORNER_MAXINDEX   ' 操作卓「入／切」      ：押下不可
        '設定ありのコーナを活性にする
        If ChkSetTaku(intLoop).Visible = True Then
            ChkSetTaku(intLoop).Enabled = False
        End If
    Next intLoop

    '確認ボタン押下用タイマを作動させる
    tmrMail.Interval = MN_MAIL_INTERVAL ' ボタン押下用タイマ時間設定
    tmrMail.Enabled = True              ' タイマ作動

'EG20 V5.13.0.1 ADD START
For intLoopMail = 0 To 1
'EG20 V5.13.0.1 ADD END

    ' メールの送信内容を編集する
    udtSendData.udtlHeader.dwId = ML_ID_REBOOT_REQ      ' メールＩＤ　=”"設定情報要求（リブート）"
    udtSendData.udtlHeader.dwSize = MlSize.REBOOT_REQ   ' メールサイズ=”"設定情報要求"
    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' 送信元プロセスＩＤ=”保守”
    udtSendData.udtlHeader.dwSubArea = 0                ' 補助情報　=　0
                                                        ' ※制御種別は押下した「確定」釦に対応
'EG20 V5.13.0.1 MOD START
'    udtSendData.dwControl = Index                       ' 制御種別（0:統合監視盤,1:操作卓）
    udtSendData.dwControl = intLoopMail                  ' 制御種別（0:統合監視盤,1:操作卓）
'EG20 V5.13.0.1 MOD END
                                                        ' ※入／切設定は釦に対応
    udtSendData.dwKanshi = ChkSet.Value                 ' 統合監視盤設定（0:切,1:入）
    For intLoop = 0 To CONECT_CORNER_MAXINDEX                    ' 操作卓設定（0:切,1:入）
        '設定ありのコーナを活性にする
        udtSendData.dwTaku(intLoop) = ChkSetTaku(intLoop).Value
    Next intLoop
    
    ' 送信サイズを設定する。
    lngSendSize = udtSendData.udtlHeader.dwSize
            
    ' 監視盤起動チェック
    If CheckAppStart(PROC_KANRI) = 0 Then
        ' /////////////////////////////////////////////////
        ' // 監視盤未起動：自力で設定値を更新
        If udtSendData.dwControl = 0 Then
            ' ////////////////////////////////////////////
            ' // 統合監視盤
            bRet = gspfSetKansiSts(ID_KANSI_SET_RBOOT_SET, ChkSet.Value)
        Else
            ' ////////////////////////////////////////////
            ' // 操作卓
            For intLoop = 0 To CONECT_CORNER_MAXINDEX
                If ChkSetTaku(intLoop).Visible = True Then
                    bRet = pfSetJikaiSts(ChkSetTaku(intLoop).Value, intLoop + 1, IdGate.ID_GATE_SET_RBOOT_SET)
                End If
            Next intLoop
        End If
    Else
        ' /////////////////////////////////////////////////
        ' // 監視盤起動中：自力で設定値を更新
        
        ' 監マに対して、設定情報要求メールを送信する。
        bRet = DssSendMail(MAIL_SLOT_KANMA, lngSendSize, udtSendData.udtlHeader)
        ' メールを正常に送信した時のログ
        If bRet = False Then
            '「設定情報要求メール送信異常」ログ出力
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
            Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
        Else
            '「設定情報要求メール送信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        End If
    End If

'EG20 V5.13.0.1 ADD START
Next intLoopMail
'EG20 V5.13.0.1 ADD END

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : cmdReturn_Click
'// 機能名称    : メニューに戻るボタン押下処理
'// 機能概要    : メニューに戻るボタン押下処理を行う
'//
'//                  型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応
'//               EG20統合監視盤USDM対応番号【Mainte_03_01】
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    On Error Resume Next

    '「リブート設定画面：戻る釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_GAMEN_END, 0)

    '画面のUnload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : Form_Activate
'// 機能名称    : Form_Activate時処理
'// 機能概要    : Form_Activate時処理を行う
'//
'//                   型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応
'//               EG20統合監視盤USDM対応番号【Mainte_03_01】
'// REVISIONS   :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    tmrMail2.Enabled = True             ' EG20 V6.8.0.1 ADD
    
    pfFormActive (hwnd)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : Form_Deactivate
'// 機能名称    : ディアクティブ時
'// 機能概要    : メール受信用のタイマ停止
'//
'//                   型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応
'//               EG20統合監視盤USDM対応番号【Mainte_03_01】
'// REVISIONS   :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   
   On Error Resume Next
    
    'タイマを停止する。
    tmrMail.Enabled = False
    
    tmrMail2.Enabled = False             ' EG20 V6.8.0.1 ADD

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : tmrMail_Timer
'// 機能名称    : タイムアウト処理
'// 機能概要    : タイマタイムアウト処理を行う
'//
'//                   型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応
'//               EG20統合監視盤USDM対応番号【Mainte_03_01】
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim intLoop     As Integer          ' ループカウンタ

    'タイマを停止
    tmrMail.Enabled = False

    ' コントロール設定
    cmdKakutei(0).Enabled = False       ' 統合監視盤「確定」    ：押下不可
    cmdKakutei(1).Enabled = False       ' 操作卓「確定」        ：押下不可
    cmdKakutei(2).Enabled = False       ' 「確定」        ：押下不可 EG20 V5.13.0.1 ADD
    cmdReturn.Enabled = True            ' 「戻る」              ：押下可能
    
    ChkSet.Enabled = True               ' 統合監視盤「入／切」  ：押下可能

    For intLoop = 0 To CONECT_CORNER_MAXINDEX    ' 操作卓「入／切」      ：押下可能
        '設定ありのコーナを活性にする
        If ChkSetTaku(intLoop).Visible = True Then
            ChkSetTaku(intLoop).Enabled = True
        End If
    Next intLoop
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : ChkSet_Click
'// 機能名称    : 統合監視盤「入／切」釦押下
'// 機能概要    : 統合監視盤「入／切」釦押下処理を行う
'//
'//                   型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応
'//               EG20統合監視盤USDM対応番号【Mainte_03_01】
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub ChkSet_Click()
    
    On Error Resume Next
    
   
    '「リブート設定画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_KANSHI_BUTTOM, 0)

    If ChkSet.Value = REBOOTSW_ON_VALUE Then
        ' /////////////////////////////////////////////////
        ' 切→入設定
        ChkSet.Caption = REBOOTSW_ON_MESSAGE
        ChkSet.BackColor = REBOOTSW_ON_COLOR
    Else
        ' /////////////////////////////////////////////////
        ' 入→切設定
        ChkSet.Caption = REBOOTSW_OFF_MESSAGE
        ChkSet.BackColor = REBOOTSW_OFF_COLOR
    End If

    cmdKakutei(0).Enabled = True       ' 統合監視盤「確定」
    cmdKakutei(2).Enabled = True       ' 「確定」        ： EG20 V5.13.0.1 ADD

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : ChkSetTaku_Click
'// 機能名称    : 操作卓「入／切」釦押下
'// 機能概要    : 操作卓「入／切」釦押下処理を行う
'//
'//                   型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20フェーズ２対応
'//               EG20統合監視盤USDM対応番号【Mainte_03_01】
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub ChkSetTaku_Click(Index As Integer)
    
    On Error Resume Next
    
   
    '「リブート設定画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_TAKU_BUTTOM, 0)

    If ChkSetTaku(Index).Value = REBOOTSW_ON_VALUE Then
        ' /////////////////////////////////////////////////
        ' 切→入設定
        ChkSetTaku(Index).Caption = REBOOTSW_ON_MESSAGE
        ChkSetTaku(Index).BackColor = REBOOTSW_ON_COLOR
    Else
        ' /////////////////////////////////////////////////
        ' 入→切設定
        ChkSetTaku(Index).Caption = REBOOTSW_OFF_MESSAGE
        ChkSetTaku(Index).BackColor = REBOOTSW_OFF_COLOR
    End If

    cmdKakutei(1).Enabled = True       ' 操作卓「確定」
    cmdKakutei(2).Enabled = True       ' 操作卓「確定」        ：EG20 V5.13.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetKansiArea_Sts
'//  機能名称  : 監視設定画面(ロード時)。
'//  機能概要  : 監視設定画面の初期処理を行う。
'//
'//              型        名称          意味
'//  引数      : Integer  intId          [IN]エリアID
'//
'//              型        値       　　 意味
'//  戻り値    : Integer                 [OUT]現在値
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetKansiArea_Sts(intId As Integer) As Integer
    
    Dim iAreaSts     As Integer     '監視設定状態値
    Dim lSts         As Long        '関数戻り値
    Dim udtAreaR255  As GATE_INFO   '読込み用エリア（255設定用）
    Dim lngSts       As Long
    Dim lngLoop1     As Long
    Dim lngHandle    As Long
    Dim FileName     As String
    Dim lngRet       As Long
    Dim bRet         As Boolean
    Dim sSetteiFile  As String      'ファイルパス
    Dim lngAplSts    As Long        'アプリ起動/未起動結果
            
    On Error Resume Next
      
    '監視盤起動有無チェック
    lngAplSts = CheckAppStart(PROC_KANRI)
    If lngAplSts = 0 Then
        '監視盤未起動時
        '監視設定ファイルをオープン
        lngHandle = CreateFile(K_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)  'V1.4.0.1　ADD
        
        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
           'オープン異常時:異常
           '「監視設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           pfGetKansiArea_Sts = HUTEI
           Exit Function
        End If
        
        '監視設定ファイル読み込み
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           '読み込み異常時：異常
           pfGetKansiArea_Sts = HUTEI
         '「監視設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           'ハンドルのクローズ
           Call CloseHandle(lngHandle)
           Exit Function
        End If
        
        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
        lngSts = KansiSerchId(udtAreaR255, CLng(intId))
        If lngSts >= 0 Then
           'IDが有った場合
           pfGetKansiArea_Sts = udtAreaR255.GateInfo(lngSts).bytDATA(0)
         Else
          ' 該当ＩＤ無しの場合:異常
          '「監視設定画面：エリア・ファイル参照異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          pfGetKansiArea_Sts = HUTEI
          Exit Function
        End If
    Else
        '監視盤起動時
        Set Idinf_KansiSettei = New IdInfProc              '監視装置設定エリア
        '共有エリアオープン
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei    '監視装置設定エリア
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
           pfGetKansiArea_Sts = HUTEI
           '「監視設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
           Exit Function
        End If
        
        '監視設定エリアをＬＯＣＫする。
        Idinf_KansiSettei.IdLock
        If Idinf_KansiSettei.Errsts <> 0 Then
          'データ参照異常時:異常
          pfGetKansiArea_Sts = HUTEI
          Idinf_KansiSettei.IdFree
          '「監視設定画面：エリア・ファイル参照異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
          Exit Function
        End If
    
        '監視設定エリアIDを設定
        Idinf_KansiSettei.id = intId
        Idinf_KansiSettei.IdGet
        If Idinf_KansiSettei.Errsts <> 0 Then
          'データ参照異常時はブランク表示設定を行う。
          pfGetKansiArea_Sts = HUTEI
          Idinf_KansiSettei.IdFree
          '「監視設定画面：エリア・ファイル参照異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
          Exit Function
        End If

        pfGetKansiArea_Sts = Idinf_KansiSettei.DataArea(0)   '設定内容
      
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
   End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : KansiSerchId
'//  機能名称  : ＩＤ検索処理
'//  機能概要  : ＩＤ検索を行う。
'//
'//              型        名称        意味
'//  引数      : GATE_INFO udtArea255 [IN]変換元データ
'//　　　　　　　Long　　　lngId　　　[IN]エリアID
'//
'//              型        値        意味
'//  戻り値    : Long　　　         　[OUT]　0以上：正常。-1以下：エラー
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-15   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function KansiSerchId(udtArea255 As GATE_INFO, lngID As Long) As Long

    Dim lngIndex As Long                '検索用インデックス
    Dim lngMin As Long                  '最小インデックス
    Dim lngMax As Long                  '最大インデックス
    Dim lngChkIndex As Long             '該当インデックス
    Dim lngWorkId   As Long             '標準ＩＤ

    On Error Resume Next
    
    '初期化
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '検索開始
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             'ＩＤ取り出し
        If lngID = lngWorkId Then                                  '同じ？
            lngChkIndex = lngIndex                                  'データ取り出し後、検索終了
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngID < lngID) Then         'データが予備か小さい
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    KansiSerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetJikaiSts
'//  機能名称  : 自改タブ表示処理(監視盤起動有無対応参照)
'//  機能概要  : 自改タブの号機別釦状態取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-18   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応(監視盤未起動時でも設定変更可)
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetJikaiSts(iJikaiSts As Integer, iGouki As Integer)
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255     As GATE_INFO        '読込み用エリア（255設定用）
    Dim lngSts          As Long             'ヒットエリアID
    Dim lngLoop1        As Long             'カウンター
    Dim lngHandle       As Long             'ハンドル
    Dim FileName        As String           'ファイル有無チェック
    Dim lngRet          As Long             '戻り値
    Dim bRet            As Boolean          '読み込み結果戻り値
    Dim sSetteiFile     As String           'ファイルパス　'V1.4.0.1　ADD
    
    On Error Resume Next
'V1.4.0.1 DEL START
'    '自改設定ファイル有無
'    FileName = Dir(G_SETTEI_FILE)
'    If FileName = "" Then
'       '無ければ参照不可のため参照異常
'       iJikaiSts = GET_CONECTSTS_ERROR
'       Exit Function
'    End If
'V1.4.0.1 DEL END
'V1.4.0.1 ADD START
   '自改設定ファイル有無
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '自改設定ファイルがない場合
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '自改設定ファイルがある場合
       sSetteiFile = G_SETTEI_FILE
    End If
'V1.4.0.1 ADD END

    '監視盤起動有無チェック
    If CheckAppStart(PROC_KANRI) = 0 Then
        
        '自改設定ファイルをオープン
'        lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 DEL
        lngHandle = CreateFile(sSetteiFile, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
           'オープン異常時は参照不可のため参照異常
'           iJikaiSts = GET_CONECTSTS_ERROR                                 ' EG20 V2.1.0.1[Mainte_03_01]削除
           iJikaiSts = REBOOTSW_OFF_VALUE                                   ' EG20 V2.1.0.1[Mainte_03_01]追加
           Exit Function
        End If
        
        '自改設定ファイル読み込み
        For lngLoop1 = 0 To iGouki - 1
            bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        Next
        
        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
'        lngSts = SerchId(udtAreaR255, IdGate.JIKAI_CONECT_SETTEI)          ' EG20 V2.1.0.1[Mainte_03_01]削除
        lngSts = SerchId(udtAreaR255, IdGate.ID_GATE_SET_RBOOT_SET)         ' EG20 V2.1.0.1[Mainte_03_01]追加
        If lngSts >= 0 Then
           'IDが有った場合
           iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
        Else
          ' 該当ＩＤ無しの場合参照異常
'          iJikaiSts = GET_CONECTSTS_ERROR                                  ' EG20 V2.1.0.1[Mainte_03_01]削除
          iJikaiSts = REBOOTSW_OFF_VALUE                                    ' EG20 V2.1.0.1[Mainte_03_01]追加
          Exit Function
        End If
        
' EG20 V2.1.0.1[Mainte_03_01]削除開始
'        Select Case iAreaSts
'           Case 1
'             '接続
'              iJikaiSts = CONECTSTS_ERROR
'              Exit Function
'           Case 0
'              iJikaiSts = CONECTSTS_END
'              Exit Function
'        End Select
' EG20 V2.1.0.1[Mainte_03_01]削除終了
' EG20 V2.1.0.1[Mainte_03_01]追加開始
        iJikaiSts = iAreaSts
        Exit Function
' EG20 V2.1.0.1[Mainte_03_01]追加終了
    
    Else
     
         Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
         '自改設定エリアをオープンする。
          Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
          Idinf_JikaiSettei.IdOpen
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
'             iJikaiSts = GET_CONECTSTS_ERROR              ' EG20 V2.1.0.1[Mainte_03_01]削除
             iJikaiSts = REBOOTSW_OFF_VALUE                ' EG20 V2.1.0.1[Mainte_03_01]追加
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
          End If
             
          '自改設定エリアをＬＯＣＫする。
          Idinf_JikaiSettei.IdLock
          If Idinf_JikaiSettei.Errsts <> 0 Then
             'データ参照異常時はブランク表示設定を行う。
'             iJikaiSts = GET_CONECTSTS_ERROR              ' EG20 V2.1.0.1[Mainte_03_01]削除
             iJikaiSts = REBOOTSW_OFF_VALUE                ' EG20 V2.1.0.1[Mainte_03_01]追加
             Idinf_JikaiSettei.IdFree
             Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
             Exit Function
           End If
              
           'エリアの内容を読み込む。
'            Idinf_JikaiSettei.id = IdGate.JIKAI_CONECT_SETTEI              ' EG20 V2.1.0.1[Mainte_03_01]削除
            Idinf_JikaiSettei.id = IdGate.ID_GATE_SET_RBOOT_SET             ' EG20 V2.1.0.1[Mainte_03_01]追加
            Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
            If Idinf_JikaiSettei.Errsts <> 0 Then
               'データ参照異常時はブランク表示設定を行う。
'                iJikaiSts = GET_CONECTSTS_ERROR              ' EG20 V2.1.0.1[Mainte_03_01]削除
                iJikaiSts = REBOOTSW_OFF_VALUE                ' EG20 V2.1.0.1[Mainte_03_01]追加
                Idinf_JikaiSettei.IdFree
                Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
                Exit Function
            End If
               
            '設定内容を取得
             iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
' EG20 V2.1.0.1[Mainte_03_01]削除開始
'             Select Case iAreaSts
'                 Case 1
'                  '接続
'                   iJikaiSts = CONECTSTS_ERROR
'                   Idinf_JikaiSettei.IdFree
'                   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
'                   Exit Function
'                 Case 0
'                   iJikaiSts = CONECTSTS_END
'                   Idinf_JikaiSettei.IdFree
'                   Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
'                   Exit Function
'             End Select
' EG20 V2.1.0.1[Mainte_03_01]削除終了
        iJikaiSts = iAreaSts                            ' EG20 V2.1.0.1[Mainte_03_01]追加
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '自改装置設定データファイル
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SerchId
'//  機能名称  : ＩＤ検索処理(全タブ専用)
'//  機能概要  : ＩＤ検索を行う。
'//
'//              型        名称        意味
'//  引数      : GATE_INFO udtArea255 [IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : Long　　　         　[OUT]　0以上：正常。-1以下：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SerchId(udtArea255 As GATE_INFO, lngID As Long) As Long

    Dim lngIndex As Long                '検索用インデックス
    Dim lngMin As Long                  '最小インデックス
    Dim lngMax As Long                  '最大インデックス
    Dim lngChkIndex As Long             '該当インデックス
    Dim lngWorkId   As Long             '標準ＩＤ

    On Error Resume Next
    
    '初期化
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '検索開始
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             'ＩＤ取り出し
        If lngID = lngWorkId Then                                  '同じ？
            lngChkIndex = lngIndex                                  'データ取り出し後、検索終了
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngID < lngID) Then         'データが予備か小さい
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    SerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ChgData
'//  機能名称  : データ変換処理処理
'//  機能概要  : データ変換処理処理を行う。
'//
'//              型        名称        意味
'//  引数      : ID_FMT 　DataArea 　[IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : String　　　        [OUT]　vbNullstring以外：正常。vbNullString    ：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function ChgData(DataArea As ID_FMT) As String

    Dim lngloop As Long
    Dim lngWork As Long
    Dim lngErrsts As Long

    On Error GoTo ChgDataErr
    
    lngErrsts = IdInfErr.OK
    
    Select Case DataArea.intType
    Case ID_TYPE.Flag   '状態
        If (DataArea.bytDATA(0) <> 255) Then
            ChgData = str$(DataArea.bytDATA(0))
            
        Else
            ChgData = "-1"                      '値が不定ならー１セット
            
        End If
            
    Case ID_TYPE.Count  '回数
        lngWork = 0                              '初期化
        For lngloop = 3 To 0 Step -1
            lngWork = lngWork * 256 + DataArea.bytDATA(lngloop)
        Next lngloop
                        
        ChgData = str$(lngWork)
    
    Case ID_TYPE.Date_Type, ID_TYPE.time_type '日付、時刻
        ChgData = StrConv(DataArea.bytDATA, vbUnicode)
        
    Case Else
        ChgData = vbNullString
        lngErrsts = IdInfErr.ID_TYPE_MISS
        Exit Function

    End Select
    
    Exit Function
    
ChgDataErr:
        ChgData = vbNullString
        lngErrsts = IdInfErr.PROC_ERR
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSetJikaiSts
'//  機能名称  : 自改設定ファイル更新処理
'//  機能概要  : 自改設定ファイル更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [IN]接続・切断タイプ
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-26   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSetJikaiSts(iJikaiSts As Integer, iGouki As Integer, iJikaiID As Long) As Boolean

    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255     As GATE_INFO        '読込み用エリア（255設定用）
    Dim lngSts          As Long             'ヒットエリアID
    Dim lngLoop1        As Long             'カウンター
    Dim lngHandle       As Long             'ハンドル
    Dim FileName        As String           'ファイル有無チェック
    Dim lngRet          As Long             '戻り値
    Dim bRet            As Boolean          '読み込み結果戻り値
    Dim sSetteiFile     As String
    Dim udtAreaR255Work As GATE_INFO        '読込み用エリア（ポインタ移動用）
    
    On Error Resume Next
    
    '自改設定ファイル有無
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '自改設定ファイルがない場合
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '自改設定ファイルがある場合
       sSetteiFile = G_SETTEI_FILE
    End If
        
    '自改設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため更新異常
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        pfSetJikaiSts = False
        Exit Function
    End If
        
    '自改設定ファイル読み込み
    For lngLoop1 = 0 To iGouki - 1
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           'ハンドルのクローズ
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
           Call CloseHandle(lngHandle)
           pfSetJikaiSts = False
           Exit Function
        End If
    Next
    
    'ハンドルのクローズ
    Call CloseHandle(lngHandle)
    
    'ID検索
    lngSts = SerchId(udtAreaR255, iJikaiID)
    If lngSts >= 0 Then
       'IDが有った場合
       SetChgData udtAreaR255.GateInfo(lngSts), iJikaiSts   'データ設定
    Else
       ' 該当ＩＤ無しの場合更新異常
        pfSetJikaiSts = False
       Exit Function
    End If
      
    '自改設定ファイルをオープン
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため更新異常
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        pfSetJikaiSts = False
        Exit Function
    End If
     
    'ファイルポインタ移動のための読み込み
     For lngLoop1 = 1 To iGouki - 1
         bRet = ReadFile(lngHandle, udtAreaR255Work, LenB(udtAreaR255Work), lngRet, 0)
         If bRet = False Then
            'ハンドルのクローズ
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
            Call CloseHandle(lngHandle)
            pfSetJikaiSts = False
            Exit Function
         End If
     Next
    
    '自改設定ファイルに書き込む
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       'ハンドルのクローズ
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       Call CloseHandle(lngHandle)
       pfSetJikaiSts = False
       Exit Function
    End If
    
    'ハンドルのクローズ
     Call CloseHandle(lngHandle)

     pfSetJikaiSts = True
     
     Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_SETTEIFILE_UPDATA_OK, 0)
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetChgData
'//  機能名称  : データ変換処理処理
'//  機能概要  : データ変換処理処理を行う。
'//
'//              型        名称        意味
'//  引数      : ID_FMT 　DataArea 　[IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : String　　　        [OUT]　vbNullstring以外：正常。vbNullString    ：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SetChgData(DataArea As ID_FMT, iSts As Integer)
   
   On Error Resume Next

   DataArea.bytDATA(0) = iSts
  
End Function

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// 関数名称    : tmrMail2_Timer
'// 機能名称    : タイムアウト処理
'// 機能概要    : タイマタイムアウト処理を行う
'//
'//                   型          名称            意味
'// 引数        :
'// 戻り値      :
'//
'// ORIGINAL    :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'// REVISIONS   :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'// 備考        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
'        AppActivate frmSystemSetteiMenu.Caption, False ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
'        pfFormActive (frmSystemSetteiMenu.hwnd)        ' EG20 V8.1.0.1【EG20_KANSI05_01】DEL
        AppActivate frmRebootTimeSettei.Caption, False  ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
        pfFormActive (frmRebootTimeSettei.hwnd)         ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If

End Sub
