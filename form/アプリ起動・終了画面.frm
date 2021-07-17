VERSION 5.00
Begin VB.Form frmAppConfig 
   BorderStyle     =   0  'なし
   Caption         =   "一括起動・終了"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrLogTimer 
      Left            =   3600
      Top             =   8040
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   1080
      Top             =   7920
   End
   Begin VB.Timer tmrMail 
      Left            =   480
      Top             =   7920
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   9793
      TabIndex        =   13
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "シャットダウン・リブート"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   242
      TabIndex        =   8
      Top             =   4440
      Width           =   11535
      Begin VB.CommandButton cmdShoutDown 
         Caption         =   "シャットダウン"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   480
         TabIndex        =   10
         Top             =   330
         Width           =   2145
      End
      Begin VB.CommandButton cmdReboot 
         Caption         =   "リブート"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   480
         TabIndex        =   9
         Top             =   1030
         Width           =   2145
      End
      Begin VB.Label lblAllEndApl 
         Caption         =   "アプリ起動中の場合は全てのアプリケーションを終了し、再起動する。"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   12
         Top             =   1200
         Width           =   7215
      End
      Begin VB.Label lblAllEndApl 
         Caption         =   "アプリ起動中の場合全てのアプリケーションを終了し、電源を切る。"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   7215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "起動・終了指定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   242
      TabIndex        =   3
      Top             =   600
      Width           =   11535
      Begin VB.Frame Frame2 
         Caption         =   "アプリケーション終了"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   10815
         Begin VB.CommandButton cmdAppEnd 
            Caption         =   "アプリ終了"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   19
            Top             =   330
            Width           =   2145
         End
         Begin VB.CommandButton cmdAppAllEnd 
            Caption         =   "アプリ完全終了"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   18
            Top             =   1030
            Width           =   2145
         End
         Begin VB.Label lblAllEndApl 
            Caption         =   "全てのアプリケーションを、統合監視盤の保守のみ残して終了する。"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   21
            Top             =   480
            Width           =   7215
         End
         Begin VB.Label lblAllEndApl 
            Caption         =   "全てのアプリケーションを終了し、Windowsのみの状態にする。"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   20
            Top             =   1200
            Width           =   7215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "アプリケーション起動"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   939
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   10815
         Begin VB.CommandButton cmdAppStart 
            Caption         =   "アプリ起動"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   15
            Top             =   330
            Width           =   2145
         End
         Begin VB.Label lblAllEndApl 
            Caption         =   "全てのアプリケーションを起動する。"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   16
            Top             =   480
            Width           =   7215
         End
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "LDU"
         Height          =   375
         Index           =   3
         Left            =   7920
         TabIndex        =   7
         Top             =   320
         Width           =   1215
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "IDU"
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   6
         Top             =   320
         Width           =   1335
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "統合監視盤"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   5
         Top             =   320
         Width           =   1695
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "全アプリ一括"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   320
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   15420
      Width           =   2895
   End
   Begin VB.ListBox LstStatus 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   242
      TabIndex        =   1
      Top             =   6360
      Width           =   11535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "アプリケーション起動・終了"
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
      TabIndex        =   2
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmAppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmAppConfig.frm
'//  パッケージ名：アプリ起動・終了画面
'//
'//  概要：アプリ起動・終了画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//     REVISIONS :(2.4.0.1) 2010-10-27   REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　八丁畷対応 不具合修正（ラジオ釦）
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ対応【残件№54】
'//                 ・ポップアップ表示メッセージ変更
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.10修正対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private iTimeCnt As Integer
Private iChoseAplEndSta As Integer  '選択ラジオ釦
Private Const AllApl = 0
Private Const KANSIApl = 1
Private Const IDUApl = 2
Private Const LDUApl = 3
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000        'アプリ起動タイマデフォルト値
Dim lngMAX_Time As Long                    'INI取得設定値
Dim lngtime     As Long                    '現在タイマ値
Private Const APL_END = 4                  'アプリ終了釦押下
Private Const APL_SHOUT_DOWN = 5           'シャットダウン釦押下
Private Const APL_REBOOT = 6               'リブート釦押下
'V1.5.0.1 ADD END
'V1.7.0.1 ADD START
Private iChoseEnd As Integer  '選択終了処理
Private Const NotEnd = -1
'V1.7.0.1 ADD END
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        'ログ起動タイマデフォルト値(30秒)
Dim lngLogMAX_Time As Long                'INI取得設定値(ログ）
'V1.20.0.1 ADD END
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : アプリ起動・終了画面(アクティブ時)
'//  機能概要  : 最前面表示を行う。
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
    pfFormActive (hwnd)
    tmrMail.Enabled = True  'V1.3.0.1 ADD    'メール受信タイマを起動する。
End Sub

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : アプリ起動・終了画面(ディアクティブ時)
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
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : アプリ起動・終了画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         フェーズ１不具合対応 アプリ起動チェック見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    On Error Resume Next
 
    '「アプリ起動・終了画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_END_GAMEN_START, 0)
    
    '初期化
    LstStatus.Clear

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    Koumoku(0).Value = True
    iChoseAplEndSta = AllApl
    iChoseEnd = NotEnd         'V1.7.0.1 ADD
    '縮退チェック
    psIDUCheck
   
    If pbIDUSts = 1 Then
      'IDU業務非表示
       Koumoku(2).Enabled = False
    End If
      
   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
   
    'V1.5.0.1 ADD START
    'INIファイルよりアプリ起動タイマ値を取得
    lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                       APL_INTERVAL, HOSHU_FILE)
    '取得値が0の場合、デフォルト値を設定
    If lngMAX_Time = 0 Then
       lngMAX_Time = APL_INTERVAL
    End If
    'タイマ値設定
    tmrAplTimer.Interval = MN_MAIL_INTERVAL
    tmrAplTimer.Enabled = False
    'V1.5.0.1 ADD END
    
    'V1.20.0.1 ADD START
    'INIファイルよりログ起動タイマ値を取得
    lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
    '取得値が0の場合、デフォルト値を設定
    If lngLogMAX_Time = 0 Then
       lngLogMAX_Time = LOG_INTERVAL
    End If
    
    tmrLogTimer.Interval = MN_MAIL_INTERVAL
    tmrLogTimer.Enabled = False
   'V1.20.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdCancel_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下時処理
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
   
   '「アプリ起動・終了画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_END_GAMEN_END, 0)
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Koumoku_Click
'//  機能名称  : 起動終了選択釦押下時処理
'//  機能概要  : 押下ラジオ釦状態の画面を表示する。
'//              [全アプリ一括][監視盤][IDU][LDU]
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　[IN]押下ラジオ釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Koumoku_Click(Index As Integer)
    
    Select Case Index
        Case 0
          lblAllEndApl(0).Caption = "全てのアプリケーションを起動する。"
'          lblAllEndApl(1).Caption = "全てのアプリケーションを、監視盤の保守のみ残して終了する。"       'EG20 V2.1.0.1 DEL 【Mainte_03_01】
          lblAllEndApl(1).Caption = "全てのアプリケーションを、統合監視盤の保守のみ残して終了する。"    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
          lblAllEndApl(2).Caption = "全てのアプリケーションを終了し、Windowsのみの状態にする。"
          cmdAppEnd.Enabled = True
          cmdAppAllEnd.Enabled = True
          iChoseAplEndSta = AllApl
        Case 1
'          lblAllEndApl(0).Caption = "監視盤アプリケーションを起動する。"           'EG20 V2.1.0.1 DEL 【Mainte_03_01】
          lblAllEndApl(0).Caption = "統合監視盤アプリケーションを起動する。"        'EG20 V2.1.0.1 ADD 【Mainte_03_01】
          lblAllEndApl(1).Caption = ""
          lblAllEndApl(2).Caption = ""
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
          iChoseAplEndSta = KANSIApl
        Case 2
          lblAllEndApl(0).Caption = "IDUアプリケーションを起動する。"
          lblAllEndApl(1).Caption = ""
          lblAllEndApl(2).Caption = "IDUアプリケーションを終了する。"
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = True
          iChoseAplEndSta = IDUApl
        Case 3
          lblAllEndApl(0).Caption = "LDUアプリケーションを起動する。"
          lblAllEndApl(1).Caption = ""
          lblAllEndApl(2).Caption = "LDUアプリケーションを終了する。"
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = True
          iChoseAplEndSta = LDUApl
    End Select
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdAppStart_Click
'//  機能名称  : アプリ起動釦押下時処理
'//  機能概要  : 対象アプリケーションを起動する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ対応【残件№54】
'//                 ・ポップアップ表示メッセージ変更
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応（プログレスバー起動対応）
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub cmdAppStart_Click()
    Dim iRet As Integer '戻り値
    
    On Error Resume Next
   
    '「アプリ起動・終了画面：アプリ起動釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_BUTTOM, 0)
      
    Select Case iChoseAplEndSta
       Case AllApl   '全アプリ一括起動
           iRet = 1
           
           If CheckAppStart(PROC_KANRI) <> 0 Then
               '2重起動
               iRet = 0
           ElseIf CheckAppStart(PROCESS_IDU_PC) <> 0 Then
               '2重起動
               iRet = 0
           ElseIf CheckAppStart(PROCESS_LDU_PC) <> 0 Then
               '2重起動
               iRet = 0
           End If
           
           If iRet = False Then
              '2重警告起動ポップアップ表示
'               iRet = MsgBox("監視盤、ID中継ユニット、LDユーティリティアプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")       'EG20 V2.1.0.1 DEL 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL START
'               iRet = MsgBox("統合監視盤、ID中継ユニット、LDユーティリティアプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
               iRet = MsgBox("統合監視盤、ＩＤＵ、ＬＤＵアプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】ADD END
               Exit Sub
           End If
           '起動確認ポップアップ表示
'           iRet = MsgBox("監視盤、ID中継ユニット、LDユーティリティアプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL START
'           iRet = MsgBox("統合監視盤、ID中継ユニット、LDユーティリティアプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")           'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
           iRet = MsgBox("統合監視盤、ＩＤＵ、ＬＤＵアプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")           'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】ADD END
           If iRet = vbCancel Then
             '[キャンセル]釦押下なら終了
             Exit Sub
           End If
            
           '画面をロックする。
           SetEnableFalse
           '「アプリ起動・終了画面：全アプリ一括起動」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_ALL, 0)

' EG20 V3.0.0.2 追加開始
            ' プログレスバー起動
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 追加終了

           '全アプリを一括起動する。
           '管理起動
             iRet = CheckAppStartComplete(FLD_KPROGNOW & "\\" & PROC_KANRI, 1)
           'IDU起動
            If pbIDUSts = 0 Then 'V2.3.0.1 ADD
             iRet = CheckAppStartComplete(PATH_IDU_APP & PATH_IDU_PROG & PROCESS_LUNCHER, 1)
             Sleep (10000)
            End If 'V2.3.0.1 ADD
           'LDU起動
             iRet = CheckAppStartComplete(PATH_LDU_APP & PATH_LDU_PROG & PROCESS_LDU_LUNCHER, 1)
             Sleep (10000)
           '全アプリの起動チェックを行う。
           '管理チェック
           If CheckAppStart(PROC_KANRI) = 0 Then
              '「アプリ起動・終了画面：アプリ起動処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'              LstStatus.AddItem ("監視盤アプリケーションの起動に失敗しました。")       'EG20 V2.1.0.1 DEL 【Mainte_03_01】
              LstStatus.AddItem ("統合監視盤アプリケーションの起動に失敗しました。")    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
              LstStatus.ListIndex = LstStatus.ListCount - 1
'          ElseIf CheckAppStart(PROCESS_IDU_PC) = 0 Then 'V2.3.0.1 DEL
           ElseIf CheckAppStart(PROCESS_IDU_PC) = 0 And pbIDUSts = 0 Then 'V2.3.0.1 ADD
             '「アプリ起動・終了画面：アプリ起動処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1【残件№54】DEL START
'              LstStatus.AddItem ("ID中継ユニットアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
              LstStatus.AddItem ("ＩＤＵアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】ADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
            ElseIf CheckAppStart(PROCESS_LDU_PC) = 0 Then
             '「アプリ起動・終了画面：アプリ起動処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1【残件№54】DEL START
'              LstStatus.AddItem ("LDユーティリティアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
              LstStatus.AddItem ("ＬＤＵアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】ADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
            Else
             '「アプリ起動・終了画面：アプリ起動処理正常」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'              LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは正常に起動しました。")     'EG20 V2.1.0.1 DEL 【Mainte_03_01】
              LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは正常に起動しました。")  'EG20 V2.1.0.1 ADD 【Mainte_03_01】
              LstStatus.ListIndex = LstStatus.ListCount - 1
           End If
            
            '画面をロックを解除する。
            SetEnableTrue
       
       Case KANSIApl '監視盤起動
            If CheckAppStart(PROC_KANRI) <> 0 Then
               '2重警告起動ポップアップ表示
'               iRet = MsgBox("監視盤アプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")     'EG20 V2.1.0.1 DEL 【Mainte_03_01】
               iRet = MsgBox("統合監視盤アプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")  'EG20 V2.1.0.1 ADD 【Mainte_03_01】
               Exit Sub
            End If
            
            '起動確認ポップアップ表示
'            iRet = MsgBox("監視盤アプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")     'EG20 V2.1.0.1 DEL 【Mainte_03_01】
            iRet = MsgBox("統合監視盤アプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")  'EG20 V2.1.0.1 ADD 【Mainte_03_01】
            If iRet = vbCancel Then
              '[キャンセル]釦押下なら終了
              Exit Sub
            End If
            
            '画面をロックする。
             SetEnableFalse
            '「アプリ起動・終了画面：監視盤アプリ起動」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_KANSI, 0)

' EG20 V3.0.0.2 追加開始
            ' プログレスバー起動
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 追加終了

            '管理起動
             iRet = CheckAppStartComplete(FLD_KPROGNOW & "\\" & PROC_KANRI, 1)
            
            '管理チェック
            If CheckAppStart(PROC_KANRI) = 0 Then
              '「アプリ起動・終了画面：アプリ起動処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'              LstStatus.AddItem ("監視盤アプリケーションの起動に失敗しました。")       'EG20 V2.1.0.1 DEL 【Mainte_03_01】
              LstStatus.AddItem ("統合監視盤アプリケーションの起動に失敗しました。")    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
              LstStatus.ListIndex = LstStatus.ListCount - 1
           Else
             '「アプリ起動・終了画面：アプリ起動処理正常」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'              LstStatus.AddItem ("監視盤アプリケーションは正常に起動しました。")       'EG20 V2.1.0.1 DEL 【Mainte_03_01】
              LstStatus.AddItem ("統合監視盤アプリケーションは正常に起動しました。")    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
              LstStatus.ListIndex = LstStatus.ListCount - 1
            End If
             '画面をロックを解除する。
             SetEnableTrue
            
             cmdAppEnd.Enabled = False
             cmdAppAllEnd.Enabled = False
        
       Case IDUApl   'IDUアプリ起動
           If CheckAppStart(PROCESS_IDU_PC) <> 0 Then
               '2重警告起動ポップアップ表示
'EG20 V2.0.1.1【残件№54】DEL START
'               iRet = MsgBox("ID中継ユニットアプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
               iRet = MsgBox("ＩＤＵアプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")
'EG20 V2.0.1.1【残件№54】ADD END
              Exit Sub
            End If
            
            '起動確認ポップアップ表示
'EG20 V2.0.1.1【残件№54】DEL START
'            iRet = MsgBox("ID中継ユニットアプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
            iRet = MsgBox("ＩＤＵアプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")
'EG20 V2.0.1.1【残件№54】ADD END
            If iRet = vbCancel Then
              '[キャンセル]釦押下なら終了
              Exit Sub
            End If
            
            '画面をロックする。
             SetEnableFalse
            '「アプリ起動・終了画面：IDUアプリ起動」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_IDU, 0)
  
' EG20 V3.0.0.2 追加開始
            ' プログレスバー起動
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 追加終了

            'IDU起動
             iRet = CheckAppStartComplete(PATH_IDU_APP & PATH_IDU_PROG & PROCESS_LUNCHER, 1)
             Sleep (10000)
             DoEvents
            'IDUチェック
            If CheckAppStart(PROCESS_IDU_PC) = 0 Then
             '「アプリ起動・終了画面：アプリ起動処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1【残件№54】DEL START
'              LstStatus.AddItem ("ID中継ユニットアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
              LstStatus.AddItem ("ＩＤＵアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】ADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
           Else
             '「アプリ起動・終了画面：アプリ起動処理正常」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'EG20 V2.0.1.1【残件№54】DEL START
'              LstStatus.AddItem ("ID中継ユニットアプリケーションは正常に起動しました。")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
              LstStatus.AddItem ("ＩＤＵアプリケーションは正常に起動しました。")
'EG20 V2.0.1.1【残件№54】ADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
            End If
             '画面をロックを解除する。
             SetEnableTrue
             cmdAppEnd.Enabled = False
                    
       Case LDUApl   'LDUアプリ起動
            If CheckAppStart(PROCESS_LDU_PC) <> 0 Then
               '2重警告起動ポップアップ表示
'EG20 V2.0.1.1【残件№54】DEL START
'               iRet = MsgBox("LDユーティリティアプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
               iRet = MsgBox("ＬＤＵアプリケーションは既に起動しています。", vbOKOnly + vbExclamation, "２重起動警告")
'EG20 V2.0.1.1【残件№54】ADD END
               Exit Sub
            End If
            
            '起動確認ポップアップ表示
'EG20 V2.0.1.1【残件№54】DEL START
'            iRet = MsgBox("LDユーティリティアプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
            iRet = MsgBox("ＬＤＵアプリケーションを起動します。よろしいですか？", vbOKCancel + vbQuestion, "起動確認")
'EG20 V2.0.1.1【残件№54】ADD END
            If iRet = vbCancel Then
              '[キャンセル]釦押下なら終了
              Exit Sub
            End If
            
            '画面をロックする。
             SetEnableFalse
            '「アプリ起動・終了画面：LDUアプリ起動」ログ出力
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_LDU, 0)
           
' EG20 V3.0.0.2 追加開始
            ' プログレスバー起動
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 追加終了
           
            'LDU起動
             iRet = CheckAppStartComplete(PATH_LDU_APP & PATH_LDU_PROG & PROCESS_LDU_LUNCHER, 1)
             Sleep (10000)

            'LDUチェック
            If CheckAppStart(PROCESS_LDU_PC) = 0 Then
             '「アプリ起動・終了画面：アプリ起動処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1【残件№54】DEL START
'              LstStatus.AddItem ("LDユーティリティアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
              LstStatus.AddItem ("ＬＤＵアプリケーションの起動に失敗しました。")
'EG20 V2.0.1.1【残件№54】ADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
           Else
             '「アプリ起動・終了画面：アプリ起動処理正常」ログ出力
               Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'EG20 V2.0.1.1【残件№54】DEL START
'               LstStatus.AddItem ("LDユーティリティアプリケーションは正常に起動しました。")
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
               LstStatus.AddItem ("ＬＤＵアプリケーションは正常に起動しました。")
'EG20 V2.0.1.1【残件№54】ADD END
               LstStatus.ListIndex = LstStatus.ListCount - 1
            End If
             '画面をロックを解除する。
             SetEnableTrue
             cmdAppEnd.Enabled = False
             
   End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdAppAllEnd_Click
'//  機能名称  : アプリ完全終了釦押下時処理
'//  機能概要  : 対象アプリケーションを完全終了する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ対応【残件№54】
'//                 ・ポップアップ表示メッセージ変更
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.10修正対応】
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub cmdAppAllEnd_Click()
    Dim iRet As Integer '戻り値
    Dim uMail As ML_KYOTU_INF                        'メール
    Dim udtMail As MAIL_IDU_LDU_APLEND_CMD           'メール
    Dim bRtn As Boolean
    Dim lExitCode As Long
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '監視盤アプリ処理結果
    Dim bIDURet As Boolean                          'IDUアプリ処理結果
    Dim bLDURet As Boolean                          'LDUアプリ処理結果
    Dim bIDULOGRet As Boolean                       'IDUログ処理結果
    Dim bLDULOGRet As Boolean                       'LDUログ処理結果
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    bIDULOGRet = False
    bLDULOGRet = False
    'V1.5.0.1 ADD END

    On Error Resume Next
    
    '「アプリ起動・終了画面：アプリ完全終了釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_ALLEND_BUTTOM, 0)

    Select Case iChoseAplEndSta
       Case AllApl   '全アプリ一括終了
           If CheckAppStart(PROC_KANRI) = 0 Then
               '終了済確認ポップアップ表示
'               iRet = MsgBox("監視盤、ID中継ユニット、LDユーティリティアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")         'EG20 V2.1.0.1 DEL 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL START
'               iRet = MsgBox("統合監視盤、ID中継ユニット、LDユーティリティアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")      'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
               iRet = MsgBox("統合監視盤、ＩＤＵ、ＬＤＵアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")      'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】ADD END
               'アプリ起動ツール起動
                Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
           
               '終了処理
                psEndHoshuProc
          
               '保守プロセス終了
                End
           End If
           
           '終了確認ポップアップ表示
'           iRet = MsgBox("監視盤、ID中継ユニット、LDユーティリティアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL START
'           iRet = MsgBox("統合監視盤、ID中継ユニット、LDユーティリティアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】DEL END
'EG20 V2.0.1.1【残件№54】ADD START
           iRet = MsgBox("統合監視盤、ＩＤＵ、ＬＤＵアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【残件№54】ADD START
           If iRet = vbCancel Then
             '[キャンセル]釦押下なら終了
             Exit Sub
           End If
            
            '画面をロックする。
            SetEnableFalse
            '「アプリ起動・終了画面：全アプリ一括終了」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_ALL, 0)
            'アプリ終了要求を管理に送信する
            uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
            uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
            uMail.udtlHeader.dwProid = RHOSHU_ID
            uMail.udtlHeader.dwSubArea = 0
            'V1.5.0.1 DEL START
            'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
            'If bRtn <> 0 Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 ADD START
            bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
            If bKansiRet <> 0 Then
            'V1.5.0.1 ADD END
              ' 「アプリ起動・終了画面：メール送信正常結果」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
            Else
              ' 「アプリ起動・終了画面：メール送信異常結果」ログ出力
              lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
              '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
              SetEnableTrue
              Exit Sub
            End If
     'V1.20.0.1 DEL START
'            'IDU/LDUログ終了要求CMD送信
'            If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'              'IDUログ終了要求CMD送信
'               'V1.5.0.1 DEL START
'               'bRtn = EndIDULog
'               'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'               bIDULOGRet = EndIDULog
'               If bIDULOGRet = False Then
'                 LstStatus.AddItem ("監視盤、ID中継ユニット、LDユーティリティアプリケーションの終了に失敗しました。")
'                 LstStatus.ListIndex = LstStatus.ListCount - 1
'               'V1.5.0.1 ADD END
'                SetEnableTrue
'                Exit Sub
'               End If
'            'V1.5.0.1 ADD START
'            Else
'               bIDULOGRet = True
'            'V1.5.0.1 ADD END
'            End If
'
'            If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'              'LDUログ終了要求CMD送信
'               'V1.5.0.1 DEL START
'               'bRtn = EndLDULog
'               'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'               bLDULOGRet = EndLDULog
'               If bLDULOGRet = False Then
'                  LstStatus.AddItem ("監視盤、ID中継ユニット、LDユーティリティアプリケーションの終了に失敗しました。")
'                  LstStatus.ListIndex = LstStatus.ListCount - 1
'               'V1.5.0.1 ADD END
'                  SetEnableTrue
'                  Exit Sub
'               End If
'            'V1.5.0.1 ADD START
'            Else
'             bLDULOGRet = True
'            'V1.5.0.1 ADD END
'            End If
     'V1.20.0.1 DEL END
'V1.5.0.1 ADD START
            '管理、IDUログ、LDUログへのメール送信正常時のみ、アプリ起動チェックタイマを起動し、
            'INIファイルより取得した時間までアプリ起動チェックを行う。
            'If bKansiRet = True And bIDULOGRet = True And bLDULOGRet = True Then       'V1.20.0.1 DEL
             If bKansiRet = True Then                                                   'V1.20.0.1 ADD
               lngtime = 0
               lngtime = MN_MAIL_INTERVAL
               tmrAplTimer.Enabled = True
               iChoseEnd = AllApl 'V1.7.0.1 ADD
            End If
'V1.5.0.1 ADD END
'           V1.5.0.1 DEL START
'           If CheckAppEndComplete(PROC_KANRI, lExitCode) = 0 _
'            And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 _
'            And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'              '管理、IDUログ、LDUログが終了していなければ、終了処理異常
'              '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'              LstStatus.AddItem ("監視盤、ID中継ユニット、LDユーティリティアプリケーションの終了に失敗しました。")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'              SetEnableTrue
'              Exit Sub
'           End If
'           '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'
'           'アプリ起動ツール起動
'           Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
'
'           '終了処理
'           psEndHoshuProc
'
'           '保守プロセス終了
'           End
'V1.5.0.1 DEL END
       Case IDUApl   'IDUアプリ
           If CheckAppStart(PROCESS_IDU_PC) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 Then
               '終了済警告ポップアップ表示
'EG20 V2.0.1.1【№54】DEL START
'               iRet = MsgBox("ID中継ユニットアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
'               iRet = MsgBox("ＩＤＵは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")                   ' EG20 V3.6.0.1削除
               iRet = MsgBox("ＩＤＵアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")    ' EG20 V3.6.0.1追加
'EG20 V2.0.1.1【№54】ADD END
              Exit Sub
           End If
            
           '終了確認ポップアップ表示
'EG20 V2.0.1.1【№54】DEL START
'           iRet = MsgBox("ID中継ユニットアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
'           iRet = MsgBox("ＩＤＵを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")                  ' EG20 V3.6.0.1削除
           iRet = MsgBox("ＩＤＵアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")   ' EG20 V3.6.0.1追加
'EG20 V2.0.1.1【№54】ADD END
           If iRet = vbCancel Then
             '[キャンセル]釦押下なら終了
             Exit Sub
           End If
           
           '画面をロックする。
            SetEnableFalse
           '「アプリ起動・終了画面：IDUアプリ完全終了」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_IDU, 0)
           'ID制にアプリ終了要求を送信する。
              udtMail.mlHeader.dwId = ML_ID_IDU_APLEND_CMD
              udtMail.mlHeader.dwSize = MlSize.IDUAPLEND_REQ
              udtMail.mlHeader.dwProid = RHOSHU_ID
              udtMail.mlHeader.dwSubArea = 0
              udtMail.dwEndType = ML_ENDTYPE_APLEND
              udtMail.dwCMDLevel = ML_CMDLEVEL_TUJYO        'V1.5.0.1 ADD
            'V1.5.0.1 DEL START
              'bRtn = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
            'If bRtn <> 0 Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 ADD START
            bIDURet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
            If bIDURet <> 0 Then
            'V1.5.0.1 ADD END
              ' 「アプリ起動・終了画面：メール送信正常結果」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
            Else
              ' 「アプリ起動・終了画面：メール送信異常結果」ログ出力
             lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
              '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
              SetEnableTrue
              cmdAppEnd.Enabled = False
              Exit Sub
            End If
   'V1.20.0.1 DEL START
'            'IDU/LDUログ終了要求CMD送信
'            If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'              'IDUログ終了要求CMD送信
'               'V1.5.0.1 DEL START
'               'bRtn = EndIDULog
'               'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'               bIDULOGRet = EndIDULog
'               If bIDULOGRet = False Then
'                  LstStatus.AddItem ("監視盤、ID中継ユニット、LDユーティリティアプリケーションの終了に失敗しました。")
'                  LstStatus.ListIndex = LstStatus.ListCount - 1
'               'V1.5.0.1 ADD END
'                  SetEnableTrue
'                  cmdAppEnd.Enabled = False
'                  Exit Sub
'               End If
'            End If
   'V1.20.0.1 DEL END
'V1.5.0.1 ADD START
            'IDUアプリ(ID制)、IDUログログへのメール送信正常時のみ、アプリ起動チェックタイマを起動し、
            'INIファイルより取得した時間までアプリ起動チェックを行う。
            'If bIDURet = True And bIDULOGRet = True Then    'V1.20.0.1 DEL
            If bIDURet = True Then                           'V1.20.0.1 ADD
               lngtime = 0
               lngtime = MN_MAIL_INTERVAL
               tmrAplTimer.Enabled = True
               iChoseEnd = IDUApl 'V1.7.0.1 ADD
            End If
'V1.5.0.1 ADD END
'V1.5.0.1　DEL　START
'           'IDUチェック
'           If CheckAppEndComplete(PROCESS_IDU_PC, lExitCode) = 0 And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 Then
'              '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'              LstStatus.AddItem ("ID中継ユニットアプリケーションの終了に失敗しました。")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           Else
'             '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'             LstStatus.AddItem ("ID中継ユニットアプリケーションの終了に成功しました。")
'             LstStatus.ListIndex = LstStatus.ListCount - 1
'           End If
'            '画面をロックを解除する。
'            SetEnableTrue
'            cmdAppEnd.Enabled = False
'V1.5.0.1　DEL　END
       Case LDUApl   'LDUアプリ起動
           If CheckAppStart(PROCESS_LDU_PC) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
              '終了済警告ポップアップ表示
'EG20 V2.0.1.1【№54】DEL START
'              iRet = MsgBox("LDユーティリティアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
              iRet = MsgBox("ＬＤＵアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告")
'EG20 V2.0.1.1【№54】ADD END
              Exit Sub
           End If
            
           '終了確認ポップアップ表示
'EG20 V2.0.1.1【№54】DEL START
'           iRet = MsgBox("LDユーティリティアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
           iRet = MsgBox("ＬＤＵアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")
'EG20 V2.0.1.1【№54】ADD END
           If iRet = vbCancel Then
             '[キャンセル]釦押下なら終了
             Exit Sub
           End If
            
           '画面をロックする。
            SetEnableFalse
           '「アプリ起動・終了画面：LDUアプリ完全終了」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_LDU, 0)
           'LD制にアプリ終了要求を送信する。
            udtMail.mlHeader.dwId = ML_ID_LDU_APLEND_CMD
            udtMail.mlHeader.dwSize = MlSize.LDUAPLEND_REQ
            udtMail.mlHeader.dwProid = RHOSHU_ID
            udtMail.mlHeader.dwSubArea = 0
            udtMail.dwEndType = ML_ENDTYPE_APLEND
            udtMail.dwCMDLevel = ML_CMDLEVEL_TUJYO        'V1.5.0.1 ADD
            'V1.5.0.1 DEL START
            'bRtn = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.mlHeader)
            'If bRtn <> 0 Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 ADD START
            bLDURet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.mlHeader)
            If bLDURet <> 0 Then
            'V1.5.0.1 ADD END
              ' 「アプリ起動・終了画面：メール送信正常結果」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
            Else
              ' 「アプリ起動・終了画面：メール送信異常結果」ログ出力
             lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
              '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
              SetEnableTrue
              cmdAppEnd.Enabled = False
              Exit Sub
            End If
 'V1.20.0.1 DEL START
'            If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'              'LDUログ終了要求CMD送信
'              'V1.5.0.1 DEL START
'               'bRtn = EndLDULog
'               'If bRtn = False Then
'              'V1.5.0.1 DEL END
'              'V1.5.0.1 ADD START
'              bLDULOGRet = EndLDULog
'              If bLDULOGRet = False Then
'                 LstStatus.AddItem ("LDユーティリティアプリケーションの終了に失敗しました。")
'                 LstStatus.ListIndex = LstStatus.ListCount - 1
'              'V1.5.0.1 ADD END
'                  SetEnableTrue
'                  cmdAppEnd.Enabled = False
'                  Exit Sub
'               End If
'            End If
 'V1.20.0.1 DEL END
'V1.5.0.1 ADD START
            'LDUアプリ(LD制)、LDUログログへのメール送信正常時のみ、アプリ起動チェックタイマを起動し、
            'INIファイルより取得した時間までアプリ起動チェックを行う。
            'If bLDURet = True And bLDULOGRet = True Then   'V1.20.0.1 DEL
            If bLDURet = True Then     'V1.20.0.1 ADD
               lngtime = 0
               lngtime = MN_MAIL_INTERVAL
               tmrAplTimer.Enabled = True
               iChoseEnd = LDUApl 'V1.7.0.1 ADD
            End If
'V1.5.0.1 ADD END
'V1.5.0.1　DEL　START
'           'LDUチェック
'           If CheckAppEndComplete(PROCESS_LDU_PC, lExitCode) = 0 And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'              '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'              LstStatus.AddItem ("LDユーティリティアプリケーションの終了に失敗しました。")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           Else
'             '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'              LstStatus.AddItem ("LDユーティリティアプリケーションの終了に成功しました。")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           End If
'            '画面をロックを解除する。
'            SetEnableTrue
'            cmdAppEnd.Enabled = False
'V1.5.0.1　DEL　END
   End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdAppEnd_Click
'//  機能名称  : アプリ終了釦押下時処理
'//  機能概要  : 対象アプリケーションを終了する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub cmdAppEnd_Click()
    Dim uMail As ML_KYOTU_INF           'メール
    Dim bRtn As Boolean                 'メールの戻り値
    Dim iRetApp As Integer              '監視盤終了確認戻り値
    Dim iRetIDUApp As Integer           'IDU終了確認戻り値
    Dim iRetLDUApp As Integer           'LDU終了確認戻り値
    Dim iRet As Integer                 'メッセージボックス戻り値
    Dim lExitCode As Long
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '監視盤アプリ処理結果
    Dim bIDURet As Boolean                          'IDUアプリ処理結果
    Dim bLDURet As Boolean                          'LDUアプリ処理結果
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1 ADD END
    
    On Error Resume Next
  
'    iChoseAplEndSta = APL_END           'V1.5.0.1 ADD 'V1.7.0.1 DEL
   
    '「アプリ起動・終了画面：アプリ終了釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_BUTTOM, 0)

    If CheckAppStart(PROC_KANRI) = 0 Then
       '終了済警告ポップアップ表示
       'iRet = MsgBox("監視盤アプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告") 'V1.8.0.1 DEL
'       iRet = MsgBox("監視盤、ID中継ユニット、LDユーティリティアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告") 'V1.8.0.1 ADD       'EG20 V2.1.0.1 DEL 【Mainte_03_01】
'EG20 V2.0.1.1【№54】DEL START
'       iRet = MsgBox("統合監視盤、ID中継ユニット、LDユーティリティアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告") 'V1.8.0.1 ADD    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
       iRet = MsgBox("統合監視盤、ＩＤＵ、ＬＤＵアプリケーションは既に終了しています。", vbOKOnly + vbExclamation, "終了済警告") 'V1.8.0.1 ADD    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【№54】ADD END
       Exit Sub
    End If
    
    '終了確認ポップアップ表示
     'iRet = MsgBox("監視盤アプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")　 ’V1.8.0.1　DEL
'     iRet = MsgBox("監視盤、ID中継ユニット、LDユーティリティアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")    'V1.8.0.1　ADD        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
'EG20 V2.0.1.1【№54】DEL START
'     iRet = MsgBox("統合監視盤、ID中継ユニット、LDユーティリティアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")    'V1.8.0.1　ADD     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
     iRet = MsgBox("統合監視盤、ＩＤＵ、ＬＤＵアプリケーションを終了します。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")    'V1.8.0.1　ADD     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
'EG20 V2.0.1.1【№54】ADD END
     If iRet = vbCancel Then
        '[キャンセル]釦押下なら終了
        Exit Sub
     End If
           
     '画面をロックする。
     SetEnableFalse
     
     '「アプリ起動・終了画面：監視盤アプリ終了」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_KANSI, 0)
     'アプリ終了要求を管理に送信する
     uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
     uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
     uMail.udtlHeader.dwProid = RHOSHU_ID
     uMail.udtlHeader.dwSubArea = 0
     'V1.5.0.1 DEL START
     'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
     'If bRtn <> 0 Then
     'V1.5.0.1 DEL END
     'V1.5.0.1 ADD START
     bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
     If bKansiRet <> 0 Then
     'V1.5.0.1 ADD END
        '「アプリ起動・終了画面：メール送信正常結果」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
        'アプリ終了確認
        'iRetApp = CheckAppEndComplete(PROC_KANRI, lExitCode) 'V1.5.0.1 DEL
  'V1.20.0.1 DEL START
'        'IDUログ確認
'        If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'           'IDUログ終了要求CMD送信
'           'V1.5.0.1 DEL START
'           'bRtn = EndIDULog
'           'If bRtn = False Then
'           'V1.5.0.1 DEL END
'           'V1.5.0.1 ADD START
'           bIDURet = EndIDULog
'           If bIDURet = False Then
'              LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           'V1.5.0.1 ADD END
'              SetEnableTrue
'              Exit Sub
'           End If
'          'IDUログプロセス終了確認
'          'iRetIDUApp = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) 'V1.5.0.1 DEL
'        Else
'           iRetIDUApp = 1
'           bIDURet = True       'V1.5.0.1 ADD
'        End If
'        'LDUログ確認
'        If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'           'LDUログ終了要求CMD送信
'            'V1.5.0.1 DEL START
'            'bRtn = EndLDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bLDURet = EndLDULog
'            If bLDURet = False Then
'               LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'           'LDUログプロセス終了確認
'            'iRetLDUApp = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)   'V1.5.0.1 DEL
'        Else
'           iRetLDUApp = 1
'           bLDURet = True    'V1.5.0.1 ADD
'        End If
  'V1.20.0.1 DEL END
     Else
        '「アプリ起動・終了画面：メール送信異常結果」ログ出力
        lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
        '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
     End If

'V1.5.0.1 ADD START
     'If bKansiRet = True And bIDURet = True And bLDURet = True Then   'V1.20.0.1 DEL
     If bKansiRet = True Then                                          'V1.20.0.1 ADD
        lngtime = 0
        lngtime = MN_MAIL_INTERVAL
        tmrAplTimer.Enabled = True
        iChoseEnd = APL_END         'V1.7.0.1 ADD
     Else
        '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'        LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")       'EG20 V2.1.0.1 DEL 【Mainte_03_01】
        LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
        LstStatus.ListIndex = LstStatus.ListCount - 1
        '画面をロックを解除する。
        SetEnableTrue
     End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'     If iRetApp = 1 And iRetIDUApp = 1 And iRetLDUApp = 1 Then
'        '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'        LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に成功しました。")
'        LstStatus.ListIndex = LstStatus.ListCount - 1
'     Else
'        '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
'        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'        LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")
'        LstStatus.ListIndex = LstStatus.ListCount - 1
'     End If
'
'     '画面をロックを解除する。
'     SetEnableTrue
'V1.5.0.1 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdShoutDown_Click
'//  機能名称  : 「シャットダウン」釦押下時処理
'//  機能概要  : OSを終了する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub cmdShoutDown_Click()
    Dim bRtn As Boolean                 'メールの戻り値
    Dim iRet As Integer                 'メッセージボックス戻り値
    Dim uMail As ML_KYOTU_INF           'メール
    Dim lExitCode As Long               'エラーコード
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '監視盤アプリ処理結果
    Dim bIDURet As Boolean                          'IDUアプリ処理結果
    Dim bLDURet As Boolean                          'LDUアプリ処理結果
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1 ADD END
    
    On Error Resume Next
  
    'iChoseAplEndSta = APL_SHOUT_DOWN           'V1.5.0.1 ADD 'V1.7.0.1 DEL

    '「アプリ起動・終了画面：シャットダウン釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_SHOUT_DOWN_BUTTOM, 0)
 
    '「シャットダウン確認」ポップアップ表示
    iRet = MsgBox("コンピュータをシャットダウンします。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")
    If iRet = vbCancel Then
      '[キャンセル]釦押下なら終了
      Exit Sub
    End If
           
    '画面をロックする。
     SetEnableFalse
     
    If CheckAppStart(PROC_KANRI) <> 0 Then
       'アプリ終了要求を管理に送信する
       uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
       uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
       uMail.udtlHeader.dwProid = RHOSHU_ID
       uMail.udtlHeader.dwSubArea = 0
       'V1.5.0.1 DEL START
       'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       'If bRtn <> 0 Then
       'V1.5.0.1 DEL END
       'V1.5.0.1 ADD START
       bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       If bKansiRet <> 0 Then
       'V1.5.0.1 ADD END
          '「アプリ起動・終了画面：メール送信正常結果」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
  'V1.20.0.1 DEL START
'          'IDUログ確認
'          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'            'IDUログ終了要求CMD送信
'            'V1.5.0.1 DEL START
'            'bRtn = EndIDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bIDURet = EndIDULog
'            If bIDURet = False Then
'               LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'         'V1.5.0.1 ADD START
'         Else
'          bIDURet = True
'         'V1.5.0.1 ADD END
'         End If
'         'LDUログ確認
'         If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'           'LDUログ終了要求CMD送信
'            'V1.5.0.1 DEL START
'            'bRtn = EndLDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bLDURet = EndLDULog
'            If bLDURet = False Then
'               LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'         'V1.5.0.1 ADD START
'         Else
'          bLDURet = True
'        'V1.5.0.1 ADD END
'        End If
 'V1.20.0.1 DEL END
       Else
          '「アプリ起動・終了画面：メール送信異常結果」ログ出力
          lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
          Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
          '「アプリ起動・終了画面：シャットダウン処理処理異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_SHOUT_DOWN_ERROR, 0)
          SetEnableTrue
          Exit Sub
       End If
'V1.5.0.1 ADD START
       'If bKansiRet = True And bIDURet = True And bLDURet = True Then 'V1.20.0.1 DEL
       If bKansiRet = True Then                                        'V1.20.0.1 ADD
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplTimer.Enabled = True
          iChoseEnd = APL_SHOUT_DOWN         'V1.7.0.1 ADD
       Else
           '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'           LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
           LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
           LstStatus.ListIndex = LstStatus.ListCount - 1
           '画面ロック解除
           'SetEnableTrue     'V1.7.0.1 DEL
           'V1.7.0.1 ADD START
           If iChoseAplEndSta = AllApl Then
              'ラジオ釦：全アプリ一括
              SetEnableTrue
           ElseIf iChoseAplEndSta = KANSIApl Then
              'ラジオ釦：監視盤
              SetEnableTrue
              cmdAppEnd.Enabled = False
              cmdAppAllEnd.Enabled = False
           ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
              'ラジオ釦：IDU又はLDU
              SetEnableTrue
              cmdAppEnd.Enabled = False
           End If
           'V1.7.0.1 ADD END
       End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'      If CheckAppEndComplete(PROC_KANRI, lExitCode) = 0 _
'          And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 _
'          And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'       End If
'
'       '終了処理
'       psEndHoshuProc
'       'シャットダウン処理
'       dllAPLEndDwon
'V1.5.0.1 DEL END
    Else
     '終了処理
     psEndHoshuProc
     'シャットダウン処理
     dllAPLEndDwon
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReboot_Click
'//  機能名称  : 「リブート」釦押下時処理
'//  機能概要  : OSを再起動する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReboot_Click()
    Dim bRtn As Boolean                 'メールの戻り値
    Dim iRet As Integer                 'メッセージボックス戻り値
    Dim uMail As ML_KYOTU_INF           'メール
    Dim lExitCode As Long                 'エラーコード
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '監視盤アプリ処理結果
    Dim bIDURet As Boolean                          'IDUアプリ処理結果
    Dim bLDURet As Boolean                          'LDUアプリ処理結果
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1 ADD END
    On Error Resume Next
   
'    iChoseAplEndSta = APL_REBOOT           'V1.5.0.1 ADD 'V1.7.0.1 DEL
     
    '「アプリ起動・終了画面：リブート釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_RBOOT_BUTTOM, 0)
    
    '「リブート確認」ポップアップ表示
    iRet = MsgBox("コンピュータをリブートします。よろしいですか？", vbOKCancel + vbQuestion, "終了確認")
    If iRet = vbCancel Then
      '[キャンセル]釦押下なら終了
      Exit Sub
    End If
           
    '画面をロックする。
    SetEnableFalse
  
    If CheckAppStart(PROC_KANRI) <> 0 Then
       'アプリ終了要求を管理に送信する
       uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
       uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
       uMail.udtlHeader.dwProid = RHOSHU_ID
       uMail.udtlHeader.dwSubArea = 0
       'V1.5.0.1 DEL START
       'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       'If bRtn <> 0 Then
       'V1.5.0.1 DEL END
       'V1.5.0.1 ADD START
       bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       If bKansiRet <> 0 Then
       'V1.5.0.1 ADD END
          '「アプリ起動・終了画面：メール送信正常結果」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
 'V1.20.0.1 DEL START
'          'IDUログ確認
'          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'            'IDUログ終了要求CMD送信
'            'V1.5.0.1 DEL START
'            'bRtn = EndIDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bIDURet = EndIDULog
'            If bIDURet = False Then
'               LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'          'V1.5.0.1 ADD START
'          Else
'           bIDURet = True
'          'V1.5.0.1 ADD END
'          End If
'          'LDUログ確認
'          If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'            'LDUログ終了要求CMD送信
'             'V1.5.0.1 DEL START
'             'bRtn = EndLDULog
'             'If bRtn = False Then
'             'V1.5.0.1 DEL END
'             'V1.5.0.1 ADDL START
'             bLDURet = EndLDULog
'             If bLDURet = False Then
'                LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")
'                LstStatus.ListIndex = LstStatus.ListCount - 1
'             'V1.5.0.1 ADD END
'                SetEnableTrue
'                Exit Sub
'             End If
'           'V1.5.0.1 ADD START
'           Else
'            bLDURet = True
'           'V1.5.0.1 ADD END
'          End If
 'V1.20.0.1 DEL END
      Else
          '「アプリ起動・終了画面：メール送信異常結果」ログ出力
          lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
          Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
          '「アプリ起動・終了画面：リブート処理処理異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_SHOUT_RBOOT_ERROR, 0)
          SetEnableTrue
          Exit Sub
      End If
'V1.5.0.1 ADD START
       'If bKansiRet = True And bIDURet = True And bLDURet = True Then 'V1.20.0.1 DEL
       If bKansiRet = True Then  'V1.20.0.1 ADD
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplTimer.Enabled = True
          iChoseEnd = APL_REBOOT         'V1.7.0.1 ADD
       Else
           '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'           LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
           LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
           LstStatus.ListIndex = LstStatus.ListCount - 1
           '画面ロック解除
           'SetEnableTrue     'V1.7.0.1 DEL
           'V1.7.0.1 ADD START
           If iChoseAplEndSta = AllApl Then
              'ラジオ釦：全アプリ一括
              SetEnableTrue
           ElseIf iChoseAplEndSta = KANSIApl Then
              'ラジオ釦：監視盤
              SetEnableTrue
              cmdAppEnd.Enabled = False
              cmdAppAllEnd.Enabled = False
           ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
              'ラジオ釦：IDU又はLDU
              SetEnableTrue
              cmdAppEnd.Enabled = False
           End If
           'V1.7.0.1 ADD END
       End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'     If CheckAppEndComplete(PROC_KANRI, lExitCode) = 0 _
'        And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 _
'        And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'     End If
'
'     '終了処理
'     psEndHoshuProc
'     'リブート処理
'     dllAPLEndReboot
'V1.5.0.1 DEL END
   Else
    '終了処理
    psEndHoshuProc
    'リブート処理
    dllAPLEndReboot
  End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.4.0.1) 2010-10-27   REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　八丁畷対応 不具合修正（ラジオ釦）
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
  Dim iCnt As Integer 'カウンター
  
  cmdAppEnd.Enabled = True
  cmdAppAllEnd.Enabled = True
  cmdShoutDown.Enabled = True
  cmdReboot.Enabled = True
  cmdCancel.Enabled = True
  cmdAppStart.Enabled = True
  
  For iCnt = 0 To 3
   Koumoku(iCnt).Enabled = True
   'V2.4.0.1 ADD START
   If pbIDUSts = 1 Then
      'IDU業務非表示
       Koumoku(2).Enabled = False
   End If
  'V2.4.0.1 ADD END
  Next
  
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
  Dim iCnt As Integer 'カウンター
  
  cmdAppEnd.Enabled = False
  cmdAppAllEnd.Enabled = False
  cmdShoutDown.Enabled = False
  cmdReboot.Enabled = False
  cmdCancel.Enabled = False
  cmdAppStart.Enabled = False
  
  For iCnt = 0 To 3
   Koumoku(iCnt).Enabled = False
  Next
  
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmAppConfig.Caption, False
        pfFormActive (frmAppConfig.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'V1.5.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplTimer_Timer()
 
  On Error Resume Next

'  Select Case iChoseAplEndSta 'V1.7.0.1 DEL
   Select Case iChoseEnd 'V1.7.0.1 ADD
    '全アプリ一括：完全終了
    Case AllApl
         '全アプリ一括完全終了処理
         ALL_APLEND
    'IDUアプリ：完全終了
    Case IDUApl
         'IDUアプリ完全終了処理
         IDU_APLEND
       
    'LDUアプリ：完全終了
    Case LDUApl
         'LDUアプリ完全終了処理
         LDU_APLEND

    '監視盤アプリ：アプリ終了
    Case APL_END
         '監視盤アプリ：アプリ終了処理
         APL_APLEND
    
    '監視盤、IDU、LDUアプリ：シャットダウン
    Case APL_SHOUT_DOWN
         '監視盤、IDU、LDUアプリ：シャットダウン終了処理
         APL_SHOUT_DOWN_END
    
    '監視盤、IDU、LDUアプリ：リブート
    Case APL_REBOOT
         '監視盤、IDU、LDUアプリ：リブート終了処理
         APL_REBOOT_END
 End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ALL_APLEND
'//  機能名称  : 全アプリ一括完全終了処理
'//  機能概要  : 全アプリ一括完全終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub ALL_APLEND()
 
 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 On Error Resume Next
 
'V1.20.0.1 DEL START
' If CheckAppStart(PROC_KANRI) <> 0 _
'    Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
'    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'V1.20.0.1 DEL END
 If CheckAppStart(PROC_KANRI) <> 0 Then  'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、ID中継ユニット、LDユーティリティアプリケーションの終了に失敗しました。")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       SetEnableTrue
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
       Exit Sub
    Else
       'タイマ張り直し
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
      lngtime = 0
      lngtime = MN_MAIL_INTERVAL
      tmrLogTimer.Enabled = True
      Exit Sub
   Else
      '管理、IDUログ、LDUログが終了していなければ、終了処理異常
      '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'      LstStatus.AddItem ("監視盤、ID中継ユニット、LDユーティリティアプリケーションの終了に失敗しました。")         'EG20 V2.1.0.1 DEL
      LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")      'EG20 V2.1.0.1 ADD
      LstStatus.ListIndex = LstStatus.ListCount - 1
      SetEnableTrue
      iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
   'V1.20.0.1 DEL START
'   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   'アプリ起動ツール起動
'   Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
'   '終了処理
'    psEndHoshuProc
'   '保守プロセス終了
'    End
   'V1.20.0.1 DEL END
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : IDU_APLEND
'//  機能名称  : IDUアプリ完全終了処理
'//  機能概要  : IDUアプリ完全終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub IDU_APLEND()
  
  Dim bIDURet As Boolean  'V1.20.0.1 ADD
 
  On Error Resume Next
 
  'If CheckAppStart(PROCESS_IDU_PC) <> 0 Or CheckAppStart(PROCESS_IDU_LOG) <> 0 Then   'V1.20.0.1 DEL
  If CheckAppStart(PROCESS_IDU_PC) <> 0 Then                                           'V1.20.0.1 ADD
     If lngtime >= lngMAX_Time Then
        tmrAplTimer.Enabled = False
        '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1【№54】DEL START
'        LstStatus.AddItem ("ID中継ユニットアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
        LstStatus.AddItem ("ＩＤＵアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】ADD END
LstStatus.ListIndex = LstStatus.ListCount - 1
        iChoseEnd = NotEnd         'V1.7.0.1 ADD
     Else
        'タイマ張り直し
        tmrAplTimer.Interval = MN_MAIL_INTERVAL
        lngtime = lngtime + MN_MAIL_INTERVAL
        Exit Sub
     End If
  Else
     tmrAplTimer.Enabled = False
     'V1.20.0.1 ADD START
     If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
         bIDURet = EndIDULog 'IDUログ起動時はIDUログに対してログ終了要求CMD送信
     Else
         bIDURet = True
     End If
   
     If bIDURet = True Then
        lngtime = 0
        lngtime = MN_MAIL_INTERVAL
        tmrLogTimer.Enabled = True
        Exit Sub
     Else
      '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1【№54】DEL START
'        LstStatus.AddItem ("ID中継ユニットアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
        LstStatus.AddItem ("ＩＤＵアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】ADD END
        LstStatus.ListIndex = LstStatus.ListCount - 1
        iChoseEnd = NotEnd
     End If
     'V1.20.0.1 ADD END
     'V1.20.0.1 DEL START
     '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'     LstStatus.AddItem ("ID中継ユニットアプリケーションの終了に成功しました。")
'     LstStatus.ListIndex = LstStatus.ListCount - 1
'     iChoseEnd = NotEnd         'V1.7.0.1 ADD
     'V1.20.0.1 DEL END
  End If
  
  '画面をロックを解除する。
  SetEnableTrue
  cmdAppEnd.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : LDU_APLEND
'//  機能名称  : LDUアプリ完全終了処理
'//  機能概要  : LDUアプリ完全終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub LDU_APLEND()
 
 Dim bLDURet As Boolean   'V1.20.0.1 ADD END

 On Error Resume Next
 
' If CheckAppStart(PROCESS_LDU_PC) <> 0 And CheckAppStart(PROCESS_LDU_LOG) <> 0 Then  'V1.20.0.1 DEL
If CheckAppStart(PROCESS_LDU_PC) <> 0 Then   'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1【№54】DEL START
'       LstStatus.AddItem ("LDユーティリティアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
       LstStatus.AddItem ("ＬＤＵアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】ADD END
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
    Else
       'タイマ張り直し
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
    tmrAplTimer.Enabled = False
    'V1.20.0.1 ADD START
    If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
       bLDURet = EndLDULog  'LDUログ起動時はLDUログに対してログ終了要求CMD送信
    Else
       bLDURet = True
    End If
    
    If bLDURet = True Then
      lngtime = 0
      lngtime = MN_MAIL_INTERVAL
      tmrLogTimer.Enabled = True
      Exit Sub
   Else
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1【№54】DEL START
'       LstStatus.AddItem ("LDユーティリティアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
       LstStatus.AddItem ("ＬＤＵアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】ADD END
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
    'V1.20.0.1 DEL START
'    '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'    LstStatus.AddItem ("LDユーティリティアプリケーションの終了に成功しました。")
'    LstStatus.ListIndex = LstStatus.ListCount - 1
'    iChoseEnd = NotEnd         'V1.7.0.1 ADD
    'V1.20.0.1 DEL END
 End If
 
 '画面をロックを解除する。
 SetEnableTrue
 cmdAppEnd.Enabled = False
 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : APL_APLEND
'//  機能名称  : 監視盤アプリ、アプリ終了処理
'//  機能概要  : 監視盤アプリ、アプリ終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub APL_APLEND()

 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 
 On Error Resume Next

'V1.20.0.1 DEL START
' If CheckAppStart(PROC_KANRI) <> 0 _
'    Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
'    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'V1.20.0.1 DEL END
 If CheckAppStart(PROC_KANRI) <> 0 Then 'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")            'EG20 V2.1.0.1 DEL 【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")         'EG20 V2.1.0.1 ADD 【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
    Else
       'タイマ張り直し
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
      lngtime = 0
      lngtime = MN_MAIL_INTERVAL
      tmrLogTimer.Enabled = True
      Exit Sub
   Else
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
   'V1.20.0.1 DEL START
'   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に成功しました。")
'   LstStatus.ListIndex = LstStatus.ListCount - 1
'   iChoseEnd = NotEnd         'V1.7.0.1 ADD
   'V1.20.0.1 DEL END
 End If
 
 '画面をロックを解除する。
 SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : APL_SHOUT_DOWN_END
'//  機能名称  : シャットダウン処理
'//  機能概要  : シャットダウン処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub APL_SHOUT_DOWN_END()
 
 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 
 On Error Resume Next

'V1.20.0.1 ADD START
' If CheckAppStart(PROC_KANRI) <> 0 _
'    Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
'    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'V1.20.0.1 ADD END
 If CheckAppStart(PROC_KANRI) <> 0 Then 'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '画面ロック解除
       'SetEnableTrue     'V1.7.0.1 DEL
       'V1.7.0.1 ADD START
       If iChoseAplEndSta = AllApl Then
          'ラジオ釦：全アプリ一括
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          'ラジオ釦：監視盤
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          'ラジオ釦：IDU又はLDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       'V1.7.0.1 ADD END
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
       Exit Sub
    Else
       'タイマ張り直し
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
      lngtime = 0
      lngtime = MN_MAIL_INTERVAL
      tmrLogTimer.Enabled = True
   Else
      '管理、IDUログ、LDUログが終了していなければ、終了処理異常
      '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'      LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")         'EG20 V2.1.0.1 DEL 【Mainte_03_01】
      LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")      'EG20 V2.1.0.1 ADD 【Mainte_03_01】
      LstStatus.ListIndex = LstStatus.ListCount - 1
      '画面ロック解除
      If iChoseAplEndSta = AllApl Then
         'ラジオ釦：全アプリ一括
         SetEnableTrue
      ElseIf iChoseAplEndSta = KANSIApl Then
         'ラジオ釦：監視盤
         SetEnableTrue
         cmdAppEnd.Enabled = False
         cmdAppAllEnd.Enabled = False
      ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
         'ラジオ釦：IDU又はLDU
         SetEnableTrue
         cmdAppEnd.Enabled = False
      End If
      iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
   'V1.20.0.1 DEL START
'   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   '終了処理
'   psEndHoshuProc
'   'シャットダウン処理
'   dllAPLEndDwon
   'V1.20.0.1 DEL END
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : APL_REBOOT_END
'//  機能名称  : リブート処理
'//  機能概要  : リブート処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                フェーズ１不具合対応　アプリ起動チェック処理見直し修正
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub APL_REBOOT_END()
 
 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 
 On Error Resume Next
 
 'V1.20.0.1 DEL START
 'If CheckAppStart(PROC_KANRI) <> 0 _
 '   Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
 '   Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
 'V1.20.0.1 DEL END
 If CheckAppStart(PROC_KANRI) <> 0 Then  'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '画面ロック解除
       'SetEnableTrue     'V1.7.0.1 DEL
       'V1.7.0.1 ADD START
       If iChoseAplEndSta = AllApl Then
          'ラジオ釦：全アプリ一括
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          'ラジオ釦：監視盤
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          'ラジオ釦：IDU又はLDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       'V1.7.0.1 ADD END
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
       Exit Sub
    Else
       'タイマ張り直し
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
      lngtime = 0
      lngtime = MN_MAIL_INTERVAL
      tmrLogTimer.Enabled = True
   Else
     '管理、IDUログ、LDUログが終了していなければ、終了処理異常
     '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'     LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")      'EG20 V2.1.0.1 DEL 【Mainte_03_01】
     LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")   'EG20 V2.1.0.1 ADD 【Mainte_03_01】
     LstStatus.ListIndex = LstStatus.ListCount - 1
     '画面ロック解除
     If iChoseAplEndSta = AllApl Then
        'ラジオ釦：全アプリ一括
        SetEnableTrue
     ElseIf iChoseAplEndSta = KANSIApl Then
        'ラジオ釦：監視盤
        SetEnableTrue
        cmdAppEnd.Enabled = False
        cmdAppAllEnd.Enabled = False
     ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
        'ラジオ釦：IDU又はLDU
        SetEnableTrue
        cmdAppEnd.Enabled = False
     End If
     iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
'V1.20.0.1 DEL START
'   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   '終了処理
'   psEndHoshuProc
'   'リブート処理
'   dllAPLEndReboot
'V1.20.0.1 DEL END
 End If
End Sub
'V1.5.0.1 ADD END

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()
    
   On Error Resume Next

   Select Case iChoseEnd
        '全アプリ一括：完全終了
        Case AllApl
             '全アプリ一括完全終了処理
             ALL_APLEND_LOG
        'IDUアプリ：完全終了
        Case IDUApl
             'IDUアプリ完全終了処理
             IDU_APLEND_LOG
           
        'LDUアプリ：完全終了
        Case LDUApl
             'LDUアプリ完全終了処理
             LDU_APLEND_LOG
    
        '監視盤アプリ：アプリ終了
        Case APL_END
             '監視盤アプリ：アプリ終了処理
             APL_APLEND_LOG
        
        '監視盤、IDU、LDUアプリ：シャットダウン
        Case APL_SHOUT_DOWN
             '監視盤、IDU、LDUアプリ：シャットダウン終了処理
             APL_SHOUT_DOWN_END_LOG
        
        '監視盤、IDU、LDUアプリ：リブート
        Case APL_REBOOT
             '監視盤、IDU、LDUアプリ：リブート終了処理
             APL_REBOOT_END_LOG
      End Select
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ALL_APLEND
'//  機能名称  : 全アプリ一括完全終了処理
'//  機能概要  : 全アプリ一括完全終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub ALL_APLEND_LOG()
 
 On Error Resume Next
 
 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       'ログ起動チェックタイマを停止する。
       tmrLogTimer.Enabled = False
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、ID中継ユニット、LDユーティリティアプリケーションの終了に失敗しました。")        'EG20 V2.1.0.1 DEL 【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD 【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       SetEnableTrue
       iChoseEnd = NotEnd
       Exit Sub
    Else
       'タイマ張り直し
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrLogTimer.Enabled = False
   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
   'アプリ起動ツール起動
   Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
   '終了処理
    psEndHoshuProc
   '保守プロセス終了
    End
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : IDU_APLEND
'//  機能名称  : IDUアプリ完全終了処理
'//  機能概要  : IDUアプリ完全終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ対応【残件№54】
'//                 ・ポップアップ表示メッセージ変更
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub IDU_APLEND_LOG()
 
  On Error Resume Next
 
  If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
     If lngtime >= lngLogMAX_Time Then
        tmrLogTimer.Enabled = False
        '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'        LstStatus.AddItem ("ID中継ユニットアプリケーションの終了に失敗しました。")     'EG20 V2.0.1.1 DEL
        LstStatus.AddItem ("ＩＤＵアプリケーションの終了に失敗しました。")              'EG20 V2.0.1.1 ADD
        LstStatus.ListIndex = LstStatus.ListCount - 1
        iChoseEnd = NotEnd
     Else
        'タイマ張り直し
        tmrLogTimer.Interval = MN_MAIL_INTERVAL
        lngtime = lngtime + MN_MAIL_INTERVAL
        Exit Sub
     End If
  Else
     tmrLogTimer.Enabled = False
     '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'     LstStatus.AddItem ("ID中継ユニットアプリケーションの終了に成功しました。")        'EG20 V2.0.1.1 DEL
     LstStatus.AddItem ("ＩＤＵアプリケーションの終了に成功しました。")                 'EG20 V2.0.1.1 ADD
     LstStatus.ListIndex = LstStatus.ListCount - 1
     iChoseEnd = NotEnd
  End If
  
  '画面をロックを解除する。
  SetEnableTrue
  cmdAppEnd.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : LDU_APLEND
'//  機能名称  : LDUアプリ完全終了処理
'//  機能概要  : LDUアプリ完全終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub LDU_APLEND_LOG()
 
 On Error Resume Next
 
 If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1【№54】DEL START
'       LstStatus.AddItem ("LDユーティリティアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
       LstStatus.AddItem ("ＬＤＵアプリケーションの終了に失敗しました。")
'EG20 V2.0.1.1【№54】ADD END
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
    Else
       'タイマ張り直し
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
    tmrLogTimer.Enabled = False
    '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'EG20 V2.0.1.1【№54】DEL START
'    LstStatus.AddItem ("LDユーティリティアプリケーションの終了に成功しました。")
'EG20 V2.0.1.1【№54】DEL END
'EG20 V2.0.1.1【№54】ADD START
    LstStatus.AddItem ("ＬＤＵアプリケーションの終了に成功しました。")
'EG20 V2.0.1.1【№54】ADD END
    LstStatus.ListIndex = LstStatus.ListCount - 1
    iChoseEnd = NotEnd
 End If
 
 '画面をロックを解除する。
 SetEnableTrue
 cmdAppEnd.Enabled = False
 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : APL_APLEND
'//  機能名称  : 監視盤アプリ、アプリ終了処理
'//  機能概要  : 監視盤アプリ、アプリ終了処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub APL_APLEND_LOG()
 
 On Error Resume Next

 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL  【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD  【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
    Else
       'タイマ張り直し
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
   tmrLogTimer.Enabled = False
   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に成功しました。")        'EG20 V2.1.0.1 DEL  【Mainte_03_01】
   LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に成功しました。")     'EG20 V2.1.0.1 ADD  【Mainte_03_01】
   LstStatus.ListIndex = LstStatus.ListCount - 1
   iChoseEnd = NotEnd
 End If
 
 '画面をロックを解除する。
 SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : APL_SHOUT_DOWN_END
'//  機能名称  : シャットダウン処理
'//  機能概要  : シャットダウン処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub APL_SHOUT_DOWN_END_LOG()
 
 On Error Resume Next

 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL  【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD  【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '画面ロック解除
       If iChoseAplEndSta = AllApl Then
          'ラジオ釦：全アプリ一括
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          'ラジオ釦：監視盤
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          'ラジオ釦：IDU又はLDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       iChoseEnd = NotEnd
       Exit Sub
    Else
       'タイマ張り直し
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrLogTimer.Enabled = False
   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
   '終了処理
   psEndHoshuProc
   'シャットダウン処理
   dllAPLEndDwon
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : APL_REBOOT_END
'//  機能名称  : リブート処理
'//  機能概要  : リブート処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　ログタイマ追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub APL_REBOOT_END_LOG()
 
 On Error Resume Next

 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '管理、IDUログ、LDUログが終了していなければ、終了処理異常
       '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("監視盤、IDU、LDUアプリケーションは終了に失敗しました。")        'EG20 V2.1.0.1 DEL  【Mainte_03_01】
       LstStatus.AddItem ("統合監視盤、IDU、LDUアプリケーションは終了に失敗しました。")     'EG20 V2.1.0.1 ADD  【Mainte_03_01】
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '画面ロック解除
       If iChoseAplEndSta = AllApl Then
          'ラジオ釦：全アプリ一括
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          'ラジオ釦：監視盤
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          'ラジオ釦：IDU又はLDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       iChoseEnd = NotEnd
       Exit Sub
    Else
       'タイマ張り直し
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrLogTimer.Enabled = False
   '「アプリ起動・終了画面：アプリ終了処理正常」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
   '終了処理
   psEndHoshuProc
   'リブート処理
   dllAPLEndReboot
 End If
End Sub
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : psfuncStartupProgressBar
'//  機能名称  : プログレスバー起動処理
'//  機能概要  : プログレスバーの起動を実行する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応（プログレスバー起動対応）
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psfuncStartupProgressBar()

    Dim iRet As Integer     ' 戻り値

    On Error Resume Next

    If CheckAppStart(PROCESS_TOOL_PROGRESSBAR) = 0 Then
        ' プログレスバー起動
        iRet = CheckAppStartComplete(FILEPATH_PROGRESSTOOL & PROCESS_TOOL_PROGRESSBAR, 1)
        If iRet <> 0 Then
            '「アプリ起動・終了画面：プログレスバー起動成功」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_STARTOK_PROGRESSBAR, 0)
        Else
            '「アプリ起動・終了画面：プログレスバー起動異常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_STARTERR_PROGRESSBAR, 0)
        End If
    End If
End Sub
