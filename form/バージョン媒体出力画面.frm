VERSION 5.00
Begin VB.Form frmVerOutput 
   BorderStyle     =   0  'なし
   Caption         =   "バージョン管理"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   14.25
      Charset         =   128
      Weight          =   700
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
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer tmrMail 
      Left            =   5880
      Top             =   4680
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "テキスト表示"
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
      Left            =   8400
      TabIndex        =   31
      Top             =   5760
      Width           =   3255
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
      Index           =   1
      Left            =   4440
      TabIndex        =   30
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "媒体出力"
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
      Left            =   240
      TabIndex        =   29
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "EG-R自改"
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   27
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "・バージョンチェックファイル："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label lblVerName 
         Caption         =   "・予備2："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   19
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "・メインCPU-OS："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   2370
      End
      Begin VB.Label lblVerName 
         Caption         =   "・メインCPU-Pro："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   17
         Top             =   705
         Width           =   2205
      End
      Begin VB.Label lblVerName 
         Caption         =   "・予備１："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   16
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "・サブCPU-Pro："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   15
         Top             =   1035
         Width           =   2175
      End
      Begin VB.Label lblVerName 
         Caption         =   "・判定CPU-Pro： "
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2190
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   13
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   12
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   11
         Top             =   705
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   10
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   9
         Top             =   1035
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraAllKansiVersion 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11535
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   20
         Left            =   8520
         TabIndex        =   34
         Top             =   650
         Width           =   2895
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   19
         Left            =   4500
         TabIndex        =   33
         Top             =   650
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   17
         Left            =   450
         TabIndex        =   32
         Top             =   650
         Width           =   2295
      End
      Begin VB.Label lblVerName 
         Caption         =   "・監視盤："
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   350
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "・ＩＤ中継ユニット："
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   3
         Top             =   350
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "・ＬＤユーティリティ："
         Height          =   375
         Index           =   3
         Left            =   8355
         TabIndex        =   2
         Top             =   350
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " 　 メニュー　   画面へ戻る"
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
   Begin VB.Label lblVerName 
      Caption         =   "Z9"
      Height          =   375
      Index           =   29
      Left            =   7320
      TabIndex        =   36
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "・磁気運賃："
      Height          =   495
      Index           =   21
      Left            =   4440
      TabIndex        =   35
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "Z9.Z9.Z9.Z9"
      Height          =   375
      Index           =   18
      Left            =   7320
      TabIndex        =   28
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   9
      Left            =   7320
      TabIndex        =   25
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   8
      Left            =   7320
      TabIndex        =   24
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "99"
      Height          =   375
      Index           =   7
      Left            =   7320
      TabIndex        =   23
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "・駅都度バージョン："
      Height          =   495
      Index           =   5
      Left            =   4440
      TabIndex        =   22
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "・NEG自改："
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   21
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "・ＩＣ共通運賃："
      Height          =   495
      Index           =   28
      Left            =   4440
      TabIndex        =   20
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "バージョン媒体出力"
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
   Begin VB.Label lblVerName 
      Caption         =   "・ＩＣ－Ｍ："
      Height          =   375
      Index           =   6
      Left            =   4440
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "frmVerOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmVerOutPut.frm
'//  パッケージ名：バージョン媒体出力画面
'//
'//  概要：バージョン媒体出力画面
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 ・フェーズ３対応　バージョン媒体出力追加
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 バージョンファイル書込み先ディレクトリ位置変更
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
Private Const PtnEkiVersion = "000002"  '駅バージョン
Dim sWriteDir As String                 '媒体出力先

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン媒体出力画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
   On Error Resume Next
    
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : バージョン媒体出力画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
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
'//  機能名称  : バージョン媒体出力画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
   Dim strWork         As String   '作業エリア
 
   On Error Resume Next
 
   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
           
   sWriteDir = ""
   
   'IDU縮退チェック
   psIDUCheck
    
   'バージョン取得処理
   psGetVersion
   
   'メール受信用のタイマ値を設定する。
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
   
   '「バージョン媒体出力画面 表示消去」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_GAMEN_START, 0)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psGetVersion
'//  機能名称  : バージョン取得処理
'//  機能概要  : バージョン取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psGetVersion()
  
  Dim sVersion  As String
  Dim sGetJikiVer As String     'V1.10.0.1 ADD
  
  On Error Resume Next

 '監視盤、EG-R全体バージョン取得
  psKansiGetVersion
 
 If pbIDUSts = 1 Then
    'IDUバージョン非表示
    lblVerName(2).Enabled = False
    lblVerName(19).Caption = ""
 Else
    '非縮退時は表示処理
    psIDUGetVersion
 End If
 
 'LDU全体バージョン取得
  psLDUVersion

 'EG-R自改バージョン取得
  '判定CPU
  sVersion = psEGRJVersion(HANTEI_CPU)
  lblVerName(10).Caption = sVersion
  'メインCPU
  sVersion = psEGRJVersion(MAIN_CPU)
  lblVerName(13).Caption = sVersion
 'サブCPU
  sVersion = psEGRJVersion(SUB_CPU)
  lblVerName(11).Caption = sVersion
 'メインOS
  sVersion = psEGRJVersion(MAIN_OS)
  lblVerName(14).Caption = sVersion
 '予備１
  sVersion = psEGRJVersion(YOBI1)
  lblVerName(12).Caption = sVersion
 '予備２
  sVersion = psEGRJVersion(YOBI2)
  lblVerName(15).Caption = sVersion
 'バージョンチェック
  sVersion = psEGRJVersion(VER_CHK)
  lblVerName(0).Caption = sVersion
  
 'NEG自改バージョン取得
  sVersion = psNEGJVersion
  lblVerName(7).Caption = sVersion
 
 '判定IC-Mバージョン取得
 If pbIDUSts = 1 Then
    '判定IC-M(IC-M)バージョン非表示
    lblVerName(6).Enabled = False
    lblVerName(8).Caption = ""
 Else
    '非縮退時は表示処理
    sVersion = psICMGetVersion
    lblVerName(8).Caption = sVersion
 End If
 
 '共通運賃バージョン取得
 If pbIDUSts = 1 Then
    '共通運賃バージョン非表示
    lblVerName(28).Enabled = False
    lblVerName(9).Caption = ""
 Else
    '非縮退時は表示処理
    sVersion = psICUnchinGetVersion
    lblVerName(9).Caption = sVersion
 End If
  
 '駅都度バージョン取得
 pfEkiVersion
 
'V1.10.0.1 ADD START
 '磁気運賃読み込み
 sGetJikiVer = psJikiUnchinVersion
 lblVerName(29).Caption = CStr(sGetJikiVer)
'V1.10.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psKansiGetVersion
'//  機能名称  : 監視装置全体、監視盤バージョン取得処理
'//  機能概要  : KansiVersion.iniよりバージョンを取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function psKansiGetVersion()
    Dim lSts As Long                                       '関数戻り値
    Dim strKansiVersion As String * VERSION_GATE_SIZE      '監視盤全体バージョン
    Dim strKansiVersion2 As String * VERSION_GATE_SIZE     '監視装置全体バージョン
    
    On Error Resume Next
    
    strKansiVersion = ""
    strKansiVersion2 = ""

    ' KansiVersion.iniから監視装置の全体バージョンを取得し表示する
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSISYSTEMVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSION_FILE)
    If lSts > 0 Then
        '取得したバージョン番号を表示
        fraAllKansiVersion.Caption = "全体バージョン： " & Left$(strKansiVersion, lSts) & ""
    Else
        'バージョン番号の取得異常の場合、「--,--,--,--」を表示
        fraAllKansiVersion.Caption = "全体バージョン：--.--.--.-- "
    End If
 
    ' KansiVersion.iniから監視盤の全体バージョンを取得し表示する
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion2, _
                                   Len(strKansiVersion2), _
                                   KANSI_VERSION_FILE)
     If lSts > 0 Then
        '取得したバージョン番号を表示
        lblVerName(17).Caption = Left$(strKansiVersion2, lSts)
    Else
        'バージョン番号の取得異常の場合、「--,--,--,--」を表示
        lblVerName(17).Caption = "--.--.--.-- "
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psIDUGetVersion
'//  機能名称  : ID中継ユニットバージョン取得処理
'//  機能概要  : ID中継ユニットバージョン管理ファイルより、
'//              ID中継ユニットのバージョンを取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function psIDUGetVersion()
    Dim strWork     As String       '作業エリア
    Dim iFileNumber As Integer      '未使用ファイル番号
    
    On Error Resume Next
    
    iFileNumber = FreeFile          '未使用のファイル番号を取得する
        
   'ID中継ユニットバージョン管理ファイルをオープン。
    Open PATH_IDU_APP & PATH_IDU_VERKANRI For Input As #iFileNumber

    '実行バージョンを取得する。
    Input #iFileNumber, strWork
    If (Trim(strWork) = "") Then
        'バージョン番号取得異常の場合
        lblVerName(19).Caption = "--.--.--.--"
    Else
       '全体バージョン文字列作成
        lblVerName(19).Caption = Trim(strWork)
    End If
      
   'ID中継ユニットバージョン管理ファイルをクローズ。
    Close #iFileNumber
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psLDUVersion
'//  機能名称  : LDユーティリティバージョン取得処理
'//  機能概要  : LDユーティリティバージョン管理ファイルより、
'//              LDユーティリティのバージョンを取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function psLDUVersion()
    Dim strWork     As String       '作業エリア
    Dim iFileNumber As Integer      '未使用ファイル番号
    
    On Error Resume Next
    
    iFileNumber = FreeFile          '未使用のファイル番号を取得する
    
   'LDユーティリティバージョン管理ファイルをオープン。
    Open PATH_LDU_APP & PATH_LDU_VERKANRI For Input As #iFileNumber

    '実行バージョンを取得する。
    Input #iFileNumber, strWork
    If (Trim(strWork) = "") Then
        'バージョン番号取得異常の場合
        lblVerName(20).Caption = "--.--.--.--"
    Else
       '全体バージョン文字列作成
        lblVerName(20).Caption = Trim(strWork)
    End If
      
   'LDユーティリティバージョン管理ファイルをクローズ。
    Close #iFileNumber

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfEkiVersion
'//  機能名称  : 駅都度バージョン取得処理
'//  機能概要  : 現在駅設定ファイルより、
'//              駅都度バージョンを取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function pfEkiVersion()

   Dim intFileNo            As Integer  'ファイル番号
   Dim intBunrui_Dai        As Integer         '大分類
   Dim intBunrui_Tyu        As Integer         '中分類
   Dim intBunrui_Sho        As Integer         '小分類
   Dim strData              As String          '設定値
   Dim strPtnNo             As String          'パターン番号
   Dim strEkiVersion        As String          '駅バージョン
   Dim iGetDataCount        As Integer         'データ取得カウンタ
   Dim strFileName          As String          'ファイル名
 
   On Error Resume Next
 
   strFileName = Dir(EKI_SETTI_FILE, vbNormal)
   
   If strFileName = "" Then
      lblVerName(18).Caption = "--.--.--.--"
      Exit Function
   End If
   
   'ファイル番号を取得する。
   intFileNo = FreeFile
    
   'ファイルオープン
   On Error GoTo FileGetError
   Open EKI_SETTI_FILE For Input As #intFileNo

   Do While Not EOF(intFileNo)
      '１ 行づつ変数読み込み
       Input #intFileNo, intBunrui_Dai, intBunrui_Tyu, intBunrui_Sho, strData
   
       'パターン番号取得
        strPtnNo = Format(intBunrui_Dai, "00") & Format(intBunrui_Tyu, "00") & Format(intBunrui_Sho, "00")
           
        Select Case strPtnNo
             
            '駅バージョン取得
             Case PtnEkiVersion
                   strEkiVersion = strData & "  "
                   iGetDataCount = iGetDataCount + 1
             
             Case Else
                   '処理なし
                    
         End Select
            
         '駅バージョンを取得したらループを抜ける
         If iGetDataCount = 1 Then Exit Do
   Loop

   'ファイルクローズ
   Close #intFileNo

   If strEkiVersion = "" Then
      lblVerName(18).Caption = "--.--.--.--"
   Else
      lblVerName(18).Caption = strEkiVersion
   End If
   
   Exit Function
FileGetError:
   lblVerName(18).Caption = "--.--.--.--"

End Function
     
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 「媒体出力」「媒体取外」「テキスト表示」を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 バージョンファイル書込み先ディレクトリ位置変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    
    Dim bRet      As Boolean         '戻り値
    Dim lRetVal   As Long            'テキスト表示処理戻り値
    Dim sCommand  As String          'コマンド文字列
    
    On Error Resume Next
 
    Select Case Index
        Case 0                                 '「媒体出力」釦
            '「媒体出力釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_OUTPUT_BUTTOM, 0)
            ' 取出し先ディレクトリを選択する
'            sWriteDir = pfDirSelection("a:", "バージョンファイル書込み先ディレクトリ選択")     'V1.12.0.1 DEL
            sWriteDir = pfDirSelection("H:", "バージョンファイル書込み先ディレクトリ選択")      'V1.12.0.1 ADD
            If sWriteDir <> "" Then
            'ディレクトリが指定されれば、バージョンファイルを取出す
                bRet = dllEGRCreateVersionFile(PATH_IDU_APP, PATH_LDU_APP)
                If bRet = False Then
                  '「ファイル作成異常」ポップアップ画面表示
                    MsgBox "ファイルの作成に失敗しました。", vbOKOnly + vbCritical, "ファイル作成異常"
                  '「ファイル作成異常」ログ出力
                  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_FILE_CREATE_ERROR, 0)
                  
                  '「媒体出力処理異常」ログ出力
                   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_OUTPUT_ERROR, 0)
                   Exit Sub
                Else
                   'ファイルコピー処理
                   fMakeOutPutFile
                End If
                
            Else
                '「媒体出力処理未実行」ログ出力
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_OUTPUT_MISHORI, 0)
            End If

        Case 1                                 '「媒体取外」釦
            '「媒体取外釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
            '媒体取外処理
            Call pfRemove(Me)
        Case 2                                 '「テキスト表示」釦
            '「テキスト表示釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_TEXT_BUTTOM, 0)
            bRet = dllEGRCreateVersionFile(PATH_IDU_APP, PATH_LDU_APP)
            If bRet = False Then
              '「ファイル作成異常」ポップアップ画面表示
               MsgBox "ファイルの作成に失敗しました。", vbOKOnly + vbCritical, "ファイル作成異常"
               
               '「ファイル作成異常」ログ出力
                   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_FILE_CREATE_ERROR, 0)
               '「テキスト表示処理異常」ログ出力
                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_TEXT_ERROR, 0)
                Exit Sub
            Else
                 'テキストファイル表示処理
                sCommand = MN_EXE_MEMO & EGR_KANSI_VERSION_FILE_PATH 'メモ帳実行コマンドを作成
                'メモ帳を起動する｡
                lRetVal = Shell(sCommand, vbMaximizedFocus)
                'メモ帳をアクティブ（前面表示）にする
                AppActivate lRetVal, True
                SendKeys "{LEFT}", True
               '「テキスト表示処理正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_TEXT_OK, 0)
            End If
    
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
     
    On Error Resume Next
    
    '「バージョン媒体出力画面 消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_GAMEN_END, 0)
 
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMakeOutPutFile
'//  機能名称  : 媒体出力処理を行う。
'//  機能概要  : 媒体出力ファイル出力を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeOutPutFile()
   Dim iResponse As Integer   'MsgBox戻り値
   Dim lngErrCode As Long     'エラーコード
   Dim fso         As New FileSystemObject   'ファイルシステムオブジェクト
   Dim strWriteDir As String               '出力先フォルダ

   On Error GoTo COPY_ERROR

   'ファイルコピー
   FileCopy EGR_KANSI_VERSION_FILE_PATH, sWriteDir & EGR_KANSI_VERSION_FILE
   
   '「媒体出力正常終了」ポップアップ画面表示
   MsgBox "媒体出力は正常終了しました。", vbOKOnly + vbInformation, "媒体出力結果"
                    
   '「バージョン媒体出力：媒体出力処理正常」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERSION_OUTPUT_OUTPUT_OK, 0)
  
   Exit Function
    
COPY_ERROR:
   '処理異常の場合、出力結果ポップアップ(異常)表示
    MsgBox "媒体出力は異常終了しました。", vbCritical, "媒体出力結果"
   '「バージョン媒体出力：媒体出力処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERSION_OUTPUT_OUTPUT_ERROR, lngErrCode)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : タイムアップ時処理
'//  機能概要  : メール受信タイムアップ時処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力画面追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  
    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmVerOutput.Caption, False
    End If

End Sub

