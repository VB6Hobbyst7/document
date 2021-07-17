VERSION 5.00
Begin VB.Form frmVersion 
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "　新幹線　改札機"
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
      Index           =   9
      Left            =   3320
      TabIndex        =   21
      Top             =   8040
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
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
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   7185
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "改札機"
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
      Left            =   4920
      TabIndex        =   17
      Top             =   7185
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＩＤＵ"
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
      Left            =   1720
      TabIndex        =   16
      Top             =   7185
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＬＤＵ"
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
      Left            =   3320
      TabIndex        =   15
      Top             =   7185
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "IC共通運賃"
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
      Left            =   120
      TabIndex        =   14
      Top             =   8040
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
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
      Height          =   855
      Index           =   6
      Left            =   1720
      TabIndex        =   13
      Top             =   8040
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "　Ver一覧　USB出力"
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
      Index           =   8
      Left            =   8120
      TabIndex        =   12
      Top             =   7185
      Width           =   1600
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
      Height          =   5100
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   11530
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
      Index           =   7
      Left            =   8120
      TabIndex        =   10
      Top             =   8040
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＩＣＭ"
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
      Left            =   6520
      TabIndex        =   9
      Top             =   7185
      Width           =   1600
   End
   Begin VB.Timer tmrMail 
      Left            =   11400
      Top             =   7320
   End
   Begin VB.Frame fraAllKansiVersion 
      Caption         =   "全体バージョン：Z9.Z9.Z9.Z9"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11535
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   5
         Left            =   8520
         TabIndex        =   8
         Top             =   650
         Width           =   2895
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   3
         Left            =   4500
         TabIndex        =   7
         Top             =   650
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "・ＩＤＵアプリケーション："
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   6
         Top             =   345
         Width           =   3255
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   1
         Left            =   450
         TabIndex        =   5
         Top             =   650
         Width           =   2295
      End
      Begin VB.Label lblVerName 
         Caption         =   "・統合監視盤： "
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   350
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "・ＬＤＵアプリケーション："
         Height          =   375
         Index           =   4
         Left            =   8355
         TabIndex        =   2
         Top             =   345
         Width           =   3015
      End
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
   Begin VB.Label lbltitle 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "タイトル"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ファイル名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   19
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "バージョン管理"
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
      TabIndex        =   4
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmVersion.frm
'//  パッケージ名：バージョン管理画面
'//
'//  概要：バージョン管理画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 ・フェーズ２対応　監視ファーム、RASマイコン追加
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 ・フェーズ３対応　バージョン媒体出力追加
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//     REVISIONS :(1.20.0.1) 2010-03-11   REVISED BY [TCC] S.Yamazaki
'//                 画面レイアウト変更
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.18修正対応】
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 【ＩＣＭバージョンファイルリスト対応】
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 【バージョン表示不正対応】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'//                 【北陸新幹線開業対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
Dim uVersion() As MN_VERSION_JIKAI      'バージョン情報格納エリア

'V1.20.0.1 ADD START
Private Type DISP_FILE_INFO      '表示バージョンcsvファイル内容
    sTitle As String             'タイトル
    sFilePath As String          'ファイルパス
    iType As Integer             '表示タイプ
    iIdu As Integer              'ＩＤＵ縮退対象ファイル有無
    iMaker As Integer            ' メーカ番号（タイプ２）             ' EG20 V5.6.0.1追加
End Type

Private Const CSV_COMMENT_CHAR = ":"  'csvファイルでコメントとする文字列
'V1.20.0.1 ADD END

' EG20 V2.1.0.1[Mainte_03_01] 追加開始
Dim FileList() As String                     'ファイル名リスト一覧格納エリア
Dim FileListType() As String                 'ファイルリスト一覧格納エリア（次世代自改タイプを含む）
' EG20 V2.1.0.1[Mainte_03_01] 追加終了


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン管理画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
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
    
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : バージョン管理画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ起動
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
'//  機能名称  : バージョン管理画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-18   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi06_004_02】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   Dim strWork         As String   '作業エリア
 
   On Error Resume Next
 
   '「バージョン管理画面：表示」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_GAMEN_START, 0)

   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
   
   'V1.4.0.1 DEL START
   'バージョン取得処理
   'psGetVersion
   'V1.4.0.1 DEL END
   
   '磁気運賃対応チェック
'   psJikiCheck                                ' EG20 V2.1.0.1[Mainte_03_01] 削除
    
   'IDU縮退チェック
   psIDUCheck
    
   If pbIDUSts = 1 Then
     'IDU業務非表示
      cmdFixedExe(1).Visible = False
'      cmdFixedExe(5).Visible = False          ' EG20 V2.1.0.1[Mainte_03_01] 削除
'      cmdFixedExe(6).Visible = False          ' EG20 V2.1.0.1[Mainte_03_01] 削除
      cmdFixedExe(4).Visible = False           ' EG20 V2.1.0.1[Mainte_03_01] 追加
      cmdFixedExe(5).Visible = False           ' EG20 V2.1.0.1[Mainte_03_01] 追加
   End If
   
   'V1.4.0.1 ADD START
   'バージョン取得処理
   psGetVersion
   'V1.4.0.1 ADD END

   'V1.3.0.1 ADD START
   'メール受信用のタイマ値を設定する。
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
   '1.3.0.1 ADD END
   'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL START
   'EG20 V30.1.0.1 ADD START
'    If fIsExistCornerType(CORNER_TYPE_ZAIRAI) = False Then
'        '在来線コーナーが一つも存在しないので、改札機釦は押下不可にする。
'        cmdFixedExe(3).Enabled = False
'    End If
'
'    If fIsExistCornerType(CORNER_TYPE_KANSEN) = False Then
'        '幹線コーナーが一つも存在しないので、新幹線改札機釦は押下不可にする。
'        cmdFixedExe(9).Enabled = False
'    End If
   'EG20 V30.1.0.1 ADD END
   'EG20 V30.3.0.1 【HKRK_Kansi06_004_02】 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各画面遷移釦押下時処理
'//  機能概要  : 釦名称画面に遷移する。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-18   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　監視ファーム、バージョン切替を追加
'//     REVISIONS :(1.6.0.1) 2009-06-18   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　バージョン媒体出力を追加
'//     REVISIONS :(1.20.0.1) 2010-03-17  REVISED BY [TCC] S.Yamazaki
'//                 バージョン切替を媒体取外に変更
'//                 バージョン媒体出力釦をVer一覧USB出力釦に変更
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.18修正対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
   Dim udtMail As ML_DISP_INF          '画面表示要求
   Dim iResponse As Integer            'メッセージボックス戻り値
   Dim bRet As Boolean                 'メール送信処理戻り値
   Dim lngErrCode As Long              'エラーコード
    
    On Error Resume Next
    
    Select Case Index
        Case 0                                 'バージョン管理（監視盤）
             '「バージョン管理画面：監視盤釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_KANSIBAN_BUTTOM, 0)
            Load frmKVer
            frmKVer.Show 1
        Case 1                                 'バージョン管理（IDU）
            '「バージョン管理画面：IDU釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_IDU_BUTTOM, 0)
            Load frmIDUVer
            frmIDUVer.Show 1
        Case 2                                 'バージョン管理（LDU）
            '「バージョン管理画面：LDU釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_LDU_BUTTOM, 0)
            Load frmLduVer
            frmLduVer.Show 1
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'        Case 3                                 'バージョン管理（EG-R自改）
'            '「バージョン管理画面：EG-R自改釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EGRJIKAI_BUTTOM, 0)
'            gStrCurrentForm = sFormName_EJVer
'            Load frmJVer
'            frmJVer.Show 1
'        Case 4                                 'バージョン管理（NEG自改）
'            '「バージョン管理画面：NEG自改釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_NEGJIKAI_BUTTOM, 0)
'            gStrCurrentForm = sFormName_NJVer
'            Load frmJVer
'            frmJVer.Show 1
'        Case 5                                 'バージョン管理（判定IC-M）
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        Case 3                                 'バージョン管理（改札機）
            '「バージョン管理画面：EG20自改釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EG20JIKAI_BUTTOM, 0)
            gStrCurrentForm = sFormName_EG20JVer
            Load frmGateVerKanri
            frmGateVerKanri.Show 1
        Case 4                                 'バージョン管理（判定IC-M）
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
            '「バージョン管理画面：判定IC-M釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_ICM_BUTTOM, 0)
            '画面表示要求(判定IC－M画面)をID制御に送信する
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_HANTEI_VER
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
               '「バージョン管理画面：保守画面表示要求メール送信異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
               '起動失敗ポップアップ表示
' EG20 V3.6.0.1【03統合TR-No.18修正対応】削除開始
'               iResponse = MsgBox("判定IC-M釦、定義エラー。" & _
'                                   Chr(vbKeyReturn) & _
'                                   "判定IC-Mバージョン管理画面を起動できません。", _
'                                   vbOKOnly, _
'                                   "画面起動エラー")
' EG20 V3.6.0.1【03統合TR-No.18修正対応】削除終了
' EG20 V3.6.0.1【03統合TR-No.18修正対応】追加開始
               iResponse = MsgBox("ＩＣＭ釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "ＩＣＭバージョン管理画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
' EG20 V3.6.0.1【03統合TR-No.18修正対応】追加終了
               Exit Sub
            End If
            '「バージョン管理画面：保守画面表示要求メール送信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
' EG20 V2.1.0.1[Mainte_03_01] 追加開始
        Case 5                                 'バージョン管理（IC共通運賃）
            '「バージョン管理画面：IC共通運賃釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_ICUNCHIN_BUTTOM, 0)
            '保守画面表示要求(IC共通運賃画面)をID制御に送信する
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_PASMO_VER
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
               '「バージョン管理画面：保守画面表示要求メール送信異常」ログ出力
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
               '起動失敗ポップアップ表示
               iResponse = MsgBox("IC共通運賃釦、定義エラー。" & _
                                  Chr(vbKeyReturn) & _
                                  "IC共通運賃データバージョン管理画面を起動できません。", _
                                  vbOKOnly, _
                                  "画面起動エラー")
                Exit Sub
            End If
            '「バージョン管理画面：保守画面表示要求メール送信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        Case 6                                 'バージョン管理（操作卓）
' EG20 V2.1.0.1[Mainte_03_01 操作卓バージョン管理画面]追加開始
            '「バージョン管理画面：EG20自改釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_TAKU_BUTTOM, 0)
            gStrCurrentForm = sFormName_EJVer
            Load frmSousaTakuVerKanri
            frmSousaTakuVerKanri.Show 1
' EG20 V2.1.0.1[Mainte_03_01 操作卓バージョン管理画面]追加終了
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
' EG20 V2.1.0.1[Mainte_03_01] 削除開始
'        Case 6                                 'バージョン管理（PASMO運賃）
'            '「バージョン管理画面：PASMO運賃釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_PASMO_BUTTOM, 0)
'            '保守画面表示要求(PASMO運賃画面)をID制御に送信する
'            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
'            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
'            udtMail.udtlHeader.dwProid = RHOSHU_ID
'            udtMail.udtlHeader.dwSubArea = 0
'            udtMail.dwDisp_Type = ML_DT_PASMO_VER
'            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
'            If bRet = False Then
'               '「バージョン管理画面：保守画面表示要求メール送信異常」ログ出力
'               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
'               '起動失敗ポップアップ表示
'               iResponse = MsgBox("PASMO共通運賃釦、定義エラー。" & _
'                                  Chr(vbKeyReturn) & _
'                                  "PASMO共通運賃データバージョン管理画面を起動できません。", _
'                                  vbOKOnly, _
'                                  "画面起動エラー")
'                Exit Sub
'            End If
'            '「バージョン管理画面：保守画面表示要求メール送信正常」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
'        Case 7                                 'バージョン管理(磁気運賃)
'            '「バージョン管理画面：磁気運賃釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_JIKIUNCHIN_BUTTOM, 0)
'            Load frmJikiUnkaiFD
'            frmJikiUnkaiFD.Show 1
'        'V1.4.0.1 ADD START
'        Case 8                                 'バージョン管理(監視ファーム)
'            '「バージョン管理画面：監視ファーム釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_FIRMWARE_BUTTOM, 0)
'            Load frmFirmWareVer
'            frmFirmWareVer.Show 1
' EG20 V2.1.0.1[Mainte_03_01] 削除終了
'        Case 9                                 'バージョン管理(バージョン切替)         ' EG20 V1.1.1.1 削除
        Case 7                                 '媒体取外                                ' EG20 V1.1.1.1 追加
            'V1.20.0.1 DEL START
'            '「バージョン管理画面：バージョン切替釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_CHANGEVER_BUTTOM, 0)
'            Load frmVerChang
'            frmVerChang.Show 1
            'V1.20.0.1 DEL END
            'V1.20.0.1 ADD START
            '「媒体取外釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
            '媒体取外処理
            Call pfRemove(Me)
            'V1.20.0.1 ADD END
        'V1.4.0.1 ADD END
        'V1.6.0.1　ADD START
'        Case 10                                 'バージョン管理(バージョン媒体出力)    ' EG20 V2.1.0.1[Mainte_03_01] 削除
        Case 8                                  'バージョン管理(バージョン媒体出力)     ' EG20 V2.1.0.1[Mainte_03_01] 追加
        'V1.20.0.1 DEL START
'           '「バージョン管理画面：バージョン媒体出力釦押下」ログ出力
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_VERSION_OUTPUT_BUTTOM, 0)
'           Load frmVerOutput
'           frmVerOutput.Show 1
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
           '「バージョン管理画面：Ver一覧USB出力釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_VER_USB_OUTPUT_BUTTOM, 0)
           Call cmdVer_Output
        'V1.20.0.1 ADD END
        'V1.6.0.1 ADD END
        'V30.1.0.1 ADD START
        Case 9                                 'バージョン管理（新幹線改札機）
            '「バージョン管理画面：新幹線自改釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EG30JIKAI_BUTTOM, 0)
            gStrCurrentForm = sFormName_EG30JVer
            Load frmKansenGateVerKanri
            frmKansenGateVerKanri.Show 1
        'V30.1.0.1 ADD END
    
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
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
Private Sub cmdReturn_Click()
    
    On Error Resume Next
    
    '「バージョン管理画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_GAMEN_END, 0)
    Unload Me
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 ・IDU縮退時非表示処理、監視ファーム、RASマイコンバージョン処理追加
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 ・フェーズ３対応　ＲＡＳマイコン表示不要のため削除
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK対応
'//     REVISIONS :(1.20.0.1) 2010-03-17  REVISED BY [TCC] S.Yamazaki
'//                 ラベルへの表示を、リストへの表示に変更
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psGetVersion()
  Dim sVersion  As String 'V1.4.0.1 ADD
  Dim sGetJikiVer As String 'V1.10.0.1 ADD
 
 '監視盤、EG-R全体バージョン取得
  psKansiGetVersion
 'IDU全体バージョン取得
 'psIDUGetVersion       'V1.4.0.1　DEL
 
 'V1.4.0.1　ADD　START
 If pbIDUSts = 1 Then
    'IDUバージョン非表示
    lblVerName(2).Enabled = False
    lblVerName(3).Caption = ""
 Else
    '非縮退時は表示処理
    psIDUGetVersion
 End If
 'V1.4.0.1　ADD　END
 
 'LDU全体バージョン取得
  psLDUVersion

'V1.4.0.1　DEL　START
' 'EG-R自改バージョン取得
'  psEGRJVersion
' 'NEG自改バージョン取得
'  psNEGJVersion
'V1.4.0.1　DEL　END

'V1.20.0.1 ADD START
Call psListVersion
'V1.20.0.1 ADD END

'V1.20.0.1 DEL START
''V1.4.0.1　ADD START
' 'EG-R自改バージョン取得
'  '判定CPU
'  sVersion = psEGRJVersion(HANTEI_CPU)
'  lblVerName(11).Caption = sVersion
'  'メインCPU
'  sVersion = psEGRJVersion(MAIN_CPU)
'  lblVerName(12).Caption = sVersion
' 'サブCPU
'  sVersion = psEGRJVersion(SUB_CPU)
'  lblVerName(13).Caption = sVersion
' 'メインOS
'  sVersion = psEGRJVersion(MAIN_OS)
'  lblVerName(14).Caption = sVersion
' '予備１
'  sVersion = psEGRJVersion(YOBI1)
'  lblVerName(15).Caption = sVersion
' '予備２
'  sVersion = psEGRJVersion(YOBI2)
'  lblVerName(16).Caption = sVersion
' 'バージョンチェック
'  sVersion = psEGRJVersion(VER_CHK)
'  lblVerName(17).Caption = sVersion
'
' 'NEG自改バージョン取得
'  sVersion = psNEGJVersion
'  lblVerName(20).Caption = sVersion
''V1.4.0.1　ADD END
'
' 'IC-Mバージョン取得
' 'psICMGetVersion     'V1.4.0.1　DEL
' 'V1.4.0.1　ADD　START
' If pbIDUSts = 1 Then
'    'IDUバージョン非表示
'    lblVerName(21).Enabled = False
'    lblVerName(31).Caption = ""
' Else
'    '非縮退時は表示処理
'    sVersion = psICMGetVersion
'    lblVerName(31).Caption = sVersion
' End If
' 'V1.4.0.1　ADD　END
'
' '共通運賃バージョン取得
' 'psICUnchinGetVersion  'V1.4.0.1　DEL
' 'V1.4.0.1　ADD　START
' If pbIDUSts = 1 Then
'    'IDUバージョン非表示
'    lblVerName(22).Enabled = False
'    lblVerName(33).Caption = ""
' Else
'    '非縮退時は表示処理
'    sVersion = psICUnchinGetVersion
'    lblVerName(33).Caption = sVersion
' End If
'
' '監視ファームバージョン表示処理
' sVersion = psKansiFirmVersion
' lblVerName(25).Caption = sVersion
'
''V1.10.0.1 ADD START
' '磁気運賃読み込み
' sGetJikiVer = psJikiUnchinVersion
' lblVerName(27).Caption = CStr(sGetJikiVer)
''V1.10.0.1 ADD END
'
' 'V1.6.0.1 DEL START
' ''RASマイコンバージョン表示処理
' 'sVersion = psRASMICOMVersion
' 'lblVerName(27).Caption = sVersion
' 'V1.6.0.1 DEL END
'
' 'V1.4.0.1　ADD　END
'V1.20.0.1 DEL END
 
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
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
        lblVerName(1).Caption = Left$(strKansiVersion2, lSts)
    Else
        'バージョン番号の取得異常の場合、「--,--,--,--」を表示
        lblVerName(1).Caption = "--.--.--.-- "
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
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
        lblVerName(3).Caption = "--.--.--.--"
    Else
       '全体バージョン文字列作成
        lblVerName(3).Caption = Trim(strWork)
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
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
        lblVerName(5).Caption = "--.--.--.--"
    Else
       '全体バージョン文字列作成
        lblVerName(5).Caption = Trim(strWork)
    End If
      
   'LDユーティリティバージョン管理ファイルをクローズ。
    Close #iFileNumber

End Function

'V1.4.0.1 DEL START
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : psEGRJVersion
''//  機能名称  : EG-R自動改札機バージョン取得処理
''//  機能概要  : GATEVER_FILE.INIファイルより、代表ファイル名を取得し、
''//              代表ファイルよりバージョンを取得する。
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Public Function psEGRJVersion()
'    Dim strWork     As String                         '作業エリア
'    Dim iFileNumber As Integer                        '未使用ファイル番号
'    Dim lSts As Long                                  '関数戻り値
'    Dim strVerFileName As String * VERSION_GATE_SIZE  '取得ファイル名
'    Dim sGetVer     As String                         '作業エリア
'    Dim lngErrCode As Long
'
'    On Error Resume Next
'
'    ' GATEVER_FILE.INIから判定データCPU-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_HANTEI_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EHAN1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 11
'    Else
'       '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(11).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIからメインCPU-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_MAIN_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EPRO1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 12
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(12).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIからサブCPU-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_SUB_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_ESCPUNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 13
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(13).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIからメインOS-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_MAIN_OS, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EOSNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 14
'    Else
'       '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(14).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIから予備1の代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_YOBI1, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'     If lSts > 0 Then
'       strWork = E_EYOBI1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'       psJVerGet strWork, 15
'     Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(15).Caption = "--"
'     End If
'
'    ' GATEVER_FILE.INIから予備2の代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_YOBI2, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EYOBI2NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 16
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(16).Caption = "--"
'    End If
'End Function
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : psNEGJVersion
''//  機能名称  : NEG自動改札機バージョン取得処理
''//  機能概要  : GATEVER_FILE.INIファイルより、代表ファイル名を取得し、
''//              代表ファイルよりバージョンを取得する。
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Public Function psNEGJVersion()
'    Dim strWork     As String                         '作業エリア
'    Dim iFileNumber As Integer                        '未使用ファイル番号
'    Dim lSts As Long                                  '関数戻り値
'    Dim strVerFileName As String * VERSION_GATE_SIZE  '取得ファイル名
'    Dim sGetVer     As String                         '作業エリア
'    Dim lngErrCode As Long
'
'     On Error Resume Next
'
'    ' GATEVER_FILE.INIから判定データCPU-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_HANTEI_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NHAN1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 17
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(17).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIからメインCPU-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_MAIN_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NPRO1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 18
'    Else
'       '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(18).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIからサブCPU-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_SUB_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NSCPUNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 19
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(19).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIからメインOS-PROの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_MAIN_OS, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NOSNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 20
'    Else
'    '「バージョン管理画面：バージョン取得異常」ログ出力
'     lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'     Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'     lblVerName(20).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIから予備1の代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_YOBI1, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NYOBI1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 21
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(21).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INIから予備2の代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_YOBI2, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NYOBI2NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 22
'    Else
'       '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(22).Caption = "--"
'    End If
'End Function

''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : psJVerGet
''//  機能名称  : 代表ファイルのバージョンを取得
''//  機能概要  : 代表ファイルのバージョンを取得し、画面表示する。
''//
''//              型        名称      意味
''//  引数      : String　　sPath　　[IN]代表ファイル名
''//  　　      : Integer　 iIndex　 [IN]表示インデックス番号
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Function psJVerGet(sPath As String, iCnt As Integer)
'
'    Dim i As Integer                    'カウンタ
'    Dim j As Integer                    'カウンタ
'    Dim iFileNumber As Integer          'ファイル番号
'    Dim lLen As Long                    'ファイルサイズ
'    Dim uFooter As MN_FOOT              'フッタ情報格納エリア
'    Dim lPos As Long                    'バージョン情報格納位置
'    Dim sDateTime As String
'    Dim lngErrCode As Long              'エラーコード
'
'On Error GoTo FileGetError
'
'    If Dir(sPath) <> "" Then            'ファイルが存在する?
'
'      lLen = FileLen(sPath)             'ファイルサイズの取得
'
'      iFileNumber = FreeFile            '未使用のファイル番号を取得する
'
'      'ファイルのオープン
'      Open sPath For Binary Access Read As #iFileNumber
'
'      'フッタ情報の取得
'      Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
'
'      'バージョン値を表示
'      lblVerName(iCnt).Caption = CStr(uFooter.sVersion)
'      Close #iFileNumber                  'ファイルを閉じます
'    Else
'      'ファイルが存在しない。
'      lblVerName(iCnt).Caption = "--"
'    End If
' Exit Function
'
'FileGetError:
'   '「バージョン管理画面：バージョン取得異常」ログ出力
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'   lblVerName(iCnt).Caption = "--"
'   Close #iFileNumber                  'ファイルを閉じます
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : psICMGetVersion
''//  機能名称  : IC-Mバージョン取得処理
''//  機能概要  : GATEVER_FILE.INIファイルより、代表ファイル名を取得し、
''//              代表ファイルよりバージョンを取得する。
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Public Function psICMGetVersion()
'    Dim strWork     As String                         '作業エリア
'    Dim iFileNumber As Integer                        '未使用ファイル番号
'    Dim lSts As Long                                  '関数戻り値
'    Dim strVerFileName As String * VERSION_GATE_SIZE  '取得ファイル名
'    Dim sGetVer     As String                         '作業エリア
'    Dim lngErrCode As Long
'
'    On Error Resume Next
'    strWork = ""
'
'    ' GATEVER_FILE.INIから判定IC-Mデータの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_ICM, _
'                                   GATE_ICM, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = PATH_IDU_APP & PATH_IDU_IC_M & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psIDUVerGet strWork, 31
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(31).Caption = "--------------------"
'    End If
'
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : psICUnchinGetVersion
''//  機能名称  : IC共通運賃バージョン取得処理
''//  機能概要  : kansi.iniファイルより、代表ファイル名を取得し、
''//              代表ファイルよりバージョンを取得する。
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Public Function psICUnchinGetVersion()
'    Dim strWork     As String                        '作業エリア
'    Dim iFileNumber As Integer                       '未使用ファイル番号
'    Dim lSts As Long                                 '関数戻り値
'    Dim strVerFileName As String * VERSION_GATE_SIZE '取得ファイル名
'    Dim sGetVer     As String                        '作業エリア
'    Dim lngErrCode As Long
'
'    strWork = ""
'
'    ' 監視盤設置構成INIファイル(kansi.ini)よりIC共通運賃データの代表ファイル名を取得する。
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(IDU_KANSI_SECTION_NAME, _
'                                   IDU_KANSI_KEY_NAME, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_IDU_APP & IDU_KANSI_INI)
'    If lSts > 0 Then
'    strWork = PATH_IDU_APP & PATH_IDU_ICUNCHIN & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psIDUVerGet strWork, 33
'    Else
'      '「バージョン管理画面：バージョン取得異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(33).Caption = "--------------------"
'    End If
'
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : psIDUVerGet
''//  機能名称  : 代表ファイルのバージョンを取得
''//  機能概要  : 代表ファイルのバージョンを取得し、画面表示する。
''//
''//              型        名称      意味
''//  引数      : String　　sPath　　[IN]代表ファイル名
''//  　　      : Integer　 iIndex　 [IN]表示インデックス番号
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Function psIDUVerGet(sPath As String, iCnt As Integer)
'
'    Dim i As Integer                    'カウンタ
'    Dim j As Integer                    'カウンタ
'    Dim sMyName As String               'ファイル名
'    Dim iFileNumber As Integer          'ファイル番号
'    Dim lLen As Long                    'ファイルサイズ
'    Dim uFooter As MN_IDU_FOOT          'フッタ情報格納エリア
'    Dim lPos As Long                    'バージョン情報格納位置
'    Dim sDateTime As String
'    Dim lngErrCode As Long              'エラーコード
'
'On Error GoTo FileGetError
'
'    If Dir(sPath) <> "" Then            'ファイルが存在する?
'
'      lLen = FileLen(sPath)             'ファイルサイズの取得
'
'      iFileNumber = FreeFile            '未使用のファイル番号を取得する
'
'        'ファイルのオープン
'        Open sPath For Binary Access Read As #iFileNumber
'        'フッタ情報の取得
'        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
'
'        'データ名＋バージョンを表示
'        lblVerName(iCnt).Caption = CStr(uFooter.sDataName) & CStr(uFooter.sVersion)
'        Close #iFileNumber                  'ファイルを閉じます
'    Else
'      'ファイルが存在しない場合
'      lblVerName(iCnt).Caption = "--------------------"
'    End If
'
'    Exit Function
'FileGetError:
'   '「バージョン管理画面：バージョン取得異常」ログ出力
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'   lblVerName(iCnt).Caption = "--------------------"
'   Close #iFileNumber                  'ファイルを閉じます
'End Function
'V1.4.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psJikiCheck
'//  機能名称  : 磁気運賃対応ユーザチェック処理
'//  機能概要  : HOSHU.INIより、磁気運賃対応ユーザであるかどうかチェックする。
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
Public Sub psJikiCheck()
    Dim iFlag As Integer '取得ユーザフラグ
 
    On Error Resume Next
 
  ' HOSHU.INIより磁気運賃対応ユーザフラグを取得する。
    iFlag = GetPrivateProfileInt(KANS_JIKI, _
                                 KANSI_JIKI_FLAG, _
                                 DEFAILT_Int, _
                                 HOSHU_FILE)
     If iFlag = 0 Then
      'フラグが0の場合「磁気運賃」釦は非表示
      cmdFixedExe(7).Visible = False
     End If
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
        AppActivate frmVersion.Caption, False
        pfFormActive (frmVersion.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdVer_Output
'//  機能名称  : 「Ver一覧　USB出力」釦押下時処理
'//  機能概要  : Ver一覧のUSB出力。バージョン媒体出力frmの媒体出力と同一処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21 REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応【残件№54】
'//                  ・バージョン一覧出力ファイル名変更
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【03統合TR-No.18修正対応】
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function cmdVer_Output() As Boolean
    
    Dim sWriteDir As String                 '媒体出力先
    Dim bRet      As Boolean         '戻り値
    Dim lRetVal   As Long            'テキスト表示処理戻り値
    Dim sCommand  As String          'コマンド文字列
    Dim iResponse As Integer   'MsgBox戻り値
    Dim lngErrCode As Long     'エラーコード
    Dim fso         As New FileSystemObject   'ファイルシステムオブジェクト
    Dim strWriteDir As String               '出力先フォルダ
' EG20 V2.0.1.1 ADD START
    Dim strStationName As String    ' 駅名
    Dim strSrcName     As String    ' コピー元ファイルパス
' EG20 V2.0.1.1 ADD END
    
   On Error GoTo COPY_ERROR

    cmdVer_Output = False

    sWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    
    If sWriteDir <> "" Then

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       
       'ディレクトリが指定されれば、バージョンファイルを取出す
        bRet = dllEGRCreateVersionFile(PATH_IDU_APP, PATH_LDU_APP)
        If bRet = False Then

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
           'プログレスバーを消去する
           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
           
           '「ファイル作成異常」ポップアップ画面表示
'           MsgBox "ファイルの作成に失敗しました。", vbOKOnly + vbCritical, "ファイル作成異常"              ' EG20 V3.6.0.1【03統合TR-No.18修正対応】削除
           MsgBox "異常終了しました。", vbCritical, "Ver一覧USB出力"                                        ' EG20 V3.6.0.1【03統合TR-No.18修正対応】追加
           '「ファイル作成異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_FILE_CREATE_ERROR, 0)

           '「Ver一覧USB出力処理異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_KANRI_MENU_VER_USB_OUTPUT_ERROR, 0)
            Exit Function
        Else
          'ファイルコピー
' EG20 V2.0.1.1 DEL START
'          FileCopy EGR_KANSI_VERSION_FILE_PATH, sWriteDir & EGR_KANSI_VERSION_FILE
' EG20 V2.0.1.1 DEL END
' EG20 V2.0.1.1 ADD START
          '駅名取得
          strStationName = gsGetStationEkiName
          ' コピー元ファイルパス
          strSrcName = PATH_HOSHU_DATA & EGR_KANSI_VERSION_FILE
          
          FileCopy strSrcName, sWriteDir & strStationName & "_" & EGR_KANSI_VERSION_FILE
' EG20 V2.0.1.1 ADD START
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
          'プログレスバーを消去する
          Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

          '「媒体出力正常終了」ポップアップ画面表示
'          MsgBox "媒体出力は正常終了しました。", vbOKOnly + vbInformation, "媒体出力結果"          ' EG20 V3.6.0.1【03統合TR-No.18修正対応】削除
          MsgBox "正常終了しました。", vbOKOnly + vbInformation, "Ver一覧USB出力"                   ' EG20 V3.6.0.1【03統合TR-No.18修正対応】追加

          '「Ver一覧USB出力処理正常」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_KANRI_MENU_VER_USB_OUTPUT_OK, 0)
        End If
     Else
         '「Ver一覧USB出力処理未実行」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_VER_USB_OUTPUT_MISHORI, 0)
     End If
  cmdVer_Output = True
  
  Exit Function
COPY_ERROR:
   
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
   
   '処理異常の場合、出力結果ポップアップ(異常)表示
'    MsgBox "媒体出力は異常終了しました。", vbCritical, "媒体出力結果"                              ' EG20 V3.6.0.1【03統合TR-No.18修正対応】削除
    MsgBox "異常終了しました。", vbCritical, "Ver一覧USB出力"                                       ' EG20 V3.6.0.1【03統合TR-No.18修正対応】追加
   
   '「Ver一覧USB出力処理異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_VER_USB_OUTPUT_ERROR, lngErrCode)
   cmdVer_Output = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : psListVersion
'//  機能名称  : リスト表示
'//  機能概要  : 表示用バージョンファイルを読み込む
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-16   CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 【ＩＣＭバージョンファイルリスト対応】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-08  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion()
    
    Dim iCnt As Integer
    Dim iMax As Integer
    Dim structDispInfo() As DISP_FILE_INFO
    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
    Dim iFileNumber As Integer
    
    Dim sLine As String             '1行分のcsv読み取りデータ
    Dim sLineSplit() As String      '1単語ずつのcsv読み取りデータ
    
    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード
    Dim sErrEventName As String        'エラーが起きたイベント名
    
    'エラートラップ
    On Error GoTo Err_FILE
    
    'リストボックス初期化
    lstKan.Clear
    
    '存在チェック
    If fsoObj.FileExists(DISP_VERFILE_FILE) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        GoTo Err_FILE
    End If
    
    '未使用のファイル番号取得
    iFileNumber = FreeFile
    
    'ファイルをオープンする。
    sErrEventName = LOG_ERR_FILE_OPEN       'ファイルオープン異常
    Open DISP_VERFILE_FILE For Input As #iFileNumber
    
    iCnt = 0
    
    sErrEventName = LOG_ERR_FILE_READ       'ファイル読込異常
    Do While Not EOF(iFileNumber)
        
        '１ 行づつ変数読み込み
        Line Input #iFileNumber, sLine
        
        'コメント行と空行じゃなければ､領域に格納
        If Trim(Left(sLine, 1)) <> CSV_COMMENT_CHAR And sLine <> "" Then
            
            sLineSplit = Split(sLine, ",")
            
'            If UBound(sLineSplit) = 3 Then                           ' EG20 V5.6.0.1削除
            If UBound(sLineSplit) = 4 Then                            ' EG20 V5.6.0.1追加
            
                ReDim Preserve structDispInfo(iCnt)
                
                structDispInfo(iCnt).sTitle = sLineSplit(0)
                structDispInfo(iCnt).sFilePath = sLineSplit(1)
                structDispInfo(iCnt).iType = sLineSplit(2)
                structDispInfo(iCnt).iIdu = sLineSplit(3)
                structDispInfo(iCnt).iMaker = sLineSplit(4)           ' EG20 V5.6.0.1追加
                
                iCnt = iCnt + 1
                
            End If
            
        End If
    Loop
    
    'ファイルをクローズする。
    sErrEventName = LOG_ERR_FILE_CLOSE      'ファイルクローズ異常
    Close #iFileNumber
    
    iMax = iCnt - 1
    
    '表示代表ファイルのエラートラップ（エラーがあっても処理は続く）
    On Error Resume Next
    
    'IDU縮退チェック
    Call psIDUCheck
    
    For iCnt = 0 To iMax
        
        '縮退機能フラグがなし、または縮退中ではないときのみ処理を行う
        If structDispInfo(iCnt).iIdu = 0 Or pbIDUSts = 0 Then
        
            Select Case structDispInfo(iCnt).iType
                Case 1
                    Call psListVersion_Type1(structDispInfo(iCnt))
                Case 2
                    Call psListVersion_Type2(structDispInfo(iCnt))
' EG20 V2.1.0.1[Mainte_03_01]追加開始
                Case 3
                    Call psListVersion_Type3(structDispInfo(iCnt))
' EG20 V2.1.0.1[Mainte_03_01]追加終了
'EG20 V30.1.0.1 ADD START
                Case 4
                    Call psListVersion_Type4(structDispInfo(iCnt))
'EG20 V30.1.0.1 ADD END
                Case Else
                    '処理なし
            End Select
        End If
    Next
    
    Set fsoObj = Nothing

    Exit Sub

Err_FILE:

    '異常ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    'ログ出力　┗ファイル名
    Call psFileNameGet(DISP_VERFILE_FILE, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    'ファイルクローズ
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    
    Set fsoObj = Nothing

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : psListVersion_Type1
'//  機能名称  : リスト表示
'//  機能概要  : 表示タイプ１の表示を行う（監視盤、RYT）
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-16   CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 【バージョン表示不正対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type1(structDispInfo As DISP_FILE_INFO)
    
    Dim lLen As Long
    Dim sDateTime As String
    Dim j As Integer
    Dim uFooter As MN_FOOT              'フッタ情報格納エリア
    Dim sTitle As String                '後ろ空白埋めしたタイトル
    Dim sDisp As String                 '表示用
    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
    Dim iFileNumber As Integer

    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード
    Dim sErrEventName As String        'エラーが起きたイベント名

    Dim bRet As Boolean                 ' 戻り値
    Dim szFileName As String            ' ファイル名
    Dim uVersion As MN_VERSION_JIKAI    ' バージョン情報格納エリア

    'エラートラップ
    On Error GoTo Err_FILE

    'タイトルの加工
    sTitle = structDispInfo.sTitle
    'タイトル後のスペース（全角の可能性があるのでFormatは使えない）
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If

    'ファイルの存在チェック。異常時は----の表示
    If fsoObj.FolderExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If

    ' ファイルリストからファイルリストの作成
    bRet = fReadFileList(structDispInfo.sFilePath & "\" & MN_FILELIST)
    If bRet <> True Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If

    szFileName = structDispInfo.sFilePath & "\" & FileList(0)   ' ファイルリストからバージョン情報を取得する
    If fsoObj.FileExists(szFileName) = True Then                ' ファイルが存在する?
        lLen = FileLen(szFileName)                              ' ファイルサイズの取得

        iFileNumber = FreeFile                                  ' 未使用のファイル番号を取得する

        Open szFileName For Binary Access Read As #iFileNumber  ' ファイルのオープン
        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter      ' フッタ情報の取得
        uVersion.sFileName = UCase(FileListType(0))             ' ファイル名を大文字にしてセット
        uVersion.sMachineName = uFooter.sKisyu                  ' 機種名セット
        uVersion.sFooterFile = uFooter.sFileName                ' ファイル名セット

        sDateTime = ""
        For j = 0 To 3
            sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
        Next
        sDateTime = sDateTime & " "
        For j = 4 To 5
            sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
        Next
        uVersion.sFileDate = sDateTime
        uVersion.sVersion = uFooter.sVersion                    ' バージョン情報セット
        uVersion.sComment = uFooter.sHyoji                      ' 表示文字列セット

        Close #iFileNumber                  'ファイルを閉じます
    End If
    
    'バージョン情報格納エリアの拡張
    sDisp = sTitle                                                                  ' タイトル
'    sDisp = sDisp & Format(Right(FileList(0), 12), "!@@@@@@@@@@@@") & Space(11)     ' ファイル名   ' EG20 V6.1.0.1削除
    sDisp = sDisp & Format(Right(FileList(0), 12), "!@@@@@@@@@@@@") & Space(7)      ' ファイル名    ' EG20 V6.1.0.1追加
    sDisp = sDisp & uFooter.sKisyu & Space(1)                                       ' 機種名
    sDisp = sDisp & uFooter.sFileName & Space(2)                                    ' ファイル
    sDisp = sDisp & sDateTime & Space(1)                                            ' 作成日時
    sDisp = sDisp & uFooter.sVersion                                                ' バージョン
    lstKan.AddItem (sDisp)

    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 32), vbUnicode)     ' コメント1
    lstKan.AddItem (Space(45) & sDisp)

    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 33, 64), vbUnicode)    ' コメント2
    lstKan.AddItem (Space(45) & sDisp)

    Set fsoObj = Nothing
    Exit Sub
    
Err_FILE:

    '異常ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    'ログ出力　┗ファイル名
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)

    'ファイルクローズ
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If

    '異常用表示
' EG20 V6.1.0.1削除開始
'    lstKan.AddItem (sTitle & "------------           -------- --------  -------- ---- --")
' EG20 V6.1.0.1削除終了
' EG20 V6.1.0.1追加開始
    lstKan.AddItem (sTitle & "------------       -------- --------  -------- ---- --")
' EG20 V6.1.0.1追加終了
    lstKan.AddItem (Space(45) & "--------------------------------")
    lstKan.AddItem (Space(45) & "--------------------------------")

    Set fsoObj = Nothing

End Sub

' EG20 V2.1.0.1[Mainte_03_01]削除開始（表示内容変更）
'Private Sub psListVersion_Type1(structDispInfo As DISP_FILE_INFO)
'
'    Dim lLen As Long
'    Dim sDateTime As String
'    Dim j As Integer
'    Dim uFooter As MN_FOOT              'フッタ情報格納エリア
'    Dim sTitle As String                '後ろ空白埋めしたタイトル
'    Dim sDisp As String                 '表示用
'    Dim sDispFile As String             '作業用
'    Dim sDispExe As String              '拡張子
'    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
'    Dim iFileNumber As Integer
'
'    Dim sErrFile As String             'エラーログ用INIファイル名
'    Dim sErrExe As String              'エラーログ用INI拡張子
'    Dim lngErrCode As Long             'エラーコード
'    Dim sErrEventName As String        'エラーが起きたイベント名
'
'    'エラートラップ
'    On Error GoTo Err_FILE
'
'    'タイトルの加工
'    sTitle = structDispInfo.sTitle
'    'タイトル後のスペース（全角の可能性があるのでFormatは使えない）
'    If LenB(StrConv(sTitle, vbFromUnicode)) < 20 Then
'        sTitle = sTitle & Space(20 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
'    Else
'        sTitle = sTitle & Space(2)
'    End If
'
'    'ファイルの存在チェック。異常時は----の表示
'    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
'        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
'        '異常
'        GoTo Err_FILE
'    End If
'
'    lLen = FileLen(structDispInfo.sFilePath)              'ファイルサイズの取得
'    If lLen < Len(uFooter) Then
'        sErrEventName = LOG_ERR_FILE_LENGTH     'ファイルレングス異常
'        '異常
'        GoTo Err_FILE
'    End If
'
'    '未使用のファイル番号取得
'    iFileNumber = FreeFile
'
'    'ファイルのオープン
'    sErrEventName = LOG_ERR_FILE_OPEN       'ファイルオープン異常
'    Open structDispInfo.sFilePath For Binary Access Read As #iFileNumber
'
'        sErrEventName = LOG_ERR_FILE_READ       'ファイル読込異常
'        'フッタ情報の取得
'        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
'
'    sErrEventName = LOG_ERR_FILE_CLOSE      'ファイルクローズ異常
'    Close #iFileNumber      'ファイルのクローズ
'
'    '作成日時の加工
'    sDateTime = ""
'    For j = 0 To 3
'        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
'    Next
'    sDateTime = sDateTime & " "
'    For j = 4 To 5
'        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
'    Next
'
'    'ファイル名の加工
'    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             'ファイルパスからファイル名を取得
'    sDispFile = UCase(sDispFile & "." & sDispExe)                                 '拡張子を結合し大文字に変換
'
'    'バージョン情報格納エリアの拡張
'    sDisp = sTitle                                                                  'タイトル
'    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(11)       'ファイル名
'    sDisp = sDisp & uFooter.sKisyu & Space(1)                                       '機種名
'    sDisp = sDisp & uFooter.sFileName & Space(2)                                    'ファイル
'    sDisp = sDisp & sDateTime & Space(1)                                            '作成日時
'    sDisp = sDisp & uFooter.sVersion                                                'バージョン
'    lstKan.AddItem (sDisp)
'
'    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 32), vbUnicode)     'コメント1
'    lstKan.AddItem (Space(45) & sDisp)
'
'    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 33, 64), vbUnicode)    'コメント2
'    lstKan.AddItem (Space(45) & sDisp)
'
'    Set fsoObj = Nothing
'
'    Exit Sub
'Err_FILE:
'
'    '異常ログ出力
'    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
'    'ログ出力　┗ファイル名
'    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
'    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
'
'    'ファイルクローズ
'    If iFileNumber > 0 Then
'        Close #iFileNumber
'    End If
'
'    '異常用表示
'    lstKan.AddItem (sTitle & "------------           -------- --------  -------- ---- --")
'    lstKan.AddItem (Space(45) & "--------------------------------")
'    lstKan.AddItem (Space(45) & "--------------------------------")
'
'    Set fsoObj = Nothing
'
'End Sub
' EG20 V2.1.0.1[Mainte_03_01]削除終了

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : psListVersion_Type2
'//  機能名称  : リスト表示
'//  機能概要  : 表示タイプ２の表示を行う（IDU）
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-16   CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 【ＩＣＭバージョンファイルリスト対応】
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 【バージョン表示不正対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type2(structDispInfo As DISP_FILE_INFO)

    Dim lLen As Long
    Dim sDateTime As String
    Dim j As Integer
    Dim uFooter As MN_IDU_FOOT          'フッタ情報格納エリア
    Dim sTitle As String                '後ろ空白埋めしたタイトル
    Dim sDisp As String                 '表示用
    Dim sDispFile As String             '作業用（ファイル表示用）
    Dim sDispDV As String               '作業用（データ＋バージョン）
    Dim sDispCom As String              '作業用（コメント）
    Dim sDispExe As String              '拡張子
    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
    Dim iFileNumber As Integer
    
    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード
    Dim sErrEventName As String        'エラーが起きたイベント名

' EG20 V5.6.0.1 追加開始
    Dim bRet As Boolean                 ' 戻り値
    Dim szFilelist As String            ' ファイルリスト名
    Dim szFileName As String            ' ファイル名
' EG20 V5.6.0.1 追加終了
    
    'エラートラップ
    On Error GoTo Err_FILE
    
    'パスを作り直し
    structDispInfo.sFilePath = PATH_IDU_APP & "\" & structDispInfo.sFilePath

    'タイトルの加工
    sTitle = structDispInfo.sTitle
    'タイトル後のスペース（全角の可能性があるのでFormatは使えない）
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If

' EG20 V5.6.0.1 追加開始

    ' ファイルリストからファイルリストの作成
    szFilelist = structDispInfo.sFilePath & "\FILELIST_" & Format(structDispInfo.iMaker, "00") & ".TXT"

    ' ファイルの存在チェック。異常時は----の表示
    If fsoObj.FileExists(szFilelist) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If

    bRet = fReadFileListIDU(szFilelist)
    If bRet <> True Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If
    structDispInfo.sFilePath = structDispInfo.sFilePath & "\" & FileList(0)   ' ファイルリストからバージョン情報を取得する

' EG20 V5.6.0.1 追加終了
    
    '--------------------------------------------
    'ファイル情報取得
    '--------------------------------------------
    'ファイルの存在チェック。異常時は----の表示
    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If
    
    lLen = FileLen(structDispInfo.sFilePath)              'ファイルサイズの取得
    If lLen < Len(uFooter) Then
        sErrEventName = LOG_ERR_FILE_LENGTH     'ファイルレングス異常
        '異常
        GoTo Err_FILE
    End If
    
    '未使用のファイル番号取得
    iFileNumber = FreeFile
    
    'ファイルのオープン
    sErrEventName = LOG_ERR_FILE_OPEN       'ファイルオープン異常
    Open structDispInfo.sFilePath For Binary Access Read As #iFileNumber
    
        sErrEventName = LOG_ERR_FILE_READ       'ファイル読込異常
        'フッタ情報の取得
        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
        
    sErrEventName = LOG_ERR_FILE_CLOSE      'ファイルクローズ異常
    Close #iFileNumber      'ファイルのクローズ
    
    '--------------------------------------------
    'バージョン情報表示部の表示テキスト作成
    '--------------------------------------------
    'タイトル
    sDisp = sTitle
    
    'ファイル名
    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             'ファイルパスからファイル名を取得
    sDispFile = sDispFile & "." & sDispExe                                        '拡張子を結合し大文字に変換
'    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(11)      ' EG20 V6.1.0.1削除
    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(7)        ' EG20 V6.1.0.1追加

    '種別
    sDisp = sDisp & LCase(Right$("0" & Hex(uFooter.bSyubetu), 2))
    
    'メーカ名
    sDisp = sDisp & uFooter.sMakerName & Space(2)
    
    'データ名＋バージョン
    sDispDV = LTrim(uFooter.sDataName) & uFooter.sVersion
    If Len(Trim(uFooter.sDataName)) = 0 And Len(Trim(uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(Trim(uFooter.sVersion) & Space(20), 20) & Space(2)
    ElseIf Len(Trim(uFooter.sDataName & uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(sDispDV & Space(20), 20) & Space(2)
    Else
        sDisp = sDisp & String(20, "-") & Space(2)
    End If
    
    '作成日時
    sDateTime = ""
    For j = 0 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    sDateTime = Format(sDateTime, "@@@@/@@/@@ @@:@@")
    sDisp = sDisp & sDateTime
    
    'リストに追加
    lstKan.AddItem (sDisp)
    
    'コメント
    '60文字で保存されているので、60バイトの領域に直す。前後の空白を取る。
    sDispCom = Trim(StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 60), vbUnicode))
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 1, 32), vbUnicode)     'コメント1
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    Else
        lstKan.AddItem (Space(45) & String(32, "-"))
    End If
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 33, 60), vbUnicode)    'コメント2
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    End If
    
    Set fsoObj = Nothing
    
    Exit Sub
Err_FILE:
    
    '異常ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    'ログ出力　┗ファイル名
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    'ファイルクローズ
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    
    '異常用表示
' EG20 V6.1.0.1削除開始
'    lstKan.AddItem (sTitle & "------------           ---  --------------------  ----/--/-- --:--")
' EG20 V6.1.0.1削除終了
' EG20 V6.1.0.1追加開始
    lstKan.AddItem (sTitle & "------------       ---  --------------------  ----/--/-- --:--")
' EG20 V6.1.0.1追加終了
    lstKan.AddItem (Space(45) & "--------------------------------")
    lstKan.AddItem (Space(45) & "--------------------------------")

    Set fsoObj = Nothing

End Sub
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : psListVersion_Type3
'//  機能名称  : リスト表示
'//  機能概要  : 表示タイプ３の表示を行う（操作卓）
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.56修正対応】
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 【バージョン表示不正対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type3(structDispInfo As DISP_FILE_INFO)

'    Dim lLen As Long                                                 ' EG20 V5.0.2.1削除
'    Dim sDateTime As String                                          ' EG20 V5.0.2.1削除
'    Dim j As Integer                                                 ' EG20 V5.0.2.1削除
'    Dim uFooter As MN_IDU_FOOT          'フッタ情報格納エリア        ' EG20 V5.0.2.1削除
    Dim sTitle As String                '後ろ空白埋めしたタイトル
    Dim sDisp As String                 '表示用
    Dim sDispFile As String             '作業用（ファイル表示用）
'    Dim sDispCom As String              '作業用（コメント）          ' EG20 V5.0.2.1削除
    Dim sDispExe As String              '拡張子
    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
'    Dim iFileNumber As Integer                                       ' EG20 V5.0.2.1削除
    
    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード
    Dim sErrEventName As String        'エラーが起きたイベント名
    
'    Dim FsoRead As TextStream                                        ' EG20 V5.0.2.1削除
'    Dim bFileOpen As Boolean            ' オープンフラグ             ' EG20 V5.0.2.1削除
'    Dim strBuffer As String             ' リードバッファ             ' EG20 V5.0.2.1削除
    Dim strVersion As String            ' バージョン文字列

' EG20 V5.0.2.1【結合TR-No.56修正対応】追加開始
    Dim lSts As Long                                       '関数戻り値
    Dim strKansiVersion2 As String * VERSION_GATE_SIZE     '監視装置全体バージョン
' EG20 V5.0.2.1【結合TR-No.56修正対応】追加終了

    
    'エラートラップ
    On Error GoTo Err_FILE
    
    ' 初期化
'    bFileOpen = False                                                ' EG20 V5.0.2.1削除
    
    
    'タイトルの加工
    sTitle = structDispInfo.sTitle
    'タイトル後のスペース（全角の可能性があるのでFormatは使えない）
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If
    
    '--------------------------------------------
    'ファイル情報取得
    '--------------------------------------------
    'ファイルの存在チェック。異常時は----の表示
    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If
    
' EG20 V5.0.2.1【結合TR-No.56修正対応】削除開始
'    Set FsoRead = fsoObj.OpenTextFile(structDispInfo.sFilePath, ForReading)
'    bFileOpen = True
'    ' ファイルから１行リード
'    strBuffer = FsoRead.ReadLine
'    strVersion = Trim(strBuffer)
'
'    FsoRead.Close
'    Set FsoRead = Nothing
'    Set fsoObj = Nothing
' EG20 V5.0.2.1【結合TR-No.56修正対応】削除終了

' EG20 V5.0.2.1【結合TR-No.56修正対応】追加開始
    Set fsoObj = Nothing
    strKansiVersion2 = ""
    strVersion = ""
    ' KansiVersion.iniから操作卓の全体バージョンを取得し表示する
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion2, _
                                   Len(strKansiVersion2), _
                                   structDispInfo.sFilePath)
     If lSts > 0 Then
        '取得したバージョン番号を表示
        strVersion = Left$(strKansiVersion2, lSts)
    End If

' EG20 V5.0.2.1【結合TR-No.56修正対応】追加終了

    '--------------------------------------------
    'バージョン情報表示部の表示テキスト作成
    '--------------------------------------------
    'タイトル
    sDisp = sTitle
    
    'ファイル名
    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             'ファイルパスからファイル名を取得
    
    sDispFile = sDispFile & "." & sDispExe                                        '拡張子を結合し大文字に変換
    
' EG20 V5.0.2.1【ファイル名は表示しない】削除開始
'    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(11)
' EG20 V5.0.2.1【ファイル名は表示しない】削除終了
' EG20 V5.0.2.1【ファイル名は表示しない】追加開始
'    sDisp = sDisp & Space(12) & Space(11)                                       ' EG20 V6.1.0.1削除
    sDisp = sDisp & Space(12) & Space(7)                                         ' EG20 V6.1.0.1追加
' EG20 V5.0.2.1【ファイル名は表示しない】追加終了

    'バージョン
    If Len(strVersion) <> 0 Then
        sDisp = sDisp & Format(Left(strVersion, 11), "!@@@@@@@@@@@")
    Else
        sDisp = sDisp & "--.--.--.--"
    End If
    
    'リストに追加
    lstKan.AddItem (sDisp)
    
    Exit Sub
Err_FILE:
    
    '異常ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    'ログ出力　┗ファイル名
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    '異常用表示
'    lstKan.AddItem (sTitle & "                       --.--.--.--")             ' EG20 V6.1.0.1削除
    lstKan.AddItem (sTitle & "                   --.--.--.--")                  ' EG20 V6.1.0.1追加
' EG20 V5.0.2.1【結合TR-No.56修正対応】削除開始
'    If bFileOpen = True Then
'        FsoRead.Close
'    End If
'    Set FsoRead = Nothing
' EG20 V5.0.2.1【結合TR-No.56修正対応】削除終了
    Set fsoObj = Nothing

End Sub
'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : psListVersion_Type4
'//  機能名称  : リスト表示
'//  機能概要  : 表示タイプ４の表示を行う（幹線自改）
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-05-08   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type4(structDispInfo As DISP_FILE_INFO)

    Dim lLen As Long
    Dim sDateTime As String
    Dim j As Integer
    Dim uFooter As MN_IDU_FOOT          'フッタ情報格納エリア(新幹線改札機はIDUのフッタフォーマットと同じなので）
    Dim sTitle As String                '後ろ空白埋めしたタイトル
    Dim sDisp As String                 '表示用
    Dim sDispFile As String             '作業用（ファイル表示用）
    Dim sDispDV As String               '作業用（データ＋バージョン）
    Dim sDispCom As String              '作業用（コメント）
    Dim sDispExe As String              '拡張子
    Dim fsoObj As New FileSystemObject  'ファイルシステムオブジェクト
    Dim iFileNumber As Integer
    
    Dim sErrFile As String             'エラーログ用INIファイル名
    Dim sErrExe As String              'エラーログ用INI拡張子
    Dim lngErrCode As Long             'エラーコード
    Dim sErrEventName As String        'エラーが起きたイベント名

    Dim bRet As Boolean                 ' 戻り値
    
    'エラートラップ
    On Error GoTo Err_FILE
    
    'タイトルの加工
    sTitle = structDispInfo.sTitle
    'タイトル後のスペース（全角の可能性があるのでFormatは使えない）
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If

    ' ファイルの存在チェック。異常時は----の表示
    If fsoObj.FolderExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If

    bRet = fReadFileList(structDispInfo.sFilePath & "\" & MN_FILELIST)
    If bRet <> True Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If
    structDispInfo.sFilePath = structDispInfo.sFilePath & "\" & FileList(0)   ' ファイルリストからバージョン情報を取得する

    '--------------------------------------------
    'ファイル情報取得
    '--------------------------------------------
    'ファイルの存在チェック。異常時は----の表示
    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     'ファイル無し
        '異常
        GoTo Err_FILE
    End If
    
    lLen = FileLen(structDispInfo.sFilePath)              'ファイルサイズの取得
    If lLen < Len(uFooter) Then
        sErrEventName = LOG_ERR_FILE_LENGTH     'ファイルレングス異常
        '異常
        GoTo Err_FILE
    End If
    
    '未使用のファイル番号取得
    iFileNumber = FreeFile
    
    'ファイルのオープン
    sErrEventName = LOG_ERR_FILE_OPEN       'ファイルオープン異常
    Open structDispInfo.sFilePath For Binary Access Read As #iFileNumber
    
        sErrEventName = LOG_ERR_FILE_READ       'ファイル読込異常
        'フッタ情報の取得
        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
        
    sErrEventName = LOG_ERR_FILE_CLOSE      'ファイルクローズ異常
    Close #iFileNumber      'ファイルのクローズ
    
    '--------------------------------------------
    'バージョン情報表示部の表示テキスト作成
    '--------------------------------------------
    'タイトル
    sDisp = sTitle
    
    'ファイル名
    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             'ファイルパスからファイル名を取得
    sDispFile = sDispFile & "." & sDispExe                                        '拡張子を結合し大文字に変換
    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(7)

    '種別
    sDisp = sDisp & LCase(Right$("0" & Hex(uFooter.bSyubetu), 2))
    
    'メーカ名
    sDisp = sDisp & uFooter.sMakerName & Space(2)
    
    'データ名＋バージョン
    uFooter.sDataName = Replace(uFooter.sDataName, vbNullChar, Space(1))
    sDispDV = LTrim(uFooter.sDataName) & uFooter.sVersion
    If Len(Trim(uFooter.sDataName)) = 0 And Len(Trim(uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(Trim(uFooter.sVersion) & Space(20), 20) & Space(2)
    ElseIf Len(Trim(uFooter.sDataName & uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(sDispDV & Space(20), 20) & Space(2)
    Else
        sDisp = sDisp & String(20, "-") & Space(2)
    End If
    
    '作成日時
    sDateTime = ""
    For j = 0 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    sDateTime = Format(sDateTime, "@@@@/@@/@@ @@:@@")
    sDisp = sDisp & sDateTime
    
    'リストに追加
    lstKan.AddItem (sDisp)
    
    'コメント
    '60文字で保存されているので、60バイトの領域に直す。前後の空白を取る。
    sDispCom = Trim(StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 60), vbUnicode))
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 1, 32), vbUnicode)     'コメント1
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    Else
        lstKan.AddItem (Space(45) & String(32, "-"))
    End If
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 33, 60), vbUnicode)    'コメント2
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    End If
    
    Set fsoObj = Nothing
    
    Exit Sub
Err_FILE:
    
    '異常ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    'ログ出力　┗ファイル名
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             'ファイルパスからファイル名を取得
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "┗File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    'ファイルクローズ
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    
    '異常用表示
    lstKan.AddItem (sTitle & "------------       ---  --------------------  ----/--/-- --:--")
    lstKan.AddItem (Space(45) & "--------------------------------")
    lstKan.AddItem (Space(45) & "--------------------------------")

    Set fsoObj = Nothing

End Sub
'V30.1.0.1 ADD END


' EG20 V2.1.0.1[Mainte_03_01] 追加開始
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fReadFileList
'//  機能名称  : ファイルリストの取得
'//  機能概要  : ファイルリストより、ファイル名を取得する。
'//
'//              型        名称      意味
'//  引数      : String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fReadFileList(sFileList As String) As Boolean
    Dim iFileNumber As Integer      'ファイル番号
    Dim sFileName As String         'ファイル名
    Dim iListCnt As Integer         'ファイル格納数

    On Error GoTo ErrorHandler      'エラーハンドル設定

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '未使用のファイル番号を取得する

    Open sFileList For Input Access Read As #iFileNumber    'ファイルリストのオープン
    Do While Not EOF(iFileNumber)                           'ファイルの終端までループを繰り返します。
        Line Input #iFileNumber, sFileName                  'データを読み込みます。
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                'ファイル名が存在する
            iListCnt = iListCnt + 1                         'ファイル数のカウンタをアップする
            ReDim Preserve FileList(iListCnt)               'ファイル名格納エリアを拡張する
            ReDim Preserve FileListType(iListCnt)           'ファイル名格納エリアを拡張する
            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
                                                            'ファイル名をファイル名格納エリアにセット
        End If
    Loop
    Close #iFileNumber      'ファイルを閉じます。

    fReadFileList = True    '戻り値を正常とする

    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    'V1.21.0.1 ADD  START
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    'V1.21.0.1 ADD  END
    fReadFileList = False   '戻り値をエラーとする
End Function
' EG20 V2.1.0.1[Mainte_03_01] 追加終了
' EG20 V5.6.0.1【ＩＣＭバージョンファイルリスト対応】追加開始
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : fReadFileListIDU
'//  機能名称  : IDUファイルリストの取得
'//  機能概要  : ファイルリストより、ファイル名を取得する。
'//
'//              型        名称      意味
'//  引数      : String　　sFileList　[IN]ファイルリストのフルパス名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 【ＩＣＭバージョンファイルリスト対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fReadFileListIDU(sFileList As String) As Boolean
    Dim iFileNumber As Integer      'ファイル番号
    Dim sFileName As String         'ファイル名
    Dim iListCnt As Integer         'ファイル格納数
    Dim nIndex As Integer           ' 文字数

    On Error GoTo ErrorHandler      'エラーハンドル設定

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '未使用のファイル番号を取得する

    Open sFileList For Input Access Read As #iFileNumber    'ファイルリストのオープン
    Do While Not EOF(iFileNumber)                           'ファイルの終端までループを繰り返します。
        Line Input #iFileNumber, sFileName                  'データを読み込みます。
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                'ファイル名が存在する
            iListCnt = iListCnt + 1                         'ファイル数のカウンタをアップする
            ReDim Preserve FileList(iListCnt)               'ファイル名格納エリアを拡張する
            ReDim Preserve FileListType(iListCnt)           'ファイル名格納エリアを拡張する

            nIndex = InStr(sFileName, " ")
            If nIndex = 0 Then
                ' スペースが含まれていない場合
                FileListType(iListCnt - 1) = UCase(Trim$(sFileName))
            Else
                ' スペースが含まれている場合
                FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, nIndex)))
            End If
            FileList(iListCnt - 1) = FileListType(iListCnt - 1)
                                                            'ファイル名をファイル名格納エリアにセット
        End If
    Loop
    Close #iFileNumber      'ファイルを閉じます。

    fReadFileListIDU = True    '戻り値を正常とする

    Exit Function           '処理を終了する

'*********************
'* エラーハンドル処理 *
'*********************
ErrorHandler:   ' エラー処理ルーチン。
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    fReadFileListIDU = False   '戻り値をエラーとする
End Function
' EG20 V5.6.0.1【ＩＣＭバージョンファイルリスト対応】追加終了

'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : fIsExistCornerType
'//  機能名称  : コーナータイプ存在チェック
'//  機能概要  : 幹線コーナーが存在するか、在来線コーナーが存在するかチェックする。
'//
'//              型        名称      意味
'//  引数      : byte  byCornerType　[IN]   0:在来線コーナー
'//                                         1:幹線コーナー
'//
'//              型        値        意味
'//  戻り値    : boolean  true/false    true:在来線コーナー有り false:幹線コーナー有り
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fIsExistCornerType(intCornerType As Integer)

    Dim intCount        As Integer
    Dim byFindFlg       As Byte
    
    byFindFlg = False

    '各コーナーのコーナータイプを取得する
    Call gsGetSettiCorner
    Call gsGetCornerType
    
    If intCornerType = CORNER_TYPE_KANSEN Then   '幹線コーナーが存在するか知りたい場合
        For intCount = 0 To UBound(gblnCornerSet)
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN And gblnCornerSet(intCount) = True Then
                byFindFlg = True        '幹線コーナーが一つでもあればOK
                Exit For
            End If
        Next intCount
    Else                                        '在来線コーナーが存在するか知りたい場合
        For intCount = 0 To UBound(gblnCornerSet)
            If gintCornerType(intCount) = CORNER_TYPE_ZAIRAI And gblnCornerSet(intCount) = True Then
                byFindFlg = True        '在来線コーナーが一つでもあればOK
                Exit For
            End If
        Next intCount
    End If
    
    fIsExistCornerType = byFindFlg
          

End Function
