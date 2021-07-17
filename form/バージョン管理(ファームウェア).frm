VERSION 5.00
Begin VB.Form frmFirmWareVer 
   BorderStyle     =   0  'なし
   Caption         =   "自動改札機バージョン管理"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   -210
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdCancel 
      Caption         =   "   メニュー     画面へ戻る"
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
      TabIndex        =   10
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdVer 
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
      Height          =   735
      Index           =   2
      Left            =   9720
      TabIndex        =   9
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
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
      Height          =   735
      Index           =   1
      Left            =   9720
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdInstall 
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
      Height          =   735
      Left            =   9720
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
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
      Height          =   7500
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8295
   End
   Begin VB.CommandButton cmdVer 
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
      Height          =   735
      Index           =   0
      Left            =   9720
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "媒体入力"
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
      Left            =   9720
      TabIndex        =   1
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Timer tmrMail 
      Left            =   8760
      Top             =   8040
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ファイル名"
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
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "Ver"
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
      Index           =   5
      Left            =   7320
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "ＲＹＴバージョン管理"
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
      TabIndex        =   11
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "作成日時"
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
      Index           =   4
      Left            =   5040
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ﾌﾟﾛｸﾞﾗﾑ名"
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
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "機種名"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmFirmWareVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmFirmWareVer.frm
'//  パッケージ名：ＲＹＴバージョン管理画面
'//
'//  概要：ＲＹＴバージョン画面
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　画面レイアウト変更による修正
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 監視ファーム書込先ディレクトリ位置変更
'//                 インストール媒体ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 ①フォルダ選択画面をOS仕様に変更
'//                 ②バージョン表示の更新処理追加
'//  備考：フェーズ１、２時は「監視ファームバージョン管理」
'//        フェーズ３にて「ＲＹＴバージョン管理」に画面名称変更のため
'//        各部のコメントについては「監視ファーム」のままとする。
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir関数をFileSystemObjectに変更
'///////////////////////////////////////////////////////////////////
Option Explicit
'V1.6.0.1 DEL START
'Private Const KANSI_FIRM = 0            '監視ファームCPU
'Private Const RAS_MICO = 1              'RASマイコン
'Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
'Private Chk_OptButtom As Integer        '選択ラジオ釦値
'V1.6.0.1 DEL END
'V1.6.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
'V1.6.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ＲＹＴバージョン管理画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
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
'//  機能名称  : ＲＹＴバージョン管理画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
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
'//  機能名称  : ＲＹＴバージョン管理画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　画面レイアウト変更による修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
'    '「監視ﾌｧｰﾑﾊﾞｰｼﾞｮﾝ管理画面：表示」ログ出力'V1.6.0.1 DEL
    '「RYTﾊﾞｰｼﾞｮﾝ管理画面：表示」ログ出力      'V1.6.0.1 ADD
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_START, 0)

   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    'V1.6.0.1 DEL START
    'optSyubetu(0).Value = True
    
    'Chk_OptButtom = KANSI_FIRM
    'V1.6.0.1 DEL END
    
    'バージョン情報表示処理
    Call psVersionDisp
    
 End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdVer_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 「表示更新」「テキスト表示」「媒体出力」「媒体入力」
'//              釦押下処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer   Index    [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-29   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　テキスト表示時にファイル有無チェックを行う。
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdVer_Click(Index As Integer)
    Dim lRetVal As Long             '戻り値
    Dim sCommand As String          'コマンド文字列
    Dim lngErrCode As Long
    Dim bRet As Boolean
    Dim sFile As String             'ファイル名

    On Error Resume Next
 
 Select Case Index
    Case 0
         '「表示更新釦：押下」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
         bRet = UpData_Info
         If bRet = True Then
            '「表示更新正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_UPDATA_OK, 0)
         Else
            '「表示更新異常」ログ出力
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_UPDATA_ERROR, lngErrCode)
         End If
      
    Case 1
         '「テキスト表示釦：押下」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_TEXT_BUTTOM, 0)
         'V1.6.0.1 ADD START
         sFile = Dir(MN_VERSI_FILE, vbNormal)
         If sFile = "" Then
            'ファイル無しログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_TEXT_ERROR, 0)
            Exit Sub
         End If
         'V1.6.0.1 ADD END
         
         'メモ帳実行コマンドを作成
         sCommand = MN_EXE_MEMO & MN_VERSI_FILE
         'メモ帳を起動する｡
         lRetVal = Shell(sCommand, vbMaximizedFocus)
         'メモ帳をアクティブ（前面表示）にする
         AppActivate lRetVal, True
         SendKeys "{LEFT}", True
    
    Case 2
         '「媒体出力釦：押下」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_OUTPUT_BUTTOM, 0)
         bRet = Text_OutPut
         If bRet = True Then
            '「媒体出力正常」ログ出力
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_UPDATA_OK, 0) 'V1.6.0.1 DEL
            Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_OUTPUT_OK, 0) 'V1.6.0.1 ADD
         Else
            '「媒体出力異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_OUTPUT_ERROR, 0)
         End If
    
    Case 3
        '「媒体入力釦：押下」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_INPUT_BUTTOM, 0)
        bRet = File_InPut
        If bRet = True Then
           '「媒体出力正常」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_INPUT_OK, 0)
        Else
           '「媒体出力異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_INPUT_ERROR, 0)
        End If
 End Select
End Sub
'V1.6.0.1 DEL START
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : optSyubetu_Click
''//  機能名称  : ラジオ釦選択処理
''//  機能概要  : 監視ファーム、RASマイコン選択情報を更新保持する。
''//
''//              型        名称      意味
''//  引数      : なし
''//
''//              型        値        意味
''//  戻り値    : なし
''//
''//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Sub optSyubetu_Click(Index As Integer)
'   On Error Resume Next
'   Chk_OptButtom = Index
'End Sub
'V1.6.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdInstall_Click
'//  機能名称  : 「媒体取外」釦押下時処理
'//  機能概要  : 媒体の取外しを行う
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdInstall_Click()
   On Error Resume Next
   
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)
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
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　画面レイアウト変更による修正
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 バージョン表示更新処理追加
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    On Error Resume Next
    
'    '「監視ファームバージョン管理画面：消去」ログ出力  'V1.6.0.1 DEL
     '「RYTﾊﾞｰｼﾞｮﾝ管理画面：消去」ログ出力              'V1.6.0.1 ADD
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_END, 0)
    
    'V1.20.0.1 ADD START
    'バージョン管理画面のバージョン表示更新処理を行う。
    frmVersion.psGetVersion
    'V1.20.0.1 ADD END
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psVersionDisp
'//  機能名称  : バージョン情報表示処理
'//  機能概要  : バージョン情報表示部の表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　画面レイアウト変更による修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Function psVersionDisp() As Boolean

    Dim strFilePath     As String   'バージョンファイルパス
    Dim bRet            As Boolean  '戻り値
    Dim intFileNo       As Integer  'ファイル番号
    Dim strWork         As String   '作業エリア
    Dim strVerData      As String   '全体バージョン
    Dim intCnt          As Integer  'カウンター
    Dim lngErrCode      As Long     'エラーコード
    Dim strFolderPath   As String   'フォルダパス
   
'*******************************
'VBエラー処理
On Error GoTo Error_psVersionDisp
'*******************************

    'リスト初期化
    lstKan.Clear

    '作業エリア初期化
    strWork = ""

    '監視ファームバージョン管理画面表示用バージョンファイルパス作成
    strFilePath = MN_VERSI_FILE
    
    'V1.6.0.1　DEL　START
    ''RAS　or　監視ファーム
    'If Chk_OptButtom = RAS_MICO Then
    '   '表示がRASマイコン
    '   strFolderPath = PATH_KANSI_FIRMWARE_RAS & "*.*"
    'Else
    '   '表示が監視ファーム
    '   strFolderPath = PATH_KANSI_FIRMWARE & "*.*"
    'End If
    'V1.6.0.1　DEL　END
    strFolderPath = PATH_KANSI_FIRMWARE & "*.*" 'V1.6.0.1　ADD
    
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ 共通DA:監視ファームバージョン管理画面表示用バージョンファイル作成
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllKANSIFRMVER(strFolderPath, lngErrCode, strFilePath)

    '監視ファームバージョン管理画面表示用バージョンファイル成功
    If lngErrCode = 1 Then
       '「監視ファームバージョン管理画面：バージョン情報ファイル作成正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    '監視ファームバージョン管理画面表示用バージョンファイル失敗
    Else
       '「監視ファームバージョン管理画面：バージョン情報ファイル作成異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
       psVersionDisp = False
       Exit Function
    End If

    '監視ファームバージョン管理画面表示用バージョンファイルの有無確認
    If Len(Trim(Dir(strFilePath))) = 0 Then
       psVersionDisp = False
       Exit Function
    End If

    '監視ファームバージョン管理画面表示用バージョンファイルのファイル番号を取得する。
    intFileNo = FreeFile

    '監視ファームバージョン管理画面表示用バージョンファイルオープン
    Open strFilePath For Input As #intFileNo

    strWork = ""

    'リスト表示分読み込み（ファイル終端までループを繰り返す）
    Do While Not EOF(1)
       
       Line Input #intFileNo, strWork

       '改行コードのみは読みとばす
       If Trim(strWork) <> "" Then
          'リストに出力
          lstKan.AddItem (strWork)
       End If
     Loop

    'ファイルクローズ
    Close #intFileNo
    
    psVersionDisp = True

    Exit Function

'*******************************
'VBエラー処理
Error_psVersionDisp:
   '「監視ファームバージョン管理画面：バージョン情報ファイル作成異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
   'ファイルクローズ
   Close #intFileNo
   psVersionDisp = False
'*******************************
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : UpData_Info
'//  機能名称  : 「表示更新」釦押下処理
'//  機能概要  : バージョン情報表示部の再描画を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function UpData_Info() As Boolean
     
     On Error Resume Next

     Dim bUpData As Boolean
     
     'バージョン表示処理を行う。
     bUpData = psVersionDisp
     
     UpData_Info = bUpData
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Text_OutPut
'//  機能名称  : 「媒体出力」釦押下処理
'//  機能概要  : バージョンテキストファイルを媒体に出力する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-23   REVISED BY [TCC] S.Terao
'//                 フェーズ２不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 監視ファーム書込先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function Text_OutPut() As Boolean

'*******************************
'VBエラー処理
On Error GoTo Error_cmdOutPut_Click
'*******************************

    Dim iRet        As Integer      '戻り値
    Dim strVerFile  As String       'ID中継ユニットファイルパス
    Dim strCopySaki As String       '出力ファイルパス
    Dim strWriteDir As String       '出力先フォルダ
    Dim strEkimei   As String       '設置駅名
    Dim strWork     As String * 256 '作業エリア
'   Dim fso         As New FileSystemObject 'ファイルシステムオブジェクト 'V1.6.0.1　DEL
    Dim lngErrCode  As Long              'エラーコード
    Dim sLzhDirName As String        '指定フォルダ　'V1.6.0.1　ADD
    
    '監視ファームバージョン管理画面表示用ファイル
    strVerFile = MN_VERSI_FILE

'V1.6.0.1 DEL START
'    'ファイルの有無確認
'    If fso.FileExists(strVerFile) = False Then
'        'ファイル無し異常ポップアップ画面表示
'        MsgBox "媒体出力するデータがありません。", vbExclamation, "データ無警告"
'        Text_OutPut = False
'        Set fso = Nothing
'        Exit Function
'    End If
'V1.6.0.1 DEL END
    
    'V1.6.0.1 ADD START
    'フォルダ選択画面を表示させ、ファイル格納ディレクトリ名を得る。
'    sLzhDirName = pfDirSelection("a:", "監視ファーム書込み先ディレクトリ選択")     'V1.12.0.1 DEL
    'sLzhDirName = pfDirSelection("H:", "監視ファーム書込み先ディレクトリ選択")      'V1.12.0.1 ADD  'V1.20.0.1 DEL
    sLzhDirName = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
    If sLzhDirName = "" Then
       '媒体フォルダ指定なし時
       Text_OutPut = True
       Exit Function
    End If
    'V1.6.0.1 ADD END

    'V1.6.0.1 DEL START
'    'フォルダ選択ポップアップ画面表示
'    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")
'
'    '指定フォルダなし
'    If Len(strWriteDir) = 0 Then
'       Text_OutPut = False
'       Set fso = Nothing
'       Exit Function
'    End If
'
'    'コピー先フォルダの有無確認
'    If fso.FolderExists(strWriteDir) = False Then
'        'コピー先フォルダ作成
'        fso.CreateFolder (strWriteDir)
'    End If
'
'    'コピー先ファイル名作成
'    strCopySaki = strWriteDir & "\" & VER_TXT_NAME
'
'   'ファイルコピー（既に存在した場合は上書きするする）
'    fso.CopyFile strVerFile, strCopySaki, True
   'V1.6.0.1 DEL END
   'V1.6.0.1 ADD START
   strCopySaki = sLzhDirName & "\" & VER_TXT_NAME
   
   FileCopy strVerFile, strCopySaki
   'V1.6.0.1 ADD END
 
    MsgBox "媒体出力は正常終了しました。", vbInformation + vbOKOnly, "媒体出力結果"
    
    Text_OutPut = True
'   Set fso = Nothing       'V1.6.0.1 DEL

    Exit Function
'*******************************
'VBエラー処理
Error_cmdOutPut_Click:
     MsgBox "媒体出力は異常終了しました。", vbCritical, "媒体出力結果"
'    Set fso = Nothing      'V1.6.0.1 DEL

     Text_OutPut = False
'*******************************
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : File_InPut
'//  機能名称  : 「媒体入力」釦押下処理
'//  機能概要  : ファイルを媒体入力する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 フェーズ３対応　画面レイアウト変更による修正
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 インストール媒体ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//                 バージョン表示更新処理追加
'//                 FileSystemObjectの使用を止め、FileCopyに変更
'//                 読み取り専用属性を変更する処理を追加
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir関数をFileSystemObjectに変更
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function File_InPut() As Boolean
  Dim lngErrCode As Long
  Dim bRet As Boolean
  Dim sLzhDirName As String
  Dim strFileName As String
'  Dim fso         As New FileSystemObject 'ファイルシステムオブジェクト        'V1.20.0.1 DEL
  Dim FolderName As String
'V2.6.0.1 ADD START
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト
    Dim objFi As File                    'ファイルオブジェクト
    Dim MyName As String
    Dim sSrcFileName As String
    Dim sDstFileName As String
'V2.6.0.1 ADD END
  
  On Error Resume Next
  
  'フォルダ選択画面を表示させ、ファイル格納ディレクトリ名を得る。
'  sLzhDirName = pfDirSelection("a:", "インストール媒体のディレクトリ選択")     'V1.12.0.1 DEL
  'sLzhDirName = pfDirSelection("H:", "インストール媒体のディレクトリ選択")      'V1.12.0.1 ADD     'V1.20.0.1 DEL
  sLzhDirName = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)      'V1.20.0.1 ADD
  If sLzhDirName = "" Then
     '媒体フォルダ指定なし時
     File_InPut = True
'     Set fso = Nothing     'V1.20.0.1 DEL
     Exit Function
  End If

  'フォルダ選択別展開処理を行う。
  'If Chk_OptButtom = KANSI_FIRM Then 'V1.6.0.1 DEL
     
     '一時フォルダを作成し、一時フォルダにコピー(リカバリ対策)
     FolderName = Mid(PATH_KANSI_FIRMWARE_WORK, 1, Len(PATH_KANSI_FIRMWARE_WORK) - 2)
     MkDir FolderName
     On Error GoTo Recovary_Error
     strFileName = Dir(PATH_KANSI_FIRMWARE & "*.*", vbNormal)
     Do While strFileName <> ""
'        fso.CopyFile PATH_KANSI_FIRMWARE & strFileName, PATH_KANSI_FIRMWARE_WORK & strFileName        'V1.20.0.1 DEL
        FileCopy PATH_KANSI_FIRMWARE & strFileName, PATH_KANSI_FIRMWARE_WORK & strFileName        'V1.20.0.1 ADD
        strFileName = Dir
     Loop
     
     'V1.6.0.1 ADD START
     strFileName = Dir(PATH_KANSI_FIRMWARE & "*.*", vbNormal)
     If strFileName <> "" Then
     'V1.6.0.1 ADD END
          Kill PATH_KANSI_FIRMWARE & "*.*"
     End If 'V1.6.0.1 ADD
     
     '媒体より、監視ファームCPUフォルダにコピー
     On Error GoTo In_Put_Error
'V2.6.0.1 DEL START
'     strFileName = Dir(sLzhDirName & "*.*", vbNormal)
'     Do While strFileName <> ""
''        fso.CopyFile sLzhDirName & strFileName, PATH_KANSI_FIRMWARE & strFileName        'V1.20.0.1 DEL
'        FileCopy sLzhDirName & strFileName, PATH_KANSI_FIRMWARE & strFileName        'V1.20.0.1 ADD
'        strFileName = Dir
'     Loop
'V2.6.0.1 DEL END
    'V2.6.0.1 ADD START
    For Each objFi In objFso.GetFolder(sLzhDirName).files   'ループを開始
        If objFso.FileExists(objFi.Path) = True Then  'ファイル名の取得チェック
           'ディレクトリ名を取得
           MyName = objFi.Name
           '媒体内ファイル名を作成
           sSrcFileName = sLzhDirName & MyName
           ' ビット単位の比較を行い、MyName がディレクトリかどうかを調べます。
           If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
               'ワークフォルダ内ファイル名を作成する
               sDstFileName = PATH_KANSI_FIRMWARE & MyName
               '媒体内のファイルをワークフォルダにコピーする
               FileCopy sSrcFileName, sDstFileName
           End If
        End If
    Next
    Set objFso = Nothing
    Set objFi = Nothing
    'V2.6.0.1 ADD END
'V1.6.0.1 DEL START
'  Else
'
'     '一時フォルダを作成し、一時フォルダにコピー(リカバリ対策)
'     FolderName = Mid(PATH_KANSI_FIRMWARE_RAS_WORK, 1, Len(PATH_KANSI_FIRMWARE_RAS_WORK) - 2)
'     MkDir FolderName
'     strFileName = Dir(PATH_KANSI_FIRMWARE_RAS & "*.*", vbNormal)
'     Do While strFileName <> ""
'        On Error GoTo Recovary_Error
'        fso.CopyFile PATH_KANSI_FIRMWARE_RAS & strFileName, PATH_KANSI_FIRMWARE_RAS_WORK & strFileName
'        strFileName = Dir
'     Loop
'
'     Kill PATH_KANSI_FIRMWARE_RAS & "*.*"
'
'     '媒体より、監視ファームCPUフォルダにコピー
'     strFileName = Dir(sLzhDirName & "*.*", vbNormal)
'     Do While strFileName <> ""
'        On Error GoTo In_Put_Error
'        fso.CopyFile sLzhDirName & strFileName, PATH_KANSI_FIRMWARE_RAS & strFileName
'        strFileName = Dir
'     Loop
'  End If
'V1.6.0.1 DEL END

  '一時フォルダを削除
    'V1.20.0.1 DEL START
'  fso.DeleteFolder FolderName, False
'  Set fso = Nothing
   'V1.20.0.1 DEL END
   
   'V1.20.0.1 ADD START
   psDeleteFolder FolderName
   
   '読み取り専用の場合に属性変更を行う
   Folder_SetAttr (PATH_KANSI_FIRMWARE)
  
  'バージョン情報表示処理
  Call psVersionDisp
  'V1.20.0.1 ADD END
    
  '媒体入力正常ポップアップ画面表示
  MsgBox "媒体入力は正常終了しました。", vbInformation + vbOKOnly, "媒体入力結果"
  File_InPut = True
  Exit Function

 
Recovary_Error:
  '媒体入力異常ポップアップ画面表示
  MsgBox "媒体入力は異常終了しました。", vbCritical, "媒体入力結果"
  File_InPut = False
  
  '一時フォルダを削除
    'V1.20.0.1 DEL START
'  fso.DeleteFolder FolderName, False
'  Set fso = Nothing
   'V1.20.0.1 DEL END
   psDeleteFolder FolderName        'V1.20.0.1 ADD
  Exit Function

In_Put_Error:
  
  'リカバリ処理を行う。
  'If Chk_OptButtom = KANSI_FIRM Then 'V1.6.0.1 DEL
  'V2.6.0.1 ADD START
   Set objFso = Nothing
   Set objFi = Nothing
  'V2.6.0.1 ADD END

     Kill PATH_KANSI_FIRMWARE & "*.*"
 
     '一時フォルダより、監視ファームCPUへコピー
     strFileName = Dir(PATH_KANSI_FIRMWARE_WORK & "*.*", vbNormal)
     Do While strFileName <> ""
        FileCopy PATH_KANSI_FIRMWARE_WORK & strFileName, PATH_KANSI_FIRMWARE & strFileName
        strFileName = Dir
     Loop
'V1.6.0.1 DEL START
'  Else
'     Kill PATH_KANSI_FIRMWARE_RAS & "*.*"
'
'     '一時フォルダより、RASマイコンへコピー
'     strFileName = Dir(PATH_KANSI_FIRMWARE_RAS_WORK & "*.*", vbNormal)
'     Do While strFileName <> ""
'        FileCopy PATH_KANSI_FIRMWARE_RAS_WORK & strFileName, PATH_KANSI_FIRMWARE_RAS & strFileName
'        strFileName = Dir
'     Loop
'  End If
'V1.6.0.1 DEL END
  
'媒体入力異常ポップアップ画面表示
  MsgBox "媒体入力は異常終了しました。", vbCritical, "媒体入力結果"
  File_InPut = False

 '一時フォルダを削除
    'V1.20.0.1 DEL START
'  fso.DeleteFolder FolderName, False
'  Set fso = Nothing
   'V1.20.0.1 DEL END
   psDeleteFolder FolderName        'V1.20.0.1 ADD

End Function

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
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    On Error Resume Next
 
    'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmFirmWareVer.Caption, False
        pfFormActive (frmFirmWareVer.hwnd)
    End If
End Sub

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Folder_SetAttr
'//  機能名称  : ファイル属性変更
'//  機能概要  : フォルダ内の読み取りファイル属性を通常に設定する
'//
'//              型      名称         意味
'//  引数      : String  sFolderName  フォルダパス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-11  CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Folder_SetAttr(sFolderName As String)
    On Error Resume Next
    
    Dim iAttr As Integer
    Dim lIndex As Long
    Dim foName As Folder
    Dim fName As File
    Dim fsoObj As FileSystemObject
    
    Set fsoObj = New FileSystemObject
    Set foName = fsoObj.GetFolder(sFolderName)
    lIndex = 0
    
    For Each fName In foName.files
        '属性を取得
        iAttr = GetAttr(fName.Path)
        '通常ファイル、またはアーカイブファイルに読み取り属性が付いていたら
        If iAttr = vbReadOnly Or iAttr = vbArchive + vbReadOnly Then
            '読み取り属性を取り除いてセット
            Call SetAttr(fName.Path, iAttr - vbReadOnly)
        End If
    lIndex = lIndex + 1
    Next fName
    
    Set fsoObj = Nothing
    Set fName = Nothing
    Set foName = Nothing
    
End Sub
'V1.20.0.1 ADD END
