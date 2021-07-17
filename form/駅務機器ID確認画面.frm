VERSION 5.00
Begin VB.Form frmEkimKikiId 
   BorderStyle     =   0  'なし
   Caption         =   "駅務機器ID確認"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9120
      Top             =   3480
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
      TabIndex        =   6
      Top             =   2400
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
      TabIndex        =   5
      Top             =   720
      Width           =   2055
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
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ListBox ListEkimId 
      Height          =   7710
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   8775
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10425
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "  機器情報設定    画面へ戻る"
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ID"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "名称"
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
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "駅務機器ID確認"
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
Attribute VB_Name = "frmEkimKikiId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmEkimKikiId.frm
'//  パッケージ名：駅務機器ID確認画面
'//
'//  概要：駅務機器ID確認画面
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 駅務機器ID書込み先ディレクトリ位置変更
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【テキスト出力、媒体出力ボタンの抑止対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
Private iSendType As Integer            '要求種別値
Private Const EKIMU_DEFU = "APL\APL_WORK"

Private Const APL = "APL"
Private Const LOG = "LOG"
Private Const Data = "DATA"
Private Const BACKUP = "BACKUP"

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 駅務機器ID確認画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【テキスト出力、媒体出力ボタンの抑止対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    On Error Resume Next
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
    
'V1.7.0.1 ADD START
    Dim bRet As Boolean                 '戻り値
    Dim bFlag As Boolean                'フラグ
    Dim lngErrCode As Long              'エラーコード
    Dim udtMail As MAIL_INFO_CMD          '画面表示要求
    Dim uMail As ML_KYOTU_INF           'メール
    Dim lLen  As Long
  
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
  
   'バッファフラッシュ要求をログプロセスに送信する
   '情報要求CMD(駅務機器ID=0)をID制御に送信する
   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
   iSendType = MailCmdType.ML_DT_EKIMU_ID
   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '「駅務機器ID確認：情報要求CMD送信異常」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
      'プログレスバーを消去する
      Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
      
      Exit Sub
   Else
      '「駅務機器ID確認：情報要求CMD送信正常」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
      '画面ロック
      cmdVer(1).Enabled = False
      cmdVer(2).Enabled = False
      cmdInstall.Enabled = False
      cmdCancel.Enabled = False
   End If
   
    'バッファフラッシュ終了通知受信
    bFlag = False
    Do Until bFlag = True
       'メール受信処理を行う
       lLen = DssMailRead(plMSlot_MN, uMail)
       If lLen > 0 Then                            '受信正常の時
         If ML_ID_INFO_RES = uMail.udtlHeader.dwId Then 'メールＩＤ
            '情報要求RESを受信したら、画面表示用ファイル作成を行う。
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
            '要求種別、処理結果を取得
            Call psDispID(uMail.lngData(1))
           '画面ロック解除
' EG20 V6.3.0.1【テキスト出力、媒体出力ボタンの抑止対応】削除開始
'           cmdVer(1).Enabled = True
'           cmdVer(2).Enabled = True
' EG20 V6.3.0.1【テキスト出力、媒体出力ボタンの抑止対応】削除終了
' EG20 V6.3.0.1【テキスト出力、媒体出力ボタンの抑止対応】追加開始
            If ListEkimId.ListCount > 0 Then
                cmdVer(1).Enabled = True
                cmdVer(2).Enabled = True
            End If
' EG20 V6.3.0.1【テキスト出力、媒体出力ボタンの抑止対応】追加終了
           cmdInstall.Enabled = True
           cmdCancel.Enabled = True
           Exit Do
         End If
        End If
        Sleep (MN_MAIL_INTERVAL)
    Loop
'V1.7.0.1 ADD END
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 駅務機器ID確認画面(ディアクティブ時)
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
'//  機能名称  : 駅務機器ID確認画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【テキスト出力、媒体出力ボタンの抑止対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
 'V1.7.0.1 DEL START
'   Dim udtMail As MAIL_INFO_CMD          '画面表示要求
'   Dim iResponse As Integer            'メッセージボックス戻り値
'   Dim bRet As Boolean                 'メール送信処理戻り値
'   Dim lngErrCode As Long              'エラーコード
'   Dim bFlag As Boolean
'   Dim lId As Long
 'V1.7.0.1 DEL END
 
   On Error Resume Next
   
   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
    
   '「駅務機器ID確認画面：表示」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_GAMEN_START, 0)
    
   'メール受信タイマのインターバルを'１秒にセット
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
 
' EG20 V6.3.0.1【テキスト出力、媒体出力ボタンの抑止対応】追加開始
    cmdVer(1).Enabled = False
    cmdVer(2).Enabled = False
' EG20 V6.3.0.1【テキスト出力、媒体出力ボタンの抑止対応】追加終了
 'V1.7.0.1 DEL START
'   '情報要求CMD(駅務機器ID=0)をID制御に送信する
'   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
'   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
'   udtMail.mlHeader.dwProid = RHOSHU_ID
'   udtMail.mlHeader.dwSubArea = 0
'   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
'   iSendType = MailCmdType.ML_DT_EKIMU_ID
'   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
'   If bRet = False Then
'      '「駅務機器ID確認：情報要求CMD送信異常」ログ出力
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
'   Else
'      '「駅務機器ID確認：情報要求CMD送信正常」ログ出力
'      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
'      '画面ロック
'      cmdVer(1).Enabled = False
'      cmdVer(2).Enabled = False
'      cmdInstall.Enabled = False
'      cmdCancel.Enabled = False
'   End If
 'V1.7.0.1 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  関数名称  : cmdCancel_Click
'//  概要     : 「メニュー画面へ戻る」釦押下処理
'//  説明     : 自画面を消去する。
'//  ﾊﾟﾗﾒｰﾀ   :
'//           :
'//
'//  ORIGINAL  ：(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
   
    On Error Resume Next
    
    '「駅務機器ID確認画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  関数名称  : cmdInstall_Click
'//  概要     : 「媒体取外」釦押下処理
'//  説明     : 媒体を取り外す。
'//  ﾊﾟﾗﾒｰﾀ   :
'//           :
'//
'//  ORIGINAL  ：(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
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
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  関数名称  : cmdVer_Click
'//  概要     : 「テキスト表示」「媒体出力」釦押下処理
'//  説明     : 各釦押下処理を行う。
'//  ﾊﾟﾗﾒｰﾀ   :
'//           :
'//
'//  ORIGINAL  ：(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdVer_Click(Index As Integer)
    Dim lRetVal As Long             '戻り値
    Dim sCommand As String          'コマンド文字列
    Dim lngErrCode As Long
    Dim bRet As Boolean
    
    On Error Resume Next
  
    Select Case Index

      Case 1
           '「テキスト表示釦：押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_TEXT_BUTTOM, 0)
           'メモ帳実行コマンドを作成
           sCommand = MN_EXE_MEMO & MN_VERSI_FILE
           'メモ帳を起動する｡
           lRetVal = Shell(sCommand, vbMaximizedFocus)
           'メモ帳をアクティブ（前面表示）にする
           AppActivate lRetVal, True
           SendKeys "{LEFT}", True
      Case 2
           '「媒体出力釦：押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_OUTPUT_BUTTOM, 0)
           bRet = Text_OutPut
           If bRet = False Then
              '「媒体出力異常」ログ出力
              Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, EKIMUKIKI_ID_OUTPUT_ERROR, 0)
           Else
              '「媒体出力正常」ログ出力
              Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, EKIMUKIKI_ID_OUTPUT_OK, 0)
           End If
           
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  関数名称  : Text_Output
'//  概要     : 「媒体出力」処理
'//  説明     : 媒体出力処理を行う。
'//  ﾊﾟﾗﾒｰﾀ   :
'//           :
'//
'//  ORIGINAL  ：(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS ：(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 駅務機器ID書込み先ディレクトリ位置変更
'//  REVISIONS ：(1.20.0.1) 2010-03-10   REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面をOS仕様に変更
'//  REVISIONS ：(EG20 V2.0.1.1) 2011-11-21  REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応【残件№54】
'//                 ・出力ファイル名変更
'//  REVISIONS ：(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS ：(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function Text_OutPut() As Boolean
    Dim sCopyfile As String         'コピー先
    Dim sCopyTargetFile As String   'コピー元
    Dim sLzhDirName As String
    Dim iResponse           As Integer          'MsgBox戻り値
    
' EG20 V2.0.1.1 ADD START
    Dim strStationName As String                ' 駅名取得エリア
' EG20 V2.0.1.1 ADD END
' EG20 V3.0.0.2追加開始
    Dim fso         As New FileSystemObject     ' ファイルシステムオブジェクト
    Dim textWrite   As TextStream               ' テキスト（ライト）
    Dim textRead    As TextStream               ' テキスト（リード）
    Dim bWOpen      As Boolean
    Dim bROpen      As Boolean
    Dim strRecord   As String                   ' ワーク
' EG20 V3.0.0.2追加終了
    
On Error GoTo FileCopyError
  
    Text_OutPut = False

' EG20 V3.0.0.2追加開始
    bWOpen = False
    bROpen = False
' EG20 V3.0.0.2追加終了
   
    'フォルダ選択画面を表示させ、ファイル格納ディレクトリ名を得る。
'    sLzhDirName = pfDirSelection("a:", "駅務機器ID書込み先のディレクトリ選択")     'V1.12.0.1 DEL
    'sLzhDirName = pfDirSelection("H:", "駅務機器ID書込み先のディレクトリ選択")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
    sLzhDirName = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
    If sLzhDirName = "" Then
       Text_OutPut = True
       Exit Function
    End If
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
' EG20 V2.0.1.1 DEL START
'    sCopyfile = sLzhDirName & EKIMU_ID_TXT
' EG20 V2.0.1.1 DEL END
' EG20 V2.0.1.1 ADD START
    '駅名取得
    strStationName = gsGetStationEkiName
    ' 出力ファイル名作成
    sCopyfile = sLzhDirName & strStationName & "_" & EKIMU_ID_TXT
' EG20 V2.0.1.1 ADD END
    
    sCopyTargetFile = MN_VERSI_FILE
    
' EG20 V3.0.0.2削除開始
'    FileCopy sCopyTargetFile, sCopyfile
' EG20 V3.0.0.2削除終了
    
' EG20 V3.0.0.2追加開始
    Set textWrite = fso.CreateTextFile(sCopyfile, True)
    bWOpen = True
    textWrite.WriteLine ("設置駅　：" & strStationName)
    textWrite.WriteBlankLines (1)
    Set textRead = fso.OpenTextFile(sCopyTargetFile, ForReading, False)
    bROpen = True
    Do Until textRead.AtEndOfStream = True
        strRecord = textRead.ReadLine
        textWrite.WriteLine strRecord
    Loop
    textWrite.Close
    bWOpen = False
    textRead.Close
    bROpen = False
    Set textWrite = Nothing
    Set textRead = Nothing
    Set fso = Nothing
' EG20 V3.0.0.2追加終了
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    iResponse = MsgBox("正常終了しました。", _
                       vbOKOnly, _
                       "媒体出力結果")
    
    
    'ディスク情報を取得
    Text_OutPut = True
    
    Exit Function

FileCopyError:
' EG20 V3.1.0.2追加開始
    If bWOpen = True Then
        textWrite.Close
        bWOpen = False
    End If
    If bROpen = True Then
        textRead.Close
        bROpen = False
    End If
    Set textWrite = Nothing
    Set textRead = Nothing
    Set fso = Nothing
' EG20 V3.1.0.2追加終了
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    iResponse = MsgBox("異常終了しました。", _
                       vbOKOnly, _
                       "媒体出力結果")
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
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
 'V1.7.0.1 DEL START
'  Dim lLen  As Long
'  Dim uMail As ML_KYOTU_INF           'メール
'
'  On Error Resume Next
'
'  'メール受信
'  lLen = DssMailRead(plMSlot_MN, uMail)
'  If lLen > 0 Then                            '受信正常の時
'
'      Select Case uMail.udtlHeader.dwId  'メールＩＤ
'        Case ML_ID_HOSHU_ACTIVE_REQ
'            '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
'            AppActivate frmEkimKikiId.Caption, False
'            pfFormActive (frmEkimKikiId.hwnd)
'        Case ML_ID_INFO_RES
'            '情報要求RESを受信したら、画面表示用ファイル作成を行う。
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
'
'            '要求種別、処理結果を取得
'            Call psDispID(uMail.lngData(1))
'        Case Else
'     End Select
'  End If
'  '画面ロック解除
'  cmdVer(1).Enabled = True
'  cmdVer(2).Enabled = True
'  cmdInstall.Enabled = True
'  cmdCancel.Enabled = True
'V1.7.0.1 DEL END
'V1.7.0.1 ADD START
    'エラールーチンを宣言
    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmEkimKikiId.Caption, False
        pfFormActive (frmEkimKikiId.hwnd)
    End If
'V1.7.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psDispID
'//  機能名称  : 画面表示処理
'//  機能概要  : 駅務機器ID情報画面表示処理を行う。
'//
'//              型        名称      意味
'//  引数      : Long     lngSts    [IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Function psDispID(lngSts As Long)
    Dim sEkimuIDFile    As String   '駅務機器IDファイルパス
    Dim iRet            As Integer  'INI取得戻り値
    Dim sFolder         As String * MAX_PATH_SIZE  'フォルダ名
    Dim sFile           As String   'ファイル名
    Dim MyName          As String   'ファイル検索結果
    Dim bRet            As Boolean  '戻り値
    Dim lngErrCode      As Long     'エラーコード
    Dim intFileNo       As Integer  'ファイル番号
    Dim strWork         As String   '作業エリア
    Dim dwErrsts        As Long
    Dim sFolderName     As String
        
    '処理結果ステータスが正常の場合、iF内処理を行う。
    If lngSts = 0 Then
      sFolder = ""
      
      '処理結果：正常時は画面表示処理
      iRet = GetPrivateProfileString(IDU_SECTION_NAME, _
                                     IDU_EKIMUID_KEY, _
                                     EKIMU_DEFU, sFolder, Len(sFolder), _
                                     PATH_IDU_INI_FILE)
      If iRet = 0 Then
        sFolder = EKIMU_DEFU
      End If
      sEkimuIDFile = ""
      '要求種別値よりファイル名作成
      sFile = Replace(EKIMU_ID_FILE, "##", Format(iSendType, "0#"))
      If iRet = 0 Then
         sFolderName = RTrim(sFolder)
      Else
         sFolderName = Mid(sFolder, 1, iRet)
      End If
      'パス変換処理
      sFolderName = pfChangeFolderName(sFolderName)
      '駅務機器IDファイルパス作成
      sEkimuIDFile = sFolderName & "\" & sFile
      'ファイル有無チェック
      If Dir(sEkimuIDFile, vbNormal) = "" Then
         Exit Function
      End If
      
      '/////////////////////////////////////////////////////////////////////
      '//保守専用関数：駅務機器ID画面表示用ファイル作成処理
      '////////////////////////////////////////////////////////////////////
      'bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE) 'V1.8.0.1 DEL
      bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE, PATH_IDU_APP) 'V1.8.0.1 ADD

      If dwErrsts = 1 Then
         'エラーコード：正常
         'リスト初期化
         ListEkimId.Clear

        'VBエラー処理
        On Error GoTo Error_psVersionDisp
    
        '駅務機器ID画面表示用ファイルのファイル番号を取得する。
        intFileNo = FreeFile
      
        '駅務機器ID画面表示用ファイルオープン
        Open MN_VERSI_FILE For Input As #intFileNo
    
        'リスト表示分読み込み（ファイル終端までループを繰り返す）
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1削除
        Do While Not EOF(intFileNo)                         ' EG20 V3.3.0.1追加
           '作業エリアを初期化
           strWork = ""

           Line Input #intFileNo, strWork
           
           '改行コードのみは読みとばす
           If Trim(strWork) <> "" Then
              'リストに出力
              ListEkimId.AddItem (strWork)
           End If
        Loop
         
        'ファイルクローズ
        Close #intFileNo
      Else
        'エラーコード：異常
        Exit Function
     End If
   Else
     '処理結果：異常時は何もしない
   End If
Exit Function

'VBエラー処理
Error_psVersionDisp:
    'V1.21.0.1 ADD  START
    If intFileNo > 0 Then
        Close #intFileNo
    End If
    'V1.21.0.1 ADD  END
    '「駅務機器ID確認画面：バージョン情報ファイル作成異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfChangeFolderName
'//  機能名称  : フォルダパス変換処理
'//  機能概要  : INIファイルより取得したフォルダ定義の変換を行う。
'//
'//              型        名称         意味
'//  引数      : String sFolderName    [IN]INI定義
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Function pfChangeFolderName(sFolderName As String) As String
   Dim iPath As Integer
   Dim sRootPath As String
   Dim sFolder As String
      
   '「￥」位置を取得
   iPath = InStr(sFolderName, "\")
   If iPath = 0 Then
     sRootPath = Mid(sFolderName, 1)
   Else
     '「￥」前文字列を取得
     sRootPath = Mid(sFolderName, 1, iPath - 1)
     '「￥」後文字列を取得
     sFolder = Mid(sFolderName, iPath + 1)
   End If
   Select Case sRootPath
      Case APL
        'アプリルート
        sRootPath = PATH_IDU_APP
      Case LOG
        'ログルート
        sRootPath = PATH_IDU_LOG
      Case Data
        'DBルート
        sRootPath = PATH_IDU_DB
      Case BACKUP
        'バックアップルート
        sRootPath = PATH_BUC
   End Select
    'パス連結
    pfChangeFolderName = sRootPath + "\" + sFolder
End Function
