VERSION 5.00
Begin VB.Form frmHoshuClear 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ｏ Ｋ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmHoshuClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmHoshuClear.frm
'//  パッケージ名：保守データクリア画面
'//
'//  概要：保守データクリア画面
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//                 フェーズ２対応　保守データクリア中画面追加
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//                 【フェーズ２対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 保守データクリア画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
  
    'クリア中のガイドを表示する｡
    lblMessage(0) = "自改保守SW設定クリア中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 保守データクリア画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'メール受信用タイマを止める
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 保守データクリア画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    On Error Resume Next
    '「自改保守データクリア中画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HODHU_SW_CLEAR_SHORI_GAMEN_START, 0)
     
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
     
    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdOK_Click
'//  機能名称  : 「OK」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    Dim iCnt As Integer     'V1.7.0.1 ADD

On Error Resume Next
    '自画面を消す。
    '「自改保守データクリア中画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HODHU_SW_CLEAR_SHORI_GAMEN_END, 0)
    'V1.7.0.1 ADD START
    For iCnt = 0 To MAX_GATE_NO + 1
       'クリア対象号機を、クリア非対象にて初期化
       gClear_Gouki(iCnt) = CLEAR_FLAG.NOT_CLEAR
    Next
    'V1.7.0.1 ADD END
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ時処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  On Error Resume Next
    
  '汎用メイル受信処理を行う
  If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
     AppActivate frmHoshuClear.Caption, False
     pfFormActive (frmHoshuClear.hwnd)
  End If
   
  'クリア処理を行う。
  psHoshuClear
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psHoshuClear
'//  機能名称  : 自改保守SW設定ファイルクリア処理
'//  機能概要  : 自改保守SW設定ファイルを削除する。
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//                 【フェーズ２対応】
'//     REVISIONS :(EG20 V6.9.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psHoshuClear()
    Dim iCnt As Integer
    Dim sMyPath As String 'GATE_ULLフォルダ内、ファイル
    Dim iGouki As Integer
    Dim sSWDataPath As String 'RMENTEフォルダパス
    Dim fso         As New FileSystemObject 'ファイルシステムオブジェクト
    Dim bRet        As Boolean '処理結果ステータス
    Dim nCorner     As Integer ' コーナ番号     ' EG20 V6.9.0.1追加

    On Error Resume Next
    
    bRet = True
    
'    For iCnt = 0 To 17         'EG20 V2.1.0.1 DEL 【フェーズ２対応】
    For iCnt = 0 To 31          'EG20 V2.1.0.1 ADD 【フェーズ２対応】
       If gClear_Gouki(iCnt) = CLEAR_FLAG.TARGET_CLEAR Then
          '「GATE_ULL」フォルダパスを作成
          sMyPath = Replace(GATE_SW_FILE, "##", Format(iCnt + 1, "0#"))
          'ファイルの有無チェックを行う。
          If Dir(sMyPath) <> "" Then
             Kill sMyPath
          End If

          If Dir(sMyPath) <> "" Then
             bRet = False
          End If
          
          '「RMENTE\本電鉄\自駅\XX号機」フォルダパスを作成
'          iGouki = pfGetGoukiNo(iCnt + 1)              ' EG20 V6.9.0.1削除
          iGouki = pfGetGoukiNo(iCnt + 1, nCorner)      ' EG20 V6.9.0.1追加
          If iGouki <> -1 Then
             sSWDataPath = PATH_RMENTE_GATE_DEN_JIEKI_GOUKI
' EG20 V6.9.0.1追加開始
             '「コーナ$」の「$」を1～6に変換する。
             sSWDataPath = Replace(sSWDataPath, "$", nCorner)
' EG20 V6.9.0.1追加終了
             sSWDataPath = Replace(sSWDataPath, "##", Format(iGouki, "0#"))
             sSWDataPath = Mid(sSWDataPath, 1, Len(sSWDataPath) - 2)
             If Dir(sSWDataPath, vbDirectory) <> "" Then
                fso.DeleteFolder (sSWDataPath)
             End If
             If Dir(sSWDataPath, vbDirectory) <> "" Then
                bRet = False
             End If
          End If
        End If
     Next
     
     Set fso = Nothing
     If bRet = False Then
        sClearEnd (1)
     Else
        sClearEnd (0)
     End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetGoukiNo
'//  機能名称  : 表示号機番号を取得する。
'//  機能概要  : GATE.INIより表示号機番号を取得する。
'//
'//              型        名称      意味
'//  引数      : Integer  iGouki    [IN]号機番号
'//              Integer  nCorner   [OUT]コーナ番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.9.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Function pfGetGoukiNo(iGouki As Integer) As Integer
Private Function pfGetGoukiNo(iGouki As Integer, nCorner As Integer) As Integer

    Dim lngRet As Long          '関数の返り値
    Dim iGate As Integer        '自改INDEX
    Dim j As Integer            'ワークINDEX
    Dim sGoukiNo As String      'GLTファイルレコードデータ(号機番号表示文字)
    Dim cWork As Byte           'ワークエリア
    Dim lngErrCode As Long      'エラーコード
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim intGLTFileNo As Integer     'ﾌｧｲﾙ番号

    On Error Resume Next

    '自動改札機情報取得
    sKeyName = "gate" & Format(iGouki, "00")
    iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                   sKeyName, _
                                   DEFAILT, sGateData, Len(sGateData), _
                                   PATH_GATE_FILE)
    If iRet = 0 Then
       '「自改保守SW設定クリア画面：自動改札機INIファイル読込異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
       pfGetGoukiNo = -1
       Exit Function
    End If
        
    If Len(sGateData) <> 0 Then
       'データの取得
       ReDim sFData(15)
       iFCnt = 1
          
       For iFLoop = 1 To Len(sGateData)
         If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
           iFLoop2 = iFLoop
              Do
               iFLoop2 = iFLoop2 + 1
               If iFLoop2 > Len(sGateData) Then
                  sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                  iFCnt = iFCnt + 1
                  If iFCnt >= 16 Then
                     Exit For
                  End If
                 
                  iFLoop = iFLoop2
              Exit Do
            End If
                       
            If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
               sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
               iFCnt = iFCnt + 1
               If iFCnt >= 16 Then
                  Exit For
               End If
                  
               iFLoop = iFLoop2
              Exit Do
           End If
          Loop
        End If
     Next
   End If
      
   If Trim(sFData(1)) <> "" Then
      pfGetGoukiNo = Trim(sFData(1))
   End If
' EG20 V6.9.0.1 【号機番号にコーナ番号を付加する対応】追加開始
   nCorner = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.9.0.1 【号機番号にコーナ番号を付加する対応】追加終了

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sClearEnd
'//  機能名称  : クリア処理結果表示処理
'//  機能概要  : 保守データクリア結果の結果文言を表示する。
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sClearEnd(iEnd As Integer)
    Dim i As Integer       'カウンタ
    Dim lngErrCode As Long 'エラーコード

    On Error Resume Next
        
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
        
    If iEnd = 0 Then
       '正常終了時の文言を表示する。
       lblMessage(0) = "正常終了しました。"
       lblMessage(1) = ""
       '「保守データ出力処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, HOSHU_SW_CLEAR_OK, 0)
    Else
       '収集失敗時の文言を表示する。
       lblMessage(0) = "異常終了しました。"
       lblMessage(1) = ""
       '「保守データ出力理異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, HOSHU_SW_CLEAR_ERROR, lngErrCode)
    End If
    cmdOK.Enabled = True
End Sub
