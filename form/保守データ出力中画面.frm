VERSION 5.00
Begin VB.Form frmSyusyuOutPut 
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
   Begin VB.Timer tmrMail2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ｏ Ｋ"
      Enabled         =   0   'False
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
Attribute VB_Name = "frmSyusyuOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmSyusyuOutPut.frm
'//  パッケージ名：保守データ出力中画面
'//
'//  概要：保守データ出力中画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 ファイル選択処理対応
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.265修正対応】
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-13  CODED BY  [TCC] H.Sugimoto
'//                 【コーナ名スペース除去対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  REVISED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi06_005_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 保守データ出力中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim intCount As Integer
    Dim blnSelected As Boolean
    
    blnSelected = False
    For intCount = 0 To UBound(gintStatus)
        If gintStatus(intCount) = TAG_STATUS.STS_SENTAKU Then
            blnSelected = True
        End If
    Next
    
    '指定号機なしの場合、メッセージボックスを表示する
    If blnSelected = False Then
        lblMessage(0) = "指定号機が選択されていません。"
        lblMessage(1) = "選択してください。"
        cmdOK.Enabled = True
        Exit Sub
    End If
    'EG20 V2.1.0.1 ADD END
    
    cmdOK.Enabled = False
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_KBN_KADO_MAINTE)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
'    出力中のガイドを表示する｡
    lblMessage(0) = "保守データを出力中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    tmrMail.Enabled = True
    tmrMail2.Enabled = True                  ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 保守データ出力中画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
  On Error Resume Next
  'メイル受信用のインタバルタイマ値を設定する。
  tmrMail.Interval = MN_MAIL_INTERVAL
  tmrMail.Enabled = False
  
  tmrMail2.Interval = MN_MAIL_INTERVAL       ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
  tmrMail2.Enabled = False                   ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 保守データ収集中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'メール受信用タイマを止める
    tmrMail.Enabled = False
    tmrMail2.Enabled = False                   ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    On Error Resume Next
    
    '自画面を消す。
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
' V1.12.0.1 ADD START
    'メール受信用タイマを止める
    tmrMail.Enabled = False
' V1.12.0.1 ADD END
' V1.3.0.1 ADD START
    On Error Resume Next
' EG20 V8.1.0.1【EG20_KANSI05_01】DEL START
'    '汎用メイル受信処理を行う
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
'        AppActivate frmSyusyuOutPut.Caption, False
'    End If
' EG20 V8.1.0.1【EG20_KANSI05_01】DEL END
' V1.3.0.1 ADD END
     '出力ファイル作成処理を行う。
    sOutPutHoshuData

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
End Sub
' EG20 V8.1.0.1【EG20_KANSI05_01】ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail2_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ時処理
'//  機能概要  : メールを受信する
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()

    On Error Resume Next
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSyusyuOutPut.Caption, False
        pfFormActive (frmSyusyuOutPut.hwnd)
    End If

End Sub
' EG20 V8.1.0.1【EG20_KANSI05_01】ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSyusyuEnd
'//  機能名称  : 出力結果表示処理
'//  機能概要  : 保守データ出力結果の結果文言を表示する。
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSyusyuEnd(iEnd As Integer)
    Dim i As Integer       'カウンタ
    Dim lngErrCode As Long 'エラーコード

    On Error Resume Next
    
    Sleep (5000)
    
    If iEnd = 0 Then
       '正常終了時の文言を表示する。
       lblMessage(0) = "正常終了しました。"
       lblMessage(1) = ""
       '「保守データ出力処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_OUTPUT_OK, 0)
    Else
       '収集失敗時の文言を表示する。
       lblMessage(0) = "異常終了しました。"
       lblMessage(1) = ""
       '「保守データ出力理異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_OUTPUT_ERROR, lngErrCode)
    End If
    cmdOK.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sOutPutHoshuData
'//  機能名称  : 保守データ出力処理
'//  機能概要  : 保守データ出力を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 フェーズ３　結合検査　不具合修正
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 ファイル選択処理対応
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.265修正対応】
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-23  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【バックアップファイル対応】
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応【HKRK_Kansi06_005_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sOutPutHoshuData()
    Dim iIniKeka As Integer '戻り値
    Dim iGate As Integer    '作成号機数
    Dim iCreate As Integer  '作成ファイル数
    Dim iCnt As Integer     'カウンター１
    Dim i As Integer        'カウンター２
    Dim sMyFromPath As String '作成元ファイル名
    Dim sMyToPath As String   '作成先ファイル名
    Dim bRet As Boolean
    Dim sFromCreateFileName As String * MAX_PATH_SIZE    '作成元ファイル名
    Dim sToCreateFileName As String * MAX_PATH_SIZE      '作成先ファイル名
    Dim sFormatCreateFileName As String * MAX_PATH_SIZE  '作成フォーマットファイル名
    Dim sFormatCreateFileNameKan As String * MAX_PATH_SIZE  '作成フォーマットファイル名（幹線用）   'EG20 V30.3.0.1 【HKRK_Kansi06_005_01】 ADD
    Dim sOut_Path As String * MAX_PATH_SIZE              '作成フォーマットファイル名
    Dim iRet As Integer
    Dim iNoFileCnt As Integer           ' V1.3.0.1 ADD
    Dim bErrChk As Boolean              ' V1.8.0.1 ADD　　'出力ファイル作成エラーチェックフラグ
    Dim nCorner As Integer              ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
    Dim nCornerGoki As Integer          ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
    Dim iStatus As Integer                              ' 処理結果          ' EG20 V5.4.0.1【バックアップファイル対応】追加
    Dim sBackupFolderName As String * MAX_PATH_SIZE     '作成元フォルダ名   ' EG20 V5.4.0.1【バックアップファイル対応】追加
      
    On Error Resume Next
    
    sOut_Path = ""
    iCreate = 0
    iGate = 0
    
    bErrChk = True                      'V1.8.0.1 ADD
    
    ' EG20 V30.3.0.1 【HKRK_Kansi06_005_01】 ADD START
    ' 各コーナの種別を取得する
    gsGetCornerType
    ' EG20 V30.3.0.1 【HKRK_Kansi06_005_01】 ADD END
    
    
    '作成ファイル数を取得する。
     iCreate = GetPrivateProfileInt(HOSHUPUT_FROM_SECTION_NAME, _
                                    HOSHUPUT_FROM_NUMBER_KEY_NAME, DEFAILT_Int, PATH_HOSHU_DATA_FILE)
    '作成号機数を取得する。
     iGate = GetPrivateProfileInt(HOSHUPUT_FROM_SECTION_NAME, _
                                    HOSHUPUT_FROM_GATE_NUMBER_KEY_NAME, DEFAILT_Int, PATH_HOSHU_DATA_FILE)
    
    'コピー先を取得する。
'V1.12.0.1 DEL START
'     iIniKeka = GetPrivateProfileString(KANSI_OUT_HOSHU_SEC, _
'                                        KANSI_OUT_HOSHU_KEY, DEFAILT, _
'                                        sOut_Path, Len(sOut_Path), _
'                                        HOSHU_FILE)
'V1.12.0.1 DEL END

'    MkDir sOut_Path                         'V1.12.0.1 DEL
     MkDir frmSyusyu.glbFilePath             'V1.12.0.1 ADD
    
     For iCnt = 1 To iCreate
         sFromCreateFileName = ""
         sToCreateFileName = ""
         sFormatCreateFileName = ""
         '作成元ファイル名を取得する。
         iIniKeka = GetPrivateProfileString(HOSHUPUT_FROM_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sFromCreateFileName, Len(sFromCreateFileName), _
                                            PATH_HOSHU_DATA_FILE)
 'V1.7.0.1 DEL START
'        '作成先ファイル名を取得する。
'         iIniKeka = GetPrivateProfileString(HOSHUPUT_TO_SECTION_NAME, _
'                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
'                                            sToCreateFileName, Len(sToCreateFileName), _
'                                            PATH_HOSHU_DATA_FILE)
 'V1.7.0.1 DEL END
        '作成フォーマットファイル名を取得する。
         iIniKeka = GetPrivateProfileString(HOSHUPUT_FORMAT_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sFormatCreateFileName, Len(sFormatCreateFileName), _
                                            PATH_HOSHU_DATA_FILE)
' EG20 V30.3.0.1 【HKRK_Kansi06_005_01】ADD START
         iIniKeka = GetPrivateProfileString(HOSHUPUT_FORMAT_KAN_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sFormatCreateFileNameKan, Len(sFormatCreateFileNameKan), _
                                            PATH_HOSHU_DATA_FILE)
' EG20 V30.3.0.1 【HKRK_Kansi06_005_01】ADD END

' EG20 V5.4.0.1【バックアップファイル対応】追加開始
        '作成フォーマットファイル名を取得する。
        iStatus = GetPrivateProfileString(HOSHUPUT_BACKUPFOLDER_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sBackupFolderName, Len(sBackupFolderName), _
                                            PATH_HOSHU_DATA_FILE)

' EG20 V5.4.0.1【バックアップファイル対応】追加終了
            
        iNoFileCnt = 0                     ' V1.3.0.1 ADD
        For i = 1 To iGate
            
            If gintStatus(i - 1) = TAG_STATUS.STS_SENTAKU Then      'EG20 V2.1.0.1 ADD 【Mainte_03_01】
                'V1.7.0.1 ADD START
                '自改別INIファイル取得処理
'                sToCreateFileName = fGetGateInfoPath(i, iCnt, iIniKeka)        ' EG20 V3.4.0.1【統合TR-No.265修正対応】削除
                sToCreateFileName = fGetGateInfoPath(i, iCnt, iIniKeka, _
                                                    nCorner, nCornerGoki)       ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
                If sToCreateFileName <> "" And iIniKeka <> 0 Then
                'V1.7.0.1 ADD END
                    sMyFromPath = ""
                    sMyToPath = ""
                    '「##」を01～32に変換する。
                    sMyFromPath = Replace(sFromCreateFileName, "##", Format(i, "0#"))
'                    sMyToPath = Replace(sToCreateFileName, "##", Format(i, "0#"))              ' EG20 V3.4.0.1【統合TR-No.265修正対応】削除
                    ' 統合監視盤論理号機番号→コーナ別論理号機番号
                    sMyToPath = Replace(sToCreateFileName, "##", Format(nCornerGoki, "0#"))     ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
                    'iRet = dllCreateCSVDataFile(sMyToPath, sFormatCreateFileName, sMyFromPath) ' EG20 V30.3.0.1 【HKRK_Kansi06_005_01】DEL
                    
' EG20 V30.3.0.1 【HKRK_Kansi06_005_01】 ADD START
                    '現在処理している号機が属するコーナ種別によりDLL関数に渡すフォーマットファイル名を切り替える。
                    If gintCornerType(nCorner - 1) = CORNER_TYPE_KANSEN Then
                        'その号機が幹線コーナの場合
                        iRet = dllCreateCSVDataFile(sMyToPath, sFormatCreateFileNameKan, sMyFromPath)
                    Else
                        'その号機が在来コーナの場合
                        iRet = dllCreateCSVDataFile(sMyToPath, sFormatCreateFileName, sMyFromPath)
                    End If
' EG20 V30.3.0.1 【HKRK_Kansi06_005_01】 ADD END
                    If iRet = 0 Then
                        sSyusyuEnd (1)
                        Exit Sub
                    End If
    ' V1.3.0.1 DEL START
    '             If iRet = 2 And iGate = i And iCnt = iCreate Then
    '                sSyusyuEnd (1)
    '                Exit Sub
    '             End If
    ' V1.3.0.1 DEL START
    ' V1.3.0.1 ADD START
                    If iRet = 2 Then
                        '作成元ファイルなし異常
                        iNoFileCnt = iNoFileCnt + 1
                        'If iGate = iNoFileCnt And iCnt = iCreate Then  'V1.8.0.1 DEL
                        bErrChk = False                                 'V1.8.0.1  ADD
                        If i = iNoFileCnt And iCnt = iCreate Then      'V1.8.0.1 ADD
                            sSyusyuEnd (1)
                            Exit Sub
                        End If
                    End If
    ' V1.3.0.1 ADD END
                    If iRet = 1 Then
        '                bRet = HoshuCopy(sOut_Path, sMyToPath)             'V1.12.0.1 DEL
'                        bRet = HoshuCopy(frmSyusyu.glbFilePath, sMyToPath)  'V1.12.0.1 ADD
'                        bRet = HoshuCopy(frmSyusyu.glbFilePath, sMyToPath, nCorner)        ' V5.4.0.1削除
                        bRet = HoshuCopy(frmSyusyu.glbFilePath, sMyToPath, nCorner, sBackupFolderName)
                        If False = bRet Then
                            sSyusyuEnd (1)
                            Exit Sub
                        End If
                    End If
                End If 'V1.7.0.1 ADD
            End If          'EG20 V2.1.0.1 ADD 【Mainte_03_01】
         Next
     Next
     
     'V1.8.0.1 ADD START
     If bErrChk = False Then
        sSyusyuEnd (1)
        Exit Sub
     End If
     'V1.8.0.1 ADD END
     
     sSyusyuEnd (0)
         
End Sub
'V1.7.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fGetGateInfoPath
'//  機能名称  : 自改別ファイルパス取得処理
'//  機能概要  : 自改別ファイルパス取得処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iGouki　　[IN]号機番号
'//              Integer　iFilType　[IN]作成ファイル種別
'//              Integer　iIniKeka　[OUT]取得文字数
'//              Integer  nCorner      [OUT]コーナ機番号
'//              Integer  nCornerGoki  [OUT]コーナ論理号機番号
'//
'//              型        値        意味
'//  戻り値    : String　　　　　　[OUT]ファイルパス
'//
'//     ORIGINAL :(1.7.0.1) 2009-07-28   CODED BY [TCC] S.Terao
'//                 保守総点検修正
'//     REVISIONS:(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.265修正対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fGetGateInfoPath(iGouki As Integer, iFilType As Integer, iIniKeka As Integer, _
                                  nCorner As Integer, nCornerGoki As Integer) As String
    Dim lngRet As Long          '関数の返り値
    Dim iGate As Integer        '自改INDEX
    Dim j As Integer            'ワークINDEX
    Dim cWork As Byte           'ワークエリア
    Dim lngErrCode As Long      'エラーコード
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim sToCreateFileName As String * MAX_PATH_SIZE      '作成先ファイル名
 
    On Error Resume Next
    
    '自動改札機情報取得
    sKeyName = "gate" & Format(iGouki, "00")
    iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                    sKeyName, _
                                    DEFAILT, sGateData, Len(sGateData), _
                                    PATH_GATE_FILE)
    If iRet = 0 Then
       '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：自動改札機INIファイル読込異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
       fGetGateInfoPath = ""
       iIniKeka = 0
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
     
    If Trim(sFData(4)) = MISETI Then
        '未設置の場合
        fGetGateInfoPath = ""
        iIniKeka = 0
        nCorner = 0                                         ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
        nCornerGoki = 0                                     ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
        Exit Function
    'EG20 V2.1.0.1 DEL START
'     ElseIf Trim(sFData(4)) = EGR Then
'        'EG-R自改の場合
'        sKeyName = HOSHUPUT_KEY_E_NAME
'     ElseIf Trim(sFData(4)) = NEG Then
'        'NEG自改の場合
'        sKeyName = HOSHUPUT_KEY_N_NAME
'     End If
    'EG20 V2.1.0.1 DEL END
    'EG20 V2.1.0.1 ADD START
    Else
        sKeyName = HOSHUPUT_KEY_NAME
    End If
    'EG20 V2.1.0.1 ADD END
     sToCreateFileName = ""
     iIniKeka = GetPrivateProfileString(HOSHUPUT_TO_SECTION_NAME, _
                                        sKeyName & iFilType, DEFAILT, _
                                        sToCreateFileName, Len(sToCreateFileName), _
                                        PATH_HOSHU_DATA_FILE)
     If iIniKeka = 0 Then
       fGetGateInfoPath = ""
     Else
       fGetGateInfoPath = sToCreateFileName
     End If

    nCorner = CInt(sFData(GATE_IDX.IDX_RONRI_CORNER))     ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
    nCornerGoki = CInt(sFData(GATE_IDX.IDX_RONRI_GOKI))   ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加

End Function
'V1.7.0.1 ADD END
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : HoshuCopy
'//  機能名称  : 保守データコピー処理
'//  機能概要  : 保守データコピーを行う。
'//
'//              型        名称      意味
'//  引数      : String　sOutPath　  [IN]出力先パス
'//              String  sFromPath   [IN]コピー元パス
'//              Integer nCorner     [IN]コーナ番号
'//              String  sBackupPath [IN]バックアップパス   ' EG20 V5.4.0.1追加
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//                 ファイルコピー成功／失敗判断追加
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【統合TR-No.265修正対応】
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-23  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-13  CODED BY  [TCC] H.Sugimoto
'//                 【コーナ名スペース除去対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Public Function HoshuCopy(sOutPath As String, sFromPath As String)                     ' EG20 V3.4.0.1【統合TR-No.265修正対応】削除
' EG20 V5.4.0.1削除開始
'Public Function HoshuCopy(sOutPath As String, sFromPath As String, nCorner As Integer)  ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
' EG20 V5.4.0.1削除終了
' EG20 V5.4.0.1追加開始
Public Function HoshuCopy(sOutPath As String, sFromPath As String, nCorner As Integer, sBackupPath As String)
' EG20 V5.4.0.1追加終了

    Dim fso         As New FileSystemObject 'ファイルシステムオブジェクト
    Dim sCopyfile As String                 'コピー先
    Dim FileName As String                  '抽出ファイル名
    Dim FileKaku As String                  '拡張子
    Dim bRet As Boolean                     '戻り値
    Dim strCorner As String                 ' コーナ名          ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加
    Dim szUnyouDate As String               ' 運用日付
    Dim szBackupFolder As String            ' バックアップフォルダのパス
    Dim nNullIndex As Integer               ' 文字数ワーク
    
    
'    On Error Resume Next       'V1.12.0.1 DEL
    On Error GoTo COPY_ERR        'V1.12.0.1 ADD
    
    HoshuCopy = False

    'コピー先フォルダの有無確認
    If fso.FolderExists(sOutPath) = False Then
        'コピー先フォルダ作成
        fso.CreateFolder (sOutPath)

    End If

' EG20 V6.1.0.1 削除開始
'' EG20 V3.4.0.1【統合TR-No.265修正対応】追加開始
'    strCorner = gstrCornerName(nCorner - 1)
'' EG20 V3.4.0.1【統合TR-No.265修正対応】追加終了
' EG20 V6.1.0.1 削除終了
' EG20 V6.1.0.1 追加開始
    strCorner = Replace(gstrCornerName(nCorner - 1), " ", "")
' EG20 V6.1.0.1 追加終了

    'ﾌｧｲﾙ名前取得
    psFileNameGet sFromPath, FileName, FileKaku

    'コピー先ファイル名作成
'    sCopyfile = sOutPath & "\" & FileName & "." & FileKaku                 ' EG20 V3.4.0.1【統合TR-No.265修正対応】削除
    sCopyfile = sOutPath & "\" & strCorner & FileName & "." & FileKaku      ' EG20 V3.4.0.1【統合TR-No.265修正対応】追加

    'ファイルコピー（既に存在した場合は上書きするする）
    fso.CopyFile sFromPath, sCopyfile, True
    
    HoshuCopy = True
    
    Set fso = Nothing

' EG20 V5.4.0.1【バックアップファイル対応】追加開始
    ' バックアップファイルの作成処理
    ' ＜入力＞運用日付
    ' ＜入力＞バックアップフォルダ（パス）
    ' ＜入力＞入力ファイル名（テキスト）
    If CheckAppStart(PROC_KANRI) <> 0 Then
        Set Idinf_KansiSettei = New IdInfProc             '監視装置設定エリア
        '参照(自改通信状態)エリア名を設定
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
            Exit Function
        End If
    
        'エリアIDの設定値を取得
        Idinf_KansiSettei.IdLock
        Idinf_KansiSettei.id = IdKansiSet.SET_ID_KANSI_SET_UNYOU_DAY
        Idinf_KansiSettei.IdGet
        szUnyouDate = Idinf_KansiSettei.DataArea(0)
        Idinf_KansiSettei.IdFree

        nNullIndex = InStr(sBackupPath, Chr(0))
        If nNullIndex <> 0 Then
                szBackupFolder = Left(sBackupPath, nNullIndex - 1)
            Else
                szBackupFolder = sBackupPath
            End If

        If Len(szUnyouDate) > 4 Then
            szBackupFolder = szBackupFolder & Right(szUnyouDate, 4) & "\"
        Else
            szBackupFolder = szBackupFolder & szUnyouDate & "\"
        End If
        ' バックアップファイル作成処理
        Call dllSaveBackupFile(sFromPath, strCorner & FileName, szBackupFolder)
    End If
' EG20 V5.4.0.1【バックアップファイル対応】追加終了

'V1.12.0.1 ADD START
    Exit Function
    
COPY_ERR:
    HoshuCopy = False
    Set fso = Nothing
'V1.12.0.1 ADD END

End Function


