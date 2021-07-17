VERSION 5.00
Begin VB.Form frmRenewOutput 
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
   Begin VB.Timer tmrOutput 
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
Attribute VB_Name = "frmRenewOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmRenewOutput.frm
'//  パッケージ名：係員設定媒体出力中画面
'//
'//  概要：係員設定媒体出力中画面
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 【圧縮フォルダ指定対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////

Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 係員設定媒体出力中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    cmdOK.Enabled = False
    
    On Error Resume Next
    
    tmrMail.Enabled = True
    
'    保存中のガイドを表示する｡
    lblMessage(0) = "設定値を出力中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    DoEvents
    
    tmrOutput.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 係員設定媒体出力中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next

    'メール受信用タイマを止める
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 係員設定媒体出力中画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim i As Integer 'カウンタ
    Dim intCount As Integer
    Dim intCount2 As Integer
    
    On Error Resume Next
    
    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    tmrOutput.Interval = 100
    tmrOutput.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()

    On Error Resume Next
    
    '自画面を消す。
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    Dim intStatus As Integer         '受信メールチェック結果

    On Error Resume Next
    
    'メールを受信する。
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
        '受信メールがあれば、メールＩＤ毎の処理をする。
        Select Case udtReadMail.udtlHeader.dwId        'メールＩＤ
            Case ML_ID_PROEND_ORD
                '「プロセス終了指示」を受信した場合、
                '「プロセス終了指示受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'プロセスの終了処理を行う
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示」を受信した場合
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '表示元画面（保守データ収集画面）をアクティブ表示する。
                AppActivate frmRenewOutput.Caption, False
                pfFormActive (frmRenewOutput.hwnd)	' EG20 V8.1.0.1 【EG20_KANSI05_01】ADD
            Case Else
                 'その他のメールを受信した場合
                 '「メールID不正」ログ出力
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sOutput_Data
'//  機能名称  : 設定値出力処理
'//  機能概要  : 設定値を編集して媒体出力する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-27   CODED   BY [TCC] M.Matsumoto
'//                【統合No56対応】
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 【圧縮フォルダ指定対応】
'//     REVISIONS :（EG20 V30.1.0.1) 2014-04-02 CODED BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sOutput_Data()

    Dim bySyoAssort As Byte             'ログ用小分類
    Dim strFilePath As String           '出力ファイルパス
    Dim strCornerPath As String         '設定ファイルパス
    Dim strStationNm As String          '駅名
    Dim strCornerNm As String           'コーナ名
    Dim intCount As Integer             'カウンタ
    Dim intCount2 As Integer            'カウンタ
    Dim intOutFile As Integer           '出力ファイル番号
    Dim intTgtFileNo As Integer         '出力対象設定ファイル番号
    Dim strTgtFileName As String        '出力対象設定ファイル
    Dim strTargetFile() As String       '出力対象ファイル
    Dim strTargetFileKan() As String    '出力対象ファイル【幹線コーナ向け】 'EG20 V30.1.0.1 ADD
    Dim intFileNum As Integer
    Dim strDefault As String
    Dim strRet As String * 32
    Dim lngRet As Long
    Dim sLzhDirName As String
    Dim sLzhFileName As String
    Dim strCabTarget As String
    Dim lngRetZip As Long
    Dim objFileObj As FileSystemObject  'ファイルシステムオブジェクト
    Const lngBufSize = 32
    Dim nIndex As Integer               ' 文字数                    ' EG20 V5.6.0.1追加
    
    On Error GoTo Err_Handler
    
    sLzhDirName = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    If sLzhDirName = "" Then
        Unload Me
        Exit Sub  'ディレクトリが指定されなければ、処理終了
    End If
    
' EG20 V5.6.0.1【圧縮フォルダ指定対応】追加開始
    ' 出力フォルダに半角スペースが含まれている場合、圧縮で異常が発生してしまうため
    ' 圧縮前にチェックして異常を表示する。
    nIndex = InStr(sLzhDirName, " ")
    If nIndex <> 0 Then
        ' 警告ポップアップウィンドウを表示する。
        Call MsgBox(CABFOLDERSELECT_ERRORMESSAGE, vbCritical, CABFOLDERSELECT_ERRORTITLE)
        Unload Me
        Exit Sub  'ディレクトリが指定されなければ、処理終了
    End If
' EG20 V5.6.0.1【圧縮フォルダ指定対応】追加終了

    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_KAKARI_OUTPUT)
    
    Set objFileObj = New FileSystemObject
    
    '出力対象設定ファイルをオープンする。
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE
    
    '出力対象設定ファイルが存在しない場合は異常終了
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_Handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '出力対象ファイル数を取得
    Input #intTgtFileNo, intFileNum
    
    '出力対象ファイルを取得
    ReDim strTargetFile(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFile)
        Input #intTgtFileNo, strTargetFile(intCount)
    Next
    
    Close #intTgtFileNo
    
    'EG20 V30.1.0.1 ADD START
    '幹線コーナーに対する出力対象ファイルの内容を確保する
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE_KAN
    
    '出力対象設定ファイルが存在しない場合は異常終了
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_Handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '出力対象ファイル数を取得
    Input #intTgtFileNo, intFileNum
    
    '出力対象ファイルを取得
    ReDim strTargetFileKan(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFileKan)
        Input #intTgtFileNo, strTargetFileKan(intCount)
    Next
    
    Close #intTgtFileNo
    'EG20 V30.1.0.1 ADD END
        
    '選択コーナについて出力処理をする
    For intCount = 0 To UBound(glngTergetCorner)
        If glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON Then
            'コーナ１
            If intCount = 0 Then
                'Iniファイルから駅名を取得
                lngRet = GetPrivateProfileString(strAppName_station, STATIONINI_KEY_EKINAME, _
                                            strDefault, strRet, lngBufSize, KANSI_STATION_INI_FILE)
            'コーナ１以外
            Else
                'Iniファイルから駅名を取得
                lngRet = GetPrivateProfileString(strAppName_station & CStr(intCount + 1), STATIONINI_KEY_EKINAME, _
                                            strDefault, strRet, lngBufSize, KANSI_STATION_INI_FILE)
            End If
            '出力ファイル名編集
            strStationNm = Trim(strRet)
            strStationNm = Replace(strStationNm, Chr(0), "")
            strStationNm = Replace(strStationNm, " ", "")           'EG20 V5.5.0.1 ADD 【統合No56対応】
            strCornerNm = gstrCornerName(intCount)
            strCornerNm = Replace(strCornerNm, Chr(0), "")
            strCornerNm = Replace(strCornerNm, " ", "")             'EG20 V5.5.0.1 ADD 【統合No56対応】
            strFilePath = strStationNm & "_" & strCornerNm & OUTPUT_LIST_FILE
            sLzhFileName = strStationNm & "_" & strCornerNm & OUTPUT_CAB_FILE
            strFilePath = sLzhDirName & strFilePath
            sLzhFileName = sLzhDirName & sLzhFileName
            
            '---- 設定一覧テキスト作成 開始
            'ファイル作成
            If objFileObj.FileExists(strFilePath) = True Then
                objFileObj.DeleteFile (strFilePath)
            End If
            Call objFileObj.CreateTextFile(strFilePath)
            
            '出力ファイルをオープンする。
            intOutFile = FreeFile
            Open strFilePath For Output As #intOutFile
    
            '設置駅・コーナ名出力
            Print #intOutFile, "設置駅：" & strStationNm
            Print #intOutFile, "設置コーナ：" & strCornerNm
            Print #intOutFile, ""
            
            'ID設定値を出力
            If gsubOutput_Id(intCount + 1, intOutFile) = False Then
                GoTo Err_Handler
            End If

            'EG20 V30.1.0.1 DEL START
            '入出場フリーファイルを出力
'            If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
'
'            '祝祭日ファイルを出力
'            If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
            'EG20 V30.1.0.1 DEL END
            
            'EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                '幹線コーナの場合

                '新幹線不正パラメータを出力
                If gsubOutput_ParaKan(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
                '在来線IC判定パラメータを出力
                If gsubOutput_ParaKan(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
                
                '在来線IC通過処理パラメータを出力
                If gsubOutput_ParaKan(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
            Else
                '在来コーナーの場合
                '入出場フリーファイルを出力
                If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
                
                '祝祭日ファイルを出力
                If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
            End If
            'EG20 V30.1.0.1 ADD END
            
            Close #intOutFile
            '---- 設定一覧テキスト作成 終了
            
            '---- 設定保存圧縮ファイル作成 開始
            'コーナ別設定ファイルパス
            strCornerPath = PATH_OPERATE_CORNER & CStr(intCount + 1) & PATH_OPERATE_SETTEI
            
            strCabTarget = Empty
            '出力対象ファイル名設定
            ' EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                '幹線コーナーの場合は幹線コーナー用の対象ファイルで処理する
                For intCount2 = 0 To UBound(strTargetFileKan)
                    strCabTarget = strCabTarget & strCornerPath & strTargetFileKan(intCount2) & " "
                Next
            Else
                '在来コーナーの場合は在来コーナー用の対象ファイルで処理する
                For intCount2 = 0 To UBound(strTargetFile)
                    strCabTarget = strCabTarget & strCornerPath & strTargetFile(intCount2) & " "
                Next
            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.1.0.1 DEL START
'            For intCount2 = 0 To UBound(strTargetFile)
'                strCabTarget = strCabTarget & strCornerPath & strTargetFile(intCount2) & " "
'            Next
            'EG20 V30.1.0.1 DEL END
            
            lngRetZip = gsubCabZip(sLzhFileName, strCabTarget)
            
            If (lngRetZip <> 0) Then   '圧縮結果が正常(0)以外
                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LZH_ERROR, 0)
                GoTo Err_Handler
            End If
            '---- 設定保存圧縮ファイル作成 終了
        End If
    Next intCount
    
    Set objFileObj = Nothing
    
    lblMessage(0).Caption = "正常終了しました。"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    DoEvents
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    Exit Sub
    
'エラー処理
Err_Handler:

    If intTgtFileNo > 0 Then
        Close #intTgtFileNo
    End If
    If intOutFile > 0 Then
        Close #intOutFile
    End If

    Set objFileObj = Nothing
    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, bySyoAssort, KAKARISET_OUTPUT_ERR, 0)
    
    lblMessage(0).Caption = "異常終了しました。"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : tmrOutput_Timer
'//  機能名称  : 出力処理実行タイマ
'//  機能概要  : 設定値を編集して媒体出力する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-02   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrOutput_Timer()

    On Error Resume Next
    
    tmrOutput.Enabled = False
    Call sOutput_Data
     
End Sub
