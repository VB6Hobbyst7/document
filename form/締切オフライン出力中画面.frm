VERSION 5.00
Begin VB.Form frmShimekiriOfflineOut 
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
      Interval        =   3000
      Left            =   360
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
Attribute VB_Name = "frmShimekiriOfflineOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  ファイル名  ：frmShimekiriOfflineOut.frm
'//  パッケージ名：締切データオフライン出力中画面
'//
'//  概要：締切データオフライン出力中画面
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.10.0.1) 2012-05-09   CODED   BY [TCC] H.Sugimoto
'//                 【保守締切機能改善】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////

Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 締切データ出力中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    On Error Resume Next
    
    ' オフライン出力中のガイドを表示する｡
    lblMessage(0) = "締切データをオフライン出力中です。"
    lblMessage(1) = "しばらくお待ち下さい。"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
    tmrMail2.Enabled = True     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 締切データ出力中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next
    
    'メール受信用タイマを止める
    tmrMail.Enabled = False
    tmrMail2.Enabled = False     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 締切データ出力中画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_SHIMEKIRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    tmrMail2.Interval = MN_MAIL_INTERVAL     ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    tmrMail2.Enabled = False                 ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
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
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()

    On Error Resume Next
    
    '自画面を消す。
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
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
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    On Error Resume Next
    
    ' メール受信用タイマを止める
    tmrMail.Enabled = False
    
' EG20 V8.1.0.1【EG20_KANSI05_01】DEL START
'    ' 汎用メイル受信処理を行う
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
'        AppActivate frmSyusyuOutPut.Caption, False
'    End If
' EG20 V8.1.0.1【EG20_KANSI05_01】DEL END
' EG20 V6.3.0.1【機能見直し】削除開始
'    ' 出力ファイル作成処理を行う。
'    frmShimekiriData.gbShimekiriResult = sOutPutOfflineData
'
' EG20 V6.3.0.1【機能見直し】削除終了
' EG20 V6.3.0.1【機能見直し】追加開始
    If frmShimekiriData.glShimekiriType = 1 Then
        ' 出力ファイル作成処理を行う。
        frmShimekiriData.gbShimekiriResult = sOutPutOfflineData
    Else
        frmShimekiriData.gbShimekiriResult = sReOutPutOfflineData
    End If
' EG20 V6.3.0.1【機能見直し】追加終了

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    If frmShimekiriData.gbShimekiriResult = True Then
        lblMessage(0) = "正常終了しました。"
        lblMessage(1) = ""
    Else
        lblMessage(0) = "異常終了しました。"
        lblMessage(1) = ""
    End If
    cmdOK.Enabled = True
    
End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
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
'//     ORIGINAL  :(EG20 V8.1.0.1) 2014-06-05  CODED  BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()

    On Error Resume Next

    ' 汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmShimekiriOfflineOut.Caption, False
        pfFormActive (frmShimekiriOfflineOut.hwnd)
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  関数名称  : sOutPutOfflineData
'//  機能名称  : オフラインデータ媒体出力処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : BOOL      TRUE      正常
'//                        FALSE     異常
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.10.0.1) 2012-05-09   CODED   BY [TCC] H.Sugimoto
'//                 【保守締切機能改善】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-01  CODED   BY [TCC]T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sOutPutOfflineData() As Boolean
            
    Dim nListCnt As Integer                             ' ファイル格納数
    Dim szFileName As String                            ' ファイル名
    Dim lResult As Long                                 ' 処理結果
    Dim dwCorner As Long                                ' コーナ
    Dim dwSequense As Long                              ' シーケンス番号
    Dim szWork As String                                ' ワーク
    Dim szNameWork As String                            ' ワーク
    
    gsGetCornerType     '各コーナーのタイプを取得       EG20 V30.1.0.1
            
    ' //////////////////////////////////////////////////////////////
    ' // ファイル作成処理
    For nListCnt = 0 To UBound(gOfflineFileList) - 1    ' ファイルリスト数
    
        szFileName = gOfflineFileList(nListCnt)         ' ファイル名の取得
        
' EG20 V5.10.0.1 削除開始
' 対象ファイル変更
' （変更前）「HOSHU_SIMEKIRI01_001.DAT」
' （変更後）「SIMEKIRI01.DAT」
'
'        ' 「HOSHU_SIMEKIRI01_001.DAT」のコーナ番号とシーケンス番号を抽出
'        szNameWork = Right(szFileName, 24)
'        szWork = Mid(szNameWork, 15, 2)
'        dwCorner = CInt(szWork)
'        szWork = Mid(szNameWork, 18, 3)
'        dwSequense = CInt(szWork)
' EG20 V5.10.0.1 削除終了
' EG20 V5.10.0.1 追加開始
        ' 「SIMEKIRI01.DAT」のコーナ番号を抽出
        szNameWork = Right(szFileName, 14)
        szWork = Mid(szNameWork, 9, 2)              ' コーナ番号
        dwCorner = CInt(szWork)
        dwSequense = 0                              ' シーケンス番号:0固定
' EG20 V5.10.0.1 追加終了

        'EG20 V30.1.0.1 DEL START
'        lResult = dllCreateShimekiriFile(dwCorner, dwSequense, _
'                                frmShimekiriData.glbFilePath, _
'                                szFileName)
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        If gintCornerType(dwCorner - 1) = CORNER_TYPE_KANSEN Then
            lResult = dllCreateShimekiriFileKan(dwCorner, dwSequense, _
                                    frmShimekiriData.glbFilePath, _
                                    szFileName)
        Else
            lResult = dllCreateShimekiriFile(dwCorner, dwSequense, _
                                    frmShimekiriData.glbFilePath, _
                                    szFileName)
        End If
        'EG20 V30.1.0.1 ADD END
        If lResult = False Then
            sOutPutOfflineData = False
            Exit Function
        End If
    Next

    sOutPutOfflineData = True
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  関数名称  : sReOutPutOfflineData
'//  機能名称  : オフラインデータ媒体再出力処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : BOOL      TRUE      正常
'//                        FALSE     異常
'//
'//     ORIGINAL  :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sReOutPutOfflineData() As Boolean
            
    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim nListCnt As Integer                             ' ファイル格納数
    Dim szSrcFileName As String                         ' 出力ファイル名
    Dim szDstFileName As String                         ' 保存先ファイル名
    Dim FileName As String                              ' ファイル名
    Dim FileKaku As String                              ' 拡張子
    
    On Error GoTo ErrorHandler                          ' エラーハンドルの登録
            
    ' //////////////////////////////////////////////////////////////
    ' // ファイル作成処理
    For nListCnt = 0 To UBound(gOfflineFileList) - 1    ' ファイルリスト数
    
        szSrcFileName = gOfflineFileList(nListCnt)      ' ファイル名の取得
        If objFso.FileExists(szSrcFileName) = True Then
            
            ' ファイル名取得
            psFileNameGet szSrcFileName, FileName, FileKaku
            
            ' コピー先ファイル名作成
            szDstFileName = frmShimekiriData.glbFilePath & "\" & FileName & "." & FileKaku
            
            'ファイルコピー（既に存在した場合は上書きするする）
            objFso.CopyFile szSrcFileName, szDstFileName, True
        
        End If
        
    Next

    sReOutPutOfflineData = True
    Set objFso = Nothing
    Exit Function
' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:

    Set objFso = Nothing
    sReOutPutOfflineData = False

End Function


