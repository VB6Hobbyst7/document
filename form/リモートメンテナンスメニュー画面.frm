VERSION 5.00
Begin VB.Form frmRmenteMenu 
   BorderStyle     =   0  'なし
   Caption         =   "リモートメンテナンス"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
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
      Enabled         =   0   'False
      Left            =   6120
      Top             =   5640
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
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   3255
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
      Index           =   1
      Left            =   6360
      TabIndex        =   1
      Top             =   960
      Width           =   3255
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
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "リモートメンテナンス"
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
      TabIndex        =   3
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmRmenteMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmRmenteMenu.frm
'//  パッケージ名：リモートメンテナンスメニュー画面
'//
'//  概要：リモートメンテナンスメニュー画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.37】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private sTOOLPass As String

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : リモートメンテナンスメニュー画面(アクティブ時)
'//  機能概要  : メール受信用タイマ、起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
     pfFormActive (hwnd)
    'タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : リモートメンテナンスメニュー画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマ、停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'タイマを止める
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : リモートメンテナンスメニュー画面(ロード時)
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
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next

   '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'IDU縮退チェック
    psIDUCheck

    If pbIDUSts = 1 Then
      'IDU業務非表示
       cmdFixedExe(1).Visible = False
    End If
    'V1.3.0.1 ADD START
    'メイル受信用のメイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称画面に遷移等を行う。
'//              「自改」「判定ＩＣ－Ｍ」
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　 [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
  On Error Resume Next
  Dim lRetVal As Double     'Shell関数戻り値

  Select Case Index
        Case 0                                 '自改
           '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：自改釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_JIKAI_BUTTOM, 0)
           Load frmRMente
           frmRMente.Show 1
        Case 1                                '判定IC-M
           '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：判定IC-M釦押下」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_ICM_BUTTOM, 0)
           fMakeICMGLTFile
           psICMRMenteTool
           If sTOOLPass = "" Then
              Exit Sub
           Else
              '判定IC-Mツール起動
            lRetVal = Shell(sTOOLPass, vbNormalFocus)
          End If
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
    
    '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psICMRMenteTool
'//  機能名称  : 判定IC-Mのリモートメンテナンスツールパスを取得処理
'//  機能概要  : 判定IC-Mのリモートメンテナンスツールパスを取得を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-01 REVISED BY [TCC] T.Koyama
'//                ＥＧ２０フェーズ２対応【残件№54、監視D-154】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Public Sub psICMRMenteTool()
 
    Dim sPath As String * MAX_PATH_SIZE
    Dim iRet As Integer
    
    Dim sMyPath As String               'EG20 V2.0.1.1 ADD
    
    On Error Resume Next
    
    ' HOSHU.INIより判定IC-Mツールパスを取得する。
    iRet = GetPrivateProfileString(KANSI_HOSHU_ICM_RMENTE_SEC, _
                                    KANSI_HOSHU_ICM_RMENTE_KEY, _
                                    DEFAILT, sPath, Len(sPath), _
                                    HOSHU_FILE)

    sMyPath = Replace(sPath, Chr(0), "")
      
      If iRet = 0 Then
        sTOOLPass = ""
      Else
'        sTOOLPass = sPath              'EG20 V2.0.1.1 DEL
        sTOOLPass = sMyPath             'EG20 V2.0.1.1 ADD
      End If
      
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMakeICMGLTFile
'//  機能名称  : 自駅.GLTファイルへの自改情報を書き込み処理
'//  機能概要  : GATE.INIを参照し、自駅.GLTファイルへ、
'//              号機番号、表示文字、IPアドレスを書き込む。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【結合TR-No.37】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeICMGLTFile() As Integer
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
    Dim szIniFilePath As String     ' INIファイルパス   ' EG20 V3.3.0.1【結合TR-No.37】追加

    On Error Resume Next
    MkDir PATH_RMENTE_ICM_DEN   '自改用電鉄フォルダを作成する。（GLTファイル用）
    'GLTファイルを開く。ファイルが存在しなければ新規に作成される。
    On Error GoTo ErrorHandlerGLTFile
    intGLTFileNo = FreeFile        ' 未使用のファイル番号を取得する。
    Open ICM_GLT_FILE_FULLPASS For Output As #intGLTFileNo     ' GLTファイルを開く。

    For iGate = CNT_MIN To MAX_GATE_NO - 1
' EG20 V3.3.0.1【結合TR-No.37】削除開始
'      '自動改札機情報取得
'      sKeyName = "gate" & Format(iGate + 1, "00")
'      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
'                                     sKeyName, _
'                                     DEFAILT, sGateData, Len(sGateData), _
'                                     PATH_GATE_FILE)
' EG20 V3.3.0.1【結合TR-No.37】削除終了
' EG20 V3.3.0.1【結合TR-No.37】追加開始
        ' IDUのICM.INIから改札機情報を取得
        szIniFilePath = PATH_IDU_APP & IDU_ICM_FILE
        sKeyName = "icm" & Format(iGate + 1, "00")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ICM, _
                                    sKeyName, _
                                    DEFAILT, sGateData, Len(sGateData), _
                                    szIniFilePath)
' EG20 V3.3.0.1【結合TR-No.37】追加終了
      
      If iRet = 0 Then
         '「ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：自動改札機INIファイル読込異常」ログ出力
         Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
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
      
      If Len(Trim(sFData(1))) = 1 Then
         '号機番号が１桁ならば、先頭に０を付加する。
         sGoukiNo = "0" & Trim(sFData(1)) & "号機"
      Else
         sGoukiNo = Trim(sFData(1)) & "号機"
      End If
        
' EG20 V3.3.0.1【結合TR-No.37】削除開始
'      If Trim(sFData(4)) <> "＊" Then
'         'Gate.iniファイルの号機番号表示文字、IPアドレスをGLTファイルに書き込む。
'         Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(14))
'      End If
' EG20 V3.3.0.1【結合TR-No.37】削除終了
' EG20 V3.3.0.1【結合TR-No.37】追加開始
    If Trim(sFData(5)) <> "＊" Then
        'ICM.iniファイルの号機番号表示文字、IPアドレスをGLTファイルに書き込む。
        Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(7))
    End If
' EG20 V3.3.0.1【結合TR-No.37】追加終了
              
    Next
    
    'GLTファイルを閉じる。
    Close #intGLTFileNo
    
    fMakeICMGLTFile = 0    '正常終了
    Exit Function

ErrorHandlerGateIni:
   '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイルアクセス異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeICMGLTFile = 1
   'GLTファイルを閉じる。
   Close #intGLTFileNo
   Exit Function
ErrorHandlerGLTFile:
   '「自動改札機ﾘﾓｰﾄﾒﾝﾃﾅﾝｽ画面：ファイルアクセス異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeICMGLTFile = 2
   'GLTファイルを閉じる。
   Close #intGLTFileNo

End Function

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
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmRmenteMenu.Caption, False
        pfFormActive (frmRmenteMenu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

