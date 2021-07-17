VERSION 5.00
Begin VB.Form frmChangeVer 
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
   Begin VB.Timer tmrAplCheck 
      Enabled         =   0   'False
      Interval        =   3000
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
Attribute VB_Name = "frmChangeVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmChangeVer.frm
'//  パッケージ名：バージョン切替中画面
'//
'//  概要：バージョン切替中画面
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//         フェーズ２対応　切替中画面追加
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

Private iChangeSts As Integer           '処理番号
Private iChangeVerFlag As Integer       '処理フラグ
Private iChengeVerApl As Integer        '監視盤=0、IDU=1
Private bChangeVerSts As Boolean        '切替処理全体の戻り値

'フォルダ構成処理分岐ステータス
Private Const DLLFILE_AtoC = 1          'バージョン１→一時フォルダ
Private Const DLLFILE_BtoA = 2          'バージョン２→バージョン１
Private Const DLLFILE_CtoB = 3          '一時フォルダ→バージョン２
Private Const PARA_BtoA = 4             'バージョン２→バージョン１(パラメフォルダ)
Private Const BACK_CtoA = 5             '一時フォルダ→バージョン１
Private Const BACK_BtoC = 6             'バージョン２→一時フォルダ
Private Const BACK_AtoB = 7             'バージョン１→バージョン２

'ファイルパス
Private DllFolderName As String         'バージョン１(本物)名
Private DllFolderName2 As String        'バージョン２(保存用)名
Private DllFolderName3 As String        '一時フォルダ名
Private ParaFolderName1 As String       'バージョン１パラメ
Private ParaFolderName2 As String       'バージョン２パラメ

'フラグ値
Private Const CHANGE_END = 0            '処理終了
Private Const CHANGE_CONTINU = 1        '処理続行
Private Const CHANGE_RENAME_ERROR = 2   'リネーム異常
Private Const CHANGE_OK = 3             '処理正常
'V1.6.0.1 ADD START
Public lngMAX_Time As Long                    'INI取得設定値
Public lngtime     As Long                    '現在タイマ値
'V1.6.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン切替中画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
  Dim bRet As Boolean
  
  On Error Resume Next
  
  bRet = True
  
  bChangeVerSts = True
 
  tmrMail.Enabled = True
  
  'アプリ起動確認を行い、起動している場合はアプリ終了処理を行う。
  bRet = pfAplEnd

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : バージョン切替中画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

  On Error Resume Next
  
  cmdOK.Visible = False
  
  '「バージョン切替中画面：表示」ログ出力
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_SHORIGAMEN_START, 0)
  
  'メイル受信用のインタバルタイマ値を設定する。
  tmrMail.Interval = MN_MAIL_INTERVAL
  tmrMail.Enabled = False
'V1.6.0.1 DEL START
'  'アプリ起動用インタバルタイマ値設定。
'  tmrAplCheck.Interval = MN_MAIL_INTERVAL
'  tmrAplCheck.Enabled = False
'V1.6.0.1 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : バージョン切替中画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
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
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    On Error Resume Next
    '「バージョン切替中画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_SHORIGAMEN_END, 0)
       
    '自画面を消す。
    Unload Me
   
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrAplCheck_Timer
'//  機能名称  : アプリ起動チェック用タイマ、タイムアップ時処理
'//  機能概要  : タイムアップ時にアプリ起動チェックを行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplCheck_Timer()

   Dim bRet As Boolean

'V1.6.0.1 DEL START
'
'   On Error Resume Next
'
'    If CheckAppStart(PROC_KANRI) = 0 And _
'         CheckAppStart(PROCESS_IDU_LOG) = 0 And _
'         CheckAppStart(PROCESS_LDU_LOG) = 0 Then
'         '管理、IDUログ、LDUログが起動していない=アプリ終了
'         tmrAplCheck.Enabled = False
'         'バージョン切替処理を行う。
'         bRet = psVersionChange
'    End If
'V1.6.0.1 DEL END
   On Error Resume Next
'V1.6.0.1 ADD START
  '待ち時間がINI定義を超えたかどうかチェック
  If lngtime <= lngMAX_Time Then
    'アプリ起動チェックを行う。全アプリが終了したときのみ、初期化処理を行う。
    If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
      'アプリ起動チェックタイマを停止する。
      tmrAplCheck.Enabled = False
      'バージョン切替処理を行う。
      bRet = psVersionChange
    Else
      '起動アプリ有りの場合、タイマを張り直す
       tmrAplCheck.Interval = MN_MAIL_INTERVAL
      '合計経過待ち時間をアップ
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI定義値を超えた場合、初期化処理異常とする。
    'アプリ起動チェックタイマを停止する。
    tmrAplCheck.Enabled = False
    '処理異常を表示
    psChangeVerEnd (1)
  End If
'V1.6.0.1 ADD END
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
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    On Error Resume Next
        
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmChangeVer.Caption, False
        pfFormActive (frmChangeVer.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psChangeVerEnd
'//  機能名称  : 切替結果表示処理
'//  機能概要  : バージョン切替結果の結果文言を表示する。
'//
'//              型        名称      意味
'//  引数      : Integer　iEnd　　　[IN]処理結果
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psChangeVerEnd(iEnd As Integer)
    Dim i As Integer       'カウンタ
    Dim lngErrCode As Long 'エラーコード

    On Error Resume Next
      
    cmdOK.Visible = True
  
    If iEnd = 0 Then
       '切替正常終了時の文言を表示する。
       lblMessage(0) = "正常終了しました。"
       lblMessage(1) = ""
       '「バージョン切替処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_CHANGE_OK, 0)
    Else
       '切替異常時の文言を表示する。
       lblMessage(0) = "異常終了しました。"
       lblMessage(1) = ""
       '「バージョン切替処理異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_CHANGE_ERROR, lngErrCode)
    End If
    
  cmdOK.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfAplEnd
'//  機能名称  : アプリ終了処理
'//  機能概要  : EG-R監視盤アプリ終了処理を行う
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 フェーズ１不具合対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfAplEnd() As Boolean
   Dim uMail As ML_KYOTU_INF           'メール
   Dim bRtn As Boolean                 'メールの戻り値
   Dim lExitCode As Long
   
   On Error Resume Next
   
   If CheckAppStart(PROC_KANRI) <> 0 Then
      'アプリ終了要求を管理に送信する
      uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
      uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
      uMail.udtlHeader.dwProid = RHOSHU_ID
      uMail.udtlHeader.dwSubArea = 0
      bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
      If bRtn <> 0 Then
         '「アプリ起動・終了画面：メール送信正常結果」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                        
         'IDUログ確認
         If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
            'IDUログ終了要求CMD送信
            bRtn = EndIDULog
            If bRtn = False Then
               pfAplEnd = False
               psChangeVerEnd (1)
               Exit Function
            End If
         End If
            
         'LDUログ確認
         If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
            'LDUログ終了要求CMD送信
            bRtn = EndLDULog
            If bRtn = False Then
               pfAplEnd = False
               psChangeVerEnd (1)
               Exit Function
            End If
         End If
         'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
         'V1.6.0.1 ADD END
       
 '        tmrAplCheck.Enabled = True  'V1.6.0.1 DEL
      Else
        '「アプリ起動・終了画面：メール送信異常結果」ログ出力
        lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
        '「アプリ起動・終了画面：アプリ終了処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
        pfAplEnd = False
        psChangeVerEnd (1)
      End If
   'IDUログ確認
   ElseIf CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
          'IDUログ終了要求CMD送信
          bRtn = EndIDULog
          If bRtn = False Then
             pfAplEnd = False
             psChangeVerEnd (1)
          End If
          
          'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
         'V1.6.0.1 ADD END

'          tmrAplCheck.Enabled = True 'V1.6.0.1 DEL
    'LDUログ確認
    ElseIf CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
           'LDUログ終了要求CMD送信
           bRtn = EndLDULog
           If bRtn = False Then
              pfAplEnd = False
              psChangeVerEnd (1)
           End If
       'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
        'V1.6.0.1 ADD END
        'tmrAplCheck.Enabled = True 'V1.6.0.1 DEL
    Else
        'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
        'V1.6.0.1 ADD END
        'tmrAplCheck.Enabled = True 'V1.6.0.1 DEL
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psVersionChange
'//  機能名称  : バージョン切替処理
'//  機能概要  : 監視盤/IDUバージョン切替処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function psVersionChange() As Boolean

    Dim bRet As Boolean
    
    On Error Resume Next
     
    iChangeVerFlag = 0
        
    bRet = False
    
    '処理フォルダパス設定
    psFolderPathSettei
        
    'フォルダ構成チェック：正常、又はリカバリ処理フォルダ構成以外は処理終了
    If iChengeVerApl = stsKansi Then
      '監視盤アプリ
      bRet = pfKansiChkFolder
    Else
      'IDUアプリ
      bRet = pfIDUChkFolder
    End If
    
    If bRet = False Then
       '処理対象外フォルダ構成状態のため処理を行わずに終了。
       psChangeVerEnd (1)
       Exit Function
    End If
    
    'フォルダ構成分岐処理
    Select Case iChangeSts
       Case DLLFILE_AtoC
            'フォルダ構成正常：正常処理シーケンスを行う。
            bRet = pfDLLFILE_AtoC
       
       Case BACK_BtoC
            'フォルダ構成異常１：バージョン２→一時フォルダへのリカバリ処理を行う。
            bRet = pfBACK_BtoC
    
       Case BACK_AtoB
            'フォルダ構成異常２：バージョン１→バージョン２へのリカバリ処理を行う。
            bRet = pfBACK_AtoB
    
       Case BACK_CtoA
            'フォルダ構成異常３：一時フォルダ→バージョン１へのリカバリ処理を行う。
            bRet = pfBACK_CtoA
    End Select
    
    If bChangeVerSts = False Then
       psChangeVerEnd (1)
       Exit Function
    End If
    
    psVersionChange = bRet

    psChangeVerEnd (0)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psFolderPathSettei
'//  機能名称  : フォルダパス設定処理
'//  機能概要  : 監視盤/IDUバージョン切替対象フォルダパス設定を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function psFolderPathSettei() As Boolean

   On Error Resume Next
 
    Select Case Change_Version
         Case EGR_CHANGE_VER
            DllFolderName = Mid(PATH_GATE_E, 1, Len(PATH_GATE_E) - 2)               'バージョン１(本物)名
            DllFolderName2 = PATH_GATE_ESAVE        'バージョン２(保存用)名
            DllFolderName3 = PATH_GATE_ETEMP        '一時フォルダ名
            ParaFolderName1 = PATH_GATE_EPARA       'バージョン１パラメ
            ParaFolderName2 = PATH_GATE_ESAVE_PARA  'バージョン２パラメ
            iChengeVerApl = stsKansi
         Case NEG_CHANGE_VER
            DllFolderName = Mid(PATH_GATE_N, 1, Len(PATH_GATE_N) - 2)             'バージョン１(本物)名
            DllFolderName2 = PATH_GATE_NSAVE        'バージョン２(保存用)名
            DllFolderName3 = PATH_GATE_NTEMP        '一時フォルダ名
            ParaFolderName1 = PATH_GATE_NPARA       'バージョン１パラメ
            ParaFolderName2 = PATH_GATE_NSAVE_PARA  'バージョン２パラメ
            iChengeVerApl = stsKansi
         Case ICM_CHANGE_VER
            DllFolderName = PATH_IDU_APP & PATH_PROHAN_FOLDER  'バージョン１(本物)名
            DllFolderName2 = PATH_IDU_APP & PATH_PROHAN_SAVE   'バージョン２(保存用)名
            DllFolderName3 = PATH_IDU_APP & PATH_PROHAN_TEMP   '一時フォルダ名
            iChengeVerApl = stsIDU
         Case PASMO_CHANGE_VER
            DllFolderName = PATH_IDU_APP & PATH_CMN_UNT_FOLDER     'バージョン１(本物)名
            DllFolderName2 = PATH_IDU_APP & PATH_CMN_UNT_SAVE     'バージョン２(保存用)名
            DllFolderName3 = PATH_IDU_APP & PATH_CMN_UNT_TEMP     '一時フォルダ名
            iChengeVerApl = stsIDU
         Case JIKIUNCHIN_CHANGE_VER
            DllFolderName = PATH_UNAKI_SUB             'バージョン１(本物)名
            DllFolderName2 = PATH_UNKAI_SUB_SAVE       'バージョン２(保存用)名
            DllFolderName3 = PATH_UNKAI_SUB_TEMP       '一時フォルダ名
            ParaFolderName1 = PATH_UNKAI_SUB_BACK      'バージョン１バックアップ
            ParaFolderName2 = PATH_UNKAI_SUB_SAVE_BACK 'バージョン２バックアップ
            iChengeVerApl = stsKansi
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfKansiChkFolder
'//  機能名称  : バージョン切替処理(フォルダチェック)
'//  機能概要  : フォルダ構成状態チェックを行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：TRUE=フォルダ有り　False=フォルダ無し
'///////////////////////////////////////////////////////////////////
Private Function pfKansiChkFolder() As Boolean
                                      
   Dim fso As New FileSystemObject
   Dim bRet1 As Boolean '検索１
   Dim bRet2 As Boolean '検索２
   Dim bRet3 As Boolean '検索３
   Dim bRet4 As Boolean '検索４
   Dim bRet5 As Boolean '検索５
   
   On Error Resume Next
       
    pfKansiChkFolder = True
   
   '「バージョン切替中画面：フォルダ状態チェック」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_FOLDER_CHK, 0)
    
   'バージョン１(本物フォルダ)有無チェック
    bRet1 = fso.FolderExists(DllFolderName)
   'バージョン２(保存用フォルダ)有無チェック
    bRet2 = fso.FolderExists(DllFolderName2)
   '一時フォルダ有無チェック
    bRet3 = fso.FolderExists(DllFolderName3)
   'パラメータフォルダ有無チェック
    bRet4 = fso.FolderExists(ParaFolderName1)
   'パラメータフォルダ２有無チェック
    bRet5 = fso.FolderExists(ParaFolderName2)
    Set fso = Nothing
 
    If bRet1 = True And bRet4 = True And bRet2 = True And bRet5 = False And bRet3 = False Then
      'バージョン１：有、バージョン２：有、一時フォルダ：無
      '結果：正常状態
      iChangeSts = DLLFILE_AtoC
      
    ElseIf bRet1 = True And bRet4 = False And bRet2 = True And bRet5 = True And bRet3 = False Then
      'バージョン１：有、パラメ：無、バージョン２：有、一時フォルダ：無
      '結果：パラメータリネーム処理異常
      iChangeSts = BACK_BtoC
    
    ElseIf bRet1 = True And bRet4 = False And bRet2 = False And bRet5 = False And bRet3 = True Then
      'バージョン１：有、バージョン２：無、一時フォルダ：有
      '結果：リネーム処理異常１
      iChangeSts = BACK_AtoB
    ElseIf bRet1 = False And bRet4 = False And bRet2 = True And bRet5 = False And bRet3 = True Then
      'バージョン１：無、バージョン２：有、一時フォルダ：有
      '結果：リネーム処理異常２
      iChangeSts = BACK_CtoA
      
    Else
       pfKansiChkFolder = False
    End If
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfIDUChkFolder
'//  機能名称  : バージョン切替処理(フォルダチェック)
'//  機能概要  : フォルダ構成状態チェックを行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：TRUE=フォルダ有り　False=フォルダ無し
'///////////////////////////////////////////////////////////////////
Private Function pfIDUChkFolder() As Boolean
                                
  Dim fso As New FileSystemObject
  Dim bRet1 As Boolean '検索１
  Dim bRet2 As Boolean '検索２
  Dim bRet3 As Boolean '検索３
   
  On Error Resume Next
  
  pfIDUChkFolder = True
   
  '「バージョン切替中画面：フォルダ状態チェック」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_FOLDER_CHK, 0)
    
   'バージョン１(本物フォルダ)有無チェック
   bRet1 = fso.FolderExists(DllFolderName)
   'バージョン２(保存用フォルダ)有無チェック
   bRet2 = fso.FolderExists(DllFolderName2)
   '一時フォルダ有無チェック
   bRet3 = fso.FolderExists(DllFolderName3)
   Set fso = Nothing
   
   If bRet1 = True And bRet2 = True And bRet3 = False Then
      'バージョン１：有、バージョン２：有、一時フォルダ：無
      '結果：正常状態
      iChangeSts = DLLFILE_AtoC
   ElseIf bRet1 = True And bRet2 = False And bRet3 = True Then
      'バージョン１：有、パラメ：無、バージョン２：有、一時フォルダ：無
      '結果：リカバリ処理２
      iChangeSts = BACK_AtoB
      
   ElseIf bRet1 = False And bRet2 = True And bRet3 = True Then
      'バージョン１：有、バージョン２：無、一時フォルダ：有
      '結果：リネーム処理１
      iChangeSts = BACK_CtoA
      
   Else
      pfIDUChkFolder = False
   End If
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfDLLFILE_AtoC
'//  機能名称  : 切替処理１
'//  機能概要  : 共通切替処理(バージョン１→一時フォルダ)リネーム処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfDLLFILE_AtoC() As Boolean
                               
  On Error Resume Next
                             
  Dim fso As New FileSystemObject 'ファイルシステムオブジェクト
  Dim bRet As Boolean
  
  bRet = True
  
  'バージョン１→一時フォルダにリネーム
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName, DllFolderName3
  Set fso = Nothing
 
  bRet = pfDLLFILE_BtoA
  If iChangeVerFlag = CHANGE_OK Then
     bRet = pfBACK_CtoA
  End If
  
  pfDLLFILE_AtoC = bRet
  Exit Function
  
FileCopyError:
  'リネーム異常時は処理終了
  '「バージョン切替中画面：一時フォルダリネーム異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DELETEFOLDER_RENAME_ERROR, 0)
  pfDLLFILE_AtoC = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfDLLFILE_BtoA
'//  機能名称  : 切替処理２
'//  機能概要  : 共通切替処理(バージョン２→バージョン１)リネーム処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfDLLFILE_BtoA() As Boolean
                                
  On Error Resume Next
                               
  Dim fso As New FileSystemObject 'ファイルシステムオブジェクト
  Dim bRet As Boolean
  
  bRet = True
  
  'バージョン２→バージョン１にリネーム
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName2, DllFolderName
  Set fso = Nothing
  
  bRet = pfDLLFILE_CtoB
  If iChangeVerFlag = CHANGE_RENAME_ERROR Then
     '一時フォルダ→バージョン２リネーム異常時処理：バージョン１→バージョン２リネーム
     bRet = True
     If bRet = True Then
        bRet = pfBACK_AtoB
     End If
     If bRet = True Then
        'バージョン１→バージョン２リネーム正常時処理：一時フォルダ→バージョン１リネーム
        bRet = pfBACK_CtoA
     End If
  End If
  
  pfDLLFILE_BtoA = bRet
  
  Exit Function
  
FileCopyError:
  pfDLLFILE_BtoA = False
  iChangeVerFlag = CHANGE_RENAME_ERROR
  '「バージョン切替中画面：DLLフォルダリネーム異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DLLFOLDER_RENAME_ERROR, 0)
  '一時フォルダ→バージョン１リネーム処理
  bRet = pfBACK_CtoA()
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfDLLFILE_CtoB
'//  機能名称  : 切替処理３
'//  機能概要  : 共通切替処理(一時フォルダ→バージョン２)リネーム処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfDLLFILE_CtoB() As Boolean
                                
   On Error Resume Next
                               
  Dim fso As New FileSystemObject 'ファイルシステムオブジェクト
  Dim bRet As Boolean
    
  bRet = True
  
  '一時フォルダ→バージョン２にリネーム
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName3, DllFolderName2
  Set fso = Nothing
  
  If iChengeVerApl = 0 Then
     '監視盤アプリのみパラメータリネームを行う。
     bRet = pfPARA_BtoA
     If iChangeVerFlag = CHANGE_RENAME_ERROR Then
        'バージョン２パラメ→バージョン１パラメ異常時処理：バージョン２→一時フォルダリネーム
        bRet = pfBACK_BtoC
     End If
  End If
  
  pfDLLFILE_CtoB = bRet

  Exit Function
  
FileCopyError:
  '「バージョン切替中画面：保存用フォルダリネーム異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_BACKUPFOLDER_RENAME_ERROR, 0)
  pfDLLFILE_CtoB = False
  iChangeVerFlag = CHANGE_RENAME_ERROR
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfPARA_BtoA
'//  機能名称  : 切替処理４
'//  機能概要  : 共通切替処理(バージョン２→バージョン１)パラメリネーム処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfPARA_BtoA() As Boolean
                                
   On Error Resume Next
                               
  Dim fso As New FileSystemObject 'ファイルシステムオブジェクト
  Dim bRet As Boolean
  
  bRet = True
  
  'バージョン２→バージョン１にパラメリネーム
  On Error GoTo FileCopyError
  fso.MoveFolder ParaFolderName2, ParaFolderName1
  Set fso = Nothing
  
  pfPARA_BtoA = bRet
  
  Exit Function
  
FileCopyError:
  '「バージョン切替中画面：サブフォルダリネーム異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_SUBFOLDER_RENAME_ERROR, 0)
  pfPARA_BtoA = False
  iChangeVerFlag = CHANGE_RENAME_ERROR
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfBACK_CtoA
'//  機能名称  : 切替処理５
'//  機能概要  : 共通切替処理(一時フォルダ→バージョン１)リネーム処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfBACK_CtoA() As Boolean
                                
  On Error Resume Next
                               
  Dim fso As New FileSystemObject 'ファイルシステムオブジェクト
  Dim bRet As Boolean
  
  bRet = True
  
  '一時フォルダ→バージョン１にリネーム
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName3, DllFolderName
  Set fso = Nothing
  
  'フォルダ構成が以下の時、C→A後に正常シーケンスを行う。
  If iChangeSts = BACK_CtoA Or iChangeSts = BACK_BtoC Then
     'フォルダ正常：正常処理を行う。
     bRet = pfDLLFILE_AtoC
  End If
  
  pfBACK_CtoA = bRet
  
  Exit Function
  
FileCopyError:
  '「バージョン切替中画面：DLLフォルダリネーム異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DLLFOLDER_RENAME_ERROR, 0)
  pfBACK_CtoA = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfBACK_BtoC
'//  機能名称  : 切替処理６
'//  機能概要  : 共通切替処理(バージョン２→一時フォルダ)リネーム処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfBACK_BtoC() As Boolean
                                  
  Dim fso As New FileSystemObject 'ファイルシステムオブジェクト
  Dim bRet As Boolean
  
  On Error Resume Next
  
  bRet = True
    
  'バージョン２→一時フォルダにリネーム
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName2, DllFolderName3
  Set fso = Nothing
  
  If iChangeSts = BACK_BtoC Then
     'バージョン１→バージョン２へのリカバリ処理
      bRet = pfBACK_AtoB

     If bRet = True Then
        '一時フォルダ→バージョン１へのリカバリ処理
        bRet = pfBACK_CtoA
     End If
  End If
  
  pfBACK_BtoC = bRet
  
  Exit Function
   
FileCopyError:
  '「バージョン切替中画面：一時フォルダリネーム異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DELETEFOLDER_RENAME_ERROR, 0)
  pfBACK_BtoC = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfBACK_AtoB
'//  機能名称  : 切替処理７
'//  機能概要  : 共通切替処理(バージョン１→バージョン２)リネーム処理
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfBACK_AtoB() As Boolean
                                                              
  Dim fso As New FileSystemObject 'ファイルシステムオブジェクト
  Dim bRet As Boolean
  
  On Error Resume Next
  
  bRet = True
  
  'バージョン１→バージョン２にパラメリネーム
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName, DllFolderName2
  Set fso = Nothing
  
  If iChangeSts = BACK_AtoB Then
     '一時フォルダ→バージョン１へのリカバリ処理を行う。
      bRet = pfBACK_CtoA
      
      If bRet = True Then
         'フォルダ正常：正常処理を行う。
         bRet = pfDLLFILE_AtoC
      End If
  End If
  
  pfBACK_AtoB = bRet
  Exit Function
  
FileCopyError:
  '「バージョン切替中画面：保存用フォルダリネーム異常」ログ出力
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_BACKUPFOLDER_RENAME_ERROR, 0)
  pfBACK_AtoB = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

