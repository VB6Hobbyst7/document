VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEkimeiFD 
   BorderStyle     =   0  'なし
   Caption         =   "駅名データ媒体投入"
   ClientHeight    =   9000
   ClientLeft      =   2160
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
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
   ScaleHeight     =   9424.084
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   12121.21
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "媒体取外"
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
      Left            =   9360
      TabIndex        =   3
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   600
      Top             =   7440
   End
   Begin VB.CommandButton cmdFDInput 
      Caption         =   "実行"
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
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
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
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "駅名データ媒体投入"
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
Attribute VB_Name = "frmEkimeiFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 ALL Rights Reserved
'//
'//  ファイル名  ：frmEkimeiFD.frm
'//  パッケージ名：駅名データ媒体投入画面
'//
'//  概要：駅名データ媒体投入画面
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

'ダイアログ表示用ＩＤ
Private Enum FDEkimei
    REBOOT = 1                          '１：正常終了ダイアログ
    FD_INPUT_ERR = 2                    '２：異常終了ダイアログ
End Enum

'ログ出力依頼用ID
Private Enum LogID
    LOG_NORMAL = 0                      '０：駅名データＦＤファイル作成正常
    FILEDELETE_ERROR = 1                '１：駅名データＦＤファイル削除異常
    FILECOPY_ERROR = 2                  '２：駅名データＦＤファイル作成異常
End Enum




'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : cmdInstall_Click
'//  概要      : 「媒体取外」釦押下処理
'//  説明      : 媒体を取り外す。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
 On Error Resume Next
   
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)

End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  概要      : 駅名データＦＤ投入画面がロードされた時のイベントプロシージャ
'//  説明      : メール受信用のタイマ値を設定する。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMEI_DATA_INPUT_GAMEN_START, 0)
    
    '----------------------------------------------------
    '画面初期値設定
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  概要      : 駅名データＦＤ投入画面が表示された時のイベントプロシージャ
'//  説明      : 「メール受信用タイマ」を起動する。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    On Error Resume Next

    'メール受信用タイマを起動する
    tmrMail.Enabled = True

End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  概要      : 駅名データＦＤ投入画面が消去された時のイベントプロシージャ
'//  説明      : 「メール受信用のタイマ」を破棄する。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next

    'メール受信用タイマを止める
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : cmdFDInput_Click
'//  概要      : 「実行」ボタン押下時のイベントプロシージャ
'//  説明      : 駅名データＦＤ投入処理を行う。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFDInput_Click()

    Dim iResponse   As Integer          'ダイアログボタンコード
    Dim bRet        As Boolean          'メール送信判定
    Dim strFilePath As String           'ファイル選択ダイアログにて選択されたファイル
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト

    On Error Resume Next

    '初期化
    bRet = False        '処理戻り値
    strFilePath = ""    'ファイル名をNULL文字で初期化   ' V6.1.0.3 ADD

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMEI_DATA_INPUT_BUTTOM, 0)

   '画面のボタンを押下不可にする
    Call sButtonEnabled(False)

    '取得ファイル名を初期化
    CommonDialog1.FileName = ""
    '初期ディレクトリを設定
    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'フォルダ選択画面デフォルトパス１が存在するか
        '存在するため、デフォルトパス１（H:）を設定
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
    Else
        '存在しないため、デフォルトパス２（C:）を設定
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
    End If
    Set objFso = Nothing
    '拡張子を設定
    CommonDialog1.Filter = "すべてのファイル(*.*)|*.*|"
    'ファイル選択画面を開く
    CommonDialog1.ShowOpen
    '選択したファイル名を取得
    strFilePath = CommonDialog1.FileName

    Call ChDrive("D")

    If strFilePath <> "" Then        'ファイル選択有
        
        'ファイルコピーを行い、駅名データＦＤファイルを作成する
        bRet = fFileCopy(strFilePath)

        'メール送信を行う
        If bRet = True Then 'ファイルコピー＝正常

            '正常終了ダイアログを表示する
            iResponse = fMessageBox(FDEkimei.REBOOT)

            If iResponse = vbOK Then    'ＯＫ押下

                'アプリ起動チェック
                If CheckAppStart(PROC_KANRI) <> 0 Then '監視盤アプリ起動中
                    'メール送信処理
                    bRet = fSendPowerOffReqMail()
                                    
                    'ログ出力依頼
                    'メール送信成功時
                    If bRet = True Then
                        sLogRequest LOG_NORMAL
                    End If

                Else '監視盤アプリ終了中
                    '保守プロセス終了処理
                    psEndHoshuProc
                    'リブート処理
                    dllAPLEndReboot
                End If

            Else                        'キャンセル：後で再起動
                '画面のボタンを押下可能にする
                Call sButtonEnabled(True)
            End If

        End If
         
    End If

    '処理キャンセルまたは、メール送信失敗
    If bRet = False Then
        '画面のボタンを押下可能にする
        Call sButtonEnabled(True)
    End If

End Sub



'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  概要      : 「メニュー画面に戻る」ボタン押下時のイベントプロシージャ
'//  説明      : 駅名データ媒体投入画面を閉じる。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    'エラールーチンを宣言
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMEI_DATA_INPUT_GAMEN_END, 0)
    
    '画面消去
    Unload Me

End Sub


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : fFileCopy
'//  概要      : 駅名データＦＤファイルを作成する。
'//  説明      : ＦＤ入力されたファイルをコピーし、駅名データＦＤファイルを作成する。
'//  ﾊﾟﾗﾒｰﾀ    :ＦＤ駅名データファイルパス
'//            :戻り値  ,コピー処理正常：True　　異常：False
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fFileCopy(strFilePath As String) As Boolean
    
    Dim iRet    As Integer                 '駅名データ有無確認
    Dim iErrorFlag As Integer              'エラーフラグ

    On Error Resume Next

    '初期化
    fFileCopy = False      '戻り値
    iRet = -1
    iErrorFlag = 0

    '駅名データＦＤファイルの有無チェック
    iRet = GetAttr(FD_EKIMEI_FILE)
    
    'エラールーチンを宣言
    On Error GoTo Err_LOG

    'エラーの分類をファイル削除エラーに設定
    iErrorFlag = 1

    'ファイルが存在する場合、駅名データＦＤファイルを削除
    If iRet <> -1 Then
        Kill FD_EKIMEI_FILE
    End If
    
    'エラーの分類をファイル作成エラーに設定
    iErrorFlag = 2
    'ＦＤ駅名データファイルから駅名データＦＤファイルにコピーする
    FileCopy strFilePath, FD_EKIMEI_FILE
    'ファイルコピー正常終了
    fFileCopy = True

    Exit Function

Err_LOG:

    'エラールーチンを宣言
    On Error Resume Next
    
    'ログ出力依頼
    If iErrorFlag = 1 Then
        sLogRequest FILEDELETE_ERROR
    ElseIf iErrorFlag = 2 Then
        sLogRequest FILECOPY_ERROR
    End If
    '異常終了ダイアログを表示
    fMessageBox (FDEkimei.FD_INPUT_ERR)
    '駅名データ削除
    Kill FD_EKIMEI_FILE

End Function
'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : fSendPowerOffReqMail
'//  概要      : 監視盤電源OFF要求メールを送信する。
'//  説明      : 監視盤電源OFF要求メールを作成、送信。
'//  ﾊﾟﾗﾒｰﾀ    :戻り値  ,メール送信正常：True　　異常：False
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSendPowerOffReqMail() As Boolean
    
    Dim lngRet      As Long                 '戻り値
    Dim udtMail     As MAIL_JIKI_UNKAI     'メール送信エリア

    On Error Resume Next

    '初期化
    fSendPowerOffReqMail = False

    'メールデータ作成
    udtMail.mlHeader.dwId = ML_ID_KAN_PW_OFF_REQ        'メールＩＤ：監視電源OFF要求
    udtMail.mlHeader.dwSize = Len(udtMail)              'メールサイズ
    udtMail.mlHeader.dwProid = RHOSHU_ID                '送信元プロセスＩＤ：保守
    udtMail.mlHeader.dwSubArea = 0                      '補助情報：０（固定）
    udtMail.dwData = Ml_SyoriType.ML_DT_REBOOT          '処理種別：リブート

    'メール送信
    lngRet = DssSendMail(MAIL_SLOT_KANRI, Len(udtMail), udtMail.mlHeader)
    
    
    If lngRet = False Then 'メール送信異常
        '「監視電源OFF要求：メール送信異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
    Else
        '「監視電源OFF要求：メール送信正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
        'メール送信正常終了
        fSendPowerOffReqMail = True
   End If

End Function


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : sButtonEnabled
'//  概要      : 画面表示中のボタンの押下可能／不可設定を行う。
'//  説明      : 画面のボタンのEnabledを設定する。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sButtonEnabled(bSet As Boolean)

    On Error Resume Next

    cmdFDInput.Enabled = bSet       '実行ボタン
    cmdReturn.Enabled = bSet        '保守画面へ戻るボタン
    cmdRemove.Enabled = bSet        '保守画面へ戻るボタン

End Sub
'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  概要      : 「メール受信用タイマ」がタイムアップした時のイベントプロシージャ
'//  説明      : メール受信処理を行う。
'//  ﾊﾟﾗﾒｰﾀ    :
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    On Error Resume Next
    
    '汎用メール受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmEkimeiFD.Caption, False
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : fMessageBox
'//  概要      : ダイアログ表示
'//  説明      : ダイアログIDにより、ダイアログを作成し表示する
'//  ﾊﾟﾗﾒｰﾀ    : ダイアログID
'//            : 戻り値  ,押下釦種別
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMessageBox(iMsgID As Integer) As Integer

    Dim strMessage  As String           'ダイアログの文言
    Dim strTitle    As String           'ダイアログのタイトル
    Dim lngOption   As Long             'ダイアログの表示ボタンとアイコン

    fMessageBox = 0

    Select Case iMsgID
        Case FDEkimei.REBOOT         '正常終了
            strMessage = "駅名データ媒体投入が正常に行われました。" & Chr(vbKeyReturn) & _
                         "監視盤を再起動しますか？"
            lngOption = vbOKCancel + vbInformation      '「ＯＫ」「キャンセル」ボタン、「情報」アイコン
            strTitle = "駅名データ媒体投入"

        Case FDEkimei.FD_INPUT_ERR   '異常終了
            strMessage = "駅名データ媒体投入に失敗しました。" & Chr(vbKeyReturn) & _
                         "媒体が異常でないかを確認し、正しい媒体を挿入し、再度実行してください。"
            lngOption = vbOKOnly + vbCritical        '「ＯＫ」ボタン、「警告」アイコン
            strTitle = "駅名データ媒体投入"

        Case Else
    End Select

    If lngOption <> 0 Then
        'メッセージボックスを表示し、戻り値をFunctionの戻り値とする。
        fMessageBox = MsgBox(strMessage, lngOption, strTitle)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  関数名称  : sLogRequest
'//  概要      : ログ出力を依頼する。
'//  説明      : 処理に関するログ出力を依頼する。
'//  ﾊﾟﾗﾒｰﾀ    :  iLog       ,I ,Integer        :ログ出力依頼用ID
'//            :
'//
'//  ORIGINAL  ：(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sLogRequest(iLogID As Integer)

    Dim udtLogParam As LOGPARAM  'ログトレース依頼ﾊﾟﾗﾒｰﾀ
    Dim lRet As Long   'ログトレース依頼関数の戻り値

'    'ログ依頼パラメータの共通部に値をセットする。
    If iLogID = LOG_NORMAL Then
    '駅名データＦＤファイル作成成功の場合、
        '「駅名ﾃﾞｰﾀFDﾌｧｲﾙ作成 正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, "駅名ﾃﾞｰﾀFDﾌｧｲﾙ作成 正常", 0)
    ElseIf iLogID = FILEDELETE_ERROR Then
    '駅名データＦＤファイル削除エラーの場合、
        '「駅名ﾃﾞｰﾀFDﾌｧｲﾙ削除 異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, "駅名ﾃﾞｰﾀFDﾌｧｲﾙ削除 異常", 0)
    ElseIf iLogID = FILECOPY_ERROR Then
    '駅名データＦＤファイルコピーエラーの場合、
        '「駅名ﾃﾞｰﾀFDﾌｧｲﾙ削除 異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "駅名ﾃﾞｰﾀFDﾌｧｲﾙ作成 異常", 0)
    Else
    End If
End Sub

