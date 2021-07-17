VERSION 5.00
Begin VB.Form frmJikiUnkaiFD 
   BorderStyle     =   0  'なし
   Caption         =   "磁 気 運 改 デ ー タ Ｆ Ｄ 投 入"
   ClientHeight    =   9000
   ClientLeft      =   2700
   ClientTop       =   2220
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdKirikae 
      Caption         =   "当日切替"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   3315
      Width           =   2175
   End
   Begin VB.CommandButton cmdFDInput 
      Caption         =   "媒体投入"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   3315
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Left            =   600
      Top             =   7440
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "      メニュー       画面へ戻る"
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
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "磁気運賃バージョン管理"
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
Attribute VB_Name = "frmJikiUnkaiFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmJikiUnkaiFD.frm
'//  パッケージ名：バージョン管理(磁気運賃)画面
'//
'//  概要：バージョン管理(磁気運賃)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 ・京王より、バージョン管理(磁気運賃)画面流用。
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.16.0.1) 2009-12-21  REVISED BY [TCC] S.Terao
'//                 不具合対応
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面の初期表示フォルダを指定
'//                「当日切替」釦の表示有無をINIファイル化
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir関数をFileSystemObjectに変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

Dim mstrCopyFile()  As String           '運改ファイル名保存エリア

'メッセージボックス表示用ＩＤ
Private Enum FDUNKAI
    FD_INSERT = 1                       '　１：ＦＤ挿入依頼ＭＳＧ
    REBOOT = 2                          '　２：再起動確認ＭＳＧ
    FD_INSERT_ERR = 11                  '１１：挿入ファイル異常ＭＳＧ
    FD_INPUT_ERR = 12                   '１２：ＦＤ入力結果確認ＭＳＧ
    TODAY_CHANGE = 21                   '２１：磁気運改当日切替処理確認ＭＳＧ
    CHANGE_OK = 22                      '２２：磁気運改当日切替処理結果ＭＳＧ（正常）
    CHANGE_ERR = 31                     '３１：磁気運改当日切替処理結果ＭＳＧ（異常）
End Enum

Private mIntFD      As Integer          'ＦＤ挿入枚数
Private mIntFDTotal As Integer          'ＦＤ総数

Private Const DEFAILT_HYOUJI_UMU = 1    '「当日切替」釦のデフォルト表示     'V1.20.0.1 ADD
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : バージョン管理(磁気運賃)画面(アクティブ時)
'//  機能概要  : メール受信用のタイマ起動
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
'//  機能名称  : バージョン管理(磁気運賃)画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
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

    'メール受信用タイマを止める
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : バージョン管理(磁気運賃)画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                「当日切替」釦の表示有無をINIファイル化
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim lSts As Long             '関数戻り値      'V1.20.0.1 ADD
    On Error Resume Next
    
    '「磁気運賃バージョン管理画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

    'V1.20.0.1 ADD START
    '保守.iniより、「当日切替」釦の表示有無を取得する。
    lSts = GetPrivateProfileInt(KANS_JIKI, _
                                   KANSI_KIRIKAE_UMU, _
                                   DEFAILT_HYOUJI_UMU, _
                                   HOSHU_FILE)
    If lSts = 1 Then
        cmdKirikae.Visible = True
    Else
        cmdKirikae.Visible = False
    End If
    'V1.20.0.1 ADD END

    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFDInput_Click
'//  機能名称  : 「媒体投入」釦押下時処理
'//  機能概要  : 媒体投入を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFDInput_Click()

    Dim iResponse   As Integer          'MsgBoxボタンコード
    Dim bRet        As Boolean          'メール送信判定

    On Error Resume Next
    
    '「磁気運賃バージョン管理画面：媒体投入釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDINPUT_BUTTOM, 0)

    '初期化
    bRet = False        '処理戻り値
    mIntFD = 0          '挿入ファイル枚数カウンタ
    mIntFDTotal = 0     'ファイル総数

    '画面のボタンを押下不可にする
    Call sButtonEnabled(False)

    'V1.16.0.1 ADD START
    '「媒体投入」時ファイル名取得処理
    m_sFILE_NAME1 = ""
    m_sFILE_NAME2 = ""
    m_sFILE_NAME1 = fGetFilName(VERJIKI_SEC1, VERJIKI_SEC1_KEY1)    '先頭ファイル名
    m_sFILE_NAME2 = fGetFilName(VERJIKI_SEC1, VERJIKI_SEC1_KEY2)    '拡張子名
    '取得結果チェック
    If m_sFILE_NAME1 = "" Or m_sFILE_NAME2 = "" Then
       '画面のボタンを押下可能にする
        Call sButtonEnabled(True)
       'メッセージを表示「磁気運賃データ入力結果　異常終了」
       fMessageBox (FDUNKAI.FD_INPUT_ERR)
       Exit Sub
    End If
    'V1.16.0.1 ADD END
    '既に運改データがある場合、破棄する
    sFileDelete

    'メッセージを表示「ＦＤ挿入　依頼」
    iResponse = fMessageBox(FDUNKAI.FD_INSERT)

    If iResponse = vbOK Then        'ＯＫ押下

        '挿入されたＦＤのファイル名をチェック＆ワークフォルダにコピーする
        bRet = fFDFileNameCheck()

        'ファイル名チェック、ファイルコピーが正常終了した時、運改ファイルを作成する
        If bRet = True Then
            '運改ファイルを作成する
            bRet = fFileJoint

            'メール送信を行う
            If bRet = True Then
                'メッセージボックス表示「再起動要求　確認」
                iResponse = fMessageBox(FDUNKAI.REBOOT)

                If iResponse = vbOK Then    'OK押下
                    'メール送信処理
                    bRet = fSendMail(MAIL_SLOT_KANRI, ML_ID_KAN_PW_OFF_REQ)
                Else                        'キャンセル：後で再起動
                    '画面のボタンを押下可能にする
                    Call sButtonEnabled(True)
                End If
            End If
        End If
    End If

    '処理キャンセルまたは、メール送信失敗
    If bRet = False Then
        '運改ファイルを全て削除
        Call sFileDelete

        '画面のボタンを押下可能にする
        Call sButtonEnabled(True)
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdKirikae_Click
'//  機能名称  : 「当日切替」釦押下時処理
'//  機能概要  : 当日切替処理を行う。
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
Private Sub cmdKirikae_Click()

    Dim iResponse   As Integer          'MsgBoxボタンコード
    Dim bSendMail   As Boolean          'メール送信判定
    On Error Resume Next
    
    '「磁気運賃バージョン管理画面：当日切替釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_KIRIKAE_BUTTOM, 0)

    '初期化
    bSendMail = False

    '画面のボタンを押下不可にする
    Call sButtonEnabled(False)

    'メッセージボックス表示「磁気運改当日切替処理　確認」
    iResponse = fMessageBox(FDUNKAI.TODAY_CHANGE)

    If iResponse = vbOK Then   'OK押下
        bSendMail = fSendMail(MAIL_SLOT_KANMA, ML_ID_HOSHU_UNKAI_DAYCHG_REQ)
    End If

    '処理キャンセルまたは、メール送信失敗
    If bSendMail = False Then
        '画面のボタンを押下可能にする
        Call sButtonEnabled(True)
       '「磁気運賃バージョン管理画面：当日切替処理異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JIKI_VERSION_KANRI_KIRIKAE_ERROR, 0)
    Else
       '「磁気運賃バージョン管理画面：当日切替処理正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_KIRIKAE_OK, 0)
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面に戻る」釦押下時処理
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
    '「磁気運賃バージョン管理画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sFileDelete
'//  機能名称  : 運改ファイルを削除する。
'//  機能概要  : 「媒体投入」釦押下時処理：存在する運改ファイルを削除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sFileDelete()

    On Error Resume Next

    'ＦＤ運改ファイル削除
'   Kill PATH_WORK & FILE_NAME1 & "*" & FILE_NAME2         'V1.16.0.1 DEL
    Kill PATH_WORK & m_sFILE_NAME1 & "*" & m_sFILE_NAME2   'V1.16.0.1 ADD
    '「磁気運賃バージョン管理画面：FD運改ファイル削除」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDUNKAI_FILE_DELETE, 0)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fFDFileNameCheck
'//  機能名称  : 運改ファイル名チェック
'//  機能概要  : 「媒体投入」釦押下時処理：投入運改ファイル名チェック
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.16.0.1) 2009-12-21  REVISED BY [TCC] S.Terao
'//                 不具合対応
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 フォルダ選択画面の初期表示フォルダを指定
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir関数をFileSystemObjectに変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fFDFileNameCheck() As Boolean

    Dim strFDFile   As String           'ファイル名
    Dim iResponse   As Integer          'MsgBoxボタンコード
    Dim intFDNum    As Integer          'ＦＤ番号
    Dim bFileCHK    As Boolean          'ファイルチェック
    Dim strWriteDir As String
   
    On Error Resume Next

    '初期化
    fFDFileNameCheck = False    '関数戻り値
    iResponse = 0               'メッセージボックスボタンコード
    intFDNum = 0                '取得するＦＤ番号

'V1.12.0.1 ADD  START
    'フォルダ選択ポップアップ画面表示
    'strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")     'V1.20.0.1 DEL
    strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)      'V1.20.0.1 ADD

    '指定フォルダなし
    If Len(strWriteDir) = 0 Then
        iResponse = vbCancel
    End If
'V1.12.0.1 ADD  END
    '処理キャンセルまたは、ファイル名取得できたときに、ループを抜ける
    Do Until mIntFD = 1 Or iResponse = vbCancel
        'ファイル名取得（１回目のファイル挿入は、必ず１枚目になる）
'        strFDFile = Dir(FDDRIVE & FILE_NAME1 & "*" & FILE_NAME2)   'V1.12.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & FILE_NAME1 & "*" & FILE_NAME2)    'V1.12.0.1 ADD  'V1.16.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & m_sFILE_NAME1 & "*" & m_sFILE_NAME2)    'V1.12.0.1 ADD  'V1.16.0.1 ADD 'V2.6.0.1 DEL
         strFDFile = sDoFileFind(strWriteDir & "\" & m_sFILE_NAME1 & "*" & m_sFILE_NAME2) 'V2.6.0.1 ADD
        'ファイル名取得チェック
        If strFDFile <> "" Then         '取得成功
            'ファイルの総数を取得する
'            mIntFDTotal = CInt(Mid(strFDFile, Len(FILE_NAME1) + 2, 1))  'V1.16.0.1 DEL
             mIntFDTotal = CInt(Mid(strFDFile, Len(m_sFILE_NAME1) + 2, 1))   ''V1.16.0.1 ADD
        End If

        '取得したファイル総数チェック
        If mIntFDTotal > 0 And mIntFDTotal < 10 Then
            mIntFD = 1          '取得成功！　ループカウンタを１にセット
        Else
            'メッセージを表示「ＦＤ挿入　ファイル異常」
            iResponse = fMessageBox(FDUNKAI.FD_INSERT_ERR)

            If iResponse = vbCancel Then
                'メッセージを表示「磁気運賃データ入力結果　異常終了」
                fMessageBox (FDUNKAI.FD_INPUT_ERR)
'            End If 'V1.12.0.1 DEL
                        'V1.12.0.1 ADD  START
            Else
                'フォルダ選択ポップアップ画面表示
                'strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")   'V1.20.0.1 DEL
                strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
            
                '指定フォルダなし
                If Len(strWriteDir) = 0 Then
                    iResponse = vbCancel
                End If
            End If
                        'V1.12.0.1 ADD  END

        End If
    Loop

    'ファイル名保存エリアの再定義
    ReDim mstrCopyFile(mIntFDTotal - 1)

    'ＦＤ番号が正しいかチェックする。処理キャンセルまたは、ファイルを全て取得できたら、ループを抜ける
    Do Until mIntFD > mIntFDTotal Or iResponse = vbCancel
        'ファイル名取得
'        strFDFile = Dir(FDDRIVE & FILE_NAME1 & mIntFD & "*" & FILE_NAME2)      'V1.12.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & FILE_NAME1 & mIntFD & "*" & FILE_NAME2)   'V1.12.0.1 ADD 'V1.16.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & m_sFILE_NAME1 & mIntFD & "*" & m_sFILE_NAME2)   'V1.12.0.1  'V1.16.0.1 ADD 'V2.6.0.1 DEL
         strFDFile = sDoFileFind(strWriteDir & "\" & m_sFILE_NAME1 & mIntFD & "*" & m_sFILE_NAME2) 'V2.6.0.1 ADD

        'ファイル名取得チェック
        'V1.16.0.1 DEL START
        'If Len(strFDFile) = Len(FILE_NAME1 & FILE_NAME2) + 2 Then      '取得成功
        '    intFDNum = CInt(Mid(strFDFile, Len(FILE_NAME1) + 1, 1))     'ＦＤ番号
        'End If
        'V1.16.0.1 DEL END
        'V1.16.0.1 ADD START
        If Len(strFDFile) = Len(m_sFILE_NAME1 & m_sFILE_NAME2) + 2 Then   '取得成功
            intFDNum = CInt(Mid(strFDFile, Len(m_sFILE_NAME1) + 1, 1))     'ＦＤ番号
        End If
        'V1.16.0.1 ADD END

        '取得したファイル番号（intFDNum）と期待するファイル番号（mIntFD）が合っているか？
        If intFDNum = mIntFD Then       'ファイル番号　正常
            'ワークフォルダに、ＦＤファイルをコピーする
'            Call FileCopy(FDDRIVE & strFDFile, PATH_WORK & strFDFile)      'V1.12.0.1 DEL
'            Call FileCopy(strWriteDir & strFDFile, PATH_WORK & strFDFile)   'V1.12.0.1 ADD 'V1.16.0.1 DEL
             Call FileCopy(strWriteDir & "\" & strFDFile, PATH_WORK & strFDFile)   'V1.16.0.1 ADD

            'ワークファイルのファイル名を保存する
            mstrCopyFile(mIntFD - 1) = PATH_WORK & strFDFile

'V1.12.0.1 DEL START
'            'ＦＤ番号が総数より少ない場合、次のＦＤの挿入を促す
'            If mIntFD < mIntFDTotal Then
'                'メッセージを表示「ＦＤ挿入　依頼」
'                iResponse = fMessageBox(FDUNKAI.FD_INSERT)
'
'                If iResponse = vbCancel Then
'                    'メッセージを表示「磁気運賃データ入力結果　異常終了」
'                    fMessageBox (FDUNKAI.FD_INPUT_ERR)
'                End If
'            End If
'V1.12.0.1 DEL END

            'ＦＤ総数をカウントアップ
            mIntFD = mIntFD + 1
        Else                            'ファイル番号　異常
            'メッセージを表示「ＦＤ挿入　ファイル異常」
            iResponse = fMessageBox(FDUNKAI.FD_INSERT_ERR)

            If iResponse = vbCancel Then
                'メッセージを表示「磁気運賃データ入力結果　異常終了」
                fMessageBox (FDUNKAI.FD_INPUT_ERR)
'            End If     'V1.12.0.1 DEL
'V1.13.0.1 ADD  START
            Else
                'フォルダ選択ポップアップ画面表示
                'strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", "")   'V1.20.0.1 DEL
                strWriteDir = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
            
                '指定フォルダなし
                If Len(strWriteDir) = 0 Then
                    iResponse = vbCancel
                    'メッセージを表示「磁気運賃データ入力結果　異常終了」
                    fMessageBox (FDUNKAI.FD_INPUT_ERR)
                End If
            End If
'V1.13.0.1 ADD  END

        End If
    Loop

    '正常終了したとき、関数の戻り値をTrueに設定する。
    If iResponse <> vbCancel Then
        fFDFileNameCheck = True
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fFileJoint
'//  機能名称  : 運改ファイルを作成する。
'//  機能概要  : 「媒体投入」釦押下時処理：
'//              投入ファイルより、運改ファイルを作成する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fFileJoint() As Boolean

    Dim intLoop         As Integer          'ループカウンタ
    Dim intReadFileNo   As Integer          '読込ファイル番号
    Dim intWriteFileNo  As Integer          '書込ファイル番号
    Dim bytReadFile()   As Byte             '読込ファイルエリア
    Dim strFile         As String           '読込ファイル取得

    On Error Resume Next

    '初期化
    fFileJoint = False      '戻り値

    '運改データ削除
    Kill FDUNKAI_FILE

    intWriteFileNo = FreeFile        '書込み専用ファイルの番号を取得する。

    On Error GoTo Err_LOG

    '書込用ファイル(FD_UNKAI.DAT)を書込み専用バイナリモードでオープンする。
    Open FDUNKAI_FILE For Binary As #intWriteFileNo
        For intLoop = 0 To mIntFDTotal - 1      'ＦＤ総数でループする。

            'ファイルが存在することを確認する。
            strFile = Dir(mstrCopyFile(intLoop))
            If strFile = "" Then                'ファイルがなかったら、エラー処理へ
                GoTo Err_LOG
            End If

            intReadFileNo = FreeFile            '読込専用のファイル番号を取得する。

            On Error GoTo Err_LOG

            '読込ファイルエリアの配列の再定義
            ReDim bytReadFile(FileLen(mstrCopyFile(intLoop)) - 1)


            '読込専用ファイルの番号を取得する。
            Open mstrCopyFile(intLoop) For Binary Access Read As intReadFileNo

            '読み込みファイルデータ取得
            Get intReadFileNo, , bytReadFile

            Close #intReadFileNo            '読込ファイルを閉じる。

            '運改データファイルに書込み
            Put #intWriteFileNo, , bytReadFile

        Next
    Close #intWriteFileNo                   '書込みファイルを閉じる。

    fFileJoint = True

    'ＦＤ運改ファイル削除
    'Kill PATH_WORK & FILE_NAME1 & "*" & FILE_NAME2       'V1.16.0.1 DEL
    Kill PATH_WORK & m_sFILE_NAME1 & "*" & m_sFILE_NAME2  'V1.16.0.1 ADD
    '「磁気運賃バージョン管理画面：FD運改ファイル削除」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDUNKAI_FILE_DELETE, 0)
    '「磁気運賃バージョン管理画面：媒体投入処理正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDINPUT_OK, 0)

    Exit Function

Err_LOG:

    'メッセージを表示「磁気運賃データ入力結果　異常終了」
    fMessageBox (FDUNKAI.FD_INPUT_ERR)

    '読込ファイルを閉じる
    If intReadFileNo > 0 Then
        Close #intReadFileNo
    End If

    '書込ファイルを閉じる
    If intWriteFileNo > 0 Then
        Close #intWriteFileNo
    End If

    'ＦＤ運改データ削除
    sFileDelete

    '運改データ削除
    Kill FDUNKAI_FILE
    '「磁気運賃バージョン管理画面：媒体投入処理異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JIKI_VERSION_KANRI_FDINPUT_ERROR, 0)

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sButtonEnabled
'//  機能名称  : 表示画面の釦コントロールを行う。
'//  機能概要  : 「媒体投入」「当日切替」釦押下時処理：釦を押下可/押下不可にする。
'//
'//              型        名称      意味
'//  引数      : Boolean　bSet　　　[IN]釦のコントロール(TRUE：押下可,FALSE：押下不可)
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sButtonEnabled(bSet As Boolean)

    On Error Resume Next

    cmdFDInput.Enabled = bSet       '媒体投入ボタン
    cmdKirikae.Enabled = bSet       '当日切替ボタン
    cmdReturn.Enabled = bSet        'メニュー画面へ戻るボタン

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMessageBox
'//  機能名称  : 画面ポップアップ表示処理
'//  機能概要  : メッセージIDによって表示ポップアップを決定/表示
'//
'//              型        名称      意味
'//  引数      : Integer　 iMsgID   [IN]メッセージID
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMessageBox(iMsgID As Integer) As Integer

    Dim strMessage  As String           'MSGBOXの文言
    Dim strTitle    As String           'MSGBOXのタイトル
    Dim lngOption   As Long             'MSGBOXの表示ボタンとアイコン

    fMessageBox = 0
   
   On Error Resume Next

    Select Case iMsgID
        Case FDUNKAI.FD_INSERT      'ＦＤ挿入　依頼
            strMessage = "磁気運賃を挿入してください。"
            lngOption = vbOKCancel + vbInformation      '「ＯＫ」「キャンセル」ボタン、「情報」アイコン
            strTitle = "媒体挿入"

        Case FDUNKAI.REBOOT         '再起動要求
            strMessage = "監視盤を再起動しますが、よろしいですか？"
            lngOption = vbOKCancel + vbExclamation      '「ＯＫ」「キャンセル」ボタン、「注意」アイコン
            strTitle = "再起動確認"

        Case FDUNKAI.FD_INSERT_ERR  'ＦＤ挿入　ファイル異常
            strMessage = "異常なファイルが挿入されました。" & Chr(vbKeyReturn) & _
                         "正しい磁気運賃を挿入してください。"
            lngOption = vbOKCancel + vbCritical         '「ＯＫ」「キャンセル」ボタン、「警告」アイコン
            strTitle = "媒体挿入"

        Case FDUNKAI.FD_INPUT_ERR   '磁気運賃データ入力結果　異常終了
            strMessage = "運賃データ入力は異常終了しました。"
            lngOption = vbOKOnly + vbInformation        '「ＯＫ」ボタン、「情報」アイコン
            strTitle = "磁気運賃データ入力結果"

        Case FDUNKAI.TODAY_CHANGE   '磁気運改当日切替処理　確認
            strMessage = "磁気運改当日切替処理を行いますが、よろしいですか？"
            lngOption = vbOKCancel + vbExclamation      '「ＯＫ」「キャンセル」ボタン、「注意」アイコン
            strTitle = "磁気運改当日切替処理確認"

        Case FDUNKAI.CHANGE_OK      '磁気運改当日切替処理　正常終了
            strMessage = "磁気運改当日切替処理を正常終了しました。"
            lngOption = vbOKOnly + vbInformation        '「ＯＫ」ボタン、「情報」アイコン
            strTitle = "磁気運改当日切替結果"

        Case FDUNKAI.CHANGE_ERR     '磁気運改当日切替処理　異常終了
            strMessage = "磁気運改当日切替処理を異常終了しました。"
            lngOption = vbOKOnly + vbInformation        '「ＯＫ」ボタン、「情報」アイコン
            strTitle = "磁気運改当日切替結果"

        Case Else
    End Select

    If lngOption <> 0 Then
        'メッセージボックスを表示し、戻り値をFunctionの戻り値とする。
        fMessageBox = MsgBox(strMessage, lngOption, strTitle)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fSendMail
'//  機能名称  : 指定プロセスにメールを送信す処理
'//  機能概要  : 指定メールスロット、メールIDで作成送信を行う。
'//
'//              型        名称      意味
'//  引数      :String　　MailSlot　[IN]メールスロット名
'//             Long      MailID    [IN]送信メールID
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fSendMail(MailSlot As String, MailID As Long) As Boolean
    Dim lngMSlot    As Long             '送信メールスロットハンドル
    Dim lngRet      As Long             '戻り値
    Dim udtMail     As MAIL_JIKI_UNKAI  '磁気運改用メール送信エリア

    On Error Resume Next

    fSendMail = False
    
    '送信メールデータ作成
     udtMail.mlHeader.dwId = MailID              'メールＩＤ：引数
     udtMail.mlHeader.dwSize = Len(udtMail)      'メールサイズ
     udtMail.mlHeader.dwProid = RHOSHU_ID        '送信元プロセスＩＤ：保守
     udtMail.mlHeader.dwSubArea = 0              '補助情報：０（固定）
     Select Case MailID                          'データ部：メールＩＤで内容設定
        Case ML_ID_HOSHU_UNKAI_DAYCHG_REQ        '保守運改当日切替要求
             udtMail.dwData = MlUnkaiJikiIC.ML_DT_UNKAI_JIKI     'データ種：磁気
        Case ML_ID_KAN_PW_OFF_REQ                '監視盤電源ＯＦＦ要求
             udtMail.dwData = Ml_SyoriType.ML_DT_REBOOT          '処理種別：リブート
        Case Else
     End Select

     'メール送信
      lngRet = DssSendMail(MailSlot, Len(udtMail), udtMail.mlHeader)
      If lngRet = False Then
         Select Case MailID                          'データ部：メールＩＤで内容設定
             Case ML_ID_HOSHU_UNKAI_DAYCHG_REQ        '保守運改当日切替要求
                  '「磁気運賃バージョン管理：メール送信異常」ログ出力
                  Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_UNKAIKIRIKAE_CMD_SEND, 0)
             Case ML_ID_KAN_PW_OFF_REQ                '監視盤電源ＯＦＦ要求
                  '「磁気運賃バージョン管理：メール送信異常」ログ出力
                  Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
             Case Else
         End Select
      Else
         Select Case MailID                          'データ部：メールＩＤで内容設定
             Case ML_ID_HOSHU_UNKAI_DAYCHG_REQ        '保守運改当日切替要求
                  '「磁気運賃バージョン管理：メール送信正常」ログ出力
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_UNKAIKIRIKAE_CMD_SEND, 0)
             Case ML_ID_KAN_PW_OFF_REQ                '監視盤電源ＯＦＦ要求
                  '「磁気運賃バージョン管理：メール送信正常」ログ出力
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
             Case Else
         End Select
         'メール送信正常終了
        fSendMail = True
      End If

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ処理
'//  機能概要  : メール受信処理を行う。
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

    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    Dim iResponse As Integer         'MsgBoxボタンコード

    On Error Resume Next

    'メールを受信する。
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
        'メール受信処理を行う
        Select Case udtReadMail.udtlHeader.dwId         'メールＩＤ
            Case ML_ID_HOSHU_ACTIVE_REQ                 '保守アクティブ要求
               '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                AppActivate frmJikiUnkaiFD.Caption, False
            
            Case ML_ID_HOSHU_UNKAI_DAYCHG_INF               '保守運改当日切替通知
               '「保守運改当日切替通知受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_UNKAIKIRIKAE_REQ_RECV, 0)
                'データ種別が“磁気"の時だけ、処理を行う
                If udtReadMail.lngData(0) = MlUnkaiJikiIC.ML_DT_UNKAI_JIKI Then
                    If udtReadMail.lngData(1) = MlUnkaiKekka.ML_DT_UNKAI_NORMAL Then
                        iResponse = fMessageBox(FDUNKAI.CHANGE_OK)      '正常終了メッセージ表示
                    Else
                        iResponse = fMessageBox(FDUNKAI.CHANGE_ERR)     '異常終了メッセージ表示
                    End If
                    '画面のボタンを押下可能にする
                    Call sButtonEnabled(True)
                End If

            Case ML_ID_PROEND_ORD                       'プロセス終了指示の場合
               '「プロセス終了指示受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                '強制終了処理を行う
                pfAbortProc
            Case Else
        End Select
  End If
End Sub
'V1.16.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fGetFilName
'//  機能名称  : 運改ファイル名取得処理
'//  機能概要  : INIファイルより投入ファイル名を取得
'//
'//              型        名称      意味
'//  引数      : String   sSecName  [IN]取得セクション名
'//  　　      : String   sKeyName  [IN]取得キー名
'//
'//              型        値        意味
'//  戻り値    :String　　　　　　　[OUT]取得ファイル名(正常)
'//                                      ブランク(異常)
'//
'//     ORIGINAL :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fGetFilName(sSecName As String, sKeyName As String) As String

    Dim lSts As Long                                 '関数戻り値
    Dim strFileName As String * MAX_PATH_SIZE     '取得ファイル名
    Dim lngErrCode As Long
    
    On Error Resume Next
  
    '磁気運賃.iniより、投入運賃データファイル名を取得する。
    strFileName = ""
    lSts = GetPrivateProfileString(sSecName, _
                                   sKeyName, _
                                   DEFAILT, _
                                   strFileName, _
                                   Len(strFileName), _
                                   JIKIUNCHIN_FILE)
    If lSts > 0 Then
       fGetFilName = Left$(strFileName, lSts)
    Else
      '「バージョン管理画面(磁気運賃)：運改ファイル名取得異常(INIファイル読み込み異常)」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_INI_READ, lngErrCode)
      fGetFilName = ""
    End If
End Function
'V1.16.0.1 ADD END
