VERSION 5.00
Begin VB.Form frmICUnkai_Type1 
   BorderStyle     =   0  'なし
   Caption         =   "運賃データＤＬＬ画面"
   ClientHeight    =   9000
   ClientLeft      =   1380
   ClientTop       =   1905
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   14.25
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
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdVer 
      Caption         =   "ＩＣ         運改データ要求"
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
      Left            =   9600
      TabIndex        =   4
      Top             =   4920
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
      Height          =   3660
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   9135
   End
   Begin VB.Timer tmrMail 
      Left            =   9360
      Top             =   7320
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "磁気         運改データ要求"
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
      Left            =   9600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "メニュー画面へ戻る"
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
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
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
      Height          =   3660
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9135
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "運賃データＤＬＬ画面"
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
      TabIndex        =   5
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmICUnkai_Type1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'*
'*   ﾓｼﾞｭｰﾙ概要  : 運賃データDLL画面のフォームモジュール
'*
'*      ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'*      REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'*                  ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'*      REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Option Explicit
'リソース定数

'選択中リソース種別 =0：集計、=1:交調
Dim iSelResource As Integer

Private Const MN_MAIL_INTERVAL = 1000 'メイルタイマのインターバル値

'ログ出力メッセージ
Private LogMsgStart(1) As String
Private LogMsgMidst(1) As String
Private LogMsgEnd(1) As String

Private gIndex As Integer           '押下された釦のINDEX
Private gSendMailKishu As Long      '送信メール：機種
Private gSendMailShubetsu As Long   '送信メール：種別
Private gSendMailShosai As Long     '送信メール：詳細

Private Enum DLL_DATA       ' DLL対象
    JIKI = 0                ' 磁気：０
    IC                      ' ＩＣ：１
End Enum
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  概要     : 運賃データDLL結果表示
'  説明     : 運賃データDLLの結果を表示する文言を作成する。
'  ﾊﾟﾗﾒｰﾀ   : strMsg, I ,string, ：DLL結果表示文言
'           :  戻り値,O ,string, ：リストボックス表示文言
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Function fMakeListbox(strMsg_1 As String, strMsg_2 As String) As String
    Dim strRet As String

    strRet = vbNullString

    strRet = Format(Now, "YYYY/MM/DD   HH:MM:SS")

    strRet = strRet & Space(3) & strMsg_1 & Space(3) & strMsg_2

    fMakeListbox = strRet

End Function

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  概要     : 「保守画面に戻る」ボタン押下時のイベントプロシージャ
'  説明     : 運賃データDLL結果表示画面を閉じる。
'  ﾊﾟﾗﾒｰﾀ   :
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub cmdReturn_Click()
   
   'V2.2.0.1 ADD START
   '「メニュー釦へ戻る」釦押下「運賃データＤＬＬ画面消去」
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UNCHINDATA_DLL_GAMEN_END, 0)
   'V2.2.0.1 ADD END
   
    '運賃データDLL結果表示画面を閉じる
    Unload Me
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'   概要    : 運改データ要求ボタン押下時のイベントプロシージャ
'   説明    : 運賃データDLL要求を監マに送信する。
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub cmdVer_Click(Index As Integer)
    Dim lngMSlot_KM As Long             '監マのメールスロットハンドル
    Dim lngRet As Long                  '戻り値
    Dim udtMail As MAIL_UNCHIN_DLL_REQ  '送信メール
    Dim iCnt As Integer                 'カウンタ
    Dim iDataSts As Integer             'データステータス　'V2.2.0.1　ADD
    
    '押下された釦のインデックス値を保存
    gIndex = Index

    'DLL開始をリストに表示する。
    lstKan(gIndex).AddItem fMakeListbox("正常", LogMsgStart(gIndex))
    
    ' ボタンを押下不可にする。
    For iCnt = 0 To cmdVer.UBound
        cmdVer(iCnt).Enabled = False
    Next

    ' 保守画面へ戻る釦を押下不可にする。
    cmdReturn.Enabled = False
'V2.2.0.1 DEL START
'    '監マへの送信メールスロットをオープンする。
'    lngMSlot_KM = DssMailOpen(MAIL_SLOT_KANMA)
'    If lngMSlot_KM <> INVALID_HANDLE_VALUE Then   '異常
'
'        gSendMailKishu = ML_DT_UNCHINDLL_KISHU     '送信メール：機種
'        gSendMailShosai = MlUnchinDllSHUBETSU.ML_DT_UNCHIN_NEW   '送信メール：詳細
'        'データ種別
'        If gIndex = DLL_DATA.JIKI Then      '磁気運賃ＤＬＬ
'            gSendMailShubetsu = MlUnchinDllData.ML_DT_UNCHIN_ICHI_FUKU     '１枚用＋複数枚用
'        Else                                'ＩＣ運賃ＤＬＬ
'            gSendMailShubetsu = MlUnchinDllData.ML_DT_UNCHIN_IC_HAN        'ＩＣ運賃＋判定プログラム
'        End If
'
'        '監マに対して運賃DLL要求を送信する。
'        udtMail.mlHeader.dwId = ML_ID_UNTIN_REQ       'メールＩＤ：運賃データＤＬＬ要求（＝９７３）
'        udtMail.mlHeader.dwSize = MlSize.UNTIN_REQ    'メールサイズ：２８
'        udtMail.mlHeader.dwProid = RHOSYU_ID          '送信元ＩＤ：保守（＝１１）
'        udtMail.mlHeader.dwSubArea = 0                '補助情報
'        udtMail.dwKishu = gSendMailKishu                '機種：自動改札機（＝１）
'        udtMail.dwData = gSendMailShubetsu              'データ種別
'        udtMail.dwSyosai = gSendMailShosai              '種別詳細：新運賃
'
'        lngRet = DssMailWrite(lngMSlot_KM, MlSize.UNTIN_REQ, udtMail.mlHeader)
'
'        '監マへの送信メールスロットをクローズする。
'        lngRet = DssMailClose(lngMSlot_KM)
'
'    End If
'V2.2.0.1 DEL END
'V2.2.0.1 ADD START
 psGetData_Type iDataSts
 
 '監視盤起動時メール送信
 gSendMailKishu = ML_DT_UNCHINDLL_KISHU     '送信メール：機種
 gSendMailShosai = MlUnchinDllSHUBETSU.ML_DT_UNCHIN_NEW   '送信メール：詳細
 'データ種別
 If gIndex = DLL_DATA.JIKI Then      '磁気運賃ＤＬＬ
    gSendMailShubetsu = iDataSts                                    '１枚用＋複数枚用
 Else                                'ＩＣ運賃ＤＬＬ
    gSendMailShubetsu = MlUnchinDllData.ML_DT_UNCHIN_IC_HAN        'ＩＣ運賃＋判定プログラム
 End If
        
 '監マに対して運賃DLL要求を送信する。
  udtMail.mlHeader.dwId = ML_ID_UNTIN_REQ       'メールＩＤ：運賃データＤＬＬ要求（＝９７３）
  udtMail.mlHeader.dwSize = MlSize.UNTIN_REQ    'メールサイズ：２８
  udtMail.mlHeader.dwProid = RHOSHU_ID          '送信元ＩＤ：保守（＝１１）
  udtMail.mlHeader.dwSubArea = 0                '補助情報
  udtMail.dwKishu = gSendMailKishu                '機種：自動改札機（＝１）
  udtMail.dwData = gSendMailShubetsu              'データ種別
  udtMail.dwSyosai = gSendMailShosai              '種別詳細：新運賃
    
  'メール送信
  lngRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.UNTIN_REQ, udtMail.mlHeader)
  If lngRet = False Then
     '送信異常
     Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, UNCHINDATA_DLL_CMD_ERROR, 0)
     
     ' ボタンを押下可能にする。
     For iCnt = 0 To cmdVer.UBound
         cmdVer(iCnt).Enabled = True
     Next
            
     ' 保守画面へ戻る釦を押下不可にする。
     cmdReturn.Enabled = True
            
     AppActivate frmICUnkai_Type1.Caption, False
  Else
    '送信正常
     Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, UNCHINDATA_DLL_CMD_OK, 0)
  End If
'V2.2.0.1 ADD END
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  概要      : 運賃データDLL結果表示画面がアクティブになった時のイベントプロシージャ
'  説明      : メイル受信用のタイマを起動する。
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub Form_Activate()
    'メール受信用のタイマを起動する。
    tmrMail.Enabled = True
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  概要     : 運賃データDLL結果表示画面がﾃﾞｨｱｸﾃｨﾌﾞになった時のｲﾍﾞﾝﾄﾌﾟﾛｼｰｼﾞｬ
'  説明     : メール受信用のタイマを止める。
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub Form_Deactivate()
    'メール受信用のタイマを止める。
    tmrMail.Enabled = False
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  概要     : 運賃データDLL結果表示画面面がロードされた時のｲﾍﾞﾝﾄﾌﾟﾛｼｰｼﾞｬ
'  説明     : 初期処理を行う
'  ﾊﾟﾗﾒｰﾀ   :
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub Form_Load()

    Dim iCnt As Integer

    'リスト表示文字列
    LogMsgStart(0) = "磁気運賃データＤＬＬ開始"
    LogMsgStart(1) = "ＩＣ運賃データＤＬＬ開始"

    LogMsgMidst(0) = "磁気運賃データＤＬＬ中"
    LogMsgMidst(1) = "ＩＣ運賃データＤＬＬ中"

    LogMsgEnd(0) = "磁気運賃データＤＬＬ終了"
    LogMsgEnd(1) = "ＩＣ運賃データＤＬＬ終了"
    
    'DLL結果表示用のリストボックスをクリアする。
    For iCnt = 0 To lstKan.UBound
        lstKan(iCnt).Clear
    Next

    'メール受信用のメール受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    'V2.2.0.1 ADD START
    '画面表示「運賃データＤＬＬ画面表示」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UNCHINDATA_DLL_GAMEN_START, 0)
    'V2.2.0.1 ADD END
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  概要     : メール受信用タイマがタイムアップした時のイベントプロシージャ
'  説明     : 受信メールの内容に基づき処理をする。
'  ﾊﾟﾗﾒｰﾀ   :
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub tmrMail_Timer()
    Dim lngLen As Long                      'メイルサイズ
    Dim bRet As Boolean                     '戻り値
    Dim uMail As MAIL_LGMINF_INF            'メイル
    Dim iCnt As Integer                     'カウンタ

    'メール受信
    Do Until fDssMailReadMN(plMSlot_MN, uMail) <= 0
        lngLen = uMail.mlHeader.dwSize    'メイルサイズ値値を設定。
       Select Case uMail.mlHeader.dwId   'メールＩＤ
        '「プロセス終了指示」を受信した場合
        Case ML_ID_PROEND_ORD
            'メール受信ログ出力
            'Call dllWriteMailLog(plMSlot_LG, "プロセス終了指示", uMail, lngLen)  'V2.2.0.1 DEL
            'V2.2.0.1 ADD START
            '「プロセス終了指示受信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            'V2.2.0.1 ADD END
            '強制終了処理を行う
            pfAbortProc

        '「運賃データDLL完了通知」を受信した場合
        Case ML_ID_UNTIN_INF
            'メール受信ログ出力
            '結果内容に基づき処理を行う。
            Dim i As Integer
            For i = 0 To 83
                Debug.Print "byTxtDat(" & i & ") : " & uMail.stMonFmt.byTxtDat(i)
            Next

            '受信メールのデータ確認を行う。機種、データ種別、種別詳細
            If uMail.stMonFmt.byTxtDat(0) <> gSendMailKishu Or _
                uMail.stMonFmt.byTxtDat(4) <> gSendMailShubetsu Or _
                uMail.stMonFmt.byTxtDat(8) <> gSendMailShosai Then

                'メール異常受信ログ出力
                'Call dllErrorMailLog(plMSlot_LG, uMail, lngLen) 'V2.2.0.1 DEL
                'V2.2.0.1 ADD START
                '「メール異常受信」ログ出力
                Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_RECV_ERROR, 0)
                'V2.2.0.1 ADD END
                '処理しないで終了
                Exit Sub
            End If

            'V2.2.0.1 ADD START
            '「運賃データＤＬＬ完了通知」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, UNCHINDATA_DLL_END_REQ, 0)
            'V2.2.0.1 ADD END

            Select Case uMail.stMonFmt.byTxtDat(12)
                Case 0
                    'テスト正常終了を表示する。
                    lstKan(gIndex).AddItem fMakeListbox("正常", LogMsgEnd(gIndex))
                Case 1
                    'テスト異常終了を表示する。
                    lstKan(gIndex).AddItem fMakeListbox("異常", LogMsgEnd(gIndex))
                Case 2
                    'テスト実行不可能を表示する。
                    lstKan(gIndex).AddItem fMakeListbox("異常", LogMsgMidst(gIndex))
            End Select

            ' ボタンを押下可能にする。
            For iCnt = 0 To cmdVer.UBound
                cmdVer(iCnt).Enabled = True
            Next
            
            ' 保守画面へ戻る釦を押下不可にする。
            cmdReturn.Enabled = True
            
            '折返しテスト画面をアクティブにする。
            'AppActivate frmICUnkai.Caption, False 'V2.2.0.1 DEL
            AppActivate frmICUnkai_Type1.Caption, False 'V2.2.0.1 ADD

        '保守画面アクティブ表示の場合
        'Case ML_ID_HOSYU_ACTIVE_REQ 'V2.2.0.1 DEL
        Case ML_ID_HOSHU_ACTIVE_REQ 'V2.2.0.1 ADD
            'V2.2.0.1 DEL START
            'メール受信ログ出力
            'Call dllWriteMailLog(plMSlot_LG, "保守画面アクティブ表示", uMail, lngLen)
            'V2.2.0.1 DEL END
            'V2.2.0.1 ADD STRT
            '「保守画面アクティブ表示要求受信正常」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
            'V2.2.0.1 ADD END
            '折返しテスト画面をアクティブにする。
            'AppActivate frmICUnkai.Caption, False 'V2.2.0.1 END
            AppActivate frmICUnkai_Type1.Caption, False 'V2.2.0.1 ADD

        '(メールＩＤ不正）
        Case Else
            'V2.2.0.1 DEL START
            'メール異常受信ログ出力
            'Call dllErrorMailLog(plMSlot_LG, uMail, lngLen)
            'V2.2.0.1 DEL END
            'V2.2.0.1 ADD START
            '「メールID不正」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
            'V2.2.0.1 ADD END
        End Select
    Loop
End Sub

'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'*
'* 　機能：メールスロットを読み込む（保守プロセス・モニタ受信専用）
'*   引数　　　　      (I/O) ｺﾒﾝﾄ
'*   lngMailHamdle      I    メールスロットハンドル
'*   vReadBuf           O  メール内容を格納するためのエリアを指すポインタ
'*   戻り値
'*   Err                  O  0以上              読込みサイズ
'*                           0                 受信データ無し
'*                           -1                エラー
'*
'*  ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'*   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'*              ＥＧＲメトロ　ＮＥＧ→ＥＧＲコンバート対応
'*  REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Function fDssMailReadMN(lngMailHamdle As Long, udtReadBuf As MAIL_LGMINF_INF) As Long
    Dim lngBool As Long             ' 処理結果
    Dim lngNextMsg As Long          ' 次のメッセージのサイズ
    Dim lngMsg As Long              ' メッセージ数
    Dim lngMailRcvLength As Long    ' 読込みサイズ
    Dim lngTraceSize As Long        ' ログにとるメールサイズ
    Dim lngErrCode As Long  'エラーコード
    
      On Error GoTo MailReadError

    lngMsg = 0
    lngBool = GetMailslotInfo(lngMailHamdle, 0, lngNextMsg, lngMsg, 0)

    ' メールに情報があれば、メール受信
    If (lngNextMsg = -1) Then
        fDssMailReadMN = 0
        Exit Function
    End If
   
    ' メールサイズが受信エリアより大きければ
    If lngNextMsg > LenB(udtReadBuf) Then
       '異常メール受信ログを出力する。
       'Call dllErrorMailLog(plMSlot_LG, udtReadBuf, lngNextMsg) 'V2.2.0.1 DEL
      'V2.2.0.1 ADD START
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MRECEIVE
      Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_RECV_SIZE_ERROR, lngErrCode)
      'V2.2.0.1 ADD END
      'MsgBox "受信できないサイズのメールが送信されたため、保守画面プロセスを異常終了します。" 'V2.2.0.1 DEL
        pfAbortProc
    End If

    On Error Resume Next
    ' メール受信処理
    lngBool = ReadFile(lngMailHamdle, udtReadBuf, lngNextMsg, lngMailRcvLength, 0)
    fDssMailReadMN = lngMailRcvLength

    Exit Function

MailReadError:
    fDssMailReadMN = -1 'INVALID_HANDLE_VALUE
End Function
