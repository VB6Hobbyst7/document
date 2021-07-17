VERSION 5.00
Begin VB.Form frmTimeDataSettei 
   BorderStyle     =   0  'なし
   Caption         =   "稼働・メンテデータ収集（次世代自動改札機）"
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
   Begin VB.Timer TmrKakunin 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   8160
   End
   Begin VB.Frame Frame2 
      Caption         =   "データ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   10815
      Begin VB.CommandButton cmdDataClear 
         Caption         =   "時間帯別データクリア"
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
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   10815
      Begin VB.CheckBox ChkSndSet 
         BackColor       =   &H0080FF80&
         Caption         =   "送信"
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
         Left            =   480
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   7
         Top             =   840
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CommandButton cmdKakutei 
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "　　　　　送信設定"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   360
      Top             =   8160
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "    システム設定      画面へ戻る"
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
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "時間帯別データ設定"
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
      TabIndex        =   1
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmTimeDataSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'*
'*    宣言名称   : 時間帯別データ設定（次世代自動改札機）
'*   ﾓｼﾞｭｰﾙ概要  : 時間帯別データ設定画面のフォームモジュール
'*
'*     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'*                 ・フェーズ２対応【Mainte_05_03】
'*     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'*                 2014年度施策 【EG20_KANSI05_01】
'*     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000     'メイルタイマのインターバル値

Private mintMaxIndex As Integer
Private mintID As Integer           'エリアID
Private Type SHUSHU_STATUS
    intStatus As Integer    'ステータス
    strCaption As String    'ボタン文言
    strColor As String      'ボタン色
    IntValue As Integer     '押下状態
End Type
Private mudtBtn_Status() As SHUSHU_STATUS

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 確定ボタンが押下された時のイベントプロシージャ
'     説明      : 送信設定エリアを更新する。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(EG20 V2.1.0.1) 2011-12-08   CODED   BY [TCC] M.Matsumoto
'                               【統-221対応】
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdKakutei_Click()

    Dim Idinf_KansiSettei As IdInfProc
    Dim lngErrCode As Long                  'エラーコード
    Dim lngRet As Long
    Dim bRet As Boolean
    
    On Error Resume Next
    
    '監視盤起動時
    If CheckAppStart(PROC_KANRI) <> 0 Then
        Set Idinf_KansiSettei = New IdInfProc             '監視装置設定エリア
        '参照(自改通信状態)エリア名を設定
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_SEND_SET_ERROR, lngErrCode)
           Exit Sub
        End If
    
        'エリアIDの設定値を更新
        Idinf_KansiSettei.IdLock
        Idinf_KansiSettei.id = mintID
        Idinf_KansiSettei.DataType = ID_TYPE.Flag
'        Call Idinf_KansiSettei.SetIDSVR(CInt(ChkSndSet.Value))     'EG20 V2.1.0.1 DEL 【統-221対応】
        Call Idinf_KansiSettei.SetIDSVR(CInt(ChkSndSet.Tag))        'EG20 V2.1.0.1 ADD 【統-221対応】
    
        If Idinf_KansiSettei.Errsts <> 0 Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_SEND_SET_ERROR, lngErrCode)
        Else
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_SEND_SET_OK, 0)
        End If
        Idinf_KansiSettei.IdFree
    
    '監視盤未起動時
    Else
'        bRet = gspfSetKansiSts(mintID, CInt(ChkSndSet.Value))       'EG20 V2.1.0.1 DEL 【統-221対応】
        bRet = gspfSetKansiSts(mintID, CInt(ChkSndSet.Tag))         'EG20 V2.1.0.1 ADD 【統-221対応】
    End If
    
    cmdDataClear.Enabled = False
    cmdKakutei.Enabled = False
    cmdReturn.Enabled = False
    ChkSndSet.Enabled = False
    
    '確認ボタン押下用タイマを作動させる
    tmrKakunin.Interval = 1000       'ボタン押下用タイマ時間設定
    tmrKakunin.Enabled = True
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 時間帯別データ設定画面がロードされた時のイベントプロシージャ
'     説明      : メイル受信用のタイマ値を設定する。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Load()

    Dim intFileNumber As Integer            'ファイル番号
    Dim strFileName As String               'ファイル名
    Dim strItmNum As String                 '設定項目数
    Dim strTemp As String
    Dim intCount As Integer                 'ループカウンタ
    Dim intStatus As Integer                'エリアID値
    Dim lngErrCode As Long                  'エラーコード
    Dim Idinf_KansiSettei As IdInfProc
    Dim intNum As Integer
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255 As GATE_INFO            '読込み用エリア（255設定用）
    Dim lngSts As Long
    Dim lngHandle As Long
    Dim lngRet As Long
    Dim bRet As Boolean
    
    On Error Resume Next
        
    tmrKakunin.Enabled = False
 
    '「時間帯別データ設定画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_START, 0)
   
    'メイル受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '未使用のファイル番号を取得します。
    intFileNumber = FreeFile

    '設定情報ファイル名を設定する。
    strFileName = TIMEDATA_STATUS_FILE

    '設定情報ファイルをオープンする。
    If strFileName <> "" Then
        Open strFileName For Input As #intFileNumber
    End If

    For intCount = 0 To 2

        '設定情報ファイル名に設定されている釦設定ファイルを読む。
        Input #intFileNumber, strItmNum, strTemp, strTemp, strTemp

        '最大コントロール数を変数に設定する。
        If intCount = 1 Then
            'エリアID
            mintID = CInt(strItmNum)
        ElseIf intCount = 2 Then
            '項目数
            mintMaxIndex = CInt(strItmNum) - 1
        End If
    Next

    ReDim mudtBtn_Status(mintMaxIndex)

    For intCount = 0 To mintMaxIndex
        '設定情報ファイル名に設定されている釦設定ファイルを読む。
        With mudtBtn_Status(intCount)
            Input #intFileNumber, .intStatus, .strCaption, .strColor, .IntValue
        End With
    Next

    Close #intFileNumber

    strFileName = Dir(K_SETTEI_FILE)
    If strFileName = "" Then
       '監視設定ファイルがない場合
       strFileName = SHOKI_K_SETTEI_FILE
    Else
       '監視設定ファイルがある場合
       strFileName = K_SETTEI_FILE
    End If
    
    '監視盤起動時は監視装置設定エリアから送信設定を取得する
    If CheckAppStart(PROC_KANRI) <> 0 Then
        Set Idinf_KansiSettei = New IdInfProc             '監視装置設定エリア
        '参照(自改通信状態)エリア名を設定
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_START, 0)
            Exit Sub
        End If
    
        'エリアIDの設定値を取得
        Idinf_KansiSettei.IdLock
        Idinf_KansiSettei.id = mintID
        Idinf_KansiSettei.IdGet
        intStatus = Idinf_KansiSettei.DataArea(0)
        Idinf_KansiSettei.IdFree
        
        cmdDataClear.Enabled = True
    '監視盤未起動の場合
    Else
    
        '監視設定ファイルをオープン
        lngHandle = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TIMEDATA_GAMEN_START, 0)
            Exit Sub
        End If
        
        '監視設定ファイル読み込み
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)

        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
        lngSts = KansiSerchId(udtAreaR255, CLng(mintID))
        If lngSts >= 0 Then
           'IDが有った場合
           intStatus = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
        Else
          ' 該当ＩＤ無しの場合参照異常
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_START, 0)
            Exit Sub
        End If
        
        cmdDataClear.Enabled = False
    End If
    
    '取得した値をTag値に設定
    ChkSndSet.Tag = CStr(intStatus)
    
    'Tag値と一致する文言、色、押下状態にする
    For intCount = 0 To UBound(mudtBtn_Status)
        If mudtBtn_Status(intCount).intStatus = CInt(ChkSndSet.Tag) Then
            ChkSndSet.Caption = mudtBtn_Status(intCount).strCaption
            ChkSndSet.BackColor = mudtBtn_Status(intCount).strColor
            ChkSndSet.Value = mudtBtn_Status(intCount).IntValue
        End If
    Next intCount
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 時間帯別データ設定画面が 表示された時のイベントプロシージャ
'     説明      : 「メール受信用タイマ」を起動する。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Activate()

    On Error Resume Next
    
    'メール受信用タイマを起動する
    tmrMail.Enabled = True
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 時間帯別データ設定画面が消去された時のイベントプロシージャ
'     説明      : 「メール受信用のタイマ」を破棄する。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Deactivate()

    On Error Resume Next
    
    'メール受信用タイマを止める
    tmrMail.Enabled = False
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 「システム設定画面に戻る」釦がクリックされた時のイベントプロシージャ
'     説明      : 画面を消去する。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdReturn_Click()

    On Error Resume Next
   '「時間帯別データ設定画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_END, 0)
 
    '自画面を消す。
    Unload Me
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'
'  概要     :確認ボタン押下用タイマイベント時処理
'  説明     :確認ボタン押下用タイマイベント発生時の処理を行う。
'            確認ボタン、その他ボタンの色を押下色から元の色に戻す
'  ﾊﾟﾗﾒｰﾀ   :
'
'    ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'    REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub tmrKakunin_Timer()
    
    On Error Resume Next

    '確認ボタン押下用タイマを停止する
    tmrKakunin.Enabled = False                   '確認ボタン押下用タイマ停止
    tmrKakunin.Interval = 0                      '確認ボタン押下用時間初期化
    
    cmdDataClear.Enabled = True
    cmdKakutei.Enabled = True
    cmdReturn.Enabled = True
    ChkSndSet.Enabled = True

End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 「メール受信用タイマ」がタイムアップした時のイベントプロシージャ
'     説明      : メール受信処理を行う。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   REVISED BY [TCC] S.Kuroda
'                 2014年度施策 【EG20_KANSI05_01】
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub tmrMail_Timer()

    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmTimeDataSettei.Caption, False
        pfFormActive (frmTimeDataSettei.hwnd)           ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If

End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 送信設定ボタン押下時のイベントプロシージャ
'     説明      : 送信設定を切り替える。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-16   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(EG20 V2.1.0.1) 2011-12-08   CODED   BY [TCC] M.Matsumoto
'                               【統-221対応】
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub ChkSndSet_Click()

    Dim intCount As Integer

'    ChkSndSet.Tag = CStr(ChkSndSet.Value)   'EG20 V2.1.0.1 DEL 【統-221対応】
    
    'Tag値と一致する文言、色、押下状態にする
    For intCount = 0 To UBound(mudtBtn_Status)
'        If mudtBtn_Status(intCount).intStatus = CInt(ChkSndSet.Tag) Then   'EG20 V2.1.0.1 DEL 【統-221対応】
        If mudtBtn_Status(intCount).IntValue = CInt(ChkSndSet.Value) Then   'EG20 V2.1.0.1 ADD 【統-221対応】
            ChkSndSet.Caption = mudtBtn_Status(intCount).strCaption
            ChkSndSet.BackColor = mudtBtn_Status(intCount).strColor
            ChkSndSet.Value = mudtBtn_Status(intCount).IntValue
            ChkSndSet.Tag = mudtBtn_Status(intCount).intStatus              'EG20 V2.1.0.1 ADD 【統-221対応】
        End If
    Next intCount

End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     概要      : 時間帯別データクリアボタン押下時のイベントプロシージャ
'     説明      : 送信設定を切り替える。
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-16   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdDataClear_Click()

    Dim iResponse As Integer
    
    '「時間帯別データ設定画面：時間帯別データクリア釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_CLEAR_BUTTOM, 0)
   
    '「時間帯別データクリア」ポップアップを表示
    iResponse = MsgBox("時間帯別データをクリアしますがよろしいですか？", _
                        vbOKCancel, "確認")
    
    'ＯＫ釦が押されたら
    If iResponse = vbOK Then
        '時間帯別データクリア中フォームをモーダルウィンドウで表示する。
        frmClearCyu.Show vbModal
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SerchId
'//  機能名称  : ＩＤ検索処理(全タブ専用)
'//  機能概要  : ＩＤ検索を行う。
'//
'//              型        名称        意味
'//  引数      : GATE_INFO udtArea255 [IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : Long　　　         　[OUT]　0以上：正常。-1以下：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function KansiSerchId(udtArea255 As GATE_INFO, lngID As Long) As Long

    Dim lngIndex As Long                '検索用インデックス
    Dim lngMin As Long                  '最小インデックス
    Dim lngMax As Long                  '最大インデックス
    Dim lngChkIndex As Long             '該当インデックス
    Dim lngWorkId   As Long             '標準ＩＤ

    On Error Resume Next
    
    '初期化
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '検索開始
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             'ＩＤ取り出し
        If lngID = lngWorkId Then                                  '同じ？
            lngChkIndex = lngIndex                                  'データ取り出し後、検索終了
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngID < lngID) Then         'データが予備か小さい
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    KansiSerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ChgData
'//  機能名称  : データ変換処理処理
'//  機能概要  : データ変換処理処理を行う。
'//
'//              型        名称        意味
'//  引数      : ID_FMT 　DataArea 　[IN]変換元データ
'//
'//              型        値        意味
'//  戻り値    : String　　　        [OUT]　vbNullstring以外：正常。vbNullString    ：エラー
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function ChgData(DataArea As ID_FMT) As String

    Dim lngloop As Long
    Dim lngWork As Long
    Dim lngErrsts As Long

    On Error GoTo ChgDataErr
    
    lngErrsts = IdInfErr.OK
    
    Select Case DataArea.intType
    Case ID_TYPE.Flag   '状態
        If (DataArea.bytDATA(0) <> 255) Then
            ChgData = str$(DataArea.bytDATA(0))
            
        Else
            ChgData = "-1"                      '値が不定ならー１セット
            
        End If
            
    Case ID_TYPE.Count  '回数
        lngWork = 0                              '初期化
        For lngloop = 3 To 0 Step -1
            lngWork = lngWork * 256 + DataArea.bytDATA(lngloop)
        Next lngloop
                        
        ChgData = str$(lngWork)
    
    Case ID_TYPE.Date_Type, ID_TYPE.time_type '日付、時刻
        ChgData = StrConv(DataArea.bytDATA, vbUnicode)
        
    Case Else
        ChgData = vbNullString
        lngErrsts = IdInfErr.ID_TYPE_MISS
        Exit Function

    End Select
    
    Exit Function
    
ChgDataErr:
        ChgData = vbNullString
        lngErrsts = IdInfErr.PROC_ERR
End Function


