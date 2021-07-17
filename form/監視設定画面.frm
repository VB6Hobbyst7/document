VERSION 5.00
Begin VB.Form frmKansiSetteiSub 
   BorderStyle     =   0  'なし
   Caption         =   "リモートメンテナンス"
   ClientHeight    =   9000
   ClientLeft      =   2160
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
   Begin Hoshu.ctlSetteiButton ctlSetteiButton1 
      Height          =   1215
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   2143
   End
   Begin VB.Timer tmrMail 
      Left            =   960
      Top             =   480
   End
   Begin VB.CommandButton cmd_Kakutei 
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
      Height          =   1095
      Left            =   7440
      TabIndex        =   2
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " 　　メニュー 　　  画面へ戻る"
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
      Caption         =   "監視設定"
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
Attribute VB_Name = "frmKansiSetteiSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmKansiSetteiSub.frm
'//  パッケージ名：監視設定画面
'//
'//  概要：監視設定画面
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//                 ・フェーズ３対応　新規追加画面
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
Private Const BUTTOM_COLOR = &H8000000F '釦色(ON時/OFF時：同一色)
Private mstrFileName     As String               'ファイル名
Private mintMaxIndex     As Integer              'Maxインデックス
Private Const KANSI_SETTEI = 1                   '監視設定
Private Const KANSI_STS = 2                      '監視状態
Private Const HUTEI = -1                         '値不定
Private Const DANKI = 0                          '暖気運転処理

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 監視設定画面(アクティブ時)
'//  機能概要  : メール受信用のタイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
   
   'メイル受信用のタイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 監視設定画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    
    On Error Resume Next
       
    'メイル受信用のタイマを止める。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 監視設定画面(ロード時)
'//  機能概要  : 監視設定画面の初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
  
    Dim intCount            As Integer          'カウンター
    Dim strTitle            As String           'タイトル
    Dim intSetteiKubun      As Integer          '監視または自改等の設定区分
    Dim intX                As Integer          'X位置
    Dim intY                As Integer          'Y位置
    Dim intId               As Integer          'ＩＤ
    Dim iShoriNo            As Integer          '処理番号
    Dim strOnMoji           As String           'ON時文字
    Dim strOffMoji          As String           'OFF時文字
    Dim iOnSts              As Integer          'ON時値
    Dim iOffSts             As Integer          'OFF時値
    Dim intBtnUmu           As Integer          '釦表示の有無
    Dim intFileNumber       As Integer          'ファイル番号
    Dim iAreaSts            As Integer          '取得値
    
    On Error Resume Next

    '「監視盤設定画面 表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SETTEI_GAMEN_START, 0)

    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '未使用のファイル番号を取得します。
    intFileNumber = FreeFile
    
    '設定情報ファイル名を設定する。
    mstrFileName = HOSHU_KANSI_SETTEI_FILE
        
    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    
    '設定情報ファイルをオープンする。
    If Len(mstrFileName) <> 0 Then
        Open mstrFileName For Input As #intFileNumber
    End If
    
    '2レコードまで読み込み、設定有効数を取得する。
    For intCount = 0 To 1
        Input #intFileNumber, intBtnUmu, intX, intY, intId, _
                              strTitle, iShoriNo, _
                              strOnMoji, strOffMoji, intSetteiKubun

        '最大コントロール数を変数に設定する。
        If intCount = 1 Then
           mintMaxIndex = intBtnUmu - 1
        End If
    Next
       
    '設定有効数分のみエリア確保。
    ReDim m_Filetest(0 To mintMaxIndex)
    
    '保守_画面専用ファイルより各釦情報を読み込み、各コントロールをLoadする。
    For intCount = 0 To mintMaxIndex
        'ファイルからデータを読む。
        '有効無効、X座標、Y座標、エリアID、タイトル
        '処理番号、ON時文字、OFF時文字、設定フラグを取得
        Input #intFileNumber, intBtnUmu, intX, intY, intId, _
                              strTitle, iShoriNo, _
                              strOnMoji, strOffMoji, intSetteiKubun
        
        '通常エラールーチンに戻る
        On Error Resume Next

       '入／切ボタンコントロールをＬＯＡＤする。
        If intCount > 0 Then
            Load ctlSetteiButton1(intCount)
        End If
        
        '釦の表示を行う場合
        If intBtnUmu = 1 Then
           
           '入／切ボタンコントロールのプロパティに値を設定する。
           '釦表示
            ctlSetteiButton1(intCount).Visible = False
            
            '設定フラグ
            ctlSetteiButton1(intCount).Settei_Flag = intSetteiKubun
        
            '対象エリアより現在値を取得する。
            If intSetteiKubun = KANSI_SETTEI Then
               iAreaSts = pfGetKansiArea_Sts(intId)
            End If
            
            If iAreaSts <> HUTEI Then
               '表示中値
               ctlSetteiButton1(intCount).pSetUp = iAreaSts
               '釦表示
               ctlSetteiButton1(intCount).Visible = True
               '釦タイトル設定
               ctlSetteiButton1(intCount).pButtonTitle = strTitle
               'エリアIDを保持
               ctlSetteiButton1(intCount).pID = intId
               'X座標設定
               ctlSetteiButton1(intCount).Top = ctlSetteiButton1(0).Top + intX
               'Y座標設定
               ctlSetteiButton1(intCount).Left = ctlSetteiButton1(0).Left + intY
               'ON時文字設定
               ctlSetteiButton1(intCount).On_Caption = strOnMoji
               'OFF時文字設定
               ctlSetteiButton1(intCount).Off_Caption = strOffMoji
               '釦背景固定設定
               ctlSetteiButton1(intCount).On_Color = BUTTOM_COLOR
               ctlSetteiButton1(intCount).Off_Color = BUTTOM_COLOR
               ctlSetteiButton1(intCount).SetteiOn_Color = BUTTOM_COLOR
               ctlSetteiButton1(intCount).SetteiOff_Color = BUTTOM_COLOR
               '処理番号
               ctlSetteiButton1(intCount).SHORI_NO = iShoriNo
               '入／切ボタンコントロールの表示処理メソッドを行う。
               '取得現在値チェックを行う。値不定時：画面非表示
               ctlSetteiButton1(intCount).psDisplay
            End If
        End If
    Next
    
    'ファイルをクローズする。
    Close #intFileNumber
      
Exit Sub

'エラー処理
Err_LOG:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next

    '「監視設定画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SETTEI_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmd_Kakutei_Click
'//  機能名称  : 「確定」釦押下
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmd_Kakutei_Click()
    
    On Error Resume Next

    '「確定釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_KENSHU_KAKUTEI_BUTTOM, 0)
    
    '画面をロックする。
    SetEnableFalse
    
    '画面設定反映処理を行う。
    psDispSettei_Hanei

    '画面のロックを解除する。
    SetEnableTrue
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetKansiArea_Sts
'//  機能名称  : 監視設定画面(ロード時)。
'//  機能概要  : 監視設定画面の初期処理を行う。
'//
'//              型        名称          意味
'//  引数      : Integer  intId          [IN]エリアID
'//
'//              型        値       　　 意味
'//  戻り値    : Integer                 [OUT]現在値
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetKansiArea_Sts(intId As Integer) As Integer
    
    Dim iAreaSts     As Integer     '監視設定状態値
    Dim lSts         As Long        '関数戻り値
    Dim udtAreaR255  As GATE_INFO   '読込み用エリア（255設定用）
    Dim lngSts       As Long
    Dim lngLoop1     As Long
    Dim lngHandle    As Long
    Dim FileName     As String
    Dim lngRet       As Long
    Dim bRet         As Boolean
    Dim sSetteiFile  As String      'ファイルパス
    Dim lngAplSts    As Long        'アプリ起動/未起動結果
            
    On Error Resume Next
      
    '監視盤起動有無チェック
    lngAplSts = CheckAppStart(PROC_KANRI)
    If lngAplSts = 0 Then
        '監視盤未起動時
        '監視設定ファイルをオープン
        lngHandle = CreateFile(K_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)  'V1.4.0.1　ADD
        
        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then
           'オープン異常時:異常
           '「監視設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           pfGetKansiArea_Sts = HUTEI
           Exit Function
        End If
        
        '監視設定ファイル読み込み
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           '読み込み異常時：異常
           pfGetKansiArea_Sts = HUTEI
         '「監視設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           'ハンドルのクローズ
           Call CloseHandle(lngHandle)
           Exit Function
        End If
        
        'ハンドルのクローズ
        Call CloseHandle(lngHandle)
        
        'ID検索
        lngSts = KansiSerchId(udtAreaR255, CLng(intId))
        If lngSts >= 0 Then
           'IDが有った場合
           pfGetKansiArea_Sts = udtAreaR255.GateInfo(lngSts).bytDATA(0)
         Else
          ' 該当ＩＤ無しの場合:異常
          '「監視設定画面：エリア・ファイル参照異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          pfGetKansiArea_Sts = HUTEI
          Exit Function
        End If
    Else
        '監視盤起動時
        Set Idinf_KansiSettei = New IdInfProc              '監視装置設定エリア
        '共有エリアオープン
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei    '監視装置設定エリア
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
           pfGetKansiArea_Sts = HUTEI
           '「監視設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
           Exit Function
        End If
        
        '監視設定エリアをＬＯＣＫする。
        Idinf_KansiSettei.IdLock
        If Idinf_KansiSettei.Errsts <> 0 Then
          'データ参照異常時:異常
          pfGetKansiArea_Sts = HUTEI
          Idinf_KansiSettei.IdFree
          '「監視設定画面：エリア・ファイル参照異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
          Exit Function
        End If
    
        '監視設定エリアIDを設定
        Idinf_KansiSettei.id = intId
        Idinf_KansiSettei.IdGet
        If Idinf_KansiSettei.Errsts <> 0 Then
          'データ参照異常時はブランク表示設定を行う。
          pfGetKansiArea_Sts = HUTEI
          Idinf_KansiSettei.IdFree
          '「監視設定画面：エリア・ファイル参照異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
          Exit Function
        End If

        pfGetKansiArea_Sts = Idinf_KansiSettei.DataArea(0)   '設定内容
      
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
   End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psDispSettei_Hanei
'//  機能名称  : 押下釦の設定(状態)を反映する。
'//  機能概要  : 押下釦状態を値を対象ファイルに反映する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psDispSettei_Hanei()

  Dim iKansiId As Long    'エリアID
  Dim iSetSts As Integer  '更新値
  Dim iSetFlag As Integer '設定フラグ
  Dim lngAplSts As Long   '監視盤アプリ起動状態
  Dim iCnt As Integer     'カウンター
  Dim bRet As Boolean     '反映処理戻り値
  Dim iRet As Integer     'メッセージボックス戻り値
  Dim iSettei_Flag As Boolean '設定変更フラグ
  
  On Error Resume Next
   
  iSettei_Flag = False
  
  '監視盤アプリ起動チェックを行う。
  lngAplSts = CheckAppStart(PROC_KANRI)
  If lngAplSts <> 0 Then
       
     '監視盤起動時:自改設定エリア更新処理を行う
      For iCnt = 0 To mintMaxIndex
          'エリアID取得
          iKansiId = ctlSetteiButton1(iCnt).pID
          '更新値を取得
          iSetSts = ctlSetteiButton1(iCnt).pSetUp
          '設定フラグを取得
          iSetFlag = ctlSetteiButton1(iCnt).Settei_Flag
          
          If iSetFlag = KANSI_SETTEI Then
             bRet = Area_Updata(iKansiId, iSetSts)
          End If
          If bRet = True Then
             iSettei_Flag = True
             'メール送信処理を行う。
             psSendMail (iCnt)
          Else
             '更新処理異常時：処理結果(異常終了)ポップアップ画面表示
             iRet = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理結果")
             Exit Sub
          End If
          iSettei_Flag = False
       Next
    Else
       '監視盤未起動時：自改設定ファイルより値取得
        For iCnt = 0 To mintMaxIndex
            'エリアID取得
            iKansiId = ctlSetteiButton1(iCnt).pID
            '更新値を取得
            iSetSts = ctlSetteiButton1(iCnt).pSetUp
            '設定フラグを取得
            iSetFlag = ctlSetteiButton1(iCnt).Settei_Flag
          
            If iSetFlag = KANSI_SETTEI Then
                bRet = Settei_Updata(iKansiId, iSetSts)
            End If
         Next
           
         If bRet = False Then
            '更新処理異常時：処理結果(異常終了)ポップアップ画面表示
            iRet = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理結果")
            Exit Sub
         End If
    End If
    
    iSettei_Flag = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Settei_Updata
'//  機能名称  : 監視設定ファイル更新処理
'//  機能概要  : 監視設定ファイル更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Long　　 iKansiId　[IN]監視設定ID
'//              Integer　iSetSts   [OUT]取得値
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]処理結果
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-15   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function Settei_Updata(iKansiId As Long, iSetSts As Integer) As Boolean

    Dim iAreaSts As Integer       '監視設定状態値
    Dim lSts            As Long   '関数戻り値
    Dim udtAreaR255 As GATE_INFO  '読込み用エリア（255設定用）
    Dim lngSts As Long
    Dim lngLoop1 As Long
    Dim lngHandle As Long
    Dim FileName As String
    Dim lngRet As Long
    Dim bRet As Boolean
    Dim sSetteiFile As String

    On Error Resume Next

    '監視設定ファイルをオープン
    lngHandle = CreateFile(K_SETTEI_FILE, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       '「監視設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Settei_Updata = False
       Exit Function
    End If

   '監視設定ファイル読み込み
    bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       '「監視設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Call CloseHandle(lngHandle)
       Settei_Updata = False
       Exit Function
    End If

   'ハンドルのクローズ
    Call CloseHandle(lngHandle)

    'ID検索
     lngSts = KansiSerchId(udtAreaR255, iKansiId)
     If lngSts >= 0 Then
        'IDが有った場合
        udtAreaR255.GateInfo(lngSts).bytDATA(0) = iSetSts
     Else
        ' 該当ＩＤ無しの場合参照異常
        '「監視設定画面：エリア・ファイル参照異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
        Settei_Updata = False
        Exit Function
     End If

    '監視設定ファイルをオープン
    lngHandle = CreateFile(K_SETTEI_FILE, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
        '「監視設定画面：エリア・ファイル参照異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
    End If

    '監視設定ファイル書込み
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       '「監視設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       'ハンドルのクローズ
       Call CloseHandle(lngHandle)
       Exit Function
    End If

   'ハンドルのクローズ
     Call CloseHandle(lngHandle)

     Settei_Updata = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Area_Updata
'//  機能名称  : 監視設定エリア更新処理
'//  機能概要  : 監視設定エリア更新処理を行う。
'//
'//              型        名称      意味
'//  引数      : Long　　 iKansiId　[IN]監視設定ID
'//              Integer　iSetSts   [OUT]取得値
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]処理結果
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-15   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function Area_Updata(iId As Long, iSts As Integer) As Boolean
    
    On Error Resume Next

    Set Idinf_KansiSettei = New IdInfProc              '監視設定エリア
    '監視設定エリアをオープンする。
    Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
    Idinf_KansiSettei.IdOpen
    If Idinf_KansiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       '「監視設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
       Exit Function
    End If
             
    '監視設定エリアをＬＯＣＫする。
    Idinf_KansiSettei.IdLock
    If Idinf_KansiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       '「監視設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
       Exit Function
    End If
              
    'エリアの内容を読み込む。
    Idinf_KansiSettei.id = iId
    Idinf_KansiSettei.IdGet
    If Idinf_KansiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       '「監視設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
       Exit Function
    End If
               
    '設定内容を取得
    Idinf_KansiSettei.SetIDSVR iSts
    If Idinf_KansiSettei.Errsts <> 0 Then
       'データ参照異常時は異常を返す。
       '「監視設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
       Exit Function
     End If
     
     Idinf_KansiSettei.IdFree
     Set Idinf_KansiSettei = Nothing               '監視装置設定データファイル
    
     Area_Updata = True
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : KansiSerchId
'//  機能名称  : ＩＤ検索処理
'//  機能概要  : ＩＤ検索を行う。
'//
'//              型        名称        意味
'//  引数      : GATE_INFO udtArea255 [IN]変換元データ
'//　　　　　　　Long　　　lngId　　　[IN]エリアID
'//
'//              型        値        意味
'//  戻り値    : Long　　　         　[OUT]　0以上：正常。-1以下：エラー
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-15   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function KansiSerchId(udtArea255 As GATE_INFO, lngId As Long) As Long

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
        If lngId = lngWorkId Then                                  '同じ？
            lngChkIndex = lngIndex                                  'データ取り出し後、検索終了
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngId < lngId) Then         'データが予備か小さい
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
'//  関数名称  : SetEnableFalse
'//  機能名称  : 画面ロック解除処理
'//  機能概要  : 画面のロックを解除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    Dim intCount As Integer 'カウンター
    
    On Error Resume Next
    
    ' ボタンの入力不可とする
    For intCount = 0 To mintMaxIndex
         ctlSetteiButton1(intCount).Enabled = False
    Next
    
    '確定釦：True(ロック)する。
    cmd_Kakutei.Enabled = False
    
    'メニュー画面へ戻る釦：False(ロック)する。
    cmdReturn.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetEnableTrue
'//  機能名称  : 画面ロック解除処理
'//  機能概要  : 画面のロックを解除する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    Dim intCount As Integer 'カウンター
    
    On Error Resume Next
    
    ' ボタンの入力不可とする
    For intCount = 0 To mintMaxIndex
         ctlSetteiButton1(intCount).Enabled = True
    Next
    
    '確定釦：True(ロック解除)する。
    cmd_Kakutei.Enabled = True
    
    'メニュー画面へ戻る釦：True(ロック解除)する。
    cmdReturn.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psSendMail
'//  機能名称  : 送信メール分岐
'//  機能概要  : 処理番号により送信メールを区別する。
'//
'//              型        名称      　意味
'//  引数      : iCnt　　　カウンター  [IN]送信対象カウンター
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub psSendMail(iCnt As Integer)
    
    On Error Resume Next
    '処理番号分岐
    If ctlSetteiButton1(iCnt).SHORI_NO = DANKI Then
       '暖気運転メール送信
       psSendDankiMail
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psSendDankiMail
'//  機能名称  : 暖機運転設定変更通知送信処理
'//  機能概要  : 暖機運転設定変更通知送信処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Public Sub psSendDankiMail()
   Dim udtMail     As MAIL_KANSI_SET_INF    '自改設定指示メール送信エリア
   Dim intCnt      As Integer              'カウンタ
   Dim bRet        As Boolean
   
   On Error Resume Next

    '共通ヘッダ編集
    udtMail.mlHeader.dwId = ML_ID_KANSI_SETTEI_INF
    udtMail.mlHeader.dwSize = MlSize.KANSI_SETTEI_INF
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequestType = MailKANSI_SET_Type.ML_DT_DANKI_UNTEN
    
    'メール送信
    bRet = DssSendMail(MAIL_SLOT_SD, MlSize.KANSI_SETTEI_INF, udtMail.mlHeader)
    If bRet = True Then
       '「監視設定画面：監視設定変更通知送信正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_MAIL, KANSI_SETTEI_DANKIMAIL_SEND_OK, 0)
    Else
       '「監視設定画面：監視設定変更通知送信異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_MAIL, KANSI_SETTEI_DANKIMAIL_SEND_ERROR, 0)
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : タイムアップ時処理
'//  機能概要  : メール受信タイムアップ時処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  
    On Error Resume Next
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmKansiSetteiSub.Caption, False
    End If

End Sub
