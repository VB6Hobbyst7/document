VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOverConectSts 
   BorderStyle     =   0  'なし
   Caption         =   "上位通信状態確認"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "表示更新"
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
      Left            =   240
      TabIndex        =   4
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   7560
      Top             =   8040
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10425
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   $"上位通信状態確認画面.frx":0000
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
      Left            =   9120
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   6355
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11218
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      RowHeightMin    =   50
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      MergeCells      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "上位通信状態確認"
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
Attribute VB_Name = "frmOverConectSts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmOverConectSts.frm
'//  パッケージ名：上位通信状態確認画面
'//
'//  概要：上位通信状態確認画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'///////////////////////////////////////////////////////////////////
'ＩＮＩファイル情報格納エリア
'///////////////////////////////////////////////////////////////////
Private iConectSts As Integer          '外部機器通信状態値
Private iSts_Naiyou As Integer         '外部機器異常状態値
Private iSts_Type As Integer           '外部機器異常種別値
Private Const CONECTSTS_NORMAL = 1     '通信正常
Private Const CONECTSTS_SOKET = 1      'ソケットレベル異常
Private Const CONECTSTS_TCP = 2        'TCPレベル異常
Private Const CONECTSTS_APL = 3        'アプリケーションレベル異常
Private Const CONECTSTS_GETERR = 4     '状態取得異常
Private udtAreaR255 As GATE_INFO                                    '読込み用エリア（255設定用）

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値

' EG20 V3.4.0.1【接続機器見直し対応】追加開始
' 上位機器設定構成
Private Type TRANSKIKI_INFO
    bStatus As Boolean              ' 設定有無（TRUE:有り,FALSE:無し）
    sGetInf As String               ' 画面表示用名称
    iAreaID As Integer              ' 対象外部機器上位機器通信状態エリアID
    nIniListNo As Integer           ' 外部機器リスト番号
    nCorner As Integer              ' コーナ番号
    nProcType As Integer            ' 処理タイプ
    iErrorInfoID As Integer         ' 通信異常状態エリアID
    iErrorTypeID As Integer         ' 通信異常種別エリアID
End Type
Private gTransKikiInfo(1 To CONECT_KIKI_INI_MAX) As TRANSKIKI_INFO

Private Const PROCTYPE_NORMAL = 0   ' 通常処理（参照エリアが上位機器通信状態エリア）
Private Const PROCTYPE_ENKAKU = 1   ' 遠隔処理（参照エリアが遠隔タイプ）
' EG20 V3.4.0.1【接続機器見直し対応】追加終了


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 上位通信状態確認画面(アクティブ時)
'//  機能概要  : 画面の最前面表示処理を行う。
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
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    'V1.3.0.1 ADD START
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
    'V1.3.0.1 ADD END
End Sub

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 上位通信状態確認画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
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
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 上位通信状態確認画面(ロード時)
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
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer
    Dim ii As Integer
    Dim iWide As Integer
    
    On Error Resume Next

   '「上位通信状態確認画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, OVER_CONECT_STS_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
      
    'V2.3.0.1 ADD START
    'IDU縮退チェック
    psIDUCheck
    'V2.3.0.1 ADD END

    '上位通信状態表示処理
    psConectSts
   
   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdCancel_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下時処理
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
Private Sub cmdCancel_Click()
   On Error Resume Next
      
   '「上位通信状態確認画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, OVER_CONECT_STS_GAMEN_END, 0)
    frmOverConectSts.ZOrder
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Command1_Click
'//  機能名称  : 「表示更新」釦押下時処理
'//  機能概要  : 上位通信状態表示処理を呼び、表示更新を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.0.1.1) 2011-11-21  CODED  BY [TCC] T.Koyama
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Command1_Click()

    On Error Resume Next

   '「上位通信状態確認画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, OVER_CONECT_STS_GAMEN_START, 0)
    
    'IDU縮退チェック
    psIDUCheck

    '上位通信状態表示処理
    psConectSts

End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psConectSts
'//  機能名称  : 上位通信状態を表示する。
'//  機能概要  : 対象上位機器の通信状態の取得表示を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R監視盤　八丁畷対応
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21   REVISED BY [TCC] T.Koyama
'//                 ＥＧ２０フェーズ２対応【残件54】
'//                 ・状態表示部へのスクロールバー追加および
'//                   表示更新釦押下時のセル制御追加
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psConectSts()
  Dim iCnt As Integer                     'INIファイル読み込みカウンター
  Dim sKey As String                      'キー名
  Dim sGetInf As String * OVERCONECT_SIZE '取得情報(表示名称)
  Dim lSts As Long                        'INI取得処理戻り値
  Dim i As Integer                        'グリッドの高さカウンター
  Dim iErrCnt   As Integer                '前詰め用カウンター
  Dim sErr_TCP As String                  'TCPレベル異常文言
  Dim sErrCode As String                  'エラーコード
  Dim iAreaID As Integer                  '取得情報(エリアID) 'V2.3.0.1 ADD
  Dim iAddRow As Integer                  ' 登録行数            ' EG20 V3.4.0.1追加
  Dim szResultName As String              ' 出力名称            ' EG20 V3.4.0.1追加
   
  On Error Resume Next
   
' EG20 V3.4.0.1追加開始
  '号機情報取得
  Call gsGetGateInfo
  ' コーナ名称設定処理
  Call gsGetCornerName
' EG20 V3.4.0.1追加終了
   
  'グリッドの変更
  With GridIni
        'グリッドの初期化
        .Clear

        'グリッドのセル数の変更
'        .Rows = 11                     ' EG20 V3.4.0.1削除
        iAddRow = 1                     ' EG20 V3.4.0.1追加
        .Rows = iAddRow                 ' EG20 V3.4.0.1追加
        .Cols = 4

        '設定値のタイトルセット
        .Row = 0
        .Col = 1: .Text = "上位機器"
        .CellAlignment = flexAlignCenterCenter
           
        .Col = 2: .Text = "通信状態"
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3: .Text = "詳細"
        .CellAlignment = flexAlignCenterCenter
        
        '外部機器名称を表示
        For iCnt = 1 To CONECT_KIKI_INI_MAX

            gTransKikiInfo(iCnt).bStatus = False               ' 設定有無（TRUE:有り,FALSE:無し）
            gTransKikiInfo(iCnt).sGetInf = ""                  ' 画面表示用名称
            gTransKikiInfo(iCnt).iAreaID = 0                   ' 対象外部機器上位機器通信状態エリアID
            gTransKikiInfo(iCnt).nIniListNo = 0                ' 外部機器リスト番号
            gTransKikiInfo(iCnt).nCorner = 0                   ' コーナ番号
            gTransKikiInfo(iCnt).nProcType = 0                 ' 処理タイプ
            gTransKikiInfo(iCnt).iErrorInfoID = 0              ' 通信異常状態エリアID
            gTransKikiInfo(iCnt).iErrorTypeID = 0              ' 通信異常種別エリアID

         'V2.3.0.1 ADD START
         ' OUTKIKI_LIST.iniから上位通信エリアIDを取得する。
         sKey = ""
         sKey = PROFILE_KEY_KIKIAREA_NAME & Format(iCnt, "00")
         iAreaID = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                        sKey, _
                                        DEFAILT_Int, _
                                        OUTKIKI_LIST_FILE)

         'IDU設置無しかつ現在の上位機器通信状態エリアがIDサーバで無い場合
         'または、IDU設置有りの場合は、以降の表示処理を行う。
         If (pbIDUSts = 1 And iAreaID <> IdKikiComSts.ID_SERVER_COM) Or _
            (pbIDUSts = 0) Then
         'V2.3.0.1 ADD END

           ' OUTKIKI_LIST.iniから表示用外部機器名称を取得する。
           sKey = PROFILE_KEY_KIKINAME & Format(iCnt, "00")
           lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                          sKey, _
                                          DEFAILT, _
                                          sGetInf, _
                                          Len(sGetInf), _
                                          OUTKIKI_LIST_FILE)
' EG20 V3.4.0.1追加開始
           If lSts <> False Then
                ' 出力名称取得処理
                Call psAddKikiCornerName(sGetInf, iAreaID, iCnt)
                If gTransKikiInfo(iCnt).bStatus = False Then
                    lSts = False
                End If
           End If
' EG20 V3.4.0.1追加終了
           If lSts = False Then
             'INI設定無しの場合、何もしない
           Else
             iAddRow = iAddRow + 1
             .Rows = iAddRow                                    ' EG20 V3.4.0.1追加
             iErrCnt = iErrCnt + 1
             .Row = iErrCnt
'             .Col = 1: .Text = sGetInf                         ' EG20 V3.4.0.1削除
             .Col = 1: .Text = gTransKikiInfo(iCnt).sGetInf     ' EG20 V3.4.0.1追加
        
' EG20 V3.4.0.1削除開始
'             '各外部機器通信状態取得処理を行う。
'             pfGetConectSts iCnt
' EG20 V3.4.0.1削除終了
' EG20 V3.4.0.1追加開始
            If gTransKikiInfo(iCnt).nProcType = PROCTYPE_NORMAL Then
                '上位機器通信状態取得処理を行う。
                pfGetConectSts iCnt
            Else
                '上位機器通信状態取得処理を行う。
                pfGetConectStsJikai iCnt
            End If
' EG20 V3.4.0.1追加終了
             
             '通信状態ステータス参照
             Select Case iConectSts
               Case CONECTSTS_NORMAL
                  '通信状態：正常
                  .Col = 2: .Text = "正常"
                  .CellAlignment = flexAlignCenterCenter
               Case CONECTSTS_GETERR
                  '通信状態：取得異常
                  .Col = 2: .Text = ""
                  .CellAlignment = flexAlignCenterCenter
               Case Else
                  '上記以外：ソケットレベル異常,TCPレベル異常,アプリレベル異常
                  .Col = 2: .Text = "異常"
                  .CellAlignment = flexAlignCenterCenter
             End Select
           
             '各外部機器詳細表示処理を行う。
             '通信状態ステータス参照
             Select Case iSts_Naiyou
              Case CONECTSTS_SOKET
                 'ソケットレベル異常
                 .Col = 3: .Text = "ソケットがつながらない"
                 .CellAlignment = flexAlignCenterCenter
              Case CONECTSTS_TCP
                 'TCPレベル異常
                 '16進数に変換。
                 sErrCode = Hex(iSts_Type)
                 sErrCode = sErrCode & "h"
                 sErr_TCP = "TCPレベルでつながらない(エラーコード:" & sErrCode & ")"
                 .Col = 3: .Text = sErr_TCP
                 .CellAlignment = flexAlignCenterCenter
              Case CONECTSTS_APL
                 'アプリケーションレベル異常
                 .Col = 3: .Text = "アプリケーションレベルでつながらない"
                 .CellAlignment = flexAlignCenterCenter
              Case Else
                 '通信状態正常/通信状態取得異常時
                 .Col = 3: .Text = ""
                 .CellAlignment = flexAlignCenterCenter
             End Select
           End If
         End If 'V2.3.0.1 ADD
        Next
   
        'グリッドの幅変更
        .ColWidth(0) = 0
' EG20 V3.4.0.1削除開始
'        .ColWidth(1) = 2500
'        .ColWidth(2) = 1500
'' EG20 V2.0.1.1 DEL START
''        .ColWidth(3) = 8000
'' EG20 V2.0.1.1 DEL END
'' EG20 V2.0.1.1 ADD START
'        .ColWidth(3) = 7775
'' EG20 V2.0.1.1 ADD END
' EG20 V3.4.0.1削除終了
' EG20 V3.4.0.1追加開始
        .ColWidth(1) = 3000
        .ColWidth(2) = 1200
        .ColWidth(3) = 7575
        For i = iAddRow To 10
            .Rows = i + 1
        Next
' EG20 V3.4.0.1追加終了
        
        For i = 0 To CONECT_KIKI_INI_MAX
        '1グリッドの高さ設定
         .RowHeight(i) = 570
        Next
         
' EG20 V2.0.1.1 ADD START
        .TopRow = 1
' EG20 V2.0.1.1 ADD END
    
    End With
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetConectSts
'//  機能名称  : 上位機器通信状態エリアより状態取得処理
'//  機能概要  : 上位機器通信状態エリアより
'//              通信状態、異常状態、異常種別の取得を行う。
'//
'//              型        名称      意味
'//  引数      : iCnt　　Integer　　[IN]取得カウンター
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetConectSts(iCnt As Integer)
    Dim iAreaID As Integer  'エリアID
    Dim sKey As String      'キー名
    Dim strMutexName    As String               'ミューテックス名
    Dim lngMuHandle     As Long                 '排他処理用ハンドル
    Dim udtMapInf       As MAP_MEM              'メモリマッピングオブジェクト
    Dim GetSts          As Long
         
    On Error Resume Next
   
    strMutexName = "Mu_" & GOverComSts
    lngMuHandle = dllOpenMutex(strMutexName)         '排他処理(OPEN)
    If lngMuHandle = 0 Then
        'データ参照異常時は状態ステータスに取得異常を設定。
        iConectSts = CONECTSTS_GETERR
        iSts_Naiyou = CONECTSTS_GETERR
        Exit Function
    End If
    
    dllCloseHandle (lngMuHandle)                 '排他処理(CLOSE)
    
    Set Idinf_Jyoui = New IdInfProc                    '上位通信状態エリア

    Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui                '上位通信状態エリア
    Idinf_Jyoui.IdOpen
    If Idinf_Jyoui.Errsts <> 0 Then
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Exit Function
    End If
    
    ' OUTKIKI_LIST.iniからエリアIDを取得する。
    sKey = PROFILE_KEY_KIKIAREA_NAME & Format(iCnt, "00")
    iAreaID = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                   sKey, _
                                   DEFAILT_Int, _
                                   OUTKIKI_LIST_FILE)
    If iAreaID = 0 Then
      '取得異常の場合次の読み込みへ。
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_Jyoui.IdFree
       Exit Function
    Else
    
        '参照(上位機器通信状態)エリア名を設定
         Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui
         Idinf_Jyoui.IdOpen
         If Idinf_Jyoui.Errsts <> 0 Then
           'データ参照異常時は状態ステータスに取得異常を設定。
           iConectSts = CONECTSTS_GETERR
           iSts_Naiyou = CONECTSTS_GETERR
           Exit Function
         End If
         
         '参照(上位機器通信状態)エリアをＬＯＣＫする。
         Idinf_Jyoui.IdLock
         If Idinf_Jyoui.Errsts <> 0 Then
           'データ参照異常時は状態ステータスに取得異常を設定。
           iConectSts = CONECTSTS_GETERR
           iSts_Naiyou = CONECTSTS_GETERR
           Idinf_Jyoui.IdFree
           Exit Function
         End If
                    
         'エリアの内容を読み込む。
         Idinf_Jyoui.id = iAreaID
           
         '通信状態を取得
         Idinf_Jyoui.GetInf (CONECT)
         If Idinf_Jyoui.Errsts <> 0 Then
            'データ参照異常時は状態ステータスに取得異常を設定。
            iConectSts = CONECTSTS_GETERR
            iSts_Naiyou = CONECTSTS_GETERR
            Idinf_Jyoui.IdFree
            Exit Function
         End If
         iConectSts = CInt(Idinf_Jyoui.DataArea(0))
           
         '異常状態を取得
         Idinf_Jyoui.GetInf (STS)
         If Idinf_Jyoui.Errsts <> 0 Then
            'データ参照異常時は状態ステータスに取得異常を設定。
            iConectSts = CONECTSTS_GETERR
            iSts_Naiyou = CONECTSTS_GETERR
            Idinf_Jyoui.IdFree
            Exit Function
         End If
         iSts_Naiyou = CInt(Idinf_Jyoui.DataArea(0))
           
         '異常種別を取得
         Idinf_Jyoui.GetInf (ERR_TYPE)
         If Idinf_Jyoui.Errsts <> 0 Then
            'データ参照異常時は状態ステータスに取得異常を設定。
             iConectSts = CONECTSTS_GETERR
             iSts_Naiyou = CONECTSTS_GETERR
             Idinf_Jyoui.IdFree
             Exit Function
          End If
           iSts_Type = CInt(Idinf_Jyoui.DataArea(0))
          
    End If
     
    Idinf_Jyoui.IdFree
    
    Set Idinf_Jyoui = Nothing                     '上位通信状態エリア
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
        AppActivate frmOverConectSts.Caption, False
        pfFormActive (frmOverConectSts.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : psAddKikiCornerName
'//  機能名称  : 上位機器コーナ名称追加処理
'//  機能概要  : 上位機器名称に対してコーナ名称を付加する必要があれば追加する。
'//
'//              型        名称      意味
'//  引数      : String 　 sName     [IN]上位機器名称
'//  引数      : Integer　 iAreaID   [IN]上位機器通信状態エリアID
'//  引数      : Integer　 nIndex    [IN]上位機器設定構成
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【接続機器見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psAddKikiCornerName(sName As String, iAreaID As Integer, nIndex As Integer)

    Dim nCorner As Integer                  ' コーナインデックス
    Dim szCornerName As String              ' コーナ名称
    Dim nNullIndex As Integer               ' 文字数ワーク
    Dim szResultName As String              ' 出力名称

    szResultName = ""
    nCorner = 0                                         ' コーナ設定不要
    gTransKikiInfo(nIndex).nIniListNo = nIndex          ' 外部機器リスト番号
    gTransKikiInfo(nIndex).nProcType = PROCTYPE_NORMAL  ' 処理タイプ
    gTransKikiInfo(nIndex).iAreaID = iAreaID            ' 参照エリア
    ' 1.対象外部機器上位機器通信状態エリアIDをチェックして
    '   接続対象を選別する。
    Select Case iAreaID
    Case IdKikiComSts.ID_DESYU_COM                                       ' 1:デ集通信状態
        nCorner = 1
    Case IdKikiComSts.ID_DESYU2_COM                                      ' 9:デ集2通信状態
        nCorner = 2
    Case IdKikiComSts.ID_DESYU3_COM                                      ' 10:デ集3通信状態
        nCorner = 3
    Case IdKikiComSts.ID_DESYU4_COM                                      ' 11:デ集4通信状態
        nCorner = 4
    Case IdKikiComSts.ID_DESYU5_COM                                      ' 12:デ集5通信状態
        nCorner = 5
    Case IdKikiComSts.ID_DESYU6_COM                                      ' 13:デ集6通信状態
        nCorner = 6
    Case IdKikiComSts.ID_ENKAKU_COM                                      ' 2:遠隔通信状態
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 1
    Case IdKikiComSts.ID_ENKAKU2_COM                                     ' 21:遠隔2通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 2
    Case IdKikiComSts.ID_ENKAKU3_COM                                     ' 22:遠隔3通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 3
    Case IdKikiComSts.ID_ENKAKU4_COM                                     ' 23:遠隔4通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 4
    Case IdKikiComSts.ID_ENKAKU5_COM                                     ' 24:遠隔5通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 5
    Case IdKikiComSts.ID_ENKAKU6_COM                                     ' 25:遠隔6通信状態（エリア定義なし）
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 6
    Case Else
    End Select

    gTransKikiInfo(nIndex).nCorner = nCorner
    If gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU Then
        gTransKikiInfo(nIndex).iAreaID = IdGate.ENKAKUKIKI_JIKAIAREAID
        gTransKikiInfo(nIndex).iErrorInfoID = IdGate.ENKAKUKIKI_JIKAIERRSTATUSID
        gTransKikiInfo(nIndex).iErrorTypeID = IdGate.ENKAKUKIKI_JIKAIERRTYPEID
    End If
    
    If nCorner <> 0 Then
        If gblnCornerSet(nCorner - 1) <> True Then
            Exit Sub
        End If
        ' コーナ名称の付加
        nNullIndex = InStr(gstrCornerName(nCorner - 1), Chr(0))
        If nNullIndex <> 0 Then
            szCornerName = vbCrLf & Left(gstrCornerName(nCorner - 1), nNullIndex - 1)
        Else
            szCornerName = vbCrLf & gstrCornerName(nCorner - 1)
        End If
    End If
    szResultName = Left(sName, InStr(sName, Chr(0)) - 1)
    szResultName = szResultName + szCornerName
    gTransKikiInfo(nIndex).sGetInf = szResultName
    gTransKikiInfo(nIndex).bStatus = True

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : pfGetConectStsJikai
'//  機能名称  : 自改状態エリアより状態取得処理
'//  機能概要  : 自改状態エリアより
'//              通信状態、異常状態、異常種別の取得を行う。
'//
'//              型        名称      意味
'//  引数      : iCnt　　Integer　　[IN]取得カウンター
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetConectStsJikai(iCnt As Integer)
    Dim iAreaSts As Integer                 ' 監視設定状態値
        
    Dim iJikaiaArea_Jyotai As Integer       ' 自改状態エリア状態値
    Dim lngMuHandle As Long                 ' 排他処理用ハンドル
    Dim strMutexName As String
        
    Dim iAreaID As Integer                  ' 通信状態エリアID
    Dim iGokiNo As Integer                  ' 自改状態の号機
    Dim iErrorInfoID As Integer             ' 通信異常状態エリアID
    Dim iErrorTypeID As Integer             ' 通信異常種別エリアID
        
    On Error Resume Next
    
    strMutexName = "Mu_" & GGateStatus
    lngMuHandle = dllOpenMutex(strMutexName)            '排他処理(OPEN)
    If lngMuHandle = 0 Then
       'エリア参照不可のため、参照異常
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Exit Function
    End If
  
    dllCloseHandle (lngMuHandle)                 '排他処理(CLOSE)
    
    ' 設定情報の取得
    iAreaID = gTransKikiInfo(iCnt).iAreaID
    iErrorInfoID = gTransKikiInfo(iCnt).iErrorInfoID
    iErrorTypeID = gTransKikiInfo(iCnt).iErrorTypeID
    iGokiNo = gTransKikiInfo(iCnt).nCorner
    
    Set Idinf_JikaiJyotai = New IdInfProc              '自改状態エリア
    '参照(自改状態)エリア名を設定
    Idinf_JikaiJyotai.ProcMode = DATA_ID.Data_Id_JkaiJyotai    '自改状態エリア
    Idinf_JikaiJyotai.IdOpen
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
        Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
    
    '参照(自改状態)エリアをＬＯＣＫする。
    Idinf_JikaiJyotai.IdLock
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 通信状態
    'エリアの内容を読み込む。
    Idinf_JikaiJyotai.id = iAreaID
    Idinf_JikaiJyotai.GetJikai_Sts iGokiNo - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
   
    '通信状態を取得
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGokiNo - 1))
    iConectSts = iJikaiaArea_Jyotai
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 通信異常状態
    'エリアの内容を読み込む。
    Idinf_JikaiJyotai.id = iErrorInfoID
    Idinf_JikaiJyotai.GetJikai_Sts iGokiNo - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
  
    '通信異常状態を取得
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGokiNo - 1))
    iSts_Naiyou = iJikaiaArea_Jyotai
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 通信異常種別
    'エリアの内容を読み込む。
    Idinf_JikaiJyotai.id = iErrorTypeID
    Idinf_JikaiJyotai.GetJikai_Sts iGokiNo - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       'データ参照異常時はブランク表示設定を行う。
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '自改状態エリア
       Exit Function
    End If
  
    '通信異常種別を取得
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGokiNo - 1))
    iSts_Type = iJikaiaArea_Jyotai
     
    Idinf_JikaiJyotai.IdFree
    Set Idinf_JikaiJyotai = Nothing                 '自改状態エリア
     
End Function


