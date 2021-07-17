VERSION 5.00
Begin VB.Form frmUtilityUSR 
   BorderStyle     =   0  'なし
   Caption         =   "ユーティリティ起動"
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
      Interval        =   3000
      Left            =   5640
      Top             =   8520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "次画面"
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
      Index           =   1
      Left            =   6840
      TabIndex        =   12
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "前画面"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   8
      Left            =   2040
      TabIndex        =   10
      Top             =   6600
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   9
      Left            =   6360
      TabIndex        =   9
      Top             =   6600
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   6
      Left            =   2040
      TabIndex        =   8
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   7
      Left            =   6360
      TabIndex        =   7
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   5
      Left            =   6360
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "○○○○○○○○○○○○"
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
      TabIndex        =   1
      Top             =   840
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
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "ユーティリティ起動"
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
      TabIndex        =   13
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmUtilityUSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmUtilityUSR.frm
'//  パッケージ名：ユーティリティ起動(一般メンテナンス)画面
'//
'//  概要：ユーティリティ起動(一般メンテナンス)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const iHoshuAplMax = 19           '登録最大件数
Private sFixedExePass(0 To 31) As String  '固定起動釦に対応したｱﾌﾟﾘﾌｧｲﾙﾊﾟｽ名（ﾖﾋﾞｴﾘｱを含む）
Private sFixedExeName(0 To 31) As String  '固定起動釦に対応した釦名称（ﾖﾋﾞｴﾘｱを含む）
Private iHyoujiCnt As Integer             '表示カウンター
Private iGamenSts As Integer              '現在表示画面数

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ユーティリティ起動（一般メンテナンス)画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
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
    'メール受信用タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : ユーティリティ起動（一般メンテナンス)画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマを停止
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
    'メール受信用タイマを止める
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : ユーティリティ起動（一般メンテナンス)画面(ロード時)
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
    Dim i As Integer    'カウンター
    
On Error Resume Next
   
   '「ﾕｰﾃｨﾘﾃｨ画面：表示」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '初期化
    iHyoujiCnt = 0    '表示カウンター
    iGamenSts = 0     '現在表示画面数
    Command1(0).Visible = False  '「前画面」釦非表示
    Command1(1).Visible = False  '「次画面」釦非表示
    
    For i = 0 To 31
        '表示名エリア初期化
        sFixedExeName(i) = ""
    Next
    For i = 0 To 31
        'ツールパスエリア初期化
        sFixedExePass(i) = ""
    Next
     
    'V1.3.0.1 ADD START
    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    '1.3.0.1 ADD END
    
    '固定釦表示初期処理
    sFixedExeDisplay
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 固定起動釦押下時処理
'//  機能概要  : 該当アプリの起動を行う。
'//
'//              型        名称      意味
'//  引数      :Integer　　Index    [IN]起動アプリ釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    Dim lRetVal As Double         'Shell関数戻り値
    Dim iResponse As Integer      'MsgBox戻り値
    Dim iSetupAplIndex As Integer '起動アプリインデックス
    
On Error GoTo ERROR_MSG
   '「ﾕｰﾃｨﾘﾃｨ画面：起動釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_KIDOU_BUTTOM, 0)
  
    '画面設定インデックスは0～9なので、釦インデックス値を算出し、
    '起動アプリのパスで起動する。
    '起動アプリインデックス=(現在画面数-1画面)×1画面最大釦数＋押下インデックス(0～9)
    '例：2画面目の押下釦インデックス3が押下された場合、起動アプリパスインデックスは13
    '13=(2-1)＊10＋3
    iSetupAplIndex = (iGamenSts - 1) * 10 + Index
    
    '該当ボタンのアプリケーションを起動する。
    lRetVal = Shell(sFixedExePass(iSetupAplIndex), vbNormalFocus)
    '「ﾕｰﾃｨﾘﾃｨ画面：ツール起動正常」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_API, UTILITY_GAMEN_TOOL_OK, 0)
 
    Exit Sub
    
ERROR_MSG:
'===アプリ起動エラーの場合、
    '「ﾕｰﾃｨﾘﾃｨ画面：ツール起動異常」ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_API, UTILITY_GAMEN_TOOL_ERROR, 0)
    '「起動失敗」ポップアップ画面を表示する。
    iResponse = MsgBox(cmdFixedExe(Index).Caption & "釦、定義エラー。" & _
                Chr(vbKeyReturn) & _
                sFixedExePass(iSetupAplIndex) & "を起動できません。", _
                vbYes, _
               "固定起動アプリ実行エラー")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Command1_Click
'//  機能名称  : 「次画面」「前画面」釦押下時処理
'//  機能概要  : 「次画面」「前画面」釦押下により、対象画面を表示する。
'//
'//              型        名称      意味
'//  引数      :Integer　　Index    [IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Command1_Click(Index As Integer)
  Dim i As Integer          'INIﾌｧｲﾙキーカウンタ：DSPi ＝起動釦INDEX
  Dim iMax As Integer       '固定起動釦INDEX最大値

On Error Resume Next

 Select Case Index
  Case 0
   '「ﾕｰﾃｨﾘﾃｨ画面：前画面釦押下」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_BACK_BUTTOM, 0)
   If iGamenSts = 2 Then
    '現在表示画面数：2画面目。「前画面」釦が押下された。
    '表示開始点は0、次表示画面数は1画面目のため、現在表示画面数に1を設定する。
    iHyoujiCnt = 0
    iGamenSts = 1
   Else
    '現在表示画面数：1画面目。「前画面」釦が押下された。
    '次表示画面数は2画面目のため、現在表示画面数に2を設定する。
    iGamenSts = 2
   End If
    
    '全ての固定起動釦について、以下を実施する。
    iMax = cmdFixedExe.UBound     '固定起動釦INDEXの最大値を得る。
    For i = 0 To iMax
     '固定起動釦を非表示にする。
      cmdFixedExe(i).Visible = False
        '起動アプリパス名と、表示釦名称の定義チェックを行う。
        If sFixedExePass(iHyoujiCnt) <> "" And sFixedExeName(iHyoujiCnt) <> "" Then
          '定義有りの場合のみ、キャプションに起動釦表示文字列を書込み、起動釦を表示する。
          cmdFixedExe(i).Visible = True
          cmdFixedExe(i).Caption = sFixedExeName(iHyoujiCnt)
        End If
         '表示カウンタアップする。
         iHyoujiCnt = iHyoujiCnt + 1
    Next i
  
  Case 1
    '「ﾕｰﾃｨﾘﾃｨ画面：次画面釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_NEXT_BUTTOM, 0)
  
    If iGamenSts = 2 Then
     '現在表示画面数：2画面目。「次画面」釦が押下された。
     '表示開始点は0、次表示画面数は1画面目のため、現在表示画面数に1を設定する。
     iGamenSts = 1
      iHyoujiCnt = 0
    Else
     '現在表示画面数：1画面目。「次画面」釦が押下された。
     '次表示画面数をカウントアップする。
     iGamenSts = iGamenSts + 1
    End If
    
    '全ての固定起動釦について、以下を実施する。
    iMax = cmdFixedExe.UBound     '固定起動釦INDEXの最大値を得る。
    For i = 0 To iMax
     '固定起動釦を非表示にする。
       cmdFixedExe(i).Visible = False
        '起動アプリパス名と、表示釦名称の定義チェックを行う。
        If sFixedExePass(iHyoujiCnt) <> "" And sFixedExeName(iHyoujiCnt) <> "" Then
           '定義有りの場合のみ、キャプションに起動釦表示文字列を書込み、起動釦を表示する。
           cmdFixedExe(i).Visible = True
           cmdFixedExe(i).Caption = sFixedExeName(iHyoujiCnt)
        End If
          '表示カウンタアップする。
          iHyoujiCnt = iHyoujiCnt + 1
    Next i
   
   Case Else
    '処理無し
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
   '「ﾕｰﾃｨﾘﾃｨ画面：消去」ログ出力
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_END, 0)
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sFixedExeDisplay
'//  機能名称  : 固定アプリ起動釦初期表示処理
'//  機能概要  : 初期表示処理を行う。
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
Private Sub sFixedExeDisplay()
Dim i As Integer                     'INIﾌｧｲﾙキーカウンタ：DSPi ＝起動釦INDEX
Dim iMax As Integer                  '固定起動釦INDEX最大値
Dim sLine As String * MAX_PATH_SIZE  '１行文の文字列。（文字列”DSPi=”を除く）
Dim lSize As Long                    '１行文のﾊﾞｲﾄ数。（文字列”DSPi=”を除く）
Dim iK As Integer                    'カンマ記述位置
Dim iKensuFlag As Integer            '件数フラグ

On Error Resume Next

'画面数を１画面目とする。
iGamenSts = 1
iMax = cmdFixedExe.UBound     '固定起動釦INDEXの最大値を得る。

 For i = 0 To iHoshuAplMax
   'アプリ起動初期値INIファイルから、１行文の文字列（DSPi=を除く）を読込む。
    lSize = GetPrivateProfileString(PROFILE_SECTION_NAME_FIXED_EXE, _
                                    PROFILE_KEY_NAME_FIXED_EXE & CStr(i), _
                                    DEFAILT, sLine, Len(sLine), HOSHUAPL_FILE)
    iK = InStr(sLine, ",")        'ファイル名（フルパス）の区切文字位置を得る。
    'INIファイルに、該当行の定義がある場合、
    If lSize > 0 And iK <> 0 Then
     'ファイル名と釦名称をを取出し、保存しておく。
      sFixedExePass(i) = Trim$(Left$(sLine, iK - 1))
      sFixedExeName(i) = Trim$(Mid$(sLine, iK + 1, lSize - iK))
    End If
Next i

'全ての固定起動釦について、以下を実施する。
 For i = 0 To iMax
   '固定起動釦を非表示にする。
    cmdFixedExe(i).Visible = False
    '起動アプリパス名と、表示釦名称の定義チェックを行う。
    If sFixedExePass(i) <> "" And sFixedExeName(i) <> "" Then
       '定義有りの場合、キャプションに起動釦表示文字列を書込み、起動釦を表示する。
       cmdFixedExe(i).Visible = True
       cmdFixedExe(i).Caption = sFixedExeName(i)
    End If
    '表示カウンタアップする。
    iHyoujiCnt = iHyoujiCnt + 1
  Next i

For i = 0 To iHoshuAplMax
   If sFixedExePass(i) <> "" And sFixedExeName(i) <> "" And i > 9 Then
      Command1(0).Visible = True
      Command1(1).Visible = True
   End If
Next i
 
End Sub

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmUtilityUSR.Caption, False
        pfFormActive (frmUtilityUSR.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
