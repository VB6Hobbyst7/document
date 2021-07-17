VERSION 5.00
Begin VB.Form frmGateHoshu 
   BorderStyle     =   0  'なし
   Caption         =   "自改保守データ"
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "変更前データ保存"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "自改保守ＳＷ設定表示(幹)"
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
      Left            =   6360
      TabIndex        =   10
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ジャーナル印字"
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
      TabIndex        =   9
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "自改保守ＳＷ設定クリア"
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
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "自改保守ＳＷ設定表示(在)"
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
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "マスタデータ内容表示"
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
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "マスタデータ入力"
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
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ＩＣメンテナンス"
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
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Timer tmrMail 
      Left            =   3960
      Top             =   7680
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "統合監視盤締切"
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
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "稼動・メンテデータ収集"
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
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "データ収集・出力"
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
Attribute VB_Name = "frmGateHoshu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmGateHoshu.frm
'//  パッケージ名：自改保守データ画面
'//
'//  概要：自改保守データ画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　自改保守データクリア画面表示処理追加
'//     REVISIONS :(1.6.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-12   CODED   BY [TCC] M.Matsumoto
'//       【統-279,281対応】
'//     REVISIONS :(EG20 V7.2.0.1) 2013-06-14  CODED   BY [TCC] T.Nakajima
'//        2013年度施策 遠隔対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-17  CODED   BY [TCC] T.Nakajima
'//        北陸新幹線フェーズ２対応
'//         【HKRK_Kansi07_005_01】
'//     REVISIONS :(EG20 V32.1.0.1) 2016-06-07  CODED   BY [TCC] T.Nakajima
'//        2016年度施策対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
'Private sHyoujiGoukiNo(0 To 18) As String        '表示号機番号格納エリア           ' EG20 V6.9.0.1削除
Private sHyoujiGoukiNo(0 To 31) As String         '表示号機番号格納エリア           ' EG20 V6.9.0.1追加
Private Const TITLENAME_CORNER = "コーナ#"        ' コーナ名                        ' EG20 V6.9.0.1追加
Private sRonriCornerNo(0 To 31) As String         '論理コーナ番号格納エリア         ' EG20 V6.9.0.1追加
Private Const DEFAILT_HYOUJI_UMU = 1    '「稼動・メンテデータ収集」釦のデフォルト表示     'V2.7.0.1 ADD
Private iToolFlg                As Integer        ' ConfigViewer 幹線or在来フラグ      ' EG20 V30.3.0.1 【HKRK_Kansi07_005_01】ADD
Private Const CONFIG_VIEWER_ZAIRAI = 0            '「自改保守SW設定表示(在)」釦押下    ' EG20 V30.3.0.1 【HKRK_Kansi07_005_01】ADD
Private Const CONFIG_VIEWER_KANSEN = 1            '「自改保守SW設定表示(幹)」釦押下    ' EG20 V30.3.0.1 【HKRK_Kansi07_005_01】ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 自改保守データ画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
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
'//  機能名称  : 自改保守データ画面(ディアクティブ時)
'//  機能概要  : メール受信タイマ起動
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
    
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 自改保守データ画面(ロード時)
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
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//       ・ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-12   CODED   BY [TCC] M.Matsumoto
'//       【統-279,281対応】
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   Dim lSts As Long             '関数戻り値      'V2.7.0.1 ADD
    
    On Error Resume Next
    
    lSts = 0    '変数の初期化 'V2.7.0.1 ADD

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '「自改保守データ画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_GAMEN_START, 0)
    
   'V1.3.0.1 ADD START
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END

    'V2.7.0.1  ADD START
    'HOSHU.INIより、「稼動・メンテデータ収集」釦の表示有無を取得する。
    lSts = GetPrivateProfileInt(KANSI_HOSHU_DATA_SEC, _
                                   KANSI_HOSHU_DATA_KEY, _
                                   DEFAILT_HYOUJI_UMU, _
                                   HOSHU_FILE)
    If lSts = 1 Then
        cmdFixedExe(0).Visible = True
    Else
        cmdFixedExe(0).Visible = False
    End If
    'V2.7.0.1  ADD END
    
    'EG20 V2.1.0.1 ADD START 【統-279,281対応】
    '監視盤未起動時は一部ボタンを押下不可とする
    If CheckAppStart(PROC_KANRI) = 0 Then
        cmdFixedExe(1).Enabled = False          '締め切り
        cmdFixedExe(3).Enabled = False          'マスタデータ入力
        cmdFixedExe(4).Enabled = False          'マスタデータ内容表示
        cmdFixedExe(7).Enabled = False          'ジャーナル印字  EG20 V7.2.0.1
    End If
    'EG20 V2.1.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 各釦押下処理
'//  機能概要  : 各釦名称処理を行う。
'//              「稼動・メンテデータ収集」「自改保守SW設定表示」
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//                 フェーズ２対応　自改保守データクリア画面表示処理追加
'//     REVISIONS :(1.6.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-23   REVISED BY [TCC] D.Yamashita
'//                 フェーズ３残件項目対応　自改保守SW設定表示時に自駅.GLT作成処理を追加
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-22   REVISED BY [TCC] T.Koyama
'//                ＥＧ２０フェーズ対応【残件54】
'//                ・マスタデータ内容表示処理追加
'//                ・自改保守SW設定表示釦、自改保守ＳＷ設定クリア釦の復活
'//     REVISIONS :(EG20 V7.2.0.1) 2013-06-14  REVISED BY [TCC] T.Nakajima
'//                2013年度施策 遠隔対応
'//                ・ジャーナル印字釦追加
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-17  REVISED BY [TCC] T.Nakajima
'//                北陸新幹線フェーズ２対応
'//                【HKRK_Kansi07_005_01】自改保守SW設定表示（ConfigVirewer)幹在混在対応
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
 Dim iResponse As Integer 'MsgBoxの戻り値
 Dim lngErrCode As Long   'エラーコード
' EG20 V2.0.1.1【残件54】ADD START
 Dim lRetVal As Double                    'Shell関数戻り値
' EG20 V2.0.1.1【残件54】ADD END
 
 On Error Resume Next
  
  Select Case Index
        Case 0                                 '稼動・メンテデータ収集画面
            '「稼動・メンテデータ収集画面：稼動・メンテデータ収集釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_KADO_MENTE_BUTTOM, 0)
            Load frmSyusyu
            frmSyusyu.Show 1
        Case 1                                 '自改保守SW設定(設定コンフィグ確認ツール起動)
        'EG20 V2.1.0.1 DEL START
'            '「稼動・メンテデータ収集画面：自改保守SW設定表示釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SWSETTEI_BUTTOM, 0)
'
'            'V1.11.0.1 ADD START
'            'GLTファイルを作成し、内容を更新する。
'            fMakeGLTFile
'            'V1.11.0.1 ADD END
'            '自改保守SWデータファイルコピー処理
'            fGetGouki '表示号機取得
'            'If sSWFileCopy > 0 Then 'V1.6.0.1 DEL
'            sSWFileCopy 'V1.6.0.1 ADD
'                'コンフィグ設定確認ツール起動処理
'                sToolOn
'            'V1.6.0.1 DEL START
'            'Else
'            '  '「自改保守データ画面：自改保守SWデータファイルコピー異常」ログ出力
'            '  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'            '  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'            'End If
'            'V1.6.0.1 DEL END
'        'V1.4.0.1　ADD START
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START
            '「稼動・メンテデータ収集画面：統合監視盤締切釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SHIMEKIRI_BUTTOM, 0)
            Load frmShimekiriData
            frmShimekiriData.Show 1
        'EG20 V2.1.0.1 ADD END
        Case 2                                 '自改保守ＳＷ設定クリア画面
        'EG20 V2.1.0.1 DEL START
'            '「稼動・メンテデータ収集画面：自改保守ＳＷ設定クリア釦押下」ログ出力
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SETTEICLEAR_BUTTOM, 0)
'            Load frmHoshuSwClear
'            frmHoshuSwClear.Show 1
'        'V1.4.0.1　ADD END
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.0.1.1【残件№54】 ADD START
            '画面表示要求（状態監視機能設定）をID制御に送信する
            If (SendMessageDispInfo(ML_DT_IC_MAINTE) = False) Then
         
                iResponse = MsgBox("ＩＣメンテナンス釦、定義エラー。" & _
                                   Chr(vbKeyReturn) & _
                                   "ＩＣメンテナンス画面を起動できません。", _
                                   vbOKOnly, _
                                   "画面起動エラー")
            End If
        'EG20 V2.0.1.1【残件№54】 ADD END
        'EG20 V2.1.0.1 ADD START
        Case 3
            '「稼動・メンテデータ収集画面：マスタデータ入力釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_MST_INPUT_BUTTOM, 0)
            Load frmInputMstData
            frmInputMstData.Show 1
        'EG20 V2.1.0.1 ADD END
        'EG20 V2.0.1.1【残件54】ADD START
        Case 4
            '「マスタデータ内容表示釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_MST_DISP_BUTTOM, 0)
            ' マスタデータ内容表示ツール起動
            lRetVal = Shell("D:\KANSI\TOOL\DataViewer\BinViewer.exe", vbNormalFocus)
        
        Case 5
            '「稼動・メンテデータ収集画面：自改保守SW設定表示釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SWSETTEI_BUTTOM, 0)
            
            iToolFlg = CONFIG_VIEWER_ZAIRAI         '在来用のConfigViewerが起動された   EG20 V30.3.0.1【HKRK_Kansi07_005_01】 ADD

            'V1.11.0.1 ADD START
            'GLTファイルを作成し、内容を更新する。
            fMakeGLTFile
            'V1.11.0.1 ADD END
            '自改保守SWデータファイルコピー処理
            fGetGouki '表示号機取得
            'If sSWFileCopy > 0 Then 'V1.6.0.1 DEL
            sSWFileCopy 'V1.6.0.1 ADD
                'コンフィグ設定確認ツール起動処理
                sToolOn
            'V1.6.0.1 DEL START
            'Else
            '  '「自改保守データ画面：自改保守SWデータファイルコピー異常」ログ出力
            '  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            '  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
            'End If
            'V1.6.0.1 DEL END
        'V1.4.0.1　ADD START
        Case 6
            '「稼動・メンテデータ収集画面：自改保守ＳＷ設定クリア釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SETTEICLEAR_BUTTOM, 0)
            Load frmHoshuSwClear
            frmHoshuSwClear.Show 1
        'EG20 V2.0.1.1【残件54】ADD START
        'EG20 V7.2.0.1 ADD START
        Case 7
            '「稼動・メンテデータ収集画面：ジャーナル印字釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_JPR_PRINT_BUTTON, 0)
            Load frmJprPrint
            frmJprPrint.Show 1
        'EG20 V7.2.0.1 ADD END
        'EG20 V30.3.0.1 【HKRK_Kansi07_005_01】ADD START
        Case 8
            '「稼動・メンテデータ収集画面：自改保守SW設定(幹)表示釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SWSETTEI_KAN_BUTTOM, 0)
            
            iToolFlg = CONFIG_VIEWER_KANSEN         '幹線用のConfigViewerが起動された

            'GLTファイルを作成し、内容を更新する。
            fMakeGLTFile
            '自改保守SWデータファイルコピー処理
            fGetGouki '表示号機取得
            
            sSWFileCopy 'V1.6.0.1 ADD
                'コンフィグ設定確認ツール起動処理
                sToolOn
        'EG20 V30.3.0.1 【HKRK_Kansi07_005_01】ADD END
        'EG30 V32.1.0.1 ADD START
        Case 9
            '「稼動・メンテデータ収集画面：変更前データ保存釦押下」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SET_BEF_BUTTON, 0)
            Load frmSetteiBefore
            frmSetteiBefore.Show 1
        'EG30 V32.1.0.1 ADD END
        
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
    
    '「自改保守データ画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSWFileCopy
'//  機能名称  : 自改保守SW設定データファイル作成処理
'//  機能概要  : 自改保守SW設定データを、自改保守SWデータファイルに
'//              コピーする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.9.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sSWFileCopy() As Integer

     Dim iCnt As Integer                     'カウンター
     Dim sSWDataPath As String               '自改保守SWデータファイル
     Dim sMyPath As String                   '自改保守SW設定データ
     
     On Error Resume Next
   
     sSWFileCopy = 0                         'ファイル存在数
    
    '自改最大数分ループする。
    For iCnt = 1 To MAX_GATE_NO
     '「GATE_SW##.dat」の「##」を01～16に変換する。
     sMyPath = Replace(GATE_SW_FILE, "##", Format(iCnt, "0#"))
     '自改保守SW設定データの検索を行う。
     If Dir(sMyPath) <> "" Then
        '自改保守SWデータファイルのパスを作成する。
        sSWDataPath = PATH_RMENTE_GATE_DEN_JIEKI_GOUKI
' EG20 V6.9.0.1追加開始
        '「コーナ$」の「$」を1～6に変換する。
        sSWDataPath = Replace(sSWDataPath, "$", sRonriCornerNo(iCnt - 1))
' EG20 V6.9.0.1追加終了
        '「##号機」の「##」を01～16に変換する。
        sSWDataPath = Replace(sSWDataPath, "##", Format(sHyoujiGoukiNo(iCnt - 1), "0#"))
        'フォルダ作成
        MkDir sSWDataPath
        sSWDataPath = sSWDataPath & TOOL_SW_File
        
        '自改保守SWデータを自改保守SWデータファイルにコピーする。
        FileCopy sMyPath, sSWDataPath
        sSWFileCopy = sSWFileCopy + 1
     End If
   Next

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sToolOn
'//  機能名称  : 自改保守SW設定ツール起動処理
'//  機能概要  : 自改保守SW設定ツールの起動を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-17  REVISED BY [TCC] T.Nakajima
'//                北陸新幹線フェーズ２対応
'//                【HKRK_Kansi07_005_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sToolOn() As Integer
    Dim lRetVal As Double                    'Shell関数戻り値
    Dim sToolName As String * MAX_PATH_SIZE  'ツールパス名
    Dim lSize As Long                        '戻り値
    
    On Error Resume Next
   
    '保守機能INIファイルから、自改保守SW設定ツールパスを取得する。
    'EG20 V30.3.0.1 DEL START【HKRK_Kansi07_005_01】
'    lSize = GetPrivateProfileString(KANSI_HOSHU_SW_TOOL_SEC, _
'                                    KANSI_HOSHU_SW_TOOL_KEY, _
'                                    DEFAILT, sToolName, Len(sToolName), HOSHU_FILE)
    'EG20 V30.3.0.1 DEL END【HKRK_Kansi07_005_01】
    'EG20 V30.3.0.1 ADD START 【HKRK_Kansi07_005_01】
    '(在)、(幹)どちらの釦が押下されているか？
    If iToolFlg = CONFIG_VIEWER_KANSEN Then
        ' (幹)が押下されているのConfigViewer2のパスを取得
        lSize = GetPrivateProfileString(KANSI_HOSHU_SW_TOOL_SEC, _
                                        KANSI_HOSHU_SW_TOOL_KEY_KAN, _
                                        DEFAILT, sToolName, Len(sToolName), HOSHU_FILE)
    Else
        '(在)が押下されているのでConfigViewerのパスを取得
        lSize = GetPrivateProfileString(KANSI_HOSHU_SW_TOOL_SEC, _
                                        KANSI_HOSHU_SW_TOOL_KEY, _
                                        DEFAILT, sToolName, Len(sToolName), HOSHU_FILE)
    End If
    'EG20 V30.3.0.1 ADD END 【HKRK_Kansi07_005_01】
    
    'INIファイルに、該当行の定義がある場合、
    If sToolName <> "" Then
        '自改保守SW設定ツールを起動する。
        lRetVal = Shell(sToolName, vbNormalFocus)
    End If
 
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fMakeGLTFile
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fGetGouki() As Integer
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

    For iGate = CNT_MIN To MAX_GATE_NO - 1
      '自動改札機情報取得
      sKeyName = "gate" & Format(iGate + 1, "00")
      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                     sKeyName, _
                                     DEFAILT, sGateData, Len(sGateData), _
                                     PATH_GATE_FILE)
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
         sGoukiNo = "0" & Trim(sFData(1))
      Else
         sGoukiNo = Trim(sFData(1))
      End If
        
      sHyoujiGoukiNo(iGate) = sGoukiNo
' EG20 V6.9.0.1 【号機番号にコーナ番号を付加する対応】追加開始
      sRonriCornerNo(iGate) = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.9.0.1 【号機番号にコーナ番号を付加する対応】追加終了

    Next
    
    fGetGouki = 0
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
        AppActivate frmGateHoshu.Caption, False
        pfFormActive (frmGateHoshu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
'V1.11.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  関数名称  : fMakeGLTFile
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
'//     ORIGINAL  :(1.11.0.1) 2009-12-23   CODED   BY [TCC] D.Yamashita
'//                 フェーズ３残件項目対応　自改保守SW設定表示時に自駅.GLT作成処理を追加
'//     REVISIONS :(EG20 V6.7.0.1)  2012-06-28  CODED BY  [TCC] H.Sugimoto
'//                 【項目チェックの対象を改札機情報のみとする修正】
'//     REVISIONS :(EG20 V30.3.0.1)  2014-09-18  CODED BY  [TCC] T.Nakajima
'//                  北陸新幹線フェーズ２対応
'//                 【HKRK_Kansi07_005_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fMakeGLTFile() As Integer
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
    Dim szCorner As String      ' コーナ番号
    Dim szTitleName As String                       ' タイトル名                    ' EG20 V6.7.0.1追加
    Dim fso As New FileSystemObject                 'ファイルシステムオブジェクト   ' EG20 V6.7.0.1追加

    On Error Resume Next
    MkDir PATH_RMENTE_GATE_DEN   '自改用電鉄フォルダを作成する。（GLTファイル用）
    
    ' EG20 V30.3.0.1 ADD START 【HKRK_Kansi07_005_01】
    ' 各コーナのコーナ種別を取得
    gsGetCornerType
    ' EG20 V30.3.0.1 ADD END 【HKRK_Kansi07_005_01】
    
' EG20 V6.7.0.1追加開始
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH_RMENTE_GATE_DEN_JIEKI) = False Then
        'コピー先フォルダ作成
        fso.CreateFolder (PATH_RMENTE_GATE_DEN_JIEKI)
    End If
    Set fso = Nothing
' EG20 V6.7.0.1追加終了
    
    'GLTファイルを開く。ファイルが存在しなければ新規に作成される。
    On Error GoTo ErrorHandlerGLTFile
    intGLTFileNo = FreeFile        ' 未使用のファイル番号を取得する。
    Open GATE_GLT_FILE_FULLPASS For Output As #intGLTFileNo     ' GLTファイルを開く。

    For iGate = CNT_MIN To MAX_GATE_NO - 1
      '自動改札機情報取得
      sKeyName = "gate" & Format(iGate + 1, "00")
      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                     sKeyName, _
                                     DEFAILT, sGateData, Len(sGateData), _
                                     PATH_GATE_FILE)
      If iRet = 0 Then
         '「自改保守データ画面：自動改札機INIファイル読込異常」ログ出力
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
'         sGoukiNo = "0" & Trim(sFData(1)) & "号機"                                 ' EG20 V6.7.0.1削除
         sGoukiNo = "0" & Trim(sFData(1))                                           ' EG20 V6.7.0.1追加
      Else
'         sGoukiNo = Trim(sFData(1)) & "号機"                                       ' EG20 V6.7.0.1削除
         sGoukiNo = Trim(sFData(1))                                                 ' EG20 V6.7.0.1追加
      End If
        
' EG20 V6.9.0.1 【号機番号にコーナ番号を付加する対応】追加開始
'      szCorner = Replace(TITLENAME_CORNER, "#", Trim(sFData(GATE_IDX.IDX_RONRI_CORNER)))   ' EG20 V6.7.0.1削除
      szCorner = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))                                    ' EG20 V6.7.0.1追加
      sRonriCornerNo(iGate) = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.9.0.1 【号機番号にコーナ番号を付加する対応】追加終了
' EG20 V6.7.0.1 【号機番号にコーナ番号を付加する対応】追加開始
      ' コーナ番号変換
      szTitleName = Replace(RMENTE_GOKITITLENAME, "$", szCorner)
      ' 号機番号変換
      szTitleName = Replace(szTitleName, "##", sGoukiNo)
' EG20 V6.7.0.1 【号機番号にコーナ番号を付加する対応】追加開始
        
      If Trim(sFData(4)) <> "＊" Then
         'Gate.iniファイルの号機番号表示文字、IPアドレスをGLTファイルに書き込む。
'         Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(5))                     ' EG20 V6.6.0.1削除
'          Print #intGLTFileNo, szCorner & "_" & sGoukiNo & "," & Trim(sFData(5))   ' EG20 V6.6.0.1追加     ' EG20 V6.7.0.1削除
         'EG20 V30.3.0.1 DEL START 【HKRK_Kansi07_005_01】
         'Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))                   ' EG20 V6.7.0.1追加
         'EG20 V30.3.0.1 DEL END 【HKRK_Kansi07_005_01】
         
         'EG20 V30.3.0.1 ADD START 【HKRK_Kansi07_005_01】
         '現在処理中の号機が属する論理コーナの種別は？
         
         'ConfigViewerかConfigViewr2どちらを起動するのか？
         If iToolFlg = CONFIG_VIEWER_KANSEN Then
            '(幹)釦が押下されているので幹線コーナの号機のみGLTファイルに更新する。
            If gintCornerType(CInt(szCorner) - 1) = CORNER_TYPE_KANSEN Then
                Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))
            Else
                '条件に当てはまらない場合は何もGLTファイルには入れない。
            End If
        Else
            '(在)釦が押下されているので在来コーナの号機のみGLTファイルに更新する。
            If gintCornerType(CInt(szCorner) - 1) = CORNER_TYPE_ZAIRAI Then
                Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))
            Else
                '条件に当てはまらない場合は何もGLTファイルには入れない。
            End If
        End If
         'EG20 V30.3.0.1 ADD END 【HKRK_Kansi07_005_01】
      End If
      
      '表示号機番号
      If Len(Trim(sFData(1))) = 1 Then
         '号機番号が１桁ならば、先頭に０を付加する。
         sHyoujiGoukiNo(iGate) = "0" & Trim(sFData(1))
      Else
         sHyoujiGoukiNo(iGate) = Trim(sFData(1))
      End If
    
    Next
    
    'GLTファイルを閉じる。
    Close #intGLTFileNo
    
    fMakeGLTFile = 0    '正常終了
    Exit Function

ErrorHandlerGateIni:
   '「自改保守データ画面：ファイルアクセス異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 1
   'GLTファイルを閉じる。
   Close #intGLTFileNo
   Exit Function
ErrorHandlerGLTFile:
   '「自改保守データ画面：ファイルアクセス異常」ログ出力
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 2
   'GLTファイルを閉じる。
   Close #intGLTFileNo

End Function
'V1.11.0.1 ADD END

' EG20 V2.0.1.1【残件№54】ADD START
'///////////////////////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  関数名称    ：SendMessageDispInfo
'//  機能名称    ：画面表示状態通知
'//  機能概要    ：画面表示状態通知を行う。
'//
'//                 型      名称                意味
'//  引数         : Long    lDispInfo           画面要求種別
'//
'//  戻り値       : TRUE    （正常）
'//                 FALSE   （異常）
'//
'//  ORIGINAL    ：(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応
'//                 EG20統合監視盤USDM対応番号【Mainte_03_01】
'//  REVISIONS    :(EG20 V2.0.1.1) 2011-11-22  CODED BY  [TCC] T.Koyama
'//                ＥＧ２０フェーズ２対応【残件№54】
'//                ・システム設定メニュー画面より流用
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  備考         :システム設定メニュー画面
'///////////////////////////////////////////////////////////////////////////////////////////////
Private Function SendMessageDispInfo(ByVal lDispInfo As Long) As Boolean

    Dim udtMail As ML_DISP_INF          '画面表示要求
    Dim bRet As Boolean                 'メール送信処理戻り値
    Dim lngErrCode As Long              'エラーコード
    
    '画面表示要求をID制御に送信する
    udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
    udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
    udtMail.udtlHeader.dwProid = RHOSHU_ID
    udtMail.udtlHeader.dwSubArea = 0
    udtMail.dwDisp_Type = lDispInfo
    bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
    If bRet = False Then
        '「画面表示要求メール送信異常」ログ出力
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
    Else
   
        '「画面表示要求メール送信正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
    End If
    
    SendMessageDispInfo = bRet

End Function

' EG20 V2.0.1.1【残件№54】ADD END
