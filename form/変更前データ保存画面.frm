VERSION 5.00
Begin VB.Form frmSetteiBefore 
   BorderStyle     =   0  'なし
   Caption         =   "各設定値保存・反映"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
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
   Begin VB.Frame Frame7 
      Caption         =   "コーナ選択"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8775
      Begin VB.CheckBox RenewChk 
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   870
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "○○○○○○○○○○○○"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   4440
         Value           =   1  'ﾁｪｯｸ
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "保存済み設定ファイル作成日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9年 Z9月 Z9日 Z9時 Z9分 Z9秒"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   16
         Top             =   915
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9年 Z9月 Z9日 Z9時 Z9分 Z9秒"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   15
         Top             =   1605
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9年 Z9月 Z9日 Z9時 Z9分 Z9秒"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   14
         Top             =   2325
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9年 Z9月 Z9日 Z9時 Z9分 Z9秒"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   13
         Top             =   3045
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9年 Z9月 Z9日 Z9時 Z9分 Z9秒"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   12
         Top             =   3765
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9年 Z9月 Z9日 Z9時 Z9分 Z9秒"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4080
         TabIndex        =   11
         Top             =   4485
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label Label2 
         Caption         =   "コーナ名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   405
         TabIndex        =   10
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "データ収集・出力  画面へ戻る"
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
      Left            =   9500
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8040
   End
   Begin VB.CommandButton cmdKeep 
      Caption         =   "保存"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "変更前データ保存"
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
Attribute VB_Name = "frmSetteiBefore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 ALL Rights Reserved
'//
'//  ファイル名  ：変更前データ保存.frm
'//  パッケージ名：変更前データ保存のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-10   CODED   BY [TCC] T.Nakajima
'//                 ・係員設定保存復元(frmRenewData.frm)を流用
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const K_SETTEI = 0                  '監視設定
Private Const G_SETTEI = 1                  '自改設定
Private Const MSG_NORMAL = "正常"           '正常終了時表示文言
Private Const MSG_ERROR = "異常"            '異常終了時表示文言
Private Const RET_ERROR = -1                '異常
Private Const RET_NASI = 0                  '変更無し
Private Const RET_ARI = 1                   '変更有り

Private Const INVALID_HANDLE_VALUE = -1     'ハンドルエラー
Private Const MN_MAIL_INTERVAL = 1000       'メイルタイマのインターバル値

Private glbSaveFoldePath    As String       '保存ファイル格納用フォルダパス
Private udtIniGate          As INI_GATE     '機器情報自動改札機エリア           'EG20 V2.1.0.1 DEL 【Mainte_03_01】
Private udtIniGateFile      As INI_GATE     '機器情報自動改札機エリア保存ファイル

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 係員設定保存復元画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    'タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 係員設定保存復元画面(ディアクティブ時:イベントプロシージャ)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    'タイマを停止する
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 係員設定保存復元画面(ロード時：イベントプロシージャ)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-02  CODED BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim lSts As Long
    Dim strPath As String * 128
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim intCount As Integer         'ループカウンタ
    Dim intIndex As Integer         'チェックボックスIndex
    Dim strSaveFile As String       '保存ファイルパス
    'EG20 V2.1.0.1 ADD END
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SET_BEF_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    glbSaveFoldePath = ""
    strPath = ""
    
    intIndex = 0
    gsGetGateInfo
    
    Call gsGetGateInfo
    Call gsGetCornerName
    Call gsGetCornerType        'コーナ種別を取得   EG20 V30.1.0.1 ADD
    
    For intCount = 0 To UBound(gblnCornerSet)
    
        '設定ありのコーナ
        If gblnCornerSet(intCount) = True Then
            'コーナー名称表示
            RenewChk(intIndex).Caption = gstrCornerName(intCount)
            'コーナIndexを記録
            RenewChk(intIndex).Tag = CStr(intCount)
            RenewChk(intIndex).Visible = True
            lblSetteDate(intIndex).Visible = True
            intIndex = intIndex + 1
        End If
    
    Next intCount
    
    ' RENEWDATAINFO.INIから保存先フォルダパスを取得する
    lSts = GetPrivateProfileString(RENEWDATA_SECTION_NAME, _
                                   FOLDER_PATH_KEY_NAME, _
                                   "", _
                                   strPath, _
                                   Len(strPath), _
                                   PATH_RENEWDATAINFO_FILE)
    'INIファイル情報取得結果
    If strPath = "" Then
        'INI情報取得異常時、設定値保存復元フォルダをデフォルト設定とする
        glbSaveFoldePath = PATH_HOSHU_RENEW_DATA
    Else
        'INI情報を設定
        glbSaveFoldePath = Left$(strPath, lSts) & "\\"
    End If
    
    '画面設定処理
    Call sFromInitialize
    
    'メイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdKeep_Click
'//  機能名称  : 「保存」釦押下時処理
'//  機能概要  : 確認メッセージ表示後、自改・監視設定ファイルを
'//              自改・監視設定保存ファイルにコピーする
'//              機器情報自動改札機エリアの情報を取得し、保存ファイルに書き込む
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 5.4.0.1) 2012-03-25  REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-22  REVISED BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdKeep_Click()
    
    Dim iResponse       As Integer      'MsgBox戻り値
    Dim strMessage      As String       'MsgBox文言
    Dim bRet            As Boolean      '戻り値
    Dim intCount        As Integer
    Dim udtSendData     As MAIL_KAKARIIN_SETTEI
    Dim lngRet          As Long
    Dim blnIsSelected   As Boolean
    Dim intGokiCount    As Integer      'そのコーナの号機数
    Dim intComSts       As Integer      'その自改の通信状態
    Dim bResult         As Boolean
    
    Dim fso             As FileSystemObject
    
    intComSts = False
    
    On Error Resume Next

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SET_BEF_GAMEN_SAVE, 0)
    
    'コーナ別の号機情報を取得する
    Call gsGetGateInfo

    Erase glngTergetCorner
    blnIsSelected = False
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnIsSelected = True
            If RenewChk(intCount).Visible = True Then
                glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
                'そのコーナに属する改札機の通信状態をチェックする
                '監視盤起動有無チェック
                If CheckAppStart(PROC_KANRI) <> 0 Then
                    For intGokiCount = 0 To gudtSettiCorner(intCount).intGokiNum - 1
                        gpfGetjikaiConectSts intComSts, gudtSettiCorner(intCount).intGateNo(intGokiCount)
                        If intComSts <> CONECTSTS_NORMAL Then
                            Exit For
                        End If
                    Next
                    '1台でも通信異常の改札機があれば、警告を表示するので、コーナ単位のループを抜ける
                    If intComSts <> CONECTSTS_NORMAL Then
                        Exit For
                    End If
                Else
                    '保守単独起動の場合は改札機保守設定データが最新では無いことを通知する
                    Exit For
                End If
            End If
        End If
    Next intCount
    '選択コーナなしの場合、メッセージを表示して終了
    If blnIsSelected = False Then
        iResponse = MsgBox("対象コーナが選択されていません。" & vbCrLf & _
                            "選択してください。", vbOKOnly + vbCritical, "コーナ未選択")
        Exit Sub
    End If
    
    'EG30 V32.1.0.1 ADD START
    If intComSts <> CONECTSTS_NORMAL Then
        iResponse = MsgBox("選択したコーナに通信異常の改札機があります。" & vbCrLf & _
                            "通信異常号機の改札機保守設定データは最新で無い可能性があります。", _
                            vbOKOnly + vbExclamation, "通信異常改札機有り")
    End If
    'EG30 V32.1.0.1 ADD END
    
    '画面をロックする
    cmdKeep.Enabled = False
    cmdReturn.Enabled = False
    
    '変更前データ保存
    pfCopySetteiFiles bResult
    
    If bResult = False Then
        iResponse = MsgBox("保存に失敗した項目があります。", vbOKOnly + vbExclamation, "保存失敗")
    Else
        iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "正常終了")
    End If
    '画面ロック解除
    cmdKeep.Enabled = True
    cmdReturn.Enabled = True
    
    '画面表示を更新
    Call sFromInitialize
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next
    
    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SET_BEF_GAMEN_END, 0)
    
    '自画面を消す。
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sFromInitialize
'//  機能名称  : 画面設定処理
'//  機能概要  : ファイルの作成日時を取得し、作成日付表示部に表示
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sFromInitialize()
    Dim strFileName(0 To 1)     As String           '作成日時
    Dim intCnt                  As Integer          'カウンタ
    Dim lngHandle               As Long             'ハンドル

    Dim lpCreatTime             As FILETIME         '作成日時
    Dim lpAccessTime            As FILETIME         '最終アクセス日時
    Dim lpLastwTime             As FILETIME         '更新日時
    Dim lpLocalTime             As FILETIME         'ローカル日時
    Dim lpSystemTime            As SYSTEMTIME       'システム時刻
    Dim bRet                    As Boolean          '戻り値
    
    Dim blnExistFile            As Boolean          '保存ファイル有無
    Dim strSaveFile             As String
    Dim intIndex                As Integer
    
    On Error Resume Next

    blnExistFile = False
    intIndex = 0
    For intCnt = 0 To UBound(gudtSettiCorner)
        If gblnCornerSet(intCnt) = True Then
            '保存ファイルの日付を取得
            strSaveFile = PATH_OPERATE & "CORNER" & CStr(intCnt + 1) & "\\SETTEI_BEF\\" & SET_BEF_DATE_FILE
            If Dir(strSaveFile) = "" Then
                lblSetteDate(intIndex).Caption = "    年   月   日   時   分   秒"
            Else
                'ファイルをオープン
                lngHandle = CreateFile(strSaveFile, GENERIC_READ, FILE_SHARE_READ, _
                                        0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)

                'ファイルオープンが正常に行われたか？
                If lngHandle = INVALID_HANDLE_VALUE Then GoTo ErrorHandler
        
                'ファイルタイムをGET
                bRet = GetFileTime(lngHandle, lpCreatTime, lpAccessTime, lpLastwTime)
                If bRet = False Then GoTo APIError                          '取得が正常に行われたか？
        
                'ファイルタイムをローカルタイムに変換
'                bRet = FileTimeToLocalFileTime(lpCreatTime, lpLocalTime)    'EG20 V2.1.0.1 DEL 【Mainte_03_01】
                bRet = FileTimeToLocalFileTime(lpLastwTime, lpLocalTime)    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
                If bRet = False Then GoTo APIError                          '変換が正常に行われたか？
        
                'ローカルタイムをシステムタイムに変換
                bRet = FileTimeToSystemTime(lpLocalTime, lpSystemTime)
                If bRet = False Then GoTo APIError                          '変換が正常に行われたか？
        
                'ハンドルのクローズ
                Call CloseHandle(lngHandle)
        
                '作成日付を表示する (YYYY年MM月DD日hh時mm分ss秒)
                lblSetteDate(intIndex).Caption = lpSystemTime.wYear & "年 " & _
                                                Right("  " & lpSystemTime.wMonth, 2) & "月 " & _
                                                Right("  " & lpSystemTime.wDay, 2) & "日 " & _
                                                Right("  " & lpSystemTime.wHour, 2) & "時 " & _
                                                Right("  " & lpSystemTime.wMinute, 2) & "分 " & _
                                                Right("  " & lpSystemTime.wSecond, 2) & "秒"
                blnExistFile = True
            End If
            
            lblSetteDate(intIndex).Visible = True
            intIndex = intIndex + 1
        Else
            lblSetteDate(intIndex).Visible = False
        End If
    Next
    
    Exit Sub

APIError:

    Call CloseHandle(lngHandle)             'ハンドルのクローズ

ErrorHandler:

    '存在しない場合またはエラーが発生した、ブランクを表示
    For intCnt = intCnt To UBound(gudtSettiCorner)
        If lblSetteDate(intCnt).Visible = True Then
            lblSetteDate(intCnt).Caption = "    年   月   日   時   分   秒"
        End If
    Next intCnt
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ処理（タイムアップ時：イベントプロシージャ）
'//  機能概要  : 汎用メイル受信処理を行う
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : Long　 　 サイズ         メール送信サイズ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSetteiBefore.Caption, False
        pfFormActive (frmSetteiBefore.hwnd)
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  関数名称  : pfCopySetteiFiles
'//  機能名称  : 操作卓設定情報、現在駅設定データ、自改保守設定データを変更前保存用とする。
'//  機能概要  :
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : boolean   TRUE/FALSE 正常：True 異常：false
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-10   CODED   BY [TCC] T.Nakajima
'//                 2016年施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub pfCopySetteiFiles(bResult As Boolean)
    Dim fso, f              As Object
    Dim i, j                As Integer
    Dim bRet                As Boolean
    Dim strSetteiBefFolder  As String       '変更前保存用フォルダパス
    Dim strSetteiBefFolderZero  As String   '変更前保存用フォルダパス(コーナ0）
    Dim strOperateSetteiFolder  As String   '操作卓設定フォルダ
    Dim strJpCfgPath            As String   '号機別設定コンフィグファイルパス
    Dim bIsFileExists           As Boolean
    Dim textFile                As TextStream
    Dim lngMuHandle             As Long     '排他処理用ハンドル
    Dim strMutexName            As String   'ミューテックス名

    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    bRet = False
    
    bResult = True
    
    '駅都度データをコーナ０にコピーする
    strSetteiBefFolderZero = PATH_OPERATE & "CORNER" & CStr(0) & "\\SETTEI_BEF\\"
    '変更前保存用フォルダのファイルをすべて削除する
    fso.DeleteFile strSetteiBefFolderZero & "*.*", True
    
    If (CopyFile(EKI_SETTI_FILE, strSetteiBefFolderZero & fso.GetFileName(EKI_SETTI_FILE), True) = False) Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & fso.GetFileName(EKI_SETTI_FILE), 0)
        bResult = False
    End If
    '変更前保存データ日付を作成する(後にこのファイルをコーナ別にコピーする)
    Set textFile = fso.CreateTextFile(strSetteiBefFolderZero & SET_BEF_DATE_FILE, True)
    textFile.Close
        
    For i = 0 To RenewChk.UBound
        '選択されたコーナ
        If RenewChk(i).Visible = True And RenewChk(i).Value = CMN_ONOFF.CMN_ON Then
            strSetteiBefFolder = PATH_OPERATE & "CORNER" & CStr(i + 1) & "\\SETTEI_BEF\\"
            '変更前保存用フォルダのファイルをすべて削除する
            fso.DeleteFile strSetteiBefFolder & "*.*", True
            
            '操作卓設定情報をコピーする
            strOperateSetteiFolder = PATH_OPERATE & "CORNER" & CStr(i + 1) & "\\SETTEI\\"
            'SETTEIフォルダに何もファイルが無ければ、コピー処理はしない。
            If fso.GetFolder(strOperateSetteiFolder).files.Count > 0 Then
                For Each f In fso.GetFolder(strOperateSetteiFolder).files
                    
                    If (CopyFile(strOperateSetteiFolder & f.Name, strSetteiBefFolder & f.Name, True) = False) Then
                        'コピー元が存在しないので取得できなかった
                        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & f.Name, 0)
                        bResult = False
                        '操作卓設定情報のいずれかが失敗した場合は操作卓設定情報のコピーを中止し、ファイルを削除
                        fso.DeleteFile strSetteiBefFolder & "*.*", True
                        Exit For
                    End If
                Next
            Else
                'コピー元フォルダにファイルがないため、エラー
                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & strOperateSetteiFolder, 0)
                bResult = False
            End If
            
            '自改保守設定データをコピーする
            If gudtSettiCorner(i).intGokiNum > 0 Then
                'そのコーナに属する改札機分をコピーする
                For j = 0 To gudtSettiCorner(i).intGokiNum - 1
                    strJpCfgPath = PATH_DATA & Replace(JP_CFG, "##", Format(gudtSettiCorner(i).intGateNo(j), "0#"))
                    
                    '排他JP_CFGファイルを作成中の場合は待つ
                    strMutexName = Replace(MU_N_CFG, "##", Format(gudtSettiCorner(i).intGateNo(j), "0#"))
                    lngMuHandle = dllOpenMutex(strMutexName)
                    
                    If lngMuHandle <> 0 Then
                        dllWaitForSingleObject (lngMuHandle)
                    End If
                    
                    If (CopyFile(strJpCfgPath, strSetteiBefFolder & fso.GetFileName(strJpCfgPath), True) = False) Then
                        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & fso.GetFileName(strJpCfgPath), 0)
                        bResult = False
                    End If
                    
                    If lngMuHandle <> 0 Then
                        dllReleaseMutex (lngMuHandle)                   '排他処理(FREE)
                        dllCloseHandle (lngMuHandle)                    '排他処理(CLOSE)
                    End If
                    
                Next j
            Else
                ' そのコーナに1台も改札機が無い場合
                ' 改札機がそもそも存在しないので、改札機保守設定データも存在しないためエラーとしない
            End If
            
            '変更前保存データ日付を作成する（コーナ０からコピーする）
            fso.CopyFile strSetteiBefFolderZero & SET_BEF_DATE_FILE, strSetteiBefFolder, True
        End If
    Next i
    
    Set fso = Nothing
    
End Sub

