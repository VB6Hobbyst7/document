VERSION 5.00
Begin VB.Form frmRenewData 
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
   Begin VB.CommandButton cmdOutput 
      Caption         =   $"係員設定保存復元画面.frx":0000
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
      TabIndex        =   19
      Top             =   3360
      Width           =   2175
   End
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
      TabIndex        =   4
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  機器情報設定    画面へ戻る"
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
      Left            =   9500
      TabIndex        =   3
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8040
   End
   Begin VB.CommandButton cmdRenew 
      Caption         =   "復元"
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
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
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
      Caption         =   "係員設定 保存／復元"
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
Attribute VB_Name = "frmRenewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：係員設定保存復元画面.frm
'//  パッケージ名：係員設定保存復元画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//                 ・阪急保守より、係員設定保存復元(frmRenewData.frm)を流用
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
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdOutput_Click
'//  機能名称  : 保存データ媒体出力釦押下時処理
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 5.4.0.1) 2012-03-25  REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()

    Dim iResponse       As Integer      'MsgBox戻り値
    Dim blnChecked      As Boolean      '対象コーナチェック有無
    Dim blnExistFile    As Boolean      '保存ファイル有無チェック
    Dim intCount        As Integer      'ループカウンタ
    
    On Error Resume Next

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_RENEW, 0)
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    '処理対象選択有無チェック
    blnChecked = False
    blnExistFile = True
    Erase glngTergetCorner
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnChecked = True
            glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
            If lblSetteDate(intCount).Caption = "    年   月   日   時   分   秒" Then
                blnExistFile = False
            End If
        End If
    Next intCount
    '選択コーナなしの場合、メッセージを表示して終了
    If blnChecked = False Then
' EG20 5.4.0.1 削除開始
'        iResponse = MsgBox("対象コーナが選択されていません。" & vbCrLf & _
'                            "選択してください。", vbOKOnly + vbExclamation, "コーナ未選択")
' EG20 5.4.0.1 削除終了
' EG20 5.4.0.1 追加開始
        iResponse = MsgBox("対象コーナが選択されていません。" & vbCrLf & _
                            "選択してください。", vbOKOnly + vbCritical, "コーナ未選択")
' EG20 5.4.0.1 追加終了
        Exit Sub
    End If
    '保存ファイル存在なしの場合、メッセージを表示して終了
    If blnExistFile = False Then
' EG20 5.4.0.1 削除開始
'        iResponse = MsgBox("媒体出力に使用するデータがありません。", vbOKOnly + vbExclamation, "データなし")
' EG20 5.4.0.1 削除終了
' EG20 5.4.0.1 追加開始
        iResponse = MsgBox("媒体出力に使用するデータがありません。", vbOKOnly + vbCritical, "データなし")
' EG20 5.4.0.1 追加終了
        Exit Sub
    End If
    'EG20 V2.1.0.1 ADD END
    
    iResponse = MsgBox("保存データを外部媒体に出力します。" & vbCrLf & _
                        "実行してもよろしいですか？", vbOKCancel + vbExclamation, "実行確認")

    '「キャンセル」ボタン押下処理は処理を終了する
    If iResponse = vbCancel Then Exit Sub

    frmRenewOutput.Show vbModal
    
End Sub

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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    glbSaveFoldePath = ""
    strPath = ""
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    
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
    'EG20 V2.1.0.1 ADD END
    
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdKeep_Click()
    
    Dim iResponse       As Integer      'MsgBox戻り値
    Dim strMessage      As String       'MsgBox文言
    Dim bRet            As Boolean      '戻り値
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim intCount        As Integer
    Dim udtSendData     As MAIL_KAKARIIN_SETTEI
    Dim lngRet          As Long
    Dim blnIsSelected   As Boolean
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SAVE, 0)
    
    '監視設定データ保存ファイル、自改設定データ保存ファイルの存在チェック
'    If Dir(glbSaveFoldePath & H_K_SETTEI_FILE) <> "" Or Dir(glbSaveFoldePath & H_G_SETTEI_FILE) <> "" Then 'EG20 V2.1.0.1 DEL 【Mainte_03_01】

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Erase glngTergetCorner
    blnIsSelected = False
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnIsSelected = True
            If RenewChk(intCount).Visible = True Then
                glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
            End If
        End If
    Next intCount
    '選択コーナなしの場合、メッセージを表示して終了
    If blnIsSelected = False Then
' EG20 5.4.0.1 削除開始
'        iResponse = MsgBox("対象コーナが選択されていません。" & vbCrLf & _
'                            "選択してください。", vbOKOnly + vbExclamation, "コーナ未選択")
' EG20 5.4.0.1 削除終了
' EG20 5.4.0.1 追加開始
        iResponse = MsgBox("対象コーナが選択されていません。" & vbCrLf & _
                            "選択してください。", vbOKOnly + vbCritical, "コーナ未選択")
' EG20 5.4.0.1 追加終了
        Exit Sub
    End If
    
    For intCount = 0 To lblSetteDate.UBound
        If lblSetteDate(intCount).Visible = True And RenewChk(intCount).Value = 1 And _
           lblSetteDate(intCount).Caption <> "    年   月   日   時   分   秒" Then
    'EG20 V2.1.0.1 ADD END
        
            '確認メッセージボックスを表示する。
            iResponse = MsgBox("設定ファイルを上書きしますがよろしいですか？", _
                                vbOKCancel + vbCritical, "上書き保存警告")
    
            '「キャンセル」ボタン押下処理は処理を終了する
            If iResponse = vbCancel Then Exit Sub
            Exit For        'EG20 V2.1.0.1 ADD 【Mainte_03_01】
        End If
    Next intCount           'EG20 V2.1.0.1 ADD 【Mainte_03_01】
    
    '初期設定
    strMessage = MSG_NORMAL '正常
    bRet = False            '異常

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    frmRenewSave.Show vbModal
    'EG20 V2.1.0.1 ADD END
    
    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
    '監視設定データファイルを監視設定データ保存ファイルとしてコピー
'    bRet = fCopySetteiFile(K_SETTEI_FILE, glbSaveFoldePath & H_K_SETTEI_FILE, MU_KSETTEI)
'    If bRet = False Then GoTo ErrorHandler          '異常の場合、コピー処理を終了
'
'    '自改設定データファイルを自改設定データ保存ファイルとしてコピー
'    bRet = fCopySetteiFile(G_SETTEI_FILE, glbSaveFoldePath & H_G_SETTEI_FILE, MU_GSETTEI)
'    If bRet = False Then GoTo ErrorHandler          '異常の場合、コピー処理を終了
'
'    '機器情報自動改札機エリア保存ファイル作成処理
'    bRet = fKeepGateIniInf

'ErrorHandler:
'
'    '処理は正常に終了したか？
'    If bRet = False Then
'        '異常処理
'        Call fDeleteKeepFile        '保存ファイルを削除
'        strMessage = MSG_ERROR      '異常文言を設定
'    End If
'
'    '処理結果メッセージボックスを表示する。
'    iResponse = MsgBox("    " & strMessage & "終了しました。    ", vbOKOnly, "保存処理結果")

    'EG20 V2.1.0.1 DEL END
    
    '画面表示を更新
    Call sFromInitialize
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdRenew_Click
'//  機能名称  : 「復元」釦押下時処理
'//  機能概要  : 機器情報自動改札機エリアと機器情報自動改札機エリア保存ファイルを比較
'//              自改・監視設定データ保存ファイルを自改・監視設定ファイルに更新する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 5.4.0.1) 2012-03-25  REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdRenew_Click()
    Dim iResponse       As Integer      'MsgBox戻り値
    Dim lngChgFlg       As Long         '比較処理結果
    Dim bRet            As Boolean      'リカバリ処理結果
    Dim lngRet          As Long         'メール送信処理結果
    Dim strKSetMessage  As String       '監視設定ファイル更新処理結果
    Dim strGSetMessage  As String       '自改設定ファイル更新処理結果
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim blnChecked      As Boolean      '対象コーナチェック有無
    Dim blnExistFile    As Boolean      '保存ファイル有無チェック
    Dim intCount        As Integer      'ループカウンタ
    Dim udtSendData     As MAIL_KAKARIIN_SETTEI
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

    '画面操作ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_RENEW, 0)
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    '処理対象選択有無チェック
    blnChecked = False
    blnExistFile = True
    Erase glngTergetCorner
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnChecked = True
            glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
            If lblSetteDate(intCount).Caption = "    年   月   日   時   分   秒" Then
                blnExistFile = False
            End If
        End If
    Next intCount
    '選択コーナなしの場合、メッセージを表示して終了
    If blnChecked = False Then
' EG20 5.4.0.1 削除開始
'        iResponse = MsgBox("対象コーナが選択されていません。" & vbCrLf & _
'                            "選択してください。", vbOKOnly + vbExclamation, "コーナ未選択")
' EG20 5.4.0.1 削除終了
' EG20 5.4.0.1 追加開始
        iResponse = MsgBox("対象コーナが選択されていません。" & vbCrLf & _
                            "選択してください。", vbOKOnly + vbCritical, "コーナ未選択")
' EG20 5.4.0.1 追加終了
        Exit Sub
    End If
    '保存ファイル存在なしの場合、メッセージを表示して終了
    If blnExistFile = False Then
' EG20 5.4.0.1 削除開始
'        iResponse = MsgBox("復元に使用するデータがありません。", vbOKOnly + vbExclamation, "データなし")
' EG20 5.4.0.1 削除終了
' EG20 5.4.0.1 追加開始
        iResponse = MsgBox("復元に使用するデータがありません。", vbOKOnly + vbCritical, "データなし")
' EG20 5.4.0.1 追加終了
        Exit Sub
    End If
    'EG20 V2.1.0.1 ADD END
    
    iResponse = MsgBox("監視盤に保存済みの設定値を反映させます。" & vbCrLf & _
                        "実行してもよろしいですか？", vbOKCancel + vbExclamation, "実行確認")

    '「キャンセル」ボタン押下処理は処理を終了する
    If iResponse = vbCancel Then Exit Sub

    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'    strKSetMessage = MSG_ERROR  '更新処理異常終了
'    strGSetMessage = MSG_ERROR  '更新処理異常終了
'
'    On Error GoTo ErrorHandler
'
'    lngChgFlg = RET_ERROR                           '初期設定（-1：異常）
'
'    lngChgFlg = fCompareGateIniInf                              '機器情報自動改札機エリア比較処理
'    If lngChgFlg <> RET_NASI Then GoTo ErrorHandler             '変更が無しの場合、更新処理を行う
'
'    bRet = False                                    '初期設定（False：異常）
'    lngRet = INVALID_HANDLE_VALUE                   '初期設定（-1：異常）
'
'    bRet = dllK_Settei_File_Recovery                            '監視装置設定データファイルリカバリ処理
'    If bRet = False Then GoTo ErrorHandler                      '異常の場合、更新処理を終了
'
'    lngRet = fKansiSetteiMailSend                               '監視設定指示メール送信
'    If lngRet = INVALID_HANDLE_VALUE Then GoTo ErrorHandler     '異常の場合、更新処理を終了
'
'    '*****監視装置設定データ更新処理正常終了
'    strKSetMessage = MSG_NORMAL                     '正常終了のメッセージを設定

'    bRet = dllG_Settei_File_Recovery                            '自改設定データファイルリカバリ処理
'    If bRet = False Then GoTo ErrorHandler                      '異常の場合、更新処理を終了

'    lngRet = fGateSetteiMailSend                                '係員設定保存要求メール送信
    'EG20 V2.1.0.1 DEL END
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    frmRenewCyu.Show vbModal
    'EG20 V2.1.0.1 ADD END
    
    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'    If lngRet = INVALID_HANDLE_VALUE Then GoTo ErrorHandler     '異常の場合、更新処理を終了
'
'    '*****自改設定データ更新処理正常終了
'    strGSetMessage = MSG_NORMAL                     '正常終了のメッセージを設定
'
'ErrorHandler:
'
'    '各保存ファイルを削除
'    Call fDeleteKeepFile
'
'    '処理結果メッセージを表示
'    If lngChgFlg = RET_ARI Then     '変更有りの場合
'        '自改構成変更有りメッセージボックスを表示する。
'        iResponse = MsgBox("自改構成が変更されたため" & vbCrLf & _
'                            "更新処理はできません。", vbOKOnly + vbExclamation, "反映処理結果")
'    Else
'        '処理結果メッセージボックスを表示する。
'        iResponse = MsgBox("    自改設定の更新は" & strGSetMessage & "終了しました。    " & vbCrLf & _
'                           "    監視設定の更新は" & strKSetMessage & "終了しました。    ", vbOKOnly, "反映処理結果")
'    End If
    'EG20 V2.1.0.1 DEL END

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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_END, 0)
    
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
    
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    Dim blnExistFile            As Boolean          '保存ファイル有無
    Dim strSaveFile             As String
    Dim intIndex                As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

                
                
    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'    '監視設定データ保存ファイル、自改設定データ保存ファイル、機器情報自動改札機エリア保存ファイルの存在チェック
'    '一つでも存在しない場合、ブランクを表示
'    If Dir(glbSaveFoldePath & H_K_SETTEI_FILE) = "" Or _
'       Dir(glbSaveFoldePath & H_G_SETTEI_FILE) = "" Or _
'       Dir(glbSaveFoldePath & H_G_INFO_FILE) = "" Then GoTo ErrorHandler
'
'    strFileName(K_SETTEI) = glbSaveFoldePath & H_K_SETTEI_FILE     '監視設定データファイル
'    strFileName(G_SETTEI) = glbSaveFoldePath & H_G_SETTEI_FILE     '自改設定データファイル
'
'    'ファイルの作成日時を取得し、作成日付表示部に表示
'    For intCnt = 0 To lblSetteDate.ubound
'
'        'ファイルをオープン
'        lngHandle = CreateFile(strFileName(intCnt), GENERIC_READ, FILE_SHARE_READ, _
'                                0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    blnExistFile = False
    intIndex = 0
    For intCnt = 0 To UBound(gudtSettiCorner)
        If gblnCornerSet(intCnt) = True Then
            '保存ファイルの日付を取得
            strSaveFile = PATH_OPERATE & "CORNER" & CStr(intCnt + 1) & "\\SETTEI\\" & CONDENSE_FILE
            If Dir(strSaveFile) = "" Then
                lblSetteDate(intIndex).Caption = "    年   月   日   時   分   秒"
            Else
                'ファイルをオープン
                lngHandle = CreateFile(strSaveFile, GENERIC_READ, FILE_SHARE_READ, _
                                        0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    'EG20 V2.1.0.1 ADD END

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
    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
                blnExistFile = True
            End If
            
            lblSetteDate(intIndex).Visible = True
            intIndex = intIndex + 1
        Else
            lblSetteDate(intIndex).Visible = False
        End If
    'EG20 V2.1.0.1 ADD END
    Next

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    '保存ファイルが１つもない場合は復元ボタン、媒体出力押下不可
    If blnExistFile = False Then
        cmdRenew.Enabled = False
        cmdOutput.Enabled = False
    Else
        cmdOutput.Enabled = True
    'EG20 V2.1.0.1 ADD END
        cmdRenew.Enabled = True     '設定値反映ボタン押下可
    End If          'EG20 V2.1.0.1 ADD 【Mainte_03_01】
        
    Exit Sub

APIError:

    Call CloseHandle(lngHandle)             'ハンドルのクローズ

ErrorHandler:

    '存在しない場合またはエラーが発生した、ブランクを表示
    'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'    lblSetteDate(K_SETTEI).Caption = "    年   月   日   時   分   秒"
'    lblSetteDate(G_SETTEI).Caption = "    年   月   日   時   分   秒"
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    For intCnt = intCnt To UBound(gudtSettiCorner)
        If lblSetteDate(intCnt).Visible = True Then
            lblSetteDate(intCnt).Caption = "    年   月   日   時   分   秒"
        End If
    Next intCnt
    'EG20 V2.1.0.1 ADD END
    
    cmdRenew.Enabled = False    '設定値反映ボタン押下不可
    cmdOutput.Enabled = False   '媒体出力ボタン押下不可     'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fCopySetteiFile
'//  機能名称  : 設定ファイルコピー処理
'//  機能概要  : 設定ファイルを指定フォルダへコピーする
'//　　　　　　　ファイルの作成日時、アクセス日時、更新日時を設定する
'//
'//              型        名称     　　　意味
'//  引数      : String　　strFromFile　　[IN]コピー元ファイル名
'//  　　　　　　String　　strToFile　　　[IN]コピー先ファイル名
'//  　　　　　　String　　strMutexName 　[IN]ミューテックス名
'//
'//              型        値        　　 意味
'//  戻り値    : Boolean　 True           正常終了
'//                     　 False          異常終了
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fCopySetteiFile(strFromFile As String, strToFile As String, strMutexName As String) As Boolean
    Dim lngHandle               As Long            'ハンドル
    Dim lngMuHandle             As Long            '排他処理用ハンドル

    Dim lpCreatTime             As FILETIME        '作成日時
    Dim lpAccessTime            As FILETIME        '最終アクセス日時
    Dim lpLastwTime             As FILETIME        '更新日時
    Dim lpLocalTime             As FILETIME        'ローカル日時
    Dim lpSystemTime            As SYSTEMTIME      'システム時刻
    Dim bRet                    As Boolean         '戻り値

    On Error Resume Next

    '初期設定
    fCopySetteiFile = False
    bRet = False

    '設定データファイルの有無チェック
    If Dir(strFromFile) = "" Then Exit Function     '存在しない場合、処理を終了

    lngMuHandle = dllOpenMutex(strMutexName)        '排他処理(OPEN)

    If lngMuHandle <> 0 Then
        dllWaitForSingleObject (lngMuHandle)        '排他処理(GET)
    End If

    bRet = CopyFile(strFromFile, strToFile, False)  '設定ファイルコピー処理
    If bRet = False Then GoTo ErrorHandler          'コピー処理は正常に行われたか？

    bRet = False    '再設定

    'ファイルの作成日時、アクセス日時、更新日時を設定
    '設定ファイルをオープン
    lngHandle = CreateFile(strToFile, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                            0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then GoTo ErrorHandler

    'ローカル時刻を取得
    Call GetLocalTime(lpSystemTime)

    'システムタイムをローカルタイムに変換
    bRet = SystemTimeToFileTime(lpSystemTime, lpLocalTime)
    If bRet = False Then GoTo APIError                          '変換が正常に行われたか？

    'ローカルタイムをファイルタイムに変換（作成日時）
    bRet = LocalFileTimeToFileTime(lpLocalTime, lpCreatTime)
    If bRet = False Then GoTo APIError                          '変換が正常に行われたか？

    'ローカルタイムをファイルタイムに変換（アクセス日時）
    bRet = LocalFileTimeToFileTime(lpLocalTime, lpAccessTime)
    If bRet = False Then GoTo APIError                          '変換が正常に行われたか？

    'ローカルタイムをファイルタイムに変換（更新日時）
    bRet = LocalFileTimeToFileTime(lpLocalTime, lpLastwTime)
    If bRet = False Then GoTo APIError                          '変換が正常に行われたか？

    'ファイルの日付を設定
    bRet = SetFileTime(lngHandle, lpCreatTime, lpAccessTime, lpLastwTime)

APIError:

    Call CloseHandle(lngHandle)                     'ハンドルのクローズ

ErrorHandler:

    If lngMuHandle <> 0 Then
        dllReleaseMutex (lngMuHandle)                   '排他処理(FREE)
        dllCloseHandle (lngMuHandle)                    '排他処理(CLOSE)
    End If

    'コピー処理は正常に行われたか？
    If bRet = False Then Exit Function              '異常の場合、処理を終了

    '正常終了
    fCopySetteiFile = True
End Function

'EG20 V2.1.0.1 DEL START 【Mainte_03_01】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fCopySetteiFile
'//  機能名称  : 機器情報自動改札機エリア取得
'//  機能概要  : 機器情報自動改札機エリア情報を取得し、
'//　　　　　　　機器情報自動改札機エリア保存ファイルを生成する
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : Boolean　 True           正常終了
'//                     　 False          異常終了
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Function fKeepGateIniInf() As Boolean
'
'    Dim bRet            As Boolean          '書込み結果
'
'    On Error Resume Next
'
'    '初期設定
'    fKeepGateIniInf = False
'
'    '機器情報自動改札機エリア読み込み処理
'    sGetGateIniInf
'
'    '機器情報自動改札機エリア保存ファイル書き込み処理
'    bRet = fWriteGateIniInf
'
'    '結果設定
'    fKeepGateIniInf = bRet
'
'End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fCompareGateIniInf
'//  機能名称  : 機器情報自動改札機エリア比較処理
'//  機能概要  : 機器情報自動改札機エリアと
'//　　　　　　　機器情報自動改札機エリア保存ファイルを比較する
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : Long　    -1             異常
'//                     　 0              変更無し
'//                     　 1              変更有り
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Function fCompareGateIniInf() As Long
'    Dim bRet            As Boolean      '読み込み処理結果
'    Dim bChgFlg         As Boolean      '変更有り
'    Dim intCnt          As Integer      'カウンタ
'    Dim lngHandle       As Long         'ハンドル
'
'    On Error GoTo ErrorHandler
'
'    '初期設定
'    fCompareGateIniInf = RET_ERROR  '異常
'
'    '機器情報自動改札機エリア保存ファイル読み込み処理
'    bRet = fReadGateIniInf
'    '保存ファイル読み込みが正常に行われたか？
'    If bRet = False Then Exit Function
'
'    '機器情報自動改札機エリア読み込み処理
'    Call sGetGateIniInf
'
'    '初期設定
'    bChgFlg = True      '変更無し
'
'    '号機数分、機器情報自動改札機エリアと機器情報自動改札機エリア保存ファイルとの比較
'    For intCnt = 0 To MAX_GATE_NO - 1
'        'NEG型/C型・集札/改札/両用
'        If udtIniGate.Gate_Set(intCnt).nGate <> udtIniGateFile.Gate_Set(intCnt).nGate Or _
'            udtIniGate.Gate_Set(intCnt).nTuuro <> udtIniGateFile.Gate_Set(intCnt).nTuuro Then
'            '自改構成変更有り
'            bChgFlg = False
'            Exit For
'        End If
'    Next
'
'    '自改構成に変更が無かったか？
'    If bChgFlg = True Then
'        '変更がなかった場合、変更無しを返す
'        fCompareGateIniInf = RET_NASI
'    Else
'        '変更があった場合、変更有りを返す
'        fCompareGateIniInf = RET_ARI
'    End If
'
'    Exit Function
'ErrorHandler:
'    '異常処理
'End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sGetGateIniInf
'//  機能名称  : 機器情報自動改札機エリア読み込み処理
'//  機能概要  : 機器情報自動改札機エリア情報を取得する
'//
'//              型         名称     　　　 意味
'//  引数      : なし
'//
'//              型         値        　　  意味
'//  戻り値    : Boolean    TRUE            正常終了
'//                         FALSE           異常終了
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub sGetGateIniInf()
'
'    Dim udtMapInf       As MAP_MEM          'メモリマッピングオブジェクト
'
'    Dim bRet            As Boolean          '書込み結果
'    Dim strIniData      As String * 1024    'INI設定値
'    Dim strKeyName      As String           'キー名
'    Dim strMutexName    As String           'ミューテックス名
'    Dim lSts            As Long             '関数戻り値
'    Dim lErrCode        As Long             'エラーコード
'    Dim lngMuHandle     As Long             '排他処理用ハンドル
'    Dim iLoopCnt        As Integer          'ループカウンタ
'
'    On Error Resume Next
'
'    strMutexName = "Mu_" & GIniGate
'    lngMuHandle = dllOpenMutex(strMutexName)         '排他処理(OPEN)
'    If lngMuHandle <> 0 Then
'
'        dllCloseHandle (lngMuHandle)                 '排他処理(CLOSE)
'
'        '機器情報定義エリアの内容を取得する。
'        'エリアの初期化
'        Call dllMemMappingInit(GIniGate, 0, MUTEXMODE_ON, udtMapInf)
'
'        'エリアの内容を取得する。
'        Call dllMemMappingRead(udtMapInf.lngpAdr, LenB(udtIniGate), MUTEXMODE_ON, udtMapInf.lnghMutex, udtIniGate)
'
'        'エリアを解放する
'        Call dllMemMappingEnd(udtMapInf, udtMapInf.lnghMutex)
'
'    Else
'
'        'エリアが存在しない場合INIファイルから取得する
'        For iLoopCnt = 0 To MAX_GATE_NO - 1
'            strIniData = ""
'            strKeyName = INI_GATE_KEY & Format(iLoopCnt + 1, "00")
'            lSts = GetPrivateProfileString(INI_GATE_SECTION, _
'                                           strKeyName, _
'                                           "Defo", _
'                                           strIniData, _
'                                           Len(strIniData), _
'                                           PATH_GATE_FILE)
'            If lSts > 0 Then
'                bRet = dllMemIniGate(udtIniGate.Gate_Set(iLoopCnt), strIniData, lErrCode)
'                If (bRet = False) Then
'                    '異常ログ出力
'                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, KAKARISET_GAMEN_GET_GATE_AREA_ERROR, lErrCode)
'                End If
'            Else
'                Exit Sub
'            End If
'        Next
'
'    End If
'
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : fWriteGateIniInf
'//  機能名称  : 機器情報自動改札機エリア保存ファイル書き込み処理
'//  機能概要  : 機器情報自動改札機エリア保存ファイルを生成する
'//
'//              型        名称     　　　意味
'//  引数      : なし
'//
'//              型        値        　　 意味
'//  戻り値    : Boolean　 True           正常終了
'//                     　 False          異常終了
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Function fWriteGateIniInf() As Boolean
'    Dim lngHandle       As Long         'ハンドル
'    Dim lngRet          As Long         '書き込まれたバイト数のアドレス
'    Dim bRet            As Boolean      '書き込み結果
'
'    On Error GoTo ErrorHandler
'
'    '初期設定
'    fWriteGateIniInf = False
'
'    'ファイルを作成
'    lngHandle = CreateFile(glbSaveFoldePath & H_G_INFO_FILE, GENERIC_WRITE, FILE_SHARE_WRITE Or FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
'
'    'ファイル作成が正常に行われたか？
'    If lngHandle = INVALID_HANDLE_VALUE Then Exit Function
'
'    'ファイルの書き込み
'    bRet = WriteFile(lngHandle, udtIniGate, LenB(udtIniGate), lngRet, 0)
'
'    'ハンドルのクローズ
'    Call CloseHandle(lngHandle)
'
'    '書き込み結果設定
'    fWriteGateIniInf = bRet
'
'    Exit Function
'
'ErrorHandler:
'    '異常処理
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : fReadGateIniInf
''//  機能名称  : 機器情報自動改札機エリア保存ファイル読み込み処理
''//  機能概要  : 機器情報自動改札機エリア保存ファイルを読み込む
''//
''//              型        名称     　　　意味
''//  引数      : なし
''//
''//              型        値        　　 意味
''//  戻り値    : Boolean　 True           正常終了
''//                     　 False          異常終了
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Function fReadGateIniInf() As Boolean
'    Dim lngHandle       As Long         'ハンドル
'    Dim lngRet          As Long         '読み込まれたバイト数のアドレス
'    Dim bRet            As Boolean      '読み込み結果
'
'    On Error GoTo ErrorHandler
'
'    '初期設定
'    fReadGateIniInf = False
'
'    'ファイルをオープン
'    lngHandle = CreateFile(glbSaveFoldePath & H_G_INFO_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
'
'    'ファイルオープンが正常に行われたか？
'    If lngHandle = INVALID_HANDLE_VALUE Then Exit Function
'
'    '機器情報自動改札機エリア保存ファイル読み込み処理
'    bRet = ReadFile(lngHandle, udtIniGateFile, LenB(udtIniGateFile), lngRet, 0)
'
'    'ハンドルのクローズ
'    Call CloseHandle(lngHandle)
'
'    '読み込み結果設定
'    fReadGateIniInf = bRet
'
'    Exit Function
'
'ErrorHandler:
'    '異常処理
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : fDeleteKeepFile
''//  機能名称  : 各保存ファイルの削除処理
''//  機能概要  : 監視設定データ保存ファイル、自改設定データ保存ファイル、
''//　　　　　　　機器情報自動改札機エリア保存ファイルが存在していた場合、削除する
''//
''//              型        名称     　　　意味
''//  引数      : なし
''//
''//              型        値        　　 意味
''//  戻り値    : Boolean　 True           正常終了
''//                     　 False          異常終了
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Function fDeleteKeepFile() As Boolean
'    Dim bRet    As Boolean      '処理結果
'    Dim lngRet  As Long
'
'    On Error GoTo ErrorHandler
'
'    '初期処理
'    fDeleteKeepFile = False
'
'    If Dir(glbSaveFoldePath & H_K_SETTEI_FILE) <> "" Then          '監視設定データ保存ファイル
'        bRet = DeleteFile(glbSaveFoldePath & H_K_SETTEI_FILE)      'ファイル削除処理
'    End If
'
'    If Dir(glbSaveFoldePath & H_G_SETTEI_FILE) <> "" Then          '自改設定データ保存ファイル
'        bRet = DeleteFile(glbSaveFoldePath & H_G_SETTEI_FILE)      'ファイル削除処理
'    End If
'
'    If Dir(glbSaveFoldePath & H_G_INFO_FILE) <> "" Then            '機器情報自動改札機エリア保存ファイル
'        bRet = DeleteFile(glbSaveFoldePath & H_G_INFO_FILE)        'ファイル削除処理
'    End If
'
'    '正常終了
'    fDeleteKeepFile = True
'
'    Exit Function
'
'ErrorHandler:
'    '異常処理
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : fKansiSetteiMailSend
''//  機能名称  : 自改定指示メール送信処理（監視設定）
''//  機能概要  : 監マプロセスへ「自改設定指示」メールを送信する
''//
''//              型        名称     　　　意味
''//  引数      : なし
''//
''//              型        値        　　 意味
''//  戻り値    : Long　 　 サイズ         メール送信サイズ
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Function fKansiSetteiMailSend() As Long
'    Dim lngMSlot_KM As Long                 '監マのメールスロットハンドル
'    Dim udtMail     As MAIL_GATE_SET_ORD    '自改設定指示メール送信エリア
'    Dim lngRet      As Long                 '関数戻り値
'    Dim intCnt      As Integer              'カウンタ
'
'    On Error Resume Next
'
'    '初期設定
'    fKansiSetteiMailSend = INVALID_HANDLE_VALUE
'
'    '共通ヘッダ編集
'    udtMail.mlHeader.dwId = ML_ID_GATE_SET_ORD
'    udtMail.mlHeader.dwSize = Len(udtMail)
'    udtMail.mlHeader.dwProid = RHOSHU_ID
'    udtMail.mlHeader.dwSubArea = 0
'
'    'エリア種別を設定
'    udtMail.dwCmnFile = K_SETTEI_FILE_NO
'
'    '設定情報
'    udtMail.dwGateSet(0) = 1
'    For intCnt = 1 To MAX_GATE_NO - 1
'        udtMail.dwGateSet(intCnt) = 0
'    Next intCnt
'
'    'メール送信
'    lngRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.GATE_SET_ORD, udtMail.mlHeader)
'
'    'メッセージ送信正常
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SENDMAIL, 0)
'
'    '処理結果を返す
'    fKansiSetteiMailSend = 1
'
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  関数名称  : fGateSetteiMailSend
''//  機能名称  : 自改定指示メール送信処理（自改設定）
''//  機能概要  : 監マプロセスへ「自改設定指示」メールを送信する
''//
''//              型        名称     　　　意味
''//  引数      : なし
''//
''//              型        値        　　 意味
''//  戻り値    : Long　 　 サイズ         メール送信サイズ
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  備考：
''///////////////////////////////////////////////////////////////////
'Private Function fGateSetteiMailSend() As Long
'    Dim lngMSlot_KM As Long                 '監マのメールスロットハンドル
'    Dim udtMail     As MAIL_GATE_SET_ORD    '自改設定指示メール送信エリア
'    Dim lngRet      As Long                 '関数戻り値
'    Dim intCnt      As Integer              'カウンタ
'
'    On Error Resume Next
'
'    '初期設定
'    fGateSetteiMailSend = INVALID_HANDLE_VALUE
'
'    '共通ヘッダ編集
'    udtMail.mlHeader.dwId = ML_ID_GATE_SET_ORD
'    udtMail.mlHeader.dwSize = Len(udtMail)
'    udtMail.mlHeader.dwProid = RHOSHU_ID
'    udtMail.mlHeader.dwSubArea = 0
'
'    'エリア種別を設定
'    udtMail.dwCmnFile = G_SETTEI_FILE_NO
'
'    '設定情報
'    For intCnt = 0 To MAX_GATE_NO - 1
'        If Not udtIniGate.Gate_Set(intCnt).intGate = GATE_NASI Then
'            udtMail.dwGateSet(intCnt) = 1
'        Else
'            udtMail.dwGateSet(intCnt) = 0
'        End If
'    Next intCnt
'
'    'メール送信
'    lngRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.GATE_SET_ORD, udtMail.mlHeader)
'
'    'メッセージ送信正常
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SENDMAIL, 0)
'
'    '処理結果を返す
'    fGateSetteiMailSend = 1
'End Function
'EG20 V2.1.0.1 DEL END

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
        AppActivate frmRenewData.Caption, False
        pfFormActive (frmRenewData.hwnd)
    End If
    
End Sub

'EG20 V2.1.0.1 ADD START 【Mainte_03_01】
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sRcv_Renew
'//  機能名称  : 係員設定復元要求RES受信処理
'//  機能概要  : 係員設定復元要求RES受信時の処理を行う
'//
'//              型           名称     　　　意味
'//  引数      : ML_KYOTU_INF udtReadMail    受信データ
'//
'//              型        値        　　 意味
'//  戻り値    : Long　 　 サイズ         メール送信サイズ
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-13   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sRcv_Renew(ByRef udtReadMail As ML_KYOTU_INF, ByVal strMsgTitle)

    Dim intCounta As Integer
    Dim blnIsErr As Boolean
    Dim intCount As Integer
    Dim iResponse As Integer
    
    On Error Resume Next
    
    blnIsErr = False
    '処理結果判定
    For intCount = 0 To lblSetteDate.UBound
        If udtReadMail.lngData(intCount) > 0 Then
            blnIsErr = True
            Exit For
        End If
    Next intCount
    
    'ファイル作成日時を更新する
    Call sFromInitialize
    
    '処理結果表示
    If blnIsErr = True Then
        iResponse = MsgBox("異常終了しました。", vbOKOnly, strMsgTitle)
    Else
        iResponse = MsgBox("正常終了しました。", vbOKOnly, strMsgTitle)
    End If
        
End Sub
'EG20 V2.1.0.1 ADD END



