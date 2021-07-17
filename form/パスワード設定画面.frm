VERSION 5.00
Begin VB.Form frmPassSet 
   BorderStyle     =   0  'なし
   Caption         =   "パスワード設定"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ClipControls    =   0   'False
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
      Left            =   960
      Top             =   960
   End
   Begin VB.Frame fraPassWord 
      Caption         =   "パスワード(業務ユーザ)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   5055
      Index           =   2
      Left            =   8040
      TabIndex        =   24
      Top             =   1980
      Width           =   3675
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   29
         Left            =   1920
         TabIndex        =   34
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   28
         Left            =   1920
         TabIndex        =   33
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   27
         Left            =   1920
         TabIndex        =   32
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   26
         Left            =   1920
         TabIndex        =   31
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   25
         Left            =   1920
         TabIndex        =   30
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   180
         TabIndex        =   29
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   180
         TabIndex        =   28
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   180
         TabIndex        =   27
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   180
         TabIndex        =   26
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   180
         TabIndex        =   25
         Top             =   720
         Width           =   1600
      End
   End
   Begin VB.Frame fraPassWord 
      Caption         =   "パスワード(特権ユーザ)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5055
      Index           =   1
      Left            =   4080
      TabIndex        =   13
      Top             =   1980
      Width           =   3735
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   1920
         TabIndex        =   23
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   1920
         TabIndex        =   22
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   1920
         TabIndex        =   21
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   1920
         TabIndex        =   20
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   1920
         TabIndex        =   19
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   180
         TabIndex        =   18
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   180
         TabIndex        =   17
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   180
         TabIndex        =   16
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   180
         TabIndex        =   15
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1600
      End
   End
   Begin VB.Frame fraPassWord 
      Caption         =   "パスワード(一般ユーザ)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   1980
      Width           =   3675
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   1920
         TabIndex        =   12
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1920
         TabIndex        =   11
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1920
         TabIndex        =   10
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   1920
         TabIndex        =   9
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   1920
         TabIndex        =   8
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   1600
      End
   End
   Begin VB.CommandButton cmdSettei 
      Caption         =   "設  定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "メンテナンス  画面へ戻る"
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
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "パスワード設定"
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
      TabIndex        =   35
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPassSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmPassSet.frm
'//  パッケージ名：パスワード設定画面
'//
'//  概要：パスワード設定画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21  REVISED BY [TCC]  T.Koyama
'//                ＥＧ２０フェーズ２対応【残件54】
'//                ・特権ユーザ時の業務ユーザパスワード入力部削除
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private bExchanged As Boolean  '変更データ有り／無し（＝True／False）
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : パスワード設定画面(アクティブ時)
'//  機能概要  : メール受信用タイマ、起動
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
    'メイル受信用のタイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : パスワード設定画面(ディアクティブ時)
'//  機能概要  : メール受信用タイマ、停止
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
    'メイル受信用のタイマをを停止する。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : パスワード設定画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21  REVISED BY [TCC]  T.Koyama
'//                ＥＧ２０フェーズ２対応【残件54】
'//                ・特権ユーザ時の業務ユーザパスワード入力部削除
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim intPassFileNo As Integer  'パスワードファイルのファイル番号
    Dim sPassword As String       'パスワードファイルの１行分のデータ
    Dim iLineSSB As Integer       '特殊ユーザテキストボックスのINDEX
    Dim iLineTSB As Integer       '特権ユーザテキストボックスのINDEX
    Dim iLineUSR As Integer       '一般保守ユーザテキストボックスのINDEX

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

 Select Case pbUserLevel
        '一般ユーザ
        Case 0
            fraPassWord(0).Caption = "パスワード"
            fraPassWord(0).Left = 4080
            fraPassWord(0).Visible = True
            fraPassWord(1).Visible = False
            fraPassWord(2).Visible = False
        '特殊ユーザ
        Case 1, 2
'EG20 V2.0.1.1 DEL START
'            fraPassWord(0).Left = 180
'            fraPassWord(1).Left = 4080
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
            fraPassWord(0).Left = 1980
            fraPassWord(1).Left = 5880
'EG20 V2.0.1.1 ADD END
            fraPassWord(2).Left = 8040
            fraPassWord(0).Visible = True
            fraPassWord(1).Visible = True
'EG20 V2.0.1.1 DEL START
'            fraPassWord(2).Visible = True
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
            fraPassWord(2).Visible = False
'EG20 V2.0.1.1 ADD END
        Case Else
    End Select
    
    On Error GoTo FileError
    iLineUSR = 0
    iLineTSB = 10
    iLineSSB = 20
    
    '保守員パスワードファイルの先頭から１行ずつ読込み、テキストボックスに表示する。
    intPassFileNo = FreeFile        ' 未使用のファイル番号を取得する。
    Open PASSWORD_FILE_FULLPASS For Input As #intPassFileNo     ' パスワードファイルを開く。
'    Do While Not EOF(1)             ' ファイルの終端まで繰り返す。             ' EG20 V3.3.0.1削除
    Do While Not EOF(intPassFileNo)  ' ファイルの終端まで繰り返す。             ' EG20 V3.3.0.1追加
        Line Input #intPassFileNo, sPassword  ' １行分読込む。
        If sPassword <> "" Then               '文字列の記述がある。
            If Left(sPassword, 1) = "0" Then        '一般保守ユーザパスワードである。
                txtPassWord(iLineUSR) = Mid(sPassword, 3, 8)
                iLineUSR = iLineUSR + 1
            ElseIf Left(sPassword, 1) = "1" Then    '特権ユーザパスワードである。
                txtPassWord(iLineTSB) = Mid(sPassword, 3, 8)
                iLineTSB = iLineTSB + 1
            Else                                    '特殊ユーザパスワードである。
                txtPassWord(iLineSSB) = Mid(sPassword, 3, 8)
                iLineSSB = iLineSSB + 1
            End If
        End If
    Loop
    Close #intPassFileNo             ' ファイルを閉じる。
    bExchanged = False               ' 変更データ無しとしておく。

    'メイル受信用のメイル受信用のタイマ値を設定する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '「パスワード設定画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_START, 0)

    Exit Sub
FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdSettei_Click
'//  機能名称  : 「設定」釦押下時処理
'//  機能概要  : 設定されたパスワードの更新を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
 Private Sub cmdSettei_Click()
  
   Dim iResponse As Integer     'ボタンコード
   Dim iLine As Integer         'テキストボックスINDEX
   Dim iLineMax As Integer      'テキストボックスの個数
   Dim iLineTSB As Integer      '特権ユーザテキストボックスの先頭INDEX
   Dim sPassword As String      'テキストボックスの表示内容
   Dim intPassFileNo As Integer 'パスワードファイルのファイル番号
   Dim bRet As Boolean          '関数戻り値
   Dim lngErrCode As Long       'エラーコード
    
    '「パスワード設定画面：設定釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_SETTEI_BUTTOM, 0)

   '更新確認メッセージを表示する。
   iResponse = MsgBox("表示中のパスワードを、登録します。" _
                       & Chr(vbKeyReturn) & " よろしいですか？", _
                       vbYesNo + vbExclamation, _
                       "パスワードの更新")
   If iResponse = vbYes Then
   ' [はい] ボタンを選択した場合
       On Error GoTo FileError

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを表示する
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
       
       '表示中のパスワードを、行をつめて保守員パスワードファイルへ書込む。
       intPassFileNo = FreeFile        ' 未使用のファイル番号を取得する。
       Open PASSWORD_FILE_FULLPASS For Output As #intPassFileNo
       iLineMax = txtPassWord.UBound 'テキストボックスの個数
       iLineTSB = (iLineMax + 1) / 2 '特権ユーザテキストボックスの先頭INDEX
       For iLine = 0 To iLineMax
           sPassword = txtPassWord(iLine)
           If sPassword <> "" Then
               If iLine < 10 Then
                   sPassword = "0," & sPassword '一般保守ユーザ用
                ElseIf iLine < 20 Then
                    sPassword = "1," & sPassword '特権ユーザ用
                Else
                    sPassword = "2," & sPassword '特殊ユーザ用
                End If
                               
                Print #intPassFileNo, sPassword
            End If
        Next
        Close #intPassFileNo
        bExchanged = False     ' 変更データ無しに戻す。
        '「パスワード設定画面：設定更新正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_SETTEI_OK, 0)
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    Else
    ' [いいえ] ボタンを選択した場合
        '何もしない。
        Exit Sub
    End If
    Exit Sub

FileError:   'パスワードファイルアクセスエラー処理ルーチン
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    MsgBox "パスワードファイルアクセスエラー：" & _
            vbCrLf & Error(Err.Number)
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '「パスワード設定画面：設定更新異常」ログ出力
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_SET_GAMEN_SETTEI_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtPassWord_DblClick
'//  機能名称  : テキストボックス、ダブルクリック時処理
'//  機能概要  : 擬似テンキー画面を表示し、パスワード設定を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　[IN]ダブルクリックされたテキストボックスインデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtPassWord_DblClick(Index As Integer)
    
    gstrTenKeyData = txtPassWord(Index)  ' 現在の行位置の表示文字を渡す
    gstrTenKeySize = 8                   '入力可能文字数を指定する。
    ' 擬似テンキー画面を表示する。
    frmTenKey.Show 1
    If gstrTenKeyData <> txtPassWord(Index) Then
    '内容が更新されていれば、
        '設定された情報で表示更新する
        txtPassWord(Index) = gstrTenKeyData
        bExchanged = True    '変更データ有りとする。
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtPassWord_KeyPress
'//  機能名称  : テキストボックス、キー入力処理
'//  機能概要  : データ変更を記録する。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　[IN]
'//  　　      : Integer　KeyAscii [IN]
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtPassWord_KeyPress(Index As Integer, KeyAscii As Integer)
    bExchanged = True '変更データ有りとする。
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メンテナンス画面へ戻る」釦押下時処理
'//  機能概要  : 設定されたパスワードの更新有無と、自画面を消去する。
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
      
    Dim iResponse As Integer     'ボタンコード
   
    On Error Resume Next

    If bExchanged = True Then
    '画面表示中の変更が登録 されていないとき、確認メッセージを表示する。
     iResponse = MsgBox("画面表示中に設定されたデータが失われます。" _
                        & Chr(vbKeyReturn) & "よろしいですか？", _
                        vbYesNo + vbExclamation, _
                        "設定データのキャンセル確認")
       If iResponse = vbYes Then
         ' [はい] ボタンを選択した場合、
         'パスワード設定画面を閉じる。
         '「パスワード設定画面：消去」ログ出力
          Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_END, 0)
          Unload Me
       Else
       ' [いいえ] ボタンを選択した場合、
       '何もしない。
         '「パスワード設定画面：消去」ログ出力
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_END, 0)
         Exit Sub
       End If
    Else
    'それ以外は、
        'パスワード設定画面を閉じる。
        '「パスワード設定画面：消去」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_END, 0)
        Unload Me
    End If
    Unload Me
End Sub

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
        AppActivate frmPassSet.Caption, False
        pfFormActive (frmPassSet.hwnd)
    End If
End Sub

