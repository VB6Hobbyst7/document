VERSION 5.00
Begin VB.Form frmTenKey 
   Caption         =   "擬似テンキー画面"
   ClientHeight    =   4665
   ClientLeft      =   7845
   ClientTop       =   1845
   ClientWidth     =   2925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4665
   ScaleWidth      =   2925
   Begin VB.TextBox txtKeyIn 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      MaxLength       =   8
      TabIndex        =   14
      Text            =   "12345678"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "０"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   1080
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   1920
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "BS"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   1920
      TabIndex        =   2
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdSet 
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
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "取消"
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
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   360
      Top             =   600
   End
End
Attribute VB_Name = "frmTenKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmTenKey.frm
'//  パッケージ名：擬似テンキー画面
'//
'//  概要：擬似テンキー画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 擬似テンキー画面(アクティブ時)
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
    'メール受信用タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 擬似テンキー画面(ディアクティブ時)
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
    'メール受信用タイマを停止する。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 擬似テンキー画面(ロード)
'//  機能概要  : 初期処理を行う。
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
Private Sub Form_Load()
On Error Resume Next
    '表示位置を設定する
    Me.Move Screen.Width - Me.Width, 0
    'メイル受信用タイマのインターバル値の設定
    tmrMail.Interval = 1000
    '表示元画面から渡された文字列を表示する。
    txtKeyIn.MaxLength = gstrTenKeySize '入力可能文字数を指定値にする。
    txtKeyIn = gstrTenKeyData

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdNumber_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称処理を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdNumber_Click(Index As Integer)
On Error Resume Next

    If Index = 10 Then
    'ＢＳキーが押下されれば、
        '最下桁の１文字を削除する。（表示文字のある場合のみ）
        If (txtKeyIn <> "") Then
            txtKeyIn = Left(txtKeyIn, Len(txtKeyIn) - 1)
        End If
        Exit Sub
    End If
    If Index = 11 Then
    'Ｃキーが押下されれば、
        '表示文字を全て削除する。
        txtKeyIn = ""
        Exit Sub
    End If
    txtKeyIn = txtKeyIn & Chr(vbKey0 + Index)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdSet_Click
'//  機能名称  : 各釦押下時処理
'//  機能概要  : 各釦名称処理を行う。
'//　　　　　　 「確定｣「取消」
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdSet_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0
        '確定釦押下時
            '表示元画面に、表示中の文字列を返す。
            gstrTenKeyData = txtKeyIn
            '自画面を消す。
            Unload Me
        Case 1
        '取消釦押下時
            '自画面を消す。
            Unload Me
    End Select
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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
On Error Resume Next
    '汎用メイル受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmTenKey.Caption, False
        pfFormActive (frmTenKey.hwnd)
    End If
End Sub
