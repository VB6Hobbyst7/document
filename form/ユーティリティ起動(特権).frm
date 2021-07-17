VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtility 
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4680
      Top             =   8160
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
      Left            =   1200
      TabIndex        =   42
      Top             =   7800
      Width           =   2295
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
      Left            =   5520
      TabIndex        =   41
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
      Height          =   615
      Index           =   9
      Left            =   8650
      TabIndex        =   40
      Top             =   6960
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
      Height          =   615
      Index           =   8
      Left            =   8650
      TabIndex        =   39
      Top             =   6240
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
      Height          =   615
      Index           =   7
      Left            =   8650
      TabIndex        =   38
      Top             =   5520
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
      Height          =   615
      Index           =   6
      Left            =   8650
      TabIndex        =   37
      Top             =   4800
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
      Height          =   615
      Index           =   5
      Left            =   8650
      TabIndex        =   36
      Top             =   4080
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
      Height          =   615
      Index           =   4
      Left            =   8650
      TabIndex        =   35
      Top             =   3360
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
      Height          =   615
      Index           =   3
      Left            =   8650
      TabIndex        =   34
      Top             =   2640
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
      Height          =   615
      Index           =   2
      Left            =   8650
      TabIndex        =   33
      Top             =   1920
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
      Height          =   615
      Index           =   1
      Left            =   8650
      TabIndex        =   32
      Top             =   1200
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
      Height          =   615
      Index           =   0
      Left            =   8650
      TabIndex        =   31
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   30
      Top             =   600
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   7680
      TabIndex        =   29
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7680
      TabIndex        =   28
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   27
      Top             =   1320
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   7680
      TabIndex        =   26
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   25
      Top             =   2040
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   7680
      TabIndex        =   24
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   23
      Top             =   2760
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   7680
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   21
      Top             =   3480
      Width           =   6135
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "設定変更"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   10
      Top             =   7080
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   7680
      TabIndex        =   9
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1440
      TabIndex        =   8
      Top             =   6360
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   7680
      TabIndex        =   7
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1440
      TabIndex        =   6
      Top             =   5640
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   7680
      TabIndex        =   5
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   4
      Top             =   4920
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   7680
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "起 動"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   7680
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   1
      Top             =   4200
      Width           =   6135
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
      TabIndex        =   43
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmUtility.frm
'//  パッケージ名：ユーティリティ起動(特権メンテナンス)画面
'//
'//  概要：ユーティリティ起動(特権メンテナンス)画面
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(2.8.0.1) 2011-02-07   REVISED BY [TCC] S.Terao
'//                 配列参照不具合修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const iHoshuAplMax = 19            '登録最大件数
Private sChangeExePass(0 To 31) As String  '変更可能固定起動釦に対応したｱﾌﾟﾘﾌｧｲﾙﾊﾟｽ名（ﾖﾋﾞｴﾘｱを含む）
Private sFixedExePass(0 To 31) As String   '固定起動釦に対応したｱﾌﾟﾘﾌｧｲﾙﾊﾟｽ名（ﾖﾋﾞｴﾘｱを含む）
Private sFixedExeName(0 To 31) As String   '固定起動釦に対応した釦名称（ﾖﾋﾞｴﾘｱを含む）
Private iGamenSts As Integer               '現在表示画面数
Private iHyoujiCnt As Integer              '表示カウンター
Private iKoteiHyouji_Flag As Integer       '固定登録数10件以上フラグ
Private iChangeHyouji_Flag As Integer      '変更登録数10件以上フラグ
Private iContinuFlag As Integer            '「次画面」「前画面」釦表示有無フラグ

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : ユーティリティ起動(特権メンテナンス)画面(アクティブ時)
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
'//  機能名称  : ユーティリティ起動(特権メンテナンス)画面(ディアクティブ時)
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
'//  機能名称  : ユーティリティ起動(特権メンテナンス)画面(ロード時)
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
    iHyoujiCnt = 0  '表示カウンター
    iGamenSts = 0 '現在表示画面数
    Command1(0).Visible = False '「前画面」釦非表示。
    Command1(1).Visible = False '「次画面」釦非表示。
    iKoteiHyouji_Flag = 0
    iChangeHyouji_Flag = 0
    
    'V1.3.0.1 ADD START
    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    '1.3.0.1 ADD END
    
    For i = 0 To 31
        '表示名エリア初期化
        sFixedExeName(i) = ""
    Next
    For i = 0 To 31
        'ツールパスエリア初期化
        sFixedExePass(i) = ""
    Next
    For i = 0 To 31
        '変更可能エリア初期化
        sChangeExePass(i) = ""
    Next
    
    
    '変更可能固定アプリ表示処理
    sFixedKoteiExeDisplay
    
    '固定アプリの情報を取得し、起動用釦を表示する。
    sFixedExeDisplay
    
    '登録件数10件以上チェックを行う。
    If iKoteiHyouji_Flag = 1 Then
     '固定アプリが10件以上ある場合
      Command1(0).Visible = True
      Command1(1).Visible = True
     '「次画面」「前画面」釦表示フラグをONにする。
      iContinuFlag = True
    End If
    
    If iChangeHyouji_Flag = 1 Then
     '変更可能固定アプリが10件以上ある場合
      Command1(0).Visible = True
      Command1(1).Visible = True
     '「次画面」「前画面」釦表示フラグをONにする。
      iContinuFlag = True
    End If
   
   '現在表示画面数を1画面目に設定する。
    iGamenSts = 1

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdChange_Click
'//  機能名称  : 「設定変更」釦押下時処理
'//  機能概要  : アプリの設定を変更するための、アプリ選択画面を表示し、
'//              設定を更新する。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdChange_Click(Index As Integer)
    Dim iResponse As Integer 'MsgBoxボタンコード
    Dim sFileName As String  '選択された実行ファイル名
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト  'V1.20.0.1 ADD
    
    On Error Resume Next
    
    '「ﾕｰﾃｨﾘﾃｨ画面：設定変更釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_CHANGE_SETTEI_BUTTOM, 0)

    
    '画面設定インデックスは0～9なので、釦インデックス値を算出し、
    '起動アプリのパスで起動する。
    '起動アプリパスインデックス=(現在画面数-1画面)×1画面最大釦数＋押下インデックス(0～9)
    '例：2画面目の押下釦インデックス3が押下された場合、起動アプリパスインデックスは13
    '13=(2-1)＊10＋3
    Index = (iGamenSts - 1) * 10 + Index
    
    'アプリ設定変更のためのファイル選択画面を出力する。
    'sFileName = pfFileSelection("D:", "*.exe;*.com;*.bat;*.cmd", _
                                        "実行ファイル選択")    'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    '取得ファイル名を初期化
    CommonDialog1.FileName = ""
    '初期ディレクトリを設定
    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'フォルダ選択画面デフォルトパス１が存在するか
        '存在するため、デフォルトパス１（H:）を設定
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
    Else
        '存在しないため、デフォルトパス２（C:）を設定
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
    End If
    Set objFso = Nothing
    '拡張子を設定
    CommonDialog1.Filter = _
        "プログラムファイル（*.exe;*.com;*.bat;*.cmd）|*.exe;*.com;*.bat;*.cmd|"
    'ファイル選択画面を開く
    CommonDialog1.ShowOpen
    '選択したファイル名を取得
    sFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
                                   
     Call ChDrive("D")  'V2.5.0.1 ADD
                              
    'ファイル選択画面でのアプリの選択有無をチェックする。
    If sFileName <> "" Then
         If iGamenSts = 2 Then
             '現在表示画面：2画面目。
             '表示部の即時反映のために、インデックス番号を1画面分(-10)して設定する。
             txtExeName(Index - 10) = sFileName
         Else
             '現在表示画面：1画面目。
             '表示部の即時反映のために設定する。
             txtExeName(Index) = sFileName
        End If
         '変更可能固定起動釦に対応したｱﾌﾟﾘﾌｧｲﾙﾊﾟｽにも変更されたﾊﾟｽ名を設定する。
         sChangeExePass(Index) = sFileName
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : txtExeName_DblClick
'//  機能名称  : ﾌｧｲﾙﾊﾟｽ表示部ダブルクリック時処理
'//  機能概要  : 削除確認メッセージを表示し、削除を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub txtExeName_DblClick(Index As Integer)
    Dim iResponse As Integer      'MsgBoxボタンコード
    Dim iSetupAplIndex As Integer '起動アプリインデックス
    
    On Error Resume Next
   
   '「ﾕｰﾃｨﾘﾃｨ画面：設定表示部ﾀﾞﾌﾞﾙｸﾘｯｸ」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_DOUBLECLICK_SETTEI, 0)

    '画面設定インデックスは0～9なので、釦インデックス値を算出し、
    '起動アプリのパスで起動する。
    '起動アプリインデックス=(現在画面数-1画面)×1画面最大釦数＋押下インデックス(0～9)
    '例：2画面目の押下釦インデックス3が押下された場合、起動アプリパスインデックスは13
    '13=(2-1)＊10＋3
   iSetupAplIndex = (iGamenSts - 1) * 10 + Index

   '変更可能固定起動釦ファイルﾊﾟｽ格納エリアの定義有無チェックを行う。
   If sChangeExePass(iSetupAplIndex) <> "" Then
        '「登録除外」ポップアップ画面を表示する。
        iResponse = MsgBox(txtExeName(Index).Text & "を登録から除外します。" _
                            & Chr(vbKeyReturn) & " よろしいですか？", _
                            vbYesNo + vbExclamation, _
                            "実行ファイル名の登録除外")
        If iResponse = vbYes Then
        ' [はい] ボタンを選択した場合
            '実行ファイル名の表示を消す。
            txtExeName(Index).Text = ""
            sChangeExePass(iSetupAplIndex) = ""
            '「ﾕｰﾃｨﾘﾃｨ画面：登録削除」ログ出力
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_TOOL_SETTEI_DELETE, 0)
        Else
        ' [いいえ] ボタンを選択した場合
            '何もしない。
            Exit Sub
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdExecute_Click
'//  機能名称  : 「起動」釦押下時処理
'//  機能概要  : アプリの起動を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 ファイル選択画面をOS仕様に変更
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 媒体取外不具合修正
'//     REVISIONS :(2.8.0.1) 2011-02-07   REVISED BY [TCC] S.Terao
'//                 配列参照不具合修正
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdExecute_Click(Index As Integer)
    Dim lRetVal As Double     'Shell関数戻り値
    Dim iResponse As Integer  'MsgBoxボタンコード
    Dim iSetupAplIndex As Integer '起動アプリインデックス
    
    Dim objFso As New FileSystemObject   'ファイルシステムオブジェクト  'V1.20.0.1 ADD
   
On Error GoTo ERROR_MSG
    '「ﾕｰﾃｨﾘﾃｨ画面：起動釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_KIDOU_BUTTOM, 0)

    '画面設定インデックスは0～9なので、釦インデックス値を算出し、
    '起動アプリのパスで起動する。
    '起動アプリインデックス=(現在画面数-1画面)×1画面最大釦数＋押下インデックス(0～9)
    '例：2画面目の押下釦インデックス3が押下された場合、起動アプリパスインデックスは13
    '13=(2-1)＊10＋3
    iSetupAplIndex = (iGamenSts - 1) * 10 + Index

    '起動対象のﾊﾟｽ名定義チェックを行う。
    If (sChangeExePass(iSetupAplIndex) = "") Then
        '起動対象ﾊﾟｽ名定義が無い場合、起動アプリ選択のためのファイル選択画面を表示する。
        'txtExeName(Index) = pfFileSelection("D:", "*.exe;*.com;*.bat;*.cmd", _
                                            "実行ファイル選択")     'V1.20.0.1 DEL
        'V1.20.0.1 ADD START
        '取得ファイル名を初期化
        CommonDialog1.FileName = ""
        '初期ディレクトリを設定
        If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'フォルダ選択画面デフォルトパス１が存在するか
            '存在するため、デフォルトパス１（H:）を設定
            CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
        Else
            '存在しないため、デフォルトパス２（C:）を設定
            CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
        End If
        Set objFso = Nothing
        '拡張子を設定
        CommonDialog1.Filter = _
            "プログラムファイル（*.exe;*.com;*.bat;*.cmd）|*.exe;*.com;*.bat;*.cmd|"
        'ファイル選択画面を開く
        CommonDialog1.ShowOpen
        '選択したファイル名を取得
        txtExeName(Index) = CommonDialog1.FileName
        'V1.20.0.1 ADD END
       sChangeExePass(iSetupAplIndex) = txtExeName(Index)
    End If
  
    '起動対象ﾊﾟｽ名定義がある場合。
    'If (sChangeExePass(Index) <> "") Then           'V2.8.0.1 DEL
    If (sChangeExePass(iSetupAplIndex) <> "") Then   'V2.8.0.1 ADD
    '設定欄にアプリケーションがあれば、
        '設定欄のアプリケーションを実行する。
        lRetVal = Shell(sChangeExePass(iSetupAplIndex), vbNormalFocus)
        '「ﾕｰﾃｨﾘﾃｨ画面：ツール起動正常」ログ出力
        Call sLogTraceReq(LTYP_NORMAL, L3AN_API, UTILITY_GAMEN_TOOL_OK, 0)
    End If
    Call ChDrive("D")  'V2.5.0.1 ADD
    Exit Sub

ERROR_MSG:
    '「ﾕｰﾃｨﾘﾃｨ画面：ツール起動異常」ログ出力
     Call sLogTraceReq(LTYP_ERROR, L3AN_API, UTILITY_GAMEN_TOOL_ERROR, 0)
    '「起動異常」ポップアップ画面を表示する。
     iResponse = MsgBox("実行するアプリケーションを" _
                        & Chr(vbKeyReturn) & "正しく設定してください", _
                        vbYes, _
                        "アプリ実行エラー")
     Exit Sub
     Set objFso = Nothing    'V1.20.0.1 ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdFixedExe_Click
'//  機能名称  : 「起動(固定)」釦押下時処理
'//  機能概要  : アプリの起動を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    Dim lRetVal As Double      'Shell関数戻り値
    Dim iResponse As Integer   'MsgBox戻り値
    Dim iSetupAplIndex As Integer '起動アプリインデックス

On Error GoTo ERROR_MSG
    '「ﾕｰﾃｨﾘﾃｨ画面：起動釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_KIDOU_BUTTOM, 0)
  
    '画面設定インデックスは0～9なので、釦インデックス値を算出し、
    '起動アプリのパスで起動する。
    '起動アプリパスインデックス=(現在画面数-1画面)×1画面最大釦数＋押下インデックス(0～9)
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
'//  引数      : Integer　Index　　　[IN]押下釦インデックス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Command1_Click(Index As Integer)

  On Error Resume Next

  Select Case Index
   Case 0
     '「ﾕｰﾃｨﾘﾃｨ画面：前画面釦押下」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_BACK_BUTTOM, 0)
     If iGamenSts = 1 Then
       '現在表示画面数：1画面目。
       '次表示画面数は2画面目のため、現在表示画面数に2を設定する。
       iGamenSts = iGamenSts + 1
     Else
       '現在表示画面数：2画面目。
       '表示開始点は0、次表示画面数は1画面目のため、現在表示画面数に1に設定する。
       iGamenSts = 1
       iHyoujiCnt = 0
     End If
      
     '固定釦、変更可能固定釦表示処理を行う。
     sSetAplClickDisplay
     sAplClick_Display
    
     '現在表示画面数：1画面時のみ、表示カウンターのカウントアップはここで行う。
     If iGamenSts = 1 Then
       iHyoujiCnt = 10
     End If
     
    Case 1
      '「ﾕｰﾃｨﾘﾃｨ画面：次画面釦押下」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_NEXT_BUTTOM, 0)
 
       If iGamenSts = 2 Then
         '現在表示画面数：2画面目。
         '表示開始点は0、次表示画面数は1画面目のため、現在表示画面数に1を設定する。
         iGamenSts = 1
         iHyoujiCnt = 0
       Else
         '現在表示画面数：1画面目。
         '次表示画面数は2画面目のため、現在表示画面数に2を設定する。
         iGamenSts = iGamenSts + 1
       End If
     
        sSetAplClickDisplay
        sAplClick_Display
     
       '現在表示画面数：1画面時のみ、表示カウンターのカウントアップはここで行う。
       If iGamenSts = 1 Then
          iHyoujiCnt = 10
       End If
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
'//  関数名称  : sFixedKoteiExeDisplay
'//  機能名称  : 変更可能固定アプリ起動釦表示初期処理
'//  機能概要  : 変更可能アプリの初期処理を行う。
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
Private Sub sFixedKoteiExeDisplay()
    Dim lSts As Long    'INIファイル設定取得関数の戻り値
    Dim lCnt As Long    'HOSHUAPL.INIの登録件数
    Dim iMax As Integer 'アプリインデックス最大値
    Dim sWork As String * UTILITY_SIZE
    Dim i As Integer

    On Error Resume Next

    iMax = txtExeName.UBound 'アプリパスINDEXの最大値を得る。

    ' 事前設定ファイルから設定変更可能アプリの「登録件数」を取出す
    lCnt = GetPrivateProfileInt(PROFILE_SECTION_NAME, PROFILE_KEY_NAME_COUNT, _
                                DEFAILT_Int, HOSHUAPL_FILE)
    If (lCnt > 0) Then
    '設定変更可能アプリがあれば、
        ' 設定変更可能アプリの登録文字列を取出す
        For i = 0 To lCnt - 1
            lSts = GetPrivateProfileString(PROFILE_SECTION_NAME, _
                                           PROFILE_KEY_NAME_HEAD & i, _
                                           DEFAILT, sWork, Len(sWork), HOSHUAPL_FILE)
             'INIファイルの取得結果チェックを行う。
             If lSts > 0 Then
               If i <= iMax Then
               'INIファイルより取得正常、また10件以内のため、画面表示する。
                txtExeName(i) = sWork
               End If
               '変更可能固定起動釦ｱﾌﾟﾘﾌｧｲﾙﾊﾟｽエリアに、取得ﾊﾟｽを格納する。
               sChangeExePass(i) = sWork
            End If
        Next i
    End If
    
    '登録件数が、1画面表示最大10以上かどうかチェックする。
    If iMax < lCnt Then
     '10件以上の場合、変更登録数10件以上フラグをONにする。
     iChangeHyouji_Flag = 1
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sFixedExeDisplay
'//  機能名称  : 固定アプリ起動釦表示初期処理
'//  機能概要  : 固定アプリの初期処理を行う。
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
Dim i As Integer          'INIﾌｧｲﾙキーカウンタ：DSPi ＝起動釦INDEX
Dim iMax As Integer       '固定起動釦INDEX最大値
Dim sLine As String * 256 '１行文の文字列。（文字列”DSPi=”を除く）
Dim lSize As Long         '１行文のﾊﾞｲﾄ数。（文字列”DSPi=”を除く）
Dim iK As Integer         'カンマ記述位置

On Error Resume Next
 
'全ての固定起動釦について、以下を実施する。
iMax = cmdFixedExe.UBound     '固定起動釦INDEXの最大値を得る。
 
 For i = 0 To iHoshuAplMax
   'アプリ起動初期値INIファイルから、１行文の文字列（DSPi=を除く）を読込む。
    lSize = GetPrivateProfileString(PROFILE_SECTION_NAME_FIXED_EXE, _
                                    PROFILE_KEY_NAME_FIXED_EXE & CStr(i), _
                                    DEFAILT, sLine, Len(sLine), HOSHUAPL_FILE)
    iK = InStr(sLine, ",")        'ファイル名（フルパス）の区切文字位置を得る。
    'INIファイルに、該当行の定義がある場合、
    If lSize > 0 And iK <> 0 Then
     'ファイル名と釦名称を取出し、保存しておく。
      sFixedExePass(i) = Trim$(Left$(sLine, iK - 1))
      sFixedExeName(i) = Trim$(Mid$(sLine, iK + 1, lSize - iK))
    End If
Next i

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
      iKoteiHyouji_Flag = 1
    End If
 Next i

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sSetAplClickDisplay
'//  機能名称  : 「次画面」「前画面」釦押下時処理。
'//  機能概要  : 固定アプリ起動部の表示処理を行う。
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
Private Sub sSetAplClickDisplay()
Dim i As Integer          'INIﾌｧｲﾙキーカウンタ：DSPi ＝起動釦INDEX
Dim iMax As Integer       '固定起動釦INDEX最大値
Dim iCnt As Integer       '内部ループカウンター

On Error Resume Next

'表示カウンターを内部ループカウンターに取得する。
iCnt = iHyoujiCnt
 
'全ての固定起動釦について、以下を実施する。
iMax = cmdFixedExe.UBound     '固定起動釦INDEXの最大値を得る。
For i = CNT_MIN To iMax
  '固定起動釦を消しておく。
     cmdFixedExe(i).Visible = False
       '起動アプリパス名と、表示釦名称の定義チェックを行う。
       If sFixedExePass(iCnt) <> "" And sFixedExeName(iCnt) <> "" Then
         '定義有りの場合、キャプションに起動釦表示文字列を書込み、起動釦を表示する。
          cmdFixedExe(i).Visible = True
          cmdFixedExe(i).Caption = sFixedExeName(iCnt)
        End If
        '内部ループカウンターをカウントアップする。
        iCnt = iCnt + 1
Next i
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sAplClick_Display
'//  機能名称  : 「次画面」「前画面」釦押下時処理。
'//  機能概要  : 変更可能固定アプリ起動部の表示処理を行う。
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
Private Sub sAplClick_Display()

Dim i As Integer          'INIﾌｧｲﾙキーカウンタ：DSPi ＝起動釦INDEX
Dim iMax As Integer       '固定起動釦INDEX最大値
Dim iCnt As Integer       '内部ループカウンター

On Error Resume Next

'表示カウンターを内部ループカウンターに取得する。
iCnt = iHyoujiCnt

iMax = txtExeName.UBound 'アプリパスINDEXの最大値を得る。
For i = CNT_MIN To iMax
  '表示処理を行う。
  txtExeName(i) = sChangeExePass(iCnt)
  '内部ループカウンターをカウントアップする。
  iCnt = iCnt + 1
Next i

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Unload
'//  機能名称  : 画面消去時の設定をINIファイルに反映する。
'//  機能概要  : 「メンテナンス画面へ戻る」釦押下時処理：
'//              HOSHUAPL.INIへの設定を更新する。
'//
'//              型        名称      意味
'//  引数      : Integer　Cancel
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer            'カウンター
    Dim l As Integer            '登録件数カウンター
    Dim iMax As Integer         '実行アプリ表示欄INDEXの最大値
    Dim lSts As Boolean         'INIファイル反映戻り値

    On Error Resume Next
    
    l = 0
    iMax = txtExeName.UBound   '実行アプリ表示欄INDEXの最大値をセットする。
   
   '「次・前画面」釦表示有無チェック。
   '「次・前画面」釦の表示がある場合、最大ループカウンターを最大20にする。
   If iContinuFlag = True Then
      iMax = (iMax + 1) * 2
   End If

    For i = CNT_MIN To iMax
      If (sChangeExePass(i) <> "") Then
        l = l + 1
      End If
       'アプリ起動初期値ファイルに起動アプリの実行ファイル名を書込む。
        lSts = WritePrivateProfileString(PROFILE_SECTION_NAME, _
               PROFILE_KEY_NAME_HEAD & CStr(i), sChangeExePass(i), HOSHUAPL_FILE)
    Next i
  
  '登録件数(定義有り件数)を更新する。
   lSts = WritePrivateProfileString(PROFILE_SECTION_NAME, _
          PROFILE_KEY_NAME_COUNT, CStr(l), HOSHUAPL_FILE)
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
        AppActivate frmUtility.Caption, False
        pfFormActive (frmUtility.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

