VERSION 5.00
Begin VB.Form frmToriatukaiKenshuModeSettei 
   BorderStyle     =   0  'なし
   Caption         =   "リモートメンテナンス"
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
      Left            =   6960
      Top             =   7680
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   31
      Left            =   10320
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   72
      Top             =   6240
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   30
      Left            =   8850
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   71
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   7410
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   70
      Top             =   6240
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   5970
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   69
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   4530
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   68
      Top             =   6240
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   3090
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   67
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   1650
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   66
      Top             =   6240
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   65
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   10290
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   56
      Top             =   2760
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   8850
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   55
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   7410
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   54
      Top             =   2760
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   5970
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   53
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   4530
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   52
      Top             =   2760
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   3090
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   51
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "不可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   1650
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   50
      Top             =   2760
      Value           =   1  'ﾁｪｯｸ
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "可"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   49
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdJIKISelect_All 
      Caption         =   " 磁気取扱 全号機可"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   3600
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   40
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdJIKISelect_All 
      Caption         =   " 磁気取扱 全号機不可"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   5280
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   39
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdICSelect_All 
      Caption         =   " ＩＣ取扱 全号機可"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   38
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdICSelect_All 
      Caption         =   "ＩＣ取扱 全号機不可"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   1920
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   37
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "磁気取扱"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   11655
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   23
         Left            =   10200
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   64
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   22
         Left            =   8730
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   63
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   21
         Left            =   7290
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   62
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   20
         Left            =   5850
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   61
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   19
         Left            =   4410
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   60
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   18
         Left            =   2970
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   59
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   17
         Left            =   1530
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   58
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   16
         Left            =   120
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   57
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   10200
         TabIndex        =   36
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   8760
         TabIndex        =   35
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   7320
         TabIndex        =   34
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   5880
         TabIndex        =   33
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   4440
         TabIndex        =   32
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   3000
         TabIndex        =   31
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   1560
         TabIndex        =   30
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   10200
         TabIndex        =   28
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   8760
         TabIndex        =   27
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   7320
         TabIndex        =   26
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   5880
         TabIndex        =   25
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   4440
         TabIndex        =   24
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ＩＣ取扱"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11655
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   10200
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   48
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   8760
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   7320
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   46
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   5880
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   4440
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   44
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   3000
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "不可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1560
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   42
         Top             =   720
         Value           =   1  'ﾁｪｯｸ
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "可"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   150
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   16
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5880
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7320
         TabIndex        =   14
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8760
         TabIndex        =   13
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   10200
         TabIndex        =   12
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   1560
         TabIndex        =   10
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3000
         TabIndex        =   9
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   4440
         TabIndex        =   8
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   5880
         TabIndex        =   7
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   7320
         TabIndex        =   6
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   8760
         TabIndex        =   5
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   10200
         TabIndex        =   4
         Top             =   1680
         Width           =   1275
      End
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
      Caption         =   "取扱券種モード設定"
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
Attribute VB_Name = "frmToriatukaiKenshuModeSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmToriatukaiKenshuModeSettei.frm
'//  パッケージ名：取扱券種モード設定画面
'//
'//  概要：取扱券種モード設定画面
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//                 ・フェーズ３対応　新規追加画面
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000             'メールタイマのインターバル値
Private iIC_KenshuMode_Sts(0 To 15) As Integer    'IC取扱券値取得エリア
Private iJIKI_KenshuMode_Sts(0 To 15) As Integer  '磁気取扱券値取得エリア
Private iICGOUKI_SETTEI(0 To 15) As Integer       'IC設定変更号機フラグ
Private iJIKIGOUKI_SETTEI(0 To 15) As Integer     '磁気設定変更号機フラグ
Private Const MAX_GOUKI = 15                      '最大号機値

Private Const MOVE_JIKI_INDEX = 16                '磁気取扱部開始インデックス値までの移動
Private Const SETTEI_ARI = 1                      '設定有
Private Const SETTEI_NASI = 0                     '設定無
Private Const HUTEI = -1                          '値不定
Private Const HUKA_STS = 1                        '不可値
Private Const KA_STS = 0                          '可値
Private Const HUKA = "不可"                       '表示文言：不可
Private Const KA = "可"                           '表示文言：可
Private Const IC_KENSHU = 0                       'IC取扱券種
Private Const JIKI_KENSHU = 1                     '磁気取扱券種
Dim bBUTTOM_STS As Boolean                        '釦押下状態：TRUE=押下中　FALSE=非押下
Dim bUpData_Flag As Boolean                       '設定更新処理有無フラグ　TRUE：更新処理有　FALSE=更新処理無

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 取扱券種モード設定画面(アクティブ時)
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
'//  機能名称  : 取扱券種モード設定画面(ディアクティブ時)
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
'//  機能名称  : 取扱券種モード設定画面(ロード時)
'//  機能概要  : 取扱券種モード設定画面の初期処理を行う。
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
    
    Dim iCnt As Integer     'カウンター
    
    On Error Resume Next

    '「取扱券種モード設定画面 表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_GAMEN_START, 0)

    'メイル受信用のインタバルタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '各エリア初期化
    For iCnt = 0 To MAX_GOUKI
      iIC_KenshuMode_Sts(iCnt) = HUTEI
      iJIKI_KenshuMode_Sts(iCnt) = HUTEI
      iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
      iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
    Next
    
    bUpData_Flag = False
    
    bBUTTOM_STS = True
    
    '画面表示処理
    pfDispSettei
    
    bBUTTOM_STS = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下時
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

    '「取扱券種モード設定画面 消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_GAMEN_END, 0)
   
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdICSelect_All_Click
'//  機能名称  : IC取扱全号機釦押下時処理
'//  機能概要  : IC取扱全号機釦押下処理(可/不可)を行う。
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
Private Sub cmdICSelect_All_Click(Index As Integer)
  
  Dim iCnt As Integer '号機カウンター
  
  On Error Resume Next
  
  bBUTTOM_STS = True

  If Index = 0 Then
    '全号機：可設定
    '「取扱券種モード設定画面:IC取扱全号機可釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_IC_ALLGOUKI_KA_BUTTOM, 0)

    For iCnt = 0 To MAX_GOUKI
        If chkMode(iCnt).Visible = True Then
           iIC_KenshuMode_Sts(iCnt) = KA_STS
           chkMode(iCnt).Caption = KA
           chkMode(iCnt).Value = 0
        End If
    Next
  Else
     '全号機：不可設定
     '「取扱券種モード設定画面:IC取扱全号機不可釦押下」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_IC_ALLGOUKI_HUKA_BUTTOM, 0)
     For iCnt = 0 To MAX_GOUKI
         If chkMode(iCnt).Visible = True Then
            iIC_KenshuMode_Sts(iCnt) = HUKA_STS
            chkMode(iCnt).Caption = HUKA
            chkMode(iCnt).Value = 1
         End If
    Next
  End If
  
  bBUTTOM_STS = False
  
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : chkMode_Click
'//  機能名称  : 各号機別釦押下時処理
'//  機能概要  : 各号号機別釦押下処理(可/不可)を行う。
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
Private Sub chkMode_Click(Index As Integer)
  
  On Error Resume Next
 
  If bBUTTOM_STS = True Then
     'IC/磁気全号機一括釦処理中は、以降処理を以降の処理を行わない。
     Exit Sub
  End If
 
  '「取扱券種モード設定画面:号機別設定変更」ログ出力
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_GOUKIBETU_BUTTOM, 0)

  '各号機別釦値を変更
  If chkMode(Index).Value = 1 Then
     chkMode(Index).Caption = HUKA
  Else
     chkMode(Index).Caption = KA
  End If
  
  '取扱エリアチェックを行い、対象エリアの値を現在値に変更
  If Index < MOVE_JIKI_INDEX Then
    'IC取扱各号機：可設定
    If chkMode(Index).Value = 1 Then
       iIC_KenshuMode_Sts(Index) = HUKA_STS
    Else
       iIC_KenshuMode_Sts(Index) = KA_STS
    End If
  Else
     '磁気取扱各号機：不可設定
    If chkMode(Index).Value = 1 Then
       iJIKI_KenshuMode_Sts(Index - MOVE_JIKI_INDEX) = HUKA_STS
    Else
       iJIKI_KenshuMode_Sts(Index - MOVE_JIKI_INDEX) = KA_STS
    End If
  End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdJIKISelect_All_Click
'//  機能名称  : 磁気取扱全号機釦押下時処理
'//  機能概要  : 磁気取扱全号機釦押下処理(可/不可)を行う。
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
Private Sub cmdJIKISelect_All_Click(Index As Integer)
 
  Dim iCnt As Integer '号機カウンター
  
  On Error Resume Next
  
  bBUTTOM_STS = True

  If Index = 0 Then
     '全号機：可設定
     '「取扱券種モード設定画面:磁気取扱全号機可釦押下」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_JIKI_ALLGOUKI_KA_BUTTOM, 0)
  
     For iCnt = 0 To MAX_GOUKI
         If chkMode(iCnt + MOVE_JIKI_INDEX).Visible = True Then
            iJIKI_KenshuMode_Sts(iCnt) = KA_STS
            chkMode(iCnt + MOVE_JIKI_INDEX).Caption = KA
            chkMode(iCnt + MOVE_JIKI_INDEX).Value = 0
         End If
    Next
  Else
     '全号機：不可設定
     '「取扱券種モード設定画面:磁気取扱全号機不可釦押下」ログ出力
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_JIKI_ALLGOUKI_HUKA_BUTTOM, 0)
   
     For iCnt = 0 To MAX_GOUKI
         If chkMode(iCnt + MOVE_JIKI_INDEX).Visible = True Then
            iJIKI_KenshuMode_Sts(iCnt) = HUKA_STS
            chkMode(iCnt + MOVE_JIKI_INDEX).Caption = HUKA
            chkMode(iCnt + MOVE_JIKI_INDEX).Value = 1
         End If
    Next
  End If
 
 bBUTTOM_STS = False
 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdKakutei_Click
'//  機能名称  : 画面設定値を反映する。
'//  機能概要  : 自改設定エリア、又は自改設定ファイルに値の反映を行う。
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
Private Sub cmdKakutei_Click()
  
  On Error Resume Next
 
  '「取扱券種モード設定画面:確定釦押下」ログ出力
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_KENSHU_KAKUTEI_BUTTOM, 0)
  
  '画面をロックする。
  SetEnableFalse
  
  '画面値設定反映処理
  psGamenSettei_Hanei
  
  '画面ロック解除。
  SetEnableTrue
  
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
        AppActivate frmToriatukaiKenshuModeSettei.Caption, False
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfDispSettei
'//  機能名称  : 取扱券種モード設定画面(ロード時)
'//  機能概要  : 取扱券種モード設定画面の初期処理を行う。
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
Private Function pfDispSettei()
    Dim iCnt As Integer '号機カウンター
    Dim iSetti_Gouki As Integer '号機設置/未設置状態フラグ
    
    On Error Resume Next
    
    For iCnt = 0 To MAX_GOUKI
        'INIファイルより号機設置/未設置情報を取得する。
        iSetti_Gouki = pfGetGoukiNo(iCnt + 1)
        If iSetti_Gouki = 1 Then
           '設置有
          lblGokiBetsuNumber(iCnt).Visible = True
          lblGokiBetsuNumber(iCnt + MOVE_JIKI_INDEX).Visible = True
          lblGokiBetsuNumber(iCnt).Caption = iCnt + 1
          lblGokiBetsuNumber(iCnt + MOVE_JIKI_INDEX).Caption = iCnt + 1
         
          '号機別釦表示処理
          pfGet_Sts iCnt
        Else
           '未設置：IC/磁気取扱の号機番号・号機別釦を非表示にする。
           lblGokiBetsuNumber(iCnt).Visible = False
           lblGokiBetsuNumber(iCnt + MOVE_JIKI_INDEX).Visible = False
           chkMode(iCnt).Visible = False
           chkMode(iCnt + MOVE_JIKI_INDEX).Visible = False
        End If
    Next
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetGoukiNo
'//  機能名称  : 設置号機を取得する。
'//  機能概要  : GATE.INIより設置号機を取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Integer            [OUT]0：未設置/取得異常
'//                                      1：設置
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetGoukiNo(iGouki As Integer) As Integer

    Dim lngRet As Long          '関数の返り値
    Dim iGate As Integer        '自改INDEX
    Dim j As Integer            'ワークINDEX
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '１行分ファイル内容取得用
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
   
    On Error Resume Next

    '自動改札機情報取得
    sKeyName = "gate" & Format(iGouki, "00")
    iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                   sKeyName, _
                                   DEFAILT, sGateData, Len(sGateData), _
                                   PATH_GATE_FILE)
    If iRet = 0 Then
       '「取扱券種モード設定画面：自動改札機INIファイル読込異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
       '取得異常
       pfGetGoukiNo = 0
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
           
    If Trim(sFData(4)) = EGR Then
       '自改タイプ：EG-R自改/設置有
       pfGetGoukiNo = 1
       Exit Function
    ElseIf Trim(sFData(4)) = NEG Then
       '自改タイプ：NEG自改/設置有
       pfGetGoukiNo = 1
       Exit Function
    Else
       '上記以外：未設置号機扱い
       pfGetGoukiNo = 0
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGet_Sts
'//  機能名称  : 自改設定ファイル/エリアより現在値を取得処理
'//  機能概要  : 自改設定ファイル/エリアより現在値の取得を行う。
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
Private Function pfGet_Sts(iGouki As Integer)
    Dim iKansiAplSts As Integer '監視盤アプリ起動状態
    Dim iRet As Integer         '値取得処理戻り値
    
    On Error Resume Next
   
    '監視盤アプリ起動チェックを行う。
    iKansiAplSts = CheckAppStart(PROC_KANRI)
    If iKansiAplSts <> 0 Then
       '監視盤起動時:自改設定エリアより値取得
        pfAreaGet_Sts iGouki
       
    Else
       '監視盤未起動時：自改設定ファイルより値取得
       pfFileGet_Sts iGouki
    End If
    
    '値取得チェック
    '取得正常：釦表示　取得異常：号機番号のみ表示
    'IC/磁気：取得正常
    If iIC_KenshuMode_Sts(iGouki) <> HUTEI And _
       iJIKI_KenshuMode_Sts(iGouki) <> HUTEI Then
       
       'IC取扱値取得処理正常：釦表示
       chkMode(iGouki).Visible = True
       If iIC_KenshuMode_Sts(iGouki) = HUKA_STS Then
          chkMode(iGouki).Caption = HUKA
          chkMode(iGouki).Value = 1
       Else
          chkMode(iGouki).Caption = KA
          chkMode(iGouki).Value = 0
       End If
       
       '磁気取扱値取得処理正常：釦表示
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = True
       If iJIKI_KenshuMode_Sts(iGouki) = HUKA_STS Then
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = HUKA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 1
       Else
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = KA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 0
       End If
       
       Exit Function
    
    'IC取得正常/磁気取得異常
    ElseIf iIC_KenshuMode_Sts(iGouki) <> HUTEI And _
           iJIKI_KenshuMode_Sts(iGouki) = HUTEI Then

       'IC取扱値取得処理正常：釦表示
       chkMode(iGouki).Visible = True
       If iIC_KenshuMode_Sts(iGouki) = HUKA_STS Then
          chkMode(iGouki).Caption = HUKA
          chkMode(iGouki).Value = 1
       Else
          chkMode(iGouki).Caption = KA
          chkMode(iGouki).Value = 0
       End If
       
       '磁気取扱部は非表示
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = False

       Exit Function
    
    'IC取得異常/磁気取得正常
    ElseIf iIC_KenshuMode_Sts(iGouki) = HUTEI And _
           iJIKI_KenshuMode_Sts(iGouki) <> HUTEI Then
       
       '磁気取扱値取得処理正常：釦表示
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = True
       If iJIKI_KenshuMode_Sts(iGouki) = HUKA_STS Then
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = HUKA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 1
       Else
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = KA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 0
       End If
       
       'IC取扱部は非表示
       chkMode(iGouki).Visible = False
       
       Exit Function
    Else
       'IC/磁気取得処理異常：釦非表示/号機番号のみ表示
       chkMode(iGouki).Visible = False
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = False
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfAreaGet_Sts
'//  機能名称  : IC/磁気取扱の現在値を取得処理(エリア参照)
'//  機能概要  : IC/磁気取扱の現在値を取得を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iICSts 　　[OUT]IC取扱現在値
'//  引数      : Integer　iJIKISts 　[OUT]磁気取扱現在値
'//              Integer　iGouki  　 [IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfAreaGet_Sts(iGouki As Integer)
    Dim strMutexName    As String           'ミューテックス名
    Dim lngMuHandle     As Long             '排他処理用ハンドル
    Dim iAreaSts        As Integer          'エリア値

    On Error Resume Next
    
    Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
    '自改設定エリアをオープンする。
    Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
    Idinf_JikaiSettei.IdOpen
    If Idinf_JikaiSettei.Errsts <> 0 Then
       Set Idinf_JikaiSettei = Nothing
       'IC/磁気取扱値取得エリアを値不定に設定
       iIC_KenshuMode_Sts(iGouki) = HUTEI
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
       '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
    End If
    
    '自改設定エリアをＬＯＣＫする。
    Idinf_JikaiSettei.IdLock
    If Idinf_JikaiSettei.Errsts <> 0 Then
       Idinf_JikaiSettei.IdFree
       'データ参照異常時
       Set Idinf_JikaiSettei = Nothing
       'IC/磁気取扱値取得エリアを値不定に設定
       iIC_KenshuMode_Sts(iGouki) = HUTEI
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
       '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
     End If
     
     'IC取扱の内容を読み込む。
     Idinf_JikaiSettei.id = IdGate.IC_TORIATUKAI_KENSHU_STS
     Idinf_JikaiSettei.GetJikai_Sts iGouki
     If Idinf_JikaiSettei.Errsts <> 0 Then
        'IC取扱取得異常：IC取扱券種取得エリアを値不定に設定
        iIC_KenshuMode_Sts(iGouki) = HUTEI
        '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
     Else
        'IC取扱取得正常：IC取扱券種取得エリアに取得値を設定
        iAreaSts = Idinf_JikaiSettei.DataArea(iGouki)
        iIC_KenshuMode_Sts(iGouki) = iAreaSts
     End If
     
     '磁気取扱の内容を読み込む。
     Idinf_JikaiSettei.id = IdGate.JIKI_TORIATUKAI_KENSHU_STS
     Idinf_JikaiSettei.GetJikai_Sts iGouki
     If Idinf_JikaiSettei.Errsts <> 0 Then
        '磁気取扱取得異常：磁気取扱券種取得エリアを値不定に設定
        iJIKI_KenshuMode_Sts(iGouki) = HUTEI
        '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
     Else
        '磁気取扱取得正常：磁気取扱券種取得エリアに取得値を設定
        iAreaSts = Idinf_JikaiSettei.DataArea(iGouki)
        iJIKI_KenshuMode_Sts(iGouki) = iAreaSts
     End If
   
     Idinf_JikaiSettei.IdFree
     Set Idinf_JikaiSettei = Nothing
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfFileGet_Sts
'//  機能名称  : IC/磁気取扱の現在値を取得処理(ファイル参照)
'//  機能概要  : IC/磁気取扱の現在値を取得を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfFileGet_Sts(iGouki As Integer)
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255     As GATE_INFO        '読込み用エリア（255設定用）
    Dim lngSts          As Long             'ヒットエリアID
    Dim lngLoop1        As Long             'カウンター
    Dim lngHandle       As Long             'ハンドル
    Dim FileName        As String           'ファイル有無チェック
    Dim lngRet          As Long             '戻り値
    Dim bRet            As Boolean          '読み込み結果戻り値
    Dim sSetteiFile     As String           'ファイルパス　'V1.4.0.1　ADD
    
    On Error Resume Next
   
     '自改設定ファイルをオープン
    lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       'IC/磁気取扱値取得エリアを値不定に設定
       iIC_KenshuMode_Sts(iGouki) = HUTEI
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
       '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
    End If
        
    '自改設定ファイル読み込み
    For lngLoop1 = 0 To iGouki
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           'ハンドルのクローズ
           Call CloseHandle(lngHandle)
           'IC/磁気取扱値取得エリアを値不定に設定
           iIC_KenshuMode_Sts(iGouki) = HUTEI
           iJIKI_KenshuMode_Sts(iGouki) = HUTEI
           '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           Exit Function
        End If
    Next
        
    'ハンドルのクローズ
    Call CloseHandle(lngHandle)
        
    'IC取扱：ID検索
    lngSts = SerchId(udtAreaR255, IdGate.IC_TORIATUKAI_KENSHU_STS)
    If lngSts >= 0 Then
       'IDが有った場合
       iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
       iIC_KenshuMode_Sts(iGouki) = iAreaSts
    Else
       ' 該当ＩＤ無しの場合参照異常
        iIC_KenshuMode_Sts(iGouki) = HUTEI
    End If
    
    '磁気取扱：ID検索
    lngSts = SerchId(udtAreaR255, IdGate.JIKI_TORIATUKAI_KENSHU_STS)
    If lngSts >= 0 Then
       'IDが有った場合
       iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         'データ変換
       iJIKI_KenshuMode_Sts(iGouki) = iAreaSts
    Else
       ' 該当ＩＤ無しの場合参照異常
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SerchId
'//  機能名称  : ＩＤ検索処理(取扱券種モード設定画面用)
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
Private Function SerchId(udtArea255 As GATE_INFO, lngId As Long) As Long

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
            
    SerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ChgData
'//  機能名称  : データ変換処理処理(取扱券種モード設定画面用)
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
    Dim iCnt As Integer
    
    On Error Resume Next

    'IC取扱全号機可/不可釦：False(ロック)する。
    cmdICSelect_All(0).Enabled = False
    cmdICSelect_All(1).Enabled = False
    
    '磁気取扱全号機可/不可釦：False(ロック)する。
    cmdJIKISelect_All(0).Enabled = False
    cmdJIKISelect_All(1).Enabled = False
    
    '確定釦：False(ロック)する。
    cmdKakutei.Enabled = False
    
    'メニュー画面へ戻る釦：False(ロック)する。
    cmdReturn.Enabled = False
    
    For iCnt = 0 To MAX_GOUKI
        'IC取扱エリア：False(ロック)する。
        chkMode(iCnt).Enabled = False
        '磁気取扱エリア：False(ロック)する。
        chkMode(iCnt + MOVE_JIKI_INDEX).Enabled = False
    Next

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
    Dim iCnt As Integer
    
    On Error Resume Next

    'IC取扱全号機可/不可釦：True(ロック解除)する。
    cmdICSelect_All(0).Enabled = True
    cmdICSelect_All(1).Enabled = True
    
    '磁気取扱全号機可/不可釦：True(ロック解除)する。
    cmdJIKISelect_All(0).Enabled = True
    cmdJIKISelect_All(1).Enabled = True
    
    '確定釦：True(ロック解除)する。
    cmdKakutei.Enabled = True
    
    'メニュー画面へ戻る釦：True(ロック解除)する。
    cmdReturn.Enabled = True
    
    For iCnt = 0 To MAX_GOUKI
        'IC取扱エリア：True(ロック解除)する。
        chkMode(iCnt).Enabled = True
        '磁気取扱エリア：True(ロック解除)する。
        chkMode(iCnt + MOVE_JIKI_INDEX).Enabled = True
    Next

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psGamenSettei_Hanei
'//  機能名称  : 画面値反映処理
'//  機能概要  : 画面値をエリア又はファイルに反映する。
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
Public Sub psGamenSettei_Hanei()
    Dim iKansiAplSts As Integer '監視盤アプリ起動状態
    Dim iCnt As Integer         'カウンター
    Dim bRet As Boolean         '反映処理戻り値
    Dim iRet As Integer         'メッセージボックス戻り値
    Dim bJikiRet As Boolean     '磁気反映処理戻り値
    Dim bICRet As Boolean       'ＩＣ反映処理戻り値
    
    On Error Resume Next
   
    '監視盤アプリ起動チェックを行う。
    iKansiAplSts = CheckAppStart(PROC_KANRI)
    If iKansiAplSts <> 0 Then
       
       '監視盤起動時:自改設定エリア更新処理を行う
        For iCnt = 0 To MAX_GOUKI
            'IC取扱券種値取得エリアチェック：値不定以外
            If iIC_KenshuMode_Sts(iCnt) <> HUTEI Then
               bRet = pfAreaSet_Sts(iCnt, IC_KENSHU)
               bUpData_Flag = True
            End If
            If iICGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '設定変更有り
               bICRet = True
            End If
                                 
            '磁気取扱券種値取得エリアチェック：値不定以外
            If iJIKI_KenshuMode_Sts(iCnt) <> HUTEI Then
               bRet = pfAreaSet_Sts(iCnt, JIKI_KENSHU)
               bUpData_Flag = True
            End If
            If iJIKIGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '設定変更有り
               bJikiRet = True
            End If
        Next
        
        If bICRet = False And bJikiRet = False And bUpData_Flag = True Then
           '更新処理異常時：処理結果(異常終了)ポップアップ画面表示
           iRet = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理結果")

           '設定変更号機フラグ：変更無しに設定
           For iCnt = 0 To MAX_GOUKI
               iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
               iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
           Next
           'エリア更新処理終了
           bUpData_Flag = False
           Exit Sub
        End If
        
        '自改設定指示を監マに送信する。
        bRet = pfSendMail
        If bRet = False Then
           '送信異常：「自改設定指示：送信異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_MAIL, KENSHUMODE_SETTEI_JIKAIMAIL_ERROR, 0)
        Else
           '送信正常：「自改設定指示：送信正常」ログ出力
           Call sLogTraceReq(LTYP_NORMAL, L3AN_MAIL, KENSHUMODE_SETTEI_JIKAIMAIL_OK, 0)
        End If
                
    Else
       '監視盤未起動時：自改設定ファイルより値取得
       For iCnt = 0 To MAX_GOUKI
           'IC取扱券種値取得エリアチェック：値不定以外
           If iIC_KenshuMode_Sts(iCnt) <> HUTEI Then
              bRet = pfFileSet_Sts(iCnt, IC_KENSHU)
              bUpData_Flag = True
           End If
           If iICGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '設定変更有り
               bICRet = True
           End If
           '磁気取扱券種値取得エリアチェック：値不定以外
           If iJIKI_KenshuMode_Sts(iCnt) <> HUTEI Then
              bRet = pfFileSet_Sts(iCnt, JIKI_KENSHU)
              bUpData_Flag = True
           End If
           If iJIKIGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '設定変更有り
               bJikiRet = True
           End If
        Next
        
        If bICRet = False And bJikiRet = False And bUpData_Flag = True Then
           '更新処理異常時：処理結果(異常終了)ポップアップ画面表示
           iRet = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "反映処理結果")
           '設定変更号機フラグ：変更無しに設定
            For iCnt = 0 To MAX_GOUKI
                iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
                iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
            Next
            'エリア更新処理終了
            bUpData_Flag = False
            Exit Sub
        End If
    End If
      
    '設定変更号機フラグ：変更無しに設定
    For iCnt = 0 To MAX_GOUKI
        iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
        iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
    Next
    'エリア更新処理終了
    bUpData_Flag = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfAreaSet_Sts
'//  機能名称  : 自改設定エリアにIC/磁気取扱の現在値を設定処理(エリア参照)
'//  機能概要  : IC/磁気取扱の現在値の設定を行う。
'//
'//              型        名称      意味
'//  引数      : Integer　iICSts 　　[OUT]IC取扱現在値
'//  引数      : Integer　iJIKISts 　[OUT]磁気取扱現在値
'//              Integer　iGouki  　 [IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfAreaSet_Sts(iGouki As Integer, iUpData_ID As Integer) As Boolean
    Dim strMutexName    As String           'ミューテックス名
    Dim lngMuHandle     As Long             '排他処理用ハンドル
    Dim iAreaSts        As Integer          'エリア値

    On Error Resume Next
    
    Set Idinf_JikaiSettei = New IdInfProc              '自改設定エリア
    '自改設定エリアをオープンする。
    Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
    Idinf_JikaiSettei.IdOpen
    If Idinf_JikaiSettei.Errsts <> 0 Then
      'データ参照異常時
      Set Idinf_JikaiSettei = Nothing
      '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
      pfAreaSet_Sts = False
      Exit Function
    End If
    
    '自改設定エリアをＬＯＣＫする。
    Idinf_JikaiSettei.IdLock
    If Idinf_JikaiSettei.Errsts <> 0 Then
       Idinf_JikaiSettei.IdFree
       'データ参照異常時
       Set Idinf_JikaiSettei = Nothing
       '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       pfAreaSet_Sts = False
       Exit Function
     End If
     
     If iUpData_ID = IC_KENSHU Then
        'IC取扱の内容を読み込む。
        Idinf_JikaiSettei.id = IdGate.IC_TORIATUKAI_KENSHU_STS
        Idinf_JikaiSettei.SetICM_Sts iGouki, iIC_KenshuMode_Sts(iGouki)
        If Idinf_JikaiSettei.Errsts <> 0 Then
           '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           pfAreaSet_Sts = False
        Else
           iICGOUKI_SETTEI(iGouki) = SETTEI_ARI
        End If
     Else
       '磁気取扱の内容を読み込む。
       Idinf_JikaiSettei.id = IdGate.JIKI_TORIATUKAI_KENSHU_STS
       Idinf_JikaiSettei.SetICM_Sts iGouki, iJIKI_KenshuMode_Sts(iGouki)
       If Idinf_JikaiSettei.Errsts <> 0 Then
          '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          pfAreaSet_Sts = False
       Else
          iJIKIGOUKI_SETTEI(iGouki) = SETTEI_ARI
       End If
     End If
     
     Idinf_JikaiSettei.IdFree
     Set Idinf_JikaiSettei = Nothing
     pfAreaSet_Sts = True
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfFileSet_Sts
'//  機能名称  : IC/磁気取扱の現在値設定処理(ファイル参照)
'//  機能概要  : IC/磁気取扱の現在値を自改設定ファイルに設定する。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfFileSet_Sts(iGouki As Integer, iUpData_ID As Integer) As Boolean
    Dim iAreaSts        As Integer          '自改設定ファイル状態値
    Dim lSts            As Long             '関数戻り値
    Dim udtAreaR255     As GATE_INFO        '読込み用エリア（255設定用）
    Dim lngSts          As Long             'ヒットエリアID
    Dim lngLoop1        As Long             'カウンター
    Dim lngHandle       As Long             'ハンドル
    Dim FileName        As String           'ファイル有無チェック
    Dim lngRet          As Long             '戻り値
    Dim bRet            As Boolean          '読み込み結果戻り値
    Dim sSetteiFile     As String           'ファイルパス
    Dim udtAreaR255Work As GATE_INFO        '読込み用エリア（ポインタ移動用）
    Dim iUpData_Sts     As Integer          '設定値
   
    On Error Resume Next
     
    '自改設定ファイルをオープン
    lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       'オープン異常時は参照不可のため参照異常
       '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       pfFileSet_Sts = False
       Exit Function
    End If
        
    '自改設定ファイル読み込み
    For lngLoop1 = 0 To iGouki
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           'ハンドルのクローズ
           Call CloseHandle(lngHandle)
           '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           pfFileSet_Sts = False
           Exit Function
        End If
    Next
        
    'ハンドルのクローズ
    Call CloseHandle(lngHandle)
        
    'IC取扱：ID検索
    If iUpData_ID = IC_KENSHU Then
       lngSts = SerchId(udtAreaR255, IdGate.IC_TORIATUKAI_KENSHU_STS)
       If lngSts >= 0 Then
       'IDが有った場合
          iUpData_Sts = iIC_KenshuMode_Sts(iGouki)
          udtAreaR255.GateInfo(lngSts).bytDATA(0) = iUpData_Sts
       Else
          ' 該当ＩＤ無しの場合：何もしない
          pfFileSet_Sts = False
       End If
    Else
      '磁気取扱：ID検索
      lngSts = SerchId(udtAreaR255, IdGate.JIKI_TORIATUKAI_KENSHU_STS)
      If lngSts >= 0 Then
         'IDが有った場合
         iUpData_Sts = iJIKI_KenshuMode_Sts(iGouki)
         udtAreaR255.GateInfo(lngSts).bytDATA(0) = iUpData_Sts
      Else
         ' 該当ＩＤ無しの場合：何もしない
         pfFileSet_Sts = False
      End If
    End If

    '自改設定ファイルをオープン
    lngHandle = CreateFile(G_SETTEI_FILE, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    'ファイルオープンが正常に行われたか？
    If lngHandle = INVALID_HANDLE_VALUE Then
       '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       pfFileSet_Sts = False
       Exit Function
    End If
     
    'ファイルポインタ移動のための読み込み
     For lngLoop1 = 0 To iGouki - 1
         bRet = ReadFile(lngHandle, udtAreaR255Work, LenB(udtAreaR255Work), lngRet, 0)
         If bRet = False Then
            '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
            Call CloseHandle(lngHandle)
            pfFileSet_Sts = False
            Exit Function
         End If
     Next
    
    '自改設定ファイルに書き込む
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       '「取扱券種モード設定画面：エリア・ファイル参照異常」ログ出力
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Call CloseHandle(lngHandle)
       pfFileSet_Sts = False
       Exit Function
    End If
    
    'ハンドルのクローズ
     Call CloseHandle(lngHandle)
    
    '設定変更号機フラグ設定有り
    If iUpData_ID = IC_KENSHU Then
       iICGOUKI_SETTEI(iGouki) = SETTEI_ARI
    Else
       iJIKIGOUKI_SETTEI(iGouki) = SETTEI_ARI
    End If

    pfFileSet_Sts = True
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfSendMail
'//  機能名称  : 「自改設定指示」送信
'//  機能概要  : IC/磁気取扱の変更を通知する。
'//
'//              型        名称      意味
'//  引数      : Integer　iJikaiSts [OUT]表示ステータス
'//              Integer　iGouki  　[IN]処理対象号機番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSendMail() As Boolean
    
    Dim udtMail     As MAIL_GATE_SET_ORD    '自改設定指示メール送信エリア
    Dim lngRet      As Long                 '関数戻り値
    Dim intCnt      As Integer              'カウンタ

    On Error Resume Next

    '共通ヘッダ編集
    udtMail.mlHeader.dwId = ML_ID_GATE_SET_ORD
    udtMail.mlHeader.dwSize = MlSize.GATE_SET_ORD
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    
    'エリア種別を設定
    udtMail.dwCmnFile = G_SETTEI_FILE_NO
    
    '設定情報
    For intCnt = 0 To MAX_GATE_NO - 1
        If iICGOUKI_SETTEI(intCnt) = SETTEI_ARI Or iJIKIGOUKI_SETTEI(intCnt) = SETTEI_ARI Then
            udtMail.dwGateSet(intCnt) = 1
        Else
            udtMail.dwGateSet(intCnt) = 0
        End If
    Next intCnt

    'メール送信
    pfSendMail = DssSendMail(MAIL_SLOT_KANMA, MlSize.GATE_SET_ORD, udtMail.mlHeader)

End Function

