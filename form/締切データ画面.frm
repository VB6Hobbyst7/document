VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmShimekiriData 
   BorderStyle     =   0  'なし
   Caption         =   "稼働・メンテデータ収集（次世代自動改札機）"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdRemove 
      Caption         =   "媒体取外"
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
      Left            =   6480
      TabIndex        =   299
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   480
      Top             =   8400
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  データ収集・出力    画面へ戻る"
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
      Left            =   9360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   10186
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   970
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " ○○○○○○ ○○○○○○"
      TabPicture(0)   =   "締切データ画面.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdShushu(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOutput(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOffLine(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdReOutput(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   " ○○○○○○ ○○○○○○"
      TabPicture(1)   =   "締切データ画面.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdReOutput(1)"
      Tab(1).Control(1)=   "cmdOffLine(1)"
      Tab(1).Control(2)=   "cmdOutput(1)"
      Tab(1).Control(3)=   "cmdShushu(1)"
      Tab(1).Control(4)=   "Frame2(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   " ○○○○○○ ○○○○○○"
      TabPicture(2)   =   "締切データ画面.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdReOutput(2)"
      Tab(2).Control(1)=   "cmdOffLine(2)"
      Tab(2).Control(2)=   "cmdOutput(2)"
      Tab(2).Control(3)=   "cmdShushu(2)"
      Tab(2).Control(4)=   "Frame2(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   " ○○○○○○ ○○○○○○"
      TabPicture(3)   =   "締切データ画面.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdReOutput(3)"
      Tab(3).Control(1)=   "cmdOffLine(3)"
      Tab(3).Control(2)=   "cmdOutput(3)"
      Tab(3).Control(3)=   "cmdShushu(3)"
      Tab(3).Control(4)=   "Frame2(3)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   " ○○○○○○ ○○○○○○"
      TabPicture(4)   =   "締切データ画面.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdReOutput(4)"
      Tab(4).Control(1)=   "cmdOffLine(4)"
      Tab(4).Control(2)=   "cmdOutput(4)"
      Tab(4).Control(3)=   "cmdShushu(4)"
      Tab(4).Control(4)=   "Frame2(4)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   " ○○○○○○ ○○○○○○"
      TabPicture(5)   =   "締切データ画面.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdReOutput(5)"
      Tab(5).Control(1)=   "cmdOffLine(5)"
      Tab(5).Control(2)=   "cmdOutput(5)"
      Tab(5).Control(3)=   "cmdShushu(5)"
      Tab(5).Control(4)=   "Frame2(5)"
      Tab(5).ControlCount=   5
      Begin VB.CommandButton cmdReOutput 
         Caption         =   $"締切データ画面.frx":00A8
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
         Index           =   5
         Left            =   -66720
         TabIndex        =   335
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOffLine 
         Caption         =   $"締切データ画面.frx":00C8
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
         Index           =   5
         Left            =   -69360
         TabIndex        =   334
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "  統合監視盤    締切処理開始"
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
         Index           =   5
         Left            =   -72000
         TabIndex        =   333
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdShushu 
         Caption         =   "収集"
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
         Index           =   5
         Left            =   -74640
         TabIndex        =   332
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdReOutput 
         Caption         =   $"締切データ画面.frx":00E6
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
         Index           =   4
         Left            =   -66720
         TabIndex        =   331
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOffLine 
         Caption         =   $"締切データ画面.frx":0106
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
         Index           =   4
         Left            =   -69360
         TabIndex        =   330
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "  統合監視盤    締切処理開始"
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
         Index           =   4
         Left            =   -72000
         TabIndex        =   329
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdShushu 
         Caption         =   "収集"
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
         Index           =   4
         Left            =   -74640
         TabIndex        =   328
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdReOutput 
         Caption         =   $"締切データ画面.frx":0124
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
         Index           =   3
         Left            =   -66720
         TabIndex        =   327
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOffLine 
         Caption         =   $"締切データ画面.frx":0144
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
         Index           =   3
         Left            =   -69360
         TabIndex        =   326
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "  統合監視盤    締切処理開始"
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
         Index           =   3
         Left            =   -72000
         TabIndex        =   325
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdShushu 
         Caption         =   "収集"
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
         Index           =   3
         Left            =   -74640
         TabIndex        =   324
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdReOutput 
         Caption         =   $"締切データ画面.frx":0162
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
         Index           =   2
         Left            =   -66720
         TabIndex        =   323
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOffLine 
         Caption         =   $"締切データ画面.frx":0182
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
         Index           =   2
         Left            =   -69360
         TabIndex        =   322
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "  統合監視盤    締切処理開始"
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
         Index           =   2
         Left            =   -72000
         TabIndex        =   321
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdShushu 
         Caption         =   "収集"
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
         Index           =   2
         Left            =   -74640
         TabIndex        =   320
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdReOutput 
         Caption         =   $"締切データ画面.frx":01A0
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
         Index           =   1
         Left            =   -66720
         TabIndex        =   319
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOffLine 
         Caption         =   $"締切データ画面.frx":01C0
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
         Index           =   1
         Left            =   -69360
         TabIndex        =   318
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "  統合監視盤    締切処理開始"
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
         Index           =   1
         Left            =   -72000
         TabIndex        =   317
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdShushu 
         Caption         =   "収集"
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
         Index           =   1
         Left            =   -74640
         TabIndex        =   316
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdReOutput 
         Caption         =   $"締切データ画面.frx":01DE
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
         Index           =   0
         Left            =   8280
         TabIndex        =   315
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOffLine 
         Caption         =   $"締切データ画面.frx":01FE
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
         Index           =   0
         Left            =   5640
         TabIndex        =   314
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "  統合監視盤    締切処理開始"
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
         Index           =   0
         Left            =   3000
         TabIndex        =   313
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdShushu 
         Caption         =   "収集"
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
         Index           =   0
         Left            =   360
         TabIndex        =   312
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "収集結果"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   5
         Left            =   -74880
         TabIndex        =   250
         Top             =   120
         Width           =   10935
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   95
            Left            =   9480
            TabIndex        =   297
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   95
            Left            =   9600
            TabIndex        =   296
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   94
            Left            =   8160
            TabIndex        =   294
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   94
            Left            =   8280
            TabIndex        =   293
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   93
            Left            =   6840
            TabIndex        =   291
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   93
            Left            =   6960
            TabIndex        =   290
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   92
            Left            =   5520
            TabIndex        =   288
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   92
            Left            =   5640
            TabIndex        =   287
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   91
            Left            =   4200
            TabIndex        =   285
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   91
            Left            =   4320
            TabIndex        =   284
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   90
            Left            =   2880
            TabIndex        =   282
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   90
            Left            =   3000
            TabIndex        =   281
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   89
            Left            =   1560
            TabIndex        =   279
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   89
            Left            =   1680
            TabIndex        =   278
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   88
            Left            =   240
            TabIndex        =   276
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   88
            Left            =   360
            TabIndex        =   275
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   87
            Left            =   9480
            TabIndex        =   273
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   87
            Left            =   9600
            TabIndex        =   272
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   86
            Left            =   8160
            TabIndex        =   270
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   86
            Left            =   8280
            TabIndex        =   269
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   85
            Left            =   6840
            TabIndex        =   267
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   85
            Left            =   6960
            TabIndex        =   266
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   84
            Left            =   5520
            TabIndex        =   264
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   84
            Left            =   5640
            TabIndex        =   263
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   83
            Left            =   4200
            TabIndex        =   261
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   83
            Left            =   4320
            TabIndex        =   260
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   82
            Left            =   2880
            TabIndex        =   258
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   82
            Left            =   3000
            TabIndex        =   257
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   81
            Left            =   1560
            TabIndex        =   255
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   81
            Left            =   1680
            TabIndex        =   254
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   80
            Left            =   240
            TabIndex        =   252
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   80
            Left            =   360
            TabIndex        =   251
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   80
            Left            =   240
            TabIndex        =   253
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   81
            Left            =   1560
            TabIndex        =   256
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   82
            Left            =   2880
            TabIndex        =   259
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   83
            Left            =   4200
            TabIndex        =   262
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   84
            Left            =   5520
            TabIndex        =   265
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   85
            Left            =   6840
            TabIndex        =   268
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   86
            Left            =   8160
            TabIndex        =   271
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   87
            Left            =   9480
            TabIndex        =   274
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   88
            Left            =   240
            TabIndex        =   277
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   89
            Left            =   1560
            TabIndex        =   280
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   90
            Left            =   2880
            TabIndex        =   283
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   91
            Left            =   4200
            TabIndex        =   286
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   92
            Left            =   5520
            TabIndex        =   289
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   93
            Left            =   6840
            TabIndex        =   292
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   94
            Left            =   8160
            TabIndex        =   295
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   95
            Left            =   9480
            TabIndex        =   298
            Top             =   2400
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "収集結果"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   4
         Left            =   -74880
         TabIndex        =   201
         Top             =   120
         Width           =   10935
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   79
            Left            =   9600
            TabIndex        =   248
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   79
            Left            =   9480
            TabIndex        =   247
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   78
            Left            =   8280
            TabIndex        =   245
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   78
            Left            =   8160
            TabIndex        =   244
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   77
            Left            =   6960
            TabIndex        =   242
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   77
            Left            =   6840
            TabIndex        =   241
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   76
            Left            =   5640
            TabIndex        =   239
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   76
            Left            =   5520
            TabIndex        =   238
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   75
            Left            =   4320
            TabIndex        =   236
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   75
            Left            =   4200
            TabIndex        =   235
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   74
            Left            =   3000
            TabIndex        =   233
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   74
            Left            =   2880
            TabIndex        =   232
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   73
            Left            =   1680
            TabIndex        =   230
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   73
            Left            =   1560
            TabIndex        =   229
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   72
            Left            =   360
            TabIndex        =   227
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   72
            Left            =   240
            TabIndex        =   226
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   71
            Left            =   9600
            TabIndex        =   224
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   71
            Left            =   9480
            TabIndex        =   223
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   70
            Left            =   8280
            TabIndex        =   221
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   70
            Left            =   8160
            TabIndex        =   220
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   69
            Left            =   6960
            TabIndex        =   218
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   69
            Left            =   6840
            TabIndex        =   217
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   68
            Left            =   5640
            TabIndex        =   215
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   68
            Left            =   5520
            TabIndex        =   214
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   67
            Left            =   4320
            TabIndex        =   212
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   67
            Left            =   4200
            TabIndex        =   211
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   66
            Left            =   3000
            TabIndex        =   209
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   66
            Left            =   2880
            TabIndex        =   208
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   65
            Left            =   1680
            TabIndex        =   206
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   65
            Left            =   1560
            TabIndex        =   205
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   64
            Left            =   360
            TabIndex        =   203
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   64
            Left            =   240
            TabIndex        =   202
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   64
            Left            =   240
            TabIndex        =   204
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   65
            Left            =   1560
            TabIndex        =   207
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   66
            Left            =   2880
            TabIndex        =   210
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   67
            Left            =   4200
            TabIndex        =   213
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   68
            Left            =   5520
            TabIndex        =   216
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   69
            Left            =   6840
            TabIndex        =   219
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   70
            Left            =   8160
            TabIndex        =   222
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   71
            Left            =   9480
            TabIndex        =   225
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   72
            Left            =   240
            TabIndex        =   228
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   73
            Left            =   1560
            TabIndex        =   231
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   74
            Left            =   2880
            TabIndex        =   234
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   75
            Left            =   4200
            TabIndex        =   237
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   76
            Left            =   5520
            TabIndex        =   240
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   77
            Left            =   6840
            TabIndex        =   243
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   78
            Left            =   8160
            TabIndex        =   246
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   79
            Left            =   9480
            TabIndex        =   249
            Top             =   2400
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "収集結果"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   3
         Left            =   -74880
         TabIndex        =   152
         Top             =   120
         Width           =   10935
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   63
            Left            =   9480
            TabIndex        =   199
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   63
            Left            =   9600
            TabIndex        =   198
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   62
            Left            =   8160
            TabIndex        =   196
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   62
            Left            =   8280
            TabIndex        =   195
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   61
            Left            =   6840
            TabIndex        =   193
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   61
            Left            =   6960
            TabIndex        =   192
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   60
            Left            =   5520
            TabIndex        =   190
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   60
            Left            =   5640
            TabIndex        =   189
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   59
            Left            =   4200
            TabIndex        =   187
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   59
            Left            =   4320
            TabIndex        =   186
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   58
            Left            =   2880
            TabIndex        =   184
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   58
            Left            =   3000
            TabIndex        =   183
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   57
            Left            =   1560
            TabIndex        =   181
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   57
            Left            =   1680
            TabIndex        =   180
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   56
            Left            =   240
            TabIndex        =   178
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   56
            Left            =   360
            TabIndex        =   177
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   55
            Left            =   9480
            TabIndex        =   175
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   55
            Left            =   9600
            TabIndex        =   174
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   54
            Left            =   8160
            TabIndex        =   172
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   54
            Left            =   8280
            TabIndex        =   171
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   53
            Left            =   6840
            TabIndex        =   169
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   53
            Left            =   6960
            TabIndex        =   168
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   52
            Left            =   5520
            TabIndex        =   166
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   52
            Left            =   5640
            TabIndex        =   165
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   51
            Left            =   4200
            TabIndex        =   163
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   51
            Left            =   4320
            TabIndex        =   162
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   50
            Left            =   2880
            TabIndex        =   160
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   50
            Left            =   3000
            TabIndex        =   159
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   49
            Left            =   1560
            TabIndex        =   157
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   49
            Left            =   1680
            TabIndex        =   156
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   48
            Left            =   240
            TabIndex        =   154
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   48
            Left            =   360
            TabIndex        =   153
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   48
            Left            =   240
            TabIndex        =   155
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   49
            Left            =   1560
            TabIndex        =   158
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   50
            Left            =   2880
            TabIndex        =   161
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   51
            Left            =   4200
            TabIndex        =   164
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   52
            Left            =   5520
            TabIndex        =   167
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   53
            Left            =   6840
            TabIndex        =   170
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   54
            Left            =   8160
            TabIndex        =   173
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   55
            Left            =   9480
            TabIndex        =   176
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   56
            Left            =   240
            TabIndex        =   179
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   57
            Left            =   1560
            TabIndex        =   182
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   58
            Left            =   2880
            TabIndex        =   185
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   59
            Left            =   4200
            TabIndex        =   188
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   60
            Left            =   5520
            TabIndex        =   191
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   61
            Left            =   6840
            TabIndex        =   194
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   62
            Left            =   8160
            TabIndex        =   197
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   63
            Left            =   9480
            TabIndex        =   200
            Top             =   2400
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "収集結果"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   2
         Left            =   -74880
         TabIndex        =   103
         Top             =   120
         Width           =   10935
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   47
            Left            =   9600
            TabIndex        =   150
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   47
            Left            =   9480
            TabIndex        =   149
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   46
            Left            =   8280
            TabIndex        =   147
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   46
            Left            =   8160
            TabIndex        =   146
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   45
            Left            =   6960
            TabIndex        =   144
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   45
            Left            =   6840
            TabIndex        =   143
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   44
            Left            =   5640
            TabIndex        =   141
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   44
            Left            =   5520
            TabIndex        =   140
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   43
            Left            =   4320
            TabIndex        =   138
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   43
            Left            =   4200
            TabIndex        =   137
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   3000
            TabIndex        =   135
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   42
            Left            =   2880
            TabIndex        =   134
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   1680
            TabIndex        =   132
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   41
            Left            =   1560
            TabIndex        =   131
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   360
            TabIndex        =   129
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   40
            Left            =   240
            TabIndex        =   128
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   9600
            TabIndex        =   126
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   39
            Left            =   9480
            TabIndex        =   125
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   8280
            TabIndex        =   123
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   38
            Left            =   8160
            TabIndex        =   122
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   6960
            TabIndex        =   120
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   37
            Left            =   6840
            TabIndex        =   119
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   5640
            TabIndex        =   117
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   36
            Left            =   5520
            TabIndex        =   116
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   4320
            TabIndex        =   114
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   35
            Left            =   4200
            TabIndex        =   113
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   3000
            TabIndex        =   111
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   34
            Left            =   2880
            TabIndex        =   110
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   1680
            TabIndex        =   108
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   33
            Left            =   1560
            TabIndex        =   107
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   360
            TabIndex        =   105
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   32
            Left            =   240
            TabIndex        =   104
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   32
            Left            =   240
            TabIndex        =   106
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   33
            Left            =   1560
            TabIndex        =   109
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   34
            Left            =   2880
            TabIndex        =   112
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   35
            Left            =   4200
            TabIndex        =   115
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   36
            Left            =   5520
            TabIndex        =   118
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   37
            Left            =   6840
            TabIndex        =   121
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   38
            Left            =   8160
            TabIndex        =   124
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   39
            Left            =   9480
            TabIndex        =   127
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   40
            Left            =   240
            TabIndex        =   130
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   41
            Left            =   1560
            TabIndex        =   133
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   42
            Left            =   2880
            TabIndex        =   136
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   43
            Left            =   4200
            TabIndex        =   139
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   44
            Left            =   5520
            TabIndex        =   142
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   45
            Left            =   6840
            TabIndex        =   145
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   46
            Left            =   8160
            TabIndex        =   148
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   47
            Left            =   9480
            TabIndex        =   151
            Top             =   2400
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "収集結果"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   1
         Left            =   -74880
         TabIndex        =   54
         Top             =   120
         Width           =   10935
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   9600
            TabIndex        =   102
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   8280
            TabIndex        =   101
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   6960
            TabIndex        =   100
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   5640
            TabIndex        =   99
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   4320
            TabIndex        =   98
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   3000
            TabIndex        =   97
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   1680
            TabIndex        =   96
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   360
            TabIndex        =   95
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   9600
            TabIndex        =   94
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   8280
            TabIndex        =   93
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   6960
            TabIndex        =   92
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   5640
            TabIndex        =   91
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   4320
            TabIndex        =   90
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   3000
            TabIndex        =   89
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   1680
            TabIndex        =   88
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   31
            Left            =   9480
            TabIndex        =   72
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   30
            Left            =   8160
            TabIndex        =   70
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   29
            Left            =   6840
            TabIndex        =   69
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   28
            Left            =   5520
            TabIndex        =   68
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   27
            Left            =   4200
            TabIndex        =   67
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   26
            Left            =   2880
            TabIndex        =   66
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   25
            Left            =   1560
            TabIndex        =   65
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   24
            Left            =   240
            TabIndex        =   64
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   23
            Left            =   9480
            TabIndex        =   63
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   22
            Left            =   8160
            TabIndex        =   62
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   21
            Left            =   6840
            TabIndex        =   61
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   20
            Left            =   5520
            TabIndex        =   60
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   19
            Left            =   4200
            TabIndex        =   59
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   18
            Left            =   2880
            TabIndex        =   58
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   17
            Left            =   1560
            TabIndex        =   57
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   360
            TabIndex        =   56
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   16
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   31
            Left            =   9480
            TabIndex        =   87
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   30
            Left            =   8160
            TabIndex        =   86
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   29
            Left            =   6840
            TabIndex        =   85
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   28
            Left            =   5520
            TabIndex        =   84
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   27
            Left            =   4200
            TabIndex        =   83
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   26
            Left            =   2880
            TabIndex        =   82
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   25
            Left            =   1560
            TabIndex        =   81
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   24
            Left            =   240
            TabIndex        =   80
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   23
            Left            =   9480
            TabIndex        =   79
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   22
            Left            =   8160
            TabIndex        =   78
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   21
            Left            =   6840
            TabIndex        =   77
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   20
            Left            =   5520
            TabIndex        =   76
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   19
            Left            =   4200
            TabIndex        =   75
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   18
            Left            =   2880
            TabIndex        =   74
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   17
            Left            =   1560
            TabIndex        =   73
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   16
            Left            =   240
            TabIndex        =   71
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "収集結果"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10935
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   15
            Left            =   9480
            TabIndex        =   53
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   9600
            TabIndex        =   51
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   14
            Left            =   8160
            TabIndex        =   50
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   8280
            TabIndex        =   48
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   13
            Left            =   6840
            TabIndex        =   47
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   6960
            TabIndex        =   45
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   12
            Left            =   5520
            TabIndex        =   44
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   5640
            TabIndex        =   42
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   11
            Left            =   4200
            TabIndex        =   41
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   4320
            TabIndex        =   39
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   10
            Left            =   2880
            TabIndex        =   38
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   3000
            TabIndex        =   36
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   9
            Left            =   1560
            TabIndex        =   35
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1680
            TabIndex        =   33
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   8
            Left            =   240
            TabIndex        =   32
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   30
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   7
            Left            =   9480
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   9600
            TabIndex        =   27
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Index           =   6
            Left            =   8160
            TabIndex        =   26
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   8280
            TabIndex        =   24
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Left            =   6840
            TabIndex        =   23
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   6960
            TabIndex        =   21
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Left            =   5520
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   5640
            TabIndex        =   18
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Left            =   4200
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   15
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Left            =   2880
            TabIndex        =   14
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   12
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Left            =   1560
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   9
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '中央揃え
            BackStyle       =   0  '透明
            Caption         =   "Z9号機"
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
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   5
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '透明
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   4
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   1
            Left            =   1560
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   2
            Left            =   2880
            TabIndex        =   13
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   3
            Left            =   4200
            TabIndex        =   16
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   4
            Left            =   5520
            TabIndex        =   19
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   5
            Left            =   6840
            TabIndex        =   22
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   6
            Left            =   8160
            TabIndex        =   25
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   7
            Left            =   9480
            TabIndex        =   28
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   8
            Left            =   240
            TabIndex        =   31
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   9
            Left            =   1560
            TabIndex        =   34
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   10
            Left            =   2880
            TabIndex        =   37
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   11
            Left            =   4200
            TabIndex        =   40
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   12
            Left            =   5520
            TabIndex        =   43
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   13
            Left            =   6840
            TabIndex        =   46
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   14
            Left            =   8160
            TabIndex        =   49
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatusBack 
            Alignment       =   2  '中央揃え
            Appearance      =   0  'ﾌﾗｯﾄ
            BorderStyle     =   1  '実線
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   15
            Left            =   9480
            TabIndex        =   52
            Top             =   2400
            Width           =   1215
         End
      End
   End
   Begin VB.Label lblMisouStatus 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "未送あり"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9960
      TabIndex        =   311
      Top             =   6960
      Width           =   1440
   End
   Begin VB.Label lblMisouMark 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   9960
      TabIndex        =   310
      Top             =   6840
      Width           =   1440
   End
   Begin VB.Label lblMisouStatus 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "未送あり"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   309
      Top             =   6960
      Width           =   1440
   End
   Begin VB.Label lblMisouMark 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   8040
      TabIndex        =   308
      Top             =   6840
      Width           =   1440
   End
   Begin VB.Label lblMisouStatus 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "未送あり"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   307
      Top             =   6960
      Width           =   1440
   End
   Begin VB.Label lblMisouMark 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   6120
      TabIndex        =   306
      Top             =   6840
      Width           =   1440
   End
   Begin VB.Label lblMisouStatus 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "未送あり"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   305
      Top             =   6960
      Width           =   1440
   End
   Begin VB.Label lblMisouMark 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   4320
      TabIndex        =   304
      Top             =   6840
      Width           =   1440
   End
   Begin VB.Label lblMisouStatus 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "未送あり"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   303
      Top             =   6960
      Width           =   1440
   End
   Begin VB.Label lblMisouMark 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   302
      Top             =   6840
      Width           =   1440
   End
   Begin VB.Label lblMisouStatus 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "未送あり"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   301
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblMisouMark 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   300
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "統合監視盤締切"
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
Attribute VB_Name = "frmShimekiriData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：frmShimekiriData.frm
'//  パッケージ名：締切画面
'//
'//  概要：締切収集画面
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//                 ・稼働・メンテデータ収集画面(frmSyusyu.frm)を流用
'//                 ・フェーズ２対応【Mainte_05_03】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000     'メールタイマのインターバル値
Public glbFilePath  As String             'ファイルパス     'V1.12.0.1 ADD

' EG20 V5.6.0.1 追加開始
Public gbShimekiriResult As Boolean       ' オフライン出力結果
' EG20 V5.6.0.1 追加終了
Public glShimekiriType As Long            ' オフライン出力種別  ' EG20 V6.3.0.1【機能見直し】追加

Private mintMaxIndex As Integer

Private Type SHIMEKIRI_STATUS
    intStatus As Integer    'ステータス
    strCaption As String    'ボタン文言
    strColor As String      'ボタン色
End Type
Private mudtBtn_Status() As SHIMEKIRI_STATUS

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : cmdOffLine_Click
'//  機能名称  : 締切データオフライン出力ボタン押下時処理
'//  機能概要  : 締切データオフライン出力ボタン押下時処理を行う。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub cmdOffLine_Click()                     ' EG20 V6.3.0.1削除
Private Sub cmdOffLine_Click(Index As Integer)      ' EG20 V6.3.0.1追加
    
    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim objFile As File                                 ' ファイルオブジェクト
    Dim bProceed As Boolean                             ' 締切処理開始フラグ
    Dim nListCnt As Integer                             ' ファイル格納数
    Dim szSaveFolder As String                          ' 保存先フォルダ
    Dim szFileName As String                            ' ファイル名
    Dim iResponse As Integer
    
    On Error GoTo ErrorHandler                          ' エラーハンドルの登録
    
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 初期化
    ReDim Preserve gOfflineFileList(0)
    bProceed = False
    nListCnt = 0
    gbShimekiriResult = False
    glShimekiriType = 0                                     ' EG20 V6.3.0.1【機能見直し】追加
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 締切出力データは存在するか？（D:\KANSI\SHUKEI\SEND_DATA\SIMEKIRI##.DAT）
' EG20 V6.3.0.1【機能見直し】追加開始
    szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(Index + 1, "0#"))
    If objFso.FileExists(szFileName) = True Then              ' ファイル名の取得チェック
        nListCnt = nListCnt + 1                             ' ファイル数のカウンタをアップする
        ReDim Preserve gOfflineFileList(nListCnt)           ' ファイル名格納エリアを拡張する
        gOfflineFileList(nListCnt - 1) = szFileName         ' ファイルパスを格納
        bProceed = True
    End If
' EG20 V6.3.0.1【機能見直し】追加終了
    
' EG20 V6.3.0.1【機能見直し】削除開始
'    For Each objFile In objFso.GetFolder(PATH_SHUKEI_SEND).files    ' ループを開始
'        If objFso.FileExists(objFile.Path) = True Then              ' ファイル名の取得チェック
'            ' ファイル名を取得
'            If InStr(objFile.Name, FILENAME_SIMEKIRIDAT) <> 0 Then
'                nListCnt = nListCnt + 1                             ' ファイル数のカウンタをアップする
'                ReDim Preserve gOfflineFileList(nListCnt)           ' ファイル名格納エリアを拡張する
'                gOfflineFileList(nListCnt - 1) = objFile.Path       ' ファイルパスを格納
'                bProceed = True
'            End If
'        End If
'    Next
' EG20 V6.3.0.1【機能見直し】削除終了
        
    If bProceed = False Then
        Call MsgBox("締切出力データが統合監視盤内にありません。" & vbCrLf & _
                    "締切データのオフライン出力処理を開始できません。", _
                    vbExclamation, "データ無警告")
        Set objFso = Nothing
        Set objFile = Nothing
        Exit Sub
    End If
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 出力先フォルダの選択
    szSaveFolder = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    ' 指定フォルダなし
    If Len(szSaveFolder) = 0 Then
        Set objFso = Nothing
        Set objFile = Nothing
        Exit Sub
    End If
    
    'コピー先フォルダの有無確認
    If objFso.FolderExists(szSaveFolder) = False Then
        'コピー先フォルダ作成
        objFso.CreateFolder (szSaveFolder)
    End If

    glbFilePath = szSaveFolder
        
    Call sCmdBtnEnabled(False)

    ' オフライン出力中画面を表示
    glShimekiriType = 1                                     ' EG20 V6.3.0.1【機能見直し】追加
    frmShimekiriOfflineOut.Show vbModal
    
    glShimekiriType = 0                                     ' EG20 V6.3.0.1【機能見直し】追加
    ' オフライン処理結果が正常の場合
    If gbShimekiriResult = True Then
        iResponse = MsgBox("統合監視盤締切データをクリアしますがよろしいですか？", _
                            vbOKCancel + vbQuestion, "締切データクリア確認")
        If iResponse = vbOK Then
            ' /////////////////////////////////////////////////////////////////
            ' // ファイル削除処理
            For nListCnt = 0 To UBound(gOfflineFileList) - 1    ' ファイルリスト数

                szFileName = gOfflineFileList(nListCnt)         ' ファイル名の取得
                Kill szFileName                                 ' ファイルの削除
            Next
' EG20 V6.3.0.1【機能見直し】削除開始
'            If objFso.FileExists(PATH_SHUKEI_SHIMEDAT01) = True Then
'                Kill PATH_SHUKEI_SHIMEDAT01
'            End If
'            If objFso.FileExists(PATH_SHUKEI_SHIMEDAT02) = True Then
'                Kill PATH_SHUKEI_SHIMEDAT02
'            End If
'            If objFso.FileExists(PATH_SHUKEI_SHIMEDAT03) = True Then
'                Kill PATH_SHUKEI_SHIMEDAT03
'            End If
'            If objFso.FileExists(PATH_SHUKEI_SHIMEDAT04) = True Then
'                Kill PATH_SHUKEI_SHIMEDAT04
'            End If
'            If objFso.FileExists(PATH_SHUKEI_SHIMEDAT05) = True Then
'                Kill PATH_SHUKEI_SHIMEDAT05
'            End If
'            If objFso.FileExists(PATH_SHUKEI_SHIMEDAT06) = True Then
'                Kill PATH_SHUKEI_SHIMEDAT06
'            End If
' EG20 V6.3.0.1【機能見直し】削除終了
' EG20 V6.3.0.1【機能見直し】追加開始
            szFileName = Replace(PATH_SHUKEI_SHIMEDAT, "##", Format(Index + 1, "0#"))
            If objFso.FileExists(szFileName) = True Then
                Kill szFileName
            End If
' EG20 V6.3.0.1【機能見直し】追加終了
            Call MsgBox("正常終了しました。", vbOKOnly + vbInformation, "クリア出力結果")
        End If
    End If
    
    Call ChDrive("D")
    Call ChDir("D:\")
    
    ReDim Preserve gOfflineFileList(0)
    Call psCheckMisouStatus                 ' EG20 V6.3.0.1追加
    Call sCmdBtnEnabled(True)
    Set objFso = Nothing
    Set objFile = Nothing
    
' EG20 V6.3.0.1削除開始
'' EG20 V5.10.0.1追加開始
'    Call funcCheckShimekiri
'    If gbShimekiriResult = True Then
'        ' オフライン出力ボタンを押下可能
'        cmdOffLine.Enabled = True
'    Else
'        ' オフライン出力ボタンを押下不可能
'        cmdOffLine.Enabled = False
'    End If
' EG20 V5.10.0.1追加終了
'' EG20 V6.3.0.1削除終了
    
    Exit Sub

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:
    Call MsgBox("異常終了しました。", vbOKOnly, "オフライン出力結果")

    Set objFso = Nothing
    Set objFile = Nothing
    glShimekiriType = 0                                     ' EG20 V6.3.0.1【機能見直し】追加
    Call sCmdBtnEnabled(True)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : CmdRemove_Click
'//  機能名称  : 「媒体取外」釦押下時処理
'//  機能概要  : 媒体の取り外しを行う。
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
Private Sub cmdRemove_Click()
   On Error Resume Next
   
   '「媒体取外釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '媒体取外処理
    Call pfRemove(Me)
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : cmdReOutput_Click
'//  機能名称  : 「締切データオフライン再出力」ボタン押下時処理
'//  機能概要  : 締切データオフライン再出力を行う。
'//
'//              型        名称      意味
'//  引数      : Integer   Index     ボタン番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReOutput_Click(Index As Integer)
    
    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim objFile As File                                 ' ファイルオブジェクト
    Dim bProceed As Boolean                             ' 締切処理開始フラグ
    Dim nListCnt As Integer                             ' ファイル格納数
    Dim szFolderName As String                          ' 保存先フォルダ
    Dim szSaveFolder As String                          ' 保存先フォルダ
    
    Dim szFileName As String                            ' ファイル名
    Dim iResponse As Integer
    
    On Error GoTo ErrorHandler                          ' エラーハンドルの登録
    
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 初期化
    ReDim Preserve gOfflineFileList(0)
    bProceed = False
    nListCnt = 0
    glShimekiriType = 0
    
    szFolderName = Replace(PATH_SIMEKIRIREOUT_FOLDER, "##", Format(Index + 1, "0#"))
    If objFso.FolderExists(szFolderName) = True Then
        ' /////////////////////////////////////////////////////////////////////////
        ' // 再出力データは存在するか？（F:\KANSI\OUT_DATA\CORNER##\SIME##.CSV）
        For Each objFile In objFso.GetFolder(szFolderName).files   ' ループを開始
            If objFso.FileExists(objFile.Path) = True Then         ' ファイル名の取得チェック
                ' ファイル名を取得
                If InStr(objFile.Name, "SIME") <> 0 Then
                    nListCnt = nListCnt + 1                             ' ファイル数のカウンタをアップする
                    ReDim Preserve gOfflineFileList(nListCnt)           ' ファイル名格納エリアを拡張する
                    gOfflineFileList(nListCnt - 1) = objFile.Path       ' ファイルパスを格納
                    bProceed = True
                End If
            End If
        Next
    End If
    
        
    If bProceed = False Then
        Call MsgBox("締切出力データが統合監視盤内にありません。" & vbCrLf & _
                    "締切データのオフライン出力処理を開始できません。", _
                    vbExclamation, "データ無警告")
        Set objFso = Nothing
        Set objFile = Nothing
        Exit Sub
    End If
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 出力先フォルダの選択
    szSaveFolder = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    ' 指定フォルダなし
    If Len(szSaveFolder) = 0 Then
        Set objFso = Nothing
        Set objFile = Nothing
        Exit Sub
    End If
    
    'コピー先フォルダの有無確認
    If objFso.FolderExists(szSaveFolder) = False Then
        'コピー先フォルダ作成
        objFso.CreateFolder (szSaveFolder)
    End If

    glbFilePath = szSaveFolder
        
    Call sCmdBtnEnabled(False)

    ' オフライン出力中画面を表示
    glShimekiriType = 2
    frmShimekiriOfflineOut.Show vbModal

    Call ChDrive("D")
    Call ChDir("D:\")
    
    glShimekiriType = 0
    ReDim Preserve gOfflineFileList(0)
    Call psCheckMisouStatus
    Call sCmdBtnEnabled(True)
    Set objFso = Nothing
    Set objFile = Nothing
    
    Exit Sub

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:
    Call MsgBox("異常終了しました。", vbOKOnly, "オフライン再出力結果")

    Set objFso = Nothing
    Set objFile = Nothing
    glShimekiriType = 0
    Call sCmdBtnEnabled(True)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdShushu_Click
'//  機能名称  : 「収集」ボタン押下時処理
'//  機能概要  : 締切データ収集を行う。
'//
'//              型        名称      意味
'//  引数      : Integer   Index     ボタン番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub cmdShushu_Click()                      ' EG20 V6.3.0.1削除
Private Sub cmdShushu_Click(Index As Integer)       ' EG20 V6.3.0.1追加
    
    Dim iResponse As Integer
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intIndex As Integer
    Dim blnErrGoki As Boolean
    
    '「キャンセル」ボタン押下処理は処理を終了する
    If iResponse = vbCancel Then Exit Sub
    
    On Error Resume Next
    
    Erase gintShimekiri
    
    '「締切データ収集画面：収集釦押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SHIMEKIRI_GAMEN_SYUSYU_BUTTOM, 0)
   
    
    '「締切データ収集」ポップアップを表示
    iResponse = MsgBox(vbCrLf & _
              "収集対象の改札機が、電源ＯＮ・通信正常でないと、" & _
              "収集に失敗します。" & _
              vbCrLf & vbCrLf & vbCrLf & _
              "確認してから「ＯＫ」ボタンを押して下さい。" & _
              vbCrLf & vbCrLf, _
              vbOKCancel, "確認")
              
    If iResponse = vbOK Then
        '号機指定確認
' EG20 V6.3.0.1【機能見直し】削除開始
'        For intCount = 0 To lblStatus.UBound
'            If lblStatusBack(intCount).Tag <> "0" Then
'                '有効な号機を選択状態とする
'                gintShimekiri(CInt(lblStatusBack(intCount).Tag) - 1) = TAG_STATUS.STS_SENTAKU
'            End If
'        Next intCount
' EG20 V6.3.0.1【機能見直し】削除終了
' EG20 V6.3.0.1【機能見直し】追加開始
        For intCount = 0 To 15
            intCount2 = (Index * 16) + intCount
            If lblStatusBack(intCount2).Tag <> "0" Then
                '有効な号機を選択状態とする
                gintShimekiri(CInt(lblStatusBack(intCount2).Tag) - 1) = TAG_STATUS.STS_SENTAKU
            End If
        Next intCount
' EG20 V6.3.0.1【機能見直し】追加終了
        
        'ＯＫ釦が押されたら、
        '締切データ収集中フォームを、モーダルウィンドウで表示する。
        frmShimekiriCyu.Show vbModal
        
        '処理結果を表示
        Call sSet_GokiStatus(SSTab1.Tab)
    
' EG20 V6.3.0.1【機能見直し】削除開始
'        blnErrGoki = False
'        '対象外号機存在確認
'        For intCount = 0 To lblStatus.UBound
'            If lblStatusBack(intCount).Tag <> "0" Then
'                '有効な号機を選択状態とする
'                If gintShimekiri(CInt(lblStatusBack(intCount).Tag) - 1) = TAG_STATUS.STS_MISENTAKU Then
'                    blnErrGoki = True
'                End If
'            End If
'        Next intCount
'
'        '対象外号機が存在しない場合、収集ボタンを押下可能にする。
'        If blnErrGoki = False Then
'            cmdOutput.Enabled = True
'        Else
'            cmdOutput.Enabled = False
'        End If
' EG20 V6.3.0.1【機能見直し】削除終了
' EG20 V6.3.0.1【機能見直し】追加開始
        Call psCheckMisouStatus
        Call sCmdBtnEnabled(True)
' EG20 V6.3.0.1【機能見直し】追加終了
        
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : 締切画面(アクティブ時)
'//  機能概要  : メール受信用タイマを起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
    'メール受信用タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : 締切画面(ディアクティブ時)
'//  機能概要  : メール受信用、タイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'タイマを停止する
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : 締切画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
        
    Dim bySyoAssort As Byte             'ログ用小分類
    Dim intFileNumber As Integer        'ファイル番号
    Dim strFileName As String           'ファイル名
    Dim intX() As Integer
    Dim intY() As Integer
    Dim strItmNum As String
    Dim strTemp As String
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCornerNo As Integer
    Dim intIndex As Integer
    Dim strCorner1 As String
    Dim strCorner2 As String
    Const lngBufSize = 32
    Dim nLoop As Integer                                ' ループ    ' EG20 V6.3.0.1追加
    
    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
' EG20 V6.3.0.1【機能見直し】削除開始
'    '収集前は出力ボタン押下不可
'    cmdOutput.Enabled = False
'' EG20 V5.6.0.1追加開始
'    '締切前はオフラインボタン押下不可
'    cmdOffLine.Enabled = False
'' EG20 V5.6.0.1追加終了
' EG20 V6.3.0.1【機能見直し】削除終了
    
    '「締切画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SHIMEKIRI_GAMEN_START, 0)
   
    'メール受信用のタイマ値を設定する。
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '号機情報取得
    Call gsGetGateInfo
    Call gsGetCornerName
    
' EG20 V6.3.0.1【機能見直し】追加開始
    For nLoop = 0 To UBound(gblnCornerSet)
        cmdShushu(nLoop).Enabled = False        ' 収集ボタン
        cmdOffLine(nLoop).Enabled = False       ' 締切データオフライン出力ボタン
        cmdReOutput(nLoop).Enabled = False      ' 締切データオフライン再出力ボタン
        cmdOutPut(nLoop).Enabled = False        ' 統合監視盤締切処理開始
    Next nLoop
    glShimekiriType = 0
' EG20 V6.3.0.1【機能見直し】追加終了

    'タブ数を設置コーナ数とする
    SSTab1.Tab = 0
    SSTab1.Tabs = gintCornerNum

    Erase gintShimekiri
   
    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    bySyoAssort = L3AN_FILE
    
    For intCount = 0 To UBound(gblnCornerSet)
    
        '設定ありのコーナを活性にする
        
        If gblnCornerSet(intCount) = True Then
            'コーナー名称表示
            strCorner1 = MidB(gstrCornerName(intCount), 1, 12)
            strCorner2 = MidB(gstrCornerName(intCount), 13, 24)
            SSTab1.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
        End If
    
    Next intCount
    
    '未使用のファイル番号を取得します。
    intFileNumber = FreeFile

    '設定情報ファイル名を設定する。
    strFileName = SHIMEKIRI_STATUS_FILE

    '設定情報ファイルをオープンする。
    If strFileName <> "" Then
        Open strFileName For Input As #intFileNumber
    End If

    For intCount = 0 To 1

        '設定情報ファイル名に設定されている釦設定ファイルを読む。
        Input #intFileNumber, strItmNum, strTemp, strTemp

        '最大コントロール数を変数に設定する。
        If intCount = 1 Then
            mintMaxIndex = CInt(strItmNum) - 1
        End If
    Next

    ReDim mudtBtn_Status(mintMaxIndex)

    For intCount = 0 To mintMaxIndex
        '設定情報ファイル名に設定されている釦設定ファイルを読む。
        With mudtBtn_Status(intCount)
            Input #intFileNumber, .intStatus, .strCaption, .strColor
        End With
    Next

    Close #intFileNumber


    intIndex = 0

    '設置コーナ数分ループ
    For intCount = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(intCount) = False Then
            SSTab1.TabVisible(intCount) = False
            Frame2(intCount).Visible = False
        End If

        '最大号機数分ループ
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + intCount2
            lblGokiNo(intIndex).Visible = False
            lblStatus(intIndex).Visible = False
            lblStatusBack(intIndex).Visible = False
            lblStatusBack(intIndex).Tag = "0"
        Next intCount2
        
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + (gudtSettiCorner(intCount).intGokiNo(intCount2) - 1)
            If gudtSettiCorner(intCount).intGokiNo(intCount2) > 0 Then
                lblGokiNo(intIndex).Caption = gudtSettiCorner(intCount).strDispGoki(intCount2) + "号機"
                lblStatusBack(intIndex).Tag = CStr(gudtSettiCorner(intCount).intGateNo(intCount2))  '号機番号を記録
                lblGokiNo(intIndex).Visible = True
                lblStatus(intIndex).Visible = True
                lblStatusBack(intIndex).Visible = True
            End If
        Next intCount2
        
    Next intCount
    
' EG20 V6.3.0.1削除開始
'' EG20 V5.10.0.1追加開始
'    ' 締切データが存在すれば、オフライン出力釦は活性化する。
'    Call funcCheckShimekiri                 ' EG20 V5.10.0.1追加
'    If gbShimekiriResult = True Then
'        '締切前はオフラインボタン押下可
'        cmdOffLine.Enabled = True
'    End If
'' EG20 V5.10.0.1追加終了
' EG20 V6.3.0.1削除終了
' EG20 V6.3.0.1【機能見直し】追加開始
        Call psCheckMisouStatus
        Call sCmdBtnEnabled(True)
' EG20 V6.3.0.1【機能見直し】追加終了

Exit Sub

'エラー処理
Err_LOG:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If

    'エラーログの出力
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, KADO_MENTE_SYUSYU_GAMEN_START, 0)
     
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「データ収集・出力画面に戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
On Error Resume Next
   '「締切画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SHIMEKIRI_GAMEN_END, 0)
 
    '自画面を消す。
    Unload Me
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : SSTab1_Click
'//  機能名称  : タブクリック処理
'//  機能概要  : 表示タブを変更する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub SSTab1_Click(PreviousTab As Integer)

    On Error Resume Next
    
' EG20 V6.3.0.1削除開始
'    '現在表示タブの更新
'    Call sSet_GokiStatus(SSTab1.Tab)
' EG20 V6.3.0.1削除終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : SSTab1_Click
'//  機能名称  : タブクリック処理
'//  機能概要  : 表示タブを変更する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub SSTab1_DblClick()

    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信用タイマ、タイムアップ時処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
On Error Resume Next
    '汎用メール受信処理を行う
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmShimekiriData.Caption, False
        pfFormActive (frmShimekiriData.hwnd)        ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdOutPut_Click
'//  機能名称  : 出力ボタン押下時処理
'//  機能概要  : 出力処理を行う。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-21   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.10.0.1) 2012-05-09   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub cmdOutPut_Click()                      ' EG20 V6.3.0.1削除
Private Sub cmdOutPut_Click(Index As Integer)       ' EG20 V6.3.0.1追加
    
    Dim iResponse As Integer
    Dim blnExistMishu As Boolean
    Dim intCount As Integer
    Dim intCount2 As Integer
    
    gbShimekiriResult = False                               ' EG20 V5.6.0.1追加
    
    '確認メッセージボックスを表示する。
    iResponse = MsgBox("未送の締切データを全て出力します。よろしいですか？", _
                        vbOKCancel, "出力確認")

    'OK釦が押された場合
    If iResponse = vbOK Then
        '未収集号機存在確認
        blnExistMishu = False
' EG20 V6.3.0.1【機能見直し】削除開始
'        For intCount = 0 To UBound(gintShimekiri)
'            If gintShimekiri(intCount) = TAG_STATUS.STS_MISHUSHU Then
'                blnExistMishu = True
'                Exit For
'            End If
'        Next
' EG20 V6.3.0.1【機能見直し】削除終了
' EG20 V6.3.0.1【機能見直し】追加開始
        For intCount = 0 To 15
            intCount2 = (Index * 16) + intCount
            If lblStatus(intCount2).Caption = "未収集" Then
                blnExistMishu = True
                Exit For
            End If
        Next intCount
' EG20 V6.3.0.1【機能見直し】追加終了
        
        If blnExistMishu = True Then
            iResponse = MsgBox("未収集の号機があります。" & vbCrLf & "出力してよろしいですか？", _
                                vbOKCancel, "未収集号機あり")
            If iResponse = vbCancel Then
                Exit Sub
            End If
        End If
        
        Call sCmdBtnEnabled(False)
    
        frmShimekiriOutPut.Show vbModal '収集中メッセージ表示
         
'        Call sSet_GokiStatus(SSTab1.Tab)    ' EG20 V6.3.0.1【機能見直し】削除
    End If
     
    Call psCheckMisouStatus                 ' EG20 V6.3.0.1追加
    Call sCmdBtnEnabled(True)
' EG20 V6.3.0.1削除開始
'' EG20 V5.6.0.1追加開始
'    Call funcCheckShimekiri                 ' EG20 V5.10.0.1追加
'    If gbShimekiriResult = True Then
'        ' オフライン出力ボタンを押下可能
'        cmdOffLine.Enabled = True
'    Else
'        ' オフライン出力ボタンを押下不可能
'        cmdOffLine.Enabled = False
'    End If
'' EG20 V5.6.0.1追加終了
' EG20 V6.3.0.1削除終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sCmdBtnEnabled
'//  機能名称  : ボタン活性／非活性処理
'//  機能概要  : ボタンの活性制御を行う。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-21   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)

' EG20 V6.3.0.1【機能見直し】削除開始
'    cmdShushu.Enabled = blnFlg
'    cmdReturn.Enabled = blnFlg
' EG20 V6.3.0.1【機能見直し】削除終了
' EG20 V6.3.0.1【機能見直し】追加開始
    Dim nLoop As Integer                                ' ループ
    
    cmdReturn.Enabled = blnFlg
    CmdRemove.Enabled = blnFlg
    SSTab1.Enabled = blnFlg
    
    If blnFlg = False Then
        For nLoop = 0 To UBound(gblnCornerSet)
            cmdShushu(nLoop).Enabled = blnFlg       ' 収集ボタン
            cmdOffLine(nLoop).Enabled = blnFlg      ' 締切データオフライン出力ボタン
            cmdReOutput(nLoop).Enabled = blnFlg     ' 締切データオフライン再出力ボタン
            cmdOutPut(nLoop).Enabled = blnFlg       ' 統合監視盤締切処理開始
        Next nLoop
    Else
        For nLoop = 0 To UBound(gblnCornerSet)
            cmdShushu(nLoop).Enabled = blnFlg       ' 収集ボタン
        Next nLoop
        Call funcCheckShimekiri                     ' 締切データオフライン出力ボタン
        Call psCheckReoutStatus                     ' 締切データオフライン再出力ボタン
        Call psCheckShimeKaishiStatus               ' 統合監視盤締切処理開始
    End If
' EG20 V6.3.0.1【機能見直し】追加終了

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sSet_GokiStatus
'//  機能名称  : 号機釦設定処理
'//  機能概要  : 各号機釦の内容を、Tagの値に従って更新する。
'//
'//              型        名称      意味
'//  引数      : Integer   intTab    更新タブ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSet_GokiStatus(ByVal intTab As Integer)

    Dim intIndex As Integer
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intStatusIdx As Integer
    Dim intStatus As Integer
    
    On Error Resume Next

    '対象タブの先頭号機釦Indexを算出
    intStIndex = intTab * 16
    intEdIndex = intStIndex + 15
    
    For intCount = intStIndex To intEdIndex
        '有効なボタンのみ
        If lblStatusBack(intCount).Tag <> "0" Then
            intStatusIdx = CInt(lblStatusBack(intCount).Tag) - 1
            intStatus = gintShimekiri(intStatusIdx)
            'Tag値と一致する文言、色にする
            For intCount2 = 0 To UBound(mudtBtn_Status)
                If mudtBtn_Status(intCount2).intStatus = gintShimekiri(intStatusIdx) Then
                    lblStatus(intCount).Caption = mudtBtn_Status(intCount2).strCaption
                    lblStatusBack(intCount).BackColor = mudtBtn_Status(intCount2).strColor
                End If
            Next intCount2
        End If
    Next intCount

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : funcCheckShimekiri
'//  機能名称  : 締切データ有無チェック
'//  機能概要  : 締切データの有無をチェックする。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V5.10.0.1) 2012-05-09   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function funcCheckShimekiri() As Boolean

' EG20 V6.3.0.1【機能見直し】追加開始
    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim nKansiShimekiriID(5) As Integer                 ' 締切ＩＤ
    Dim nLoop As Integer                                ' ループ
    Dim bEnable As Boolean                              ' ボタン状態
    Dim szFileName As String

    On Error GoTo ErrorHandler                          ' エラーハンドルの登録

    nKansiShimekiriID(0) = IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING01
    nKansiShimekiriID(1) = IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING02
    nKansiShimekiriID(2) = IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING03
    nKansiShimekiriID(3) = IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING04
    nKansiShimekiriID(4) = IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING05
    nKansiShimekiriID(5) = IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING06

    For nLoop = 0 To UBound(gblnCornerSet)

        bEnable = False
        If gblnCornerSet(nLoop) = True Then
            If pfGetKansiShimeJyotai(nKansiShimekiriID(nLoop)) = 0 Then
                ' /////////////////////////////////////////////////////////////////////////
                ' // 締切出力データは存在するか？（D:\KANSI\SHUKEI\SEND_DATA\SIMEKIRI##.DAT）
                szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(nLoop + 1, "0#"))
                If objFso.FileExists(szFileName) = True Then
                    bEnable = True
                End If
            End If
        End If
        cmdOffLine(nLoop).Enabled = bEnable
    Next nLoop
    
    Set objFso = Nothing
    funcCheckShimekiri = True
    Exit Function

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:

    Set objFso = Nothing
    cmdOffLine(0).Enabled = False
    cmdOffLine(1).Enabled = False
    cmdOffLine(2).Enabled = False
    cmdOffLine(3).Enabled = False
    cmdOffLine(4).Enabled = False
    cmdOffLine(5).Enabled = False
    funcCheckShimekiri = False
    Exit Function

End Function
' EG20 V6.3.0.1【機能見直し】追加終了
    
' EG20 V6.3.0.1【機能見直し】削除開始
'    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
'    Dim objFile As File                                 ' ファイルオブジェクト
'
'    On Error GoTo ErrorHandler                          ' エラーハンドルの登録
'
'
'    ' /////////////////////////////////////////////////////////////////////////
'    ' // 初期化
'    gbShimekiriResult = False
'
'    If pfGetKansiShimeJyotai(IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING01) <> 0 Then
'        GoTo ErrorHandler
'    End If
'    If pfGetKansiShimeJyotai(IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING02) <> 0 Then
'        GoTo ErrorHandler
'    End If
'    If pfGetKansiShimeJyotai(IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING03) <> 0 Then
'        GoTo ErrorHandler
'    End If
'    If pfGetKansiShimeJyotai(IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING04) <> 0 Then
'        GoTo ErrorHandler
'    End If
'    If pfGetKansiShimeJyotai(IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING05) <> 0 Then
'        GoTo ErrorHandler
'    End If
'    If pfGetKansiShimeJyotai(IdKansiSts.ID_KANSI_STS_SHIME_PROCESSING06) <> 0 Then
'        GoTo ErrorHandler
'    End If
'
'    ' /////////////////////////////////////////////////////////////////////////
'    ' // 締切出力データは存在するか？（D:\KANSI\SHUKEI\SEND_DATA\HOSHU_SIMEKIRI**_***.DAT）
'    For Each objFile In objFso.GetFolder(PATH_SHUKEI_SEND).files    ' ループを開始
'        If objFso.FileExists(objFile.Path) = True Then              ' ファイル名の取得チェック
'            ' ファイル名を取得
'            If InStr(objFile.Name, FILENAME_SIMEKIRIDAT) <> 0 Then
'                gbShimekiriResult = True
'            End If
'        End If
'    Next
'
'    funcCheckShimekiri = gbShimekiriResult
'    Set objFso = Nothing
'    Set objFile = Nothing
'    Exit Function
'
'' /////////////////////////////////////////////////////////
'' // エラー処理
'ErrorHandler:
'
'    Set objFso = Nothing
'    Set objFile = Nothing
'    gbShimekiriResult = False
'    funcCheckShimekiri = gbShimekiriResult
'    Exit Function
'
'End Function
' EG20 V6.3.0.1【機能見直し】削除終了

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : pfGetKansiShimeJyotai
'//  機能名称  : 監視状態ファイル取得処理
'//  機能概要  : 監視状態ファイル取得処理を行う。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V3.1.0.1) 2011-11-17  CODED BY  [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetKansiShimeJyotai(iAreId As Integer) As Integer

    On Error Resume Next
    
    pfGetKansiShimeJyotai = 99     '初期値
    
    '監視盤起動有無チェック
    If CheckAppStart(PROC_KANRI) <> 0 Then
    '起動ありの場合
     
        Set Idinf_KansiJyotai = New IdInfProc              '監視装置状態エリア
        '共有エリアオープン
        Idinf_KansiJyotai.ProcMode = DATA_ID.Data_Id_KansiJyotai    '監視装置状態エリア
        Idinf_KansiJyotai.IdOpen
        If Idinf_KansiJyotai.Errsts <> 0 Then
           '「監視状態画面：エリア・ファイル参照異常」ログ出力
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
           Exit Function
        End If
        
        '監視状態エリアをＬＯＣＫする。
        Idinf_KansiJyotai.IdLock
        If Idinf_KansiJyotai.Errsts <> 0 Then
            'データ参照異常時:異常
            Idinf_KansiJyotai.IdFree
            '「監視状態画面：エリア・ファイル参照異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
            Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
            Exit Function
        End If
    
        '監視状態エリアIDを設定
        Idinf_KansiJyotai.id = iAreId
        Idinf_KansiJyotai.IdGet
        If Idinf_KansiJyotai.Errsts <> 0 Then
            'データ参照異常時はブランク表示設定を行う。
            Idinf_KansiJyotai.IdFree
            '「監視状態画面：エリア・ファイル参照異常」ログ出力
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
            Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
            Exit Function
        End If
    
        pfGetKansiShimeJyotai = Idinf_KansiJyotai.DataArea(0)   '設定内容
      
        Idinf_KansiJyotai.IdFree
        Set Idinf_KansiJyotai = Nothing               '監視装置設定データファイル
        
    Else
        pfGetKansiShimeJyotai = 0
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : psCheckMisouStatus
'//  機能名称  : 未送状態チェック処理
'//  機能概要  : コーナ別の未送状態をチェックする。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psCheckMisouStatus()

    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim nLoop As Integer                                ' ループ
    Dim bEnable As Boolean                              ' ボタン状態
    Dim szFileName As String

    On Error GoTo ErrorHandler                          ' エラーハンドルの登録

    For nLoop = 0 To UBound(gblnCornerSet)

        bEnable = False
        If gblnCornerSet(nLoop) = True Then
            ' /////////////////////////////////////////////////////////////////////////
            ' // 締切出力データは存在するか？（D:\KANSI\SHUKEI\SEND_DATA\SIMEKIRI##.DAT）
            szFileName = Replace(PATH_SHUKEI_SHIMEDAT, "##", Format(nLoop + 1, "0#"))
            If objFso.FileExists(szFileName) = True Then
                bEnable = True
            End If
        End If
        
        ' 未送あり状態を更新
        lblMisouMark(nLoop).Visible = bEnable
        lblMisouStatus(nLoop).Visible = bEnable
    Next nLoop
    
    Set objFso = Nothing
    Exit Sub

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:
    Set objFso = Nothing
    Exit Sub
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : psCheckReoutStatus
'//  機能名称  : 再出力状態チェック処理
'//  機能概要  : コーナ別の再出力ファイルをチェックする。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psCheckReoutStatus()

    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim objFile As File                                 ' ファイルオブジェクト
    Dim nLoop As Integer                                ' ループ
    Dim bEnable As Boolean                              ' ボタン状態
    Dim szFolderName As String                          ' フォルダ名

    On Error GoTo ErrorHandler                          ' エラーハンドルの登録

    For nLoop = 0 To UBound(gblnCornerSet)

        bEnable = False
        If gblnCornerSet(nLoop) = True Then
            szFolderName = Replace(PATH_SIMEKIRIREOUT_FOLDER, "##", Format(nLoop + 1, "0#"))
            If objFso.FolderExists(szFolderName) = True Then
                ' /////////////////////////////////////////////////////////////////////////
                ' // 再出力データは存在するか？（F:\KANSI\OUT_DATA\CORNER##\SIME##.CSV）
                For Each objFile In objFso.GetFolder(szFolderName).files   ' ループを開始
                    If objFso.FileExists(objFile.Path) = True Then         ' ファイル名の取得チェック
                        ' ファイル名を取得
                        If InStr(objFile.Name, "SIME") <> 0 Then
                            bEnable = True
                        End If
                    End If
                Next
            End If
        End If
        
        ' 未送あり状態を更新
        cmdReOutput(nLoop).Enabled = bEnable
    Next nLoop
    
    Set objFso = Nothing
    Set objFile = Nothing
    Exit Sub

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:
    Set objFso = Nothing
    Set objFile = Nothing
    cmdReOutput(0).Enabled = False
    cmdReOutput(1).Enabled = False
    cmdReOutput(2).Enabled = False
    cmdReOutput(3).Enabled = False
    cmdReOutput(4).Enabled = False
    cmdReOutput(5).Enabled = False
    Exit Sub
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  関数名称  : psCheckShimeKaishiStatus
'//  機能名称  : 締切開始状態チェック処理
'//  機能概要  : コーナ別の統合監視盤締切処理開始の押下可否をチェックする。
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 【機能見直し】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psCheckShimeKaishiStatus()

    Dim nLoop As Integer                                ' ループ
    Dim bEnable As Boolean                              ' ボタン状態
    Dim intCount As Integer
    Dim intCount2 As Integer

    On Error GoTo ErrorHandler                          ' エラーハンドルの登録

    For nLoop = 0 To UBound(gblnCornerSet)

        bEnable = False
        If gblnCornerSet(nLoop) = True Then
            For intCount = 0 To 15
                intCount2 = (nLoop * 16) + intCount
                If lblStatus(intCount2).Caption <> "" Then
                    bEnable = True
                    Exit For
                End If
            Next intCount
        End If
        ' 統合監視盤締切処理開始を更新
        cmdOutPut(nLoop).Enabled = bEnable
    Next nLoop
    
    Exit Sub

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:
    cmdOutPut(0).Enabled = False
    cmdOutPut(1).Enabled = False
    cmdOutPut(2).Enabled = False
    cmdOutPut(3).Enabled = False
    cmdOutPut(4).Enabled = False
    cmdOutPut(5).Enabled = False
    Exit Sub
End Sub


