VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInputMstData 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'なし
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExtMstInput 
      Caption         =   "   外部マスタ   入力      "
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   12
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdUSBRemove 
      Caption         =   "媒体取外"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   4
      Top             =   3495
      Width           =   2415
   End
   Begin VB.CommandButton cmdMasterInput 
      Caption         =   "媒体入力"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   3
      Top             =   2415
      Width           =   2415
   End
   Begin VB.CommandButton cmdKoshin 
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
      Height          =   975
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdModoru_Menu 
      Caption         =   " データ収集・出力 画面へ戻る"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   0
      Top             =   7080
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15266
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   794
      TabMaxWidth     =   3475
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(0)   =   "マスタデータ入力画面.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tmrMail"
      Tab(0).Control(1)=   "dlgSelectFile"
      Tab(0).Control(2)=   "grdData(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(1)   =   "マスタデータ入力画面.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdData(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(2)   =   "マスタデータ入力画面.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdData(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(3)   =   "マスタデータ入力画面.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdData(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(4)   =   "マスタデータ入力画面.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdData(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "   ○○○○○○　 ○○○○○○"
      TabPicture(5)   =   "マスタデータ入力画面.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "grdData(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Timer tmrMail 
         Left            =   -74520
         Top             =   240
      End
      Begin MSComDlg.CommonDialog dlgSelectFile 
         Left            =   -73800
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "マスタデータ入力画面.frx":00A8
         Height          =   4770
         Index           =   0
         Left            =   -74640
         TabIndex        =   6
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "マスタデータ入力画面.frx":00BE
         Height          =   4770
         Index           =   1
         Left            =   -74640
         TabIndex        =   7
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "マスタデータ入力画面.frx":00D4
         Height          =   4770
         Index           =   2
         Left            =   -74640
         TabIndex        =   8
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "マスタデータ入力画面.frx":00EA
         Height          =   4770
         Index           =   3
         Left            =   -74640
         TabIndex        =   9
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "マスタデータ入力画面.frx":0100
         Height          =   4770
         Index           =   4
         Left            =   -74640
         TabIndex        =   10
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "マスタデータ入力画面.frx":0116
         Height          =   4770
         Index           =   5
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "マスタデータ入力"
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
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmInputMstData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  ファイル名  ：マスタデータ入力.frm
'//  パッケージ名：マスタデータ入力画面のフォームモジュール
'//
'//  概要：パスワード入力画面
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-24  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.2.0.1) 2014-06-25  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応２
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const DispKensu = 20                'グリッド表示行数
Private Const GRID_TITLE = "<　　 　　|　　　　　 　 ﾏｽﾀ名称 　　　　　　|　ﾊﾞｰｼﾞｮﾝ　|　　　　受信日時　　　　"
Private Const MN_MAIL_INTERVAL = 1000   'メイルタイマのインターバル値
Private mlngHandle(19)          As Long     'EG20 V30.0.1.1 ADD


'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : cmdExtMstInput_Click
'//  機能名称  : 「外部マスタ入力」釦押下時処理
'//  機能概要  : 外部マスタ入力画面を表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-24   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdExtMstInput_Click()

    '「マスタデータ入力画面：外部マスタ入力押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_EXTMST_BUTTON, 0)
    
    '外部マスタ入力画面を表示
    Load frmExMasterInput
    frmExMasterInput.Show 1

End Sub
'EG20 V30.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdKoshin_Click
'//  機能名称  : 「表示更新」釦押下時処理
'//  機能概要  : マスタデータを再表示する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-26  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdKoshin_Click()

    '「マスタ入力画面：表示更新押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_KOUSHIN_BUTTOM, 0)
   
    Call sDisp_MasterData(SSTab1.Tab)
    Call sDisp_ParaData(SSTab1.Tab)     'EG20 V30.1.0.1 ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdMasterInput_Click
'//  機能名称  : 「媒体入力」釦押下時処理
'//  機能概要  : マスタデータをインストールする。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V6.2.0.1) 2012-06-15 REVISED BY  [TCC] H.Sugimoto
'//                 【メッセージボックスボタンコード不正対応】
'//     REVISIONS :(EG20 V6.5.0.1) 2012-06-18 REVISED BY  [TCC] H.Sugimoto
'//                 【ファイルの選択方法を改善】
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-09 REVISED BY  [TCC] T.Nakajima
'//                  北陸新幹線開業対応
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 北陸新幹線フェーズ３対応【HKRK_kansi02_001_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdMasterInput_Click()

    Dim iResponse As Integer            'メッセージ応答結果
    Dim strToPath As String             'コピー先ファイルパス
    Dim lngErrCode As Long              'エラーコード
    Dim fso As New FileSystemObject     'ファイルシステムオブジェクト
    
    Dim szInputPath As String           ' コピー元フォルダパス      ' EG20 V6.5.0.1追加
    
    On Error Resume Next
    
    '「マスタ入力画面：表示更新押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_INSTALL_BUTTOM, 0)
    
    iResponse = MsgBox("現在選択中のコーナのマスタデータを入力します。" & vbCrLf & "よろしいですか？", _
                        vbOKCancel + vbQuestion, "マスタデータ入力確認")
    
    If iResponse = vbCancel Then
        Exit Sub
    End If
    
' EG20 V6.5.0.1【ファイルの選択方法を改善】削除開始
'    'フォルダ選択画面
'    dlgSelectFile.FileName = ""
'    dlgSelectFile.Filter = "MST ファイル (*.MST)|*.MST|"
'    dlgSelectFile.DialogTitle = "フォルダを指定してください"
'    dlgSelectFile.InitDir = SHOWFOLDER_DEFAULTFOLDER
'    dlgSelectFile.Flags = dlgSelectFile.Flags Or cdlOFNNoChangeDir
'    dlgSelectFile.ShowOpen
'
'    '指定フォルダなし
'    If Len(dlgSelectFile.FileName) = 0 Then
'         Exit Sub
'    End If
' EG20 V6.5.0.1【ファイルの選択方法を改善】削除終了
' EG20 V6.5.0.1【ファイルの選択方法を改善】追加開始
    ' ファイル選択方式からフォルダ選択方式へ変更し、
    ' 固定ファイル名に対して処理を行う。
    szInputPath = ShowFolders(Me.hwnd, "フォルダを指定してください", SHOWFOLDER_DEFAULTFOLDER)
    ' 指定フォルダなし
    If Len(szInputPath) = 0 Then
        Exit Sub
    End If
    'szInputPath = szInputPath & USB_MASTER_FILE    'EG20 V30.0.1.1 DEL
    'EG20 V30.1.0.1 ADD START
    If gintCornerType(SSTab1.Tab) = CORNER_TYPE_KANSEN Then
        szInputPath = szInputPath & USB_MASTER_FILE_KAN
    Else
        szInputPath = szInputPath & USB_MASTER_FILE
    End If
    'EG20 V30.1.0.1 ADD END
    If fso.FileExists(szInputPath) = False Then
        ' コピー元にマスタファイルが存在しない場合は異常
        Call MsgBox("異常終了しました。", vbOKOnly + vbCritical, "マスタデータ更新結果")
        Set fso = Nothing
        Exit Sub
    End If
' EG20 V6.5.0.1【ファイルの選択方法を改善】追加終了
        
    '画面をロックする
    Call sSetEnable(False)
    
    On Error GoTo Err_Handler
    
    strToPath = PATH_KANSI & "DESHU" & Format(SSTab1.Tab + 1, "00") & DIR_MASTER_V
    
    'コピー先フォルダの有無確認
    If fso.FolderExists(strToPath) = False Then
        'コピー先フォルダ作成
        fso.CreateFolder (strToPath)
    End If
    strToPath = strToPath & USB_MASTER_FILE
' EG20 V6.5.0.1【ファイルの選択方法を改善】削除開始
'    fso.CopyFile dlgSelectFile.FileName, strToPath, True
'    dlgSelectFile.InitDir = ""
'    dlgSelectFile.FileName = ""
' EG20 V6.5.0.1【ファイルの選択方法を改善】削除終了
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL Start
'' EG20 V6.5.0.1【ファイルの選択方法を改善】追加開始
'    fso.CopyFile szInputPath, strToPath, True
'' EG20 V6.5.0.1【ファイルの選択方法を改善】追加終了
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 DEL End
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    '一時保存フォルダにデータをコピーし読取専用を解除する
    If pfChangeAttrNormal(szInputPath, PATH_HOSHUTMP_MST_DATA, strToPath) = False Then
        '異常処理へ
        GoTo Err_Handler
    End If
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End

    ChDir$ "C:\"
    Set fso = Nothing
    
    On Error Resume Next
    
    iResponse = MsgBox("入力されたマスタデータを適用します。" & vbCrLf & "よろしいですか？", _
                        vbOKCancel + vbQuestion, "マスタデータ適用確認")
    
    If iResponse = vbCancel Then
        Call sSetEnable(True)
        Exit Sub
    End If
    
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを表示する
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_SHIMEKIRI)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
    'マスタ更新要求送信
    If fCDATAMailSend = False Then
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
        'プログレスバーを消去する
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了

'        iResponse = MsgBox("異常終了しました。", vbOK + vbCritical, "マスタデータ更新結果")     ' EG20 V6.2.0.1削除
        iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "マスタデータ更新結果")  ' EG20 V6.2.0.1追加
        Call sSetEnable(True)
        Exit Sub
    End If
    
    '受信待ち。
    
    Exit Sub

Err_Handler:
    Set fso = Nothing
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD Start
    '一時保存フォルダを削除する
    psDeleteFolder PATH_HOSHUTMP
' EG20 V30.4.0.1【HKRK_kansi02_001_01】 ADD End
    Call sSetEnable(True)
    '「マスタ入力画面：マスタデータ入力異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FWRITE
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_INSTALL_ERROR, lngErrCode)

' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
    'プログレスバーを消去する
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
    
'    iResponse = MsgBox("異常終了しました。", vbOK + vbCritical, "マスタデータ入力結果")         ' EG20 V6.2.0.1削除
    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "マスタデータ入力結果")      ' EG20 V6.2.0.1追加
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdUSBRemove_Click
'//  機能名称  : 「媒体取外」釦押下時処理
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdUSBRemove_Click()

    On Error Resume Next
    
    '「マスタ入力画面：媒体取外押下」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_EJECT_BUTTOM, 0)
    
   '媒体取外処理
    Call pfRemove(Me)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面に戻る」釦押下時処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Menu_Click()

    On Error Resume Next
    
   '「マスタデータ入力画面：終了」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_GAMEN_END, 0)
 
    '自画面を消す。
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : マスタ入力画面(アクティブ時)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-26  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'タイマを起動する
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = True
    
    'EG20  V30.1.0.1 ADD START
    '外部マスタ入力画面から戻ってきたときにも表示できるようにActivateイベントで表示処理を行うようにした。
    Call sDisp_MasterData(SSTab1.Tab)
    Call sDisp_ParaData(SSTab1.Tab)
    'EG20 V30.1.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : マスタ入力画面(ディアクティブ時:イベントプロシージャ)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    'タイマを停止する
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : マスタ入力画面(ロード時：イベントプロシージャ)
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-20  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim intCount As Integer
    Dim bySyoAssort As Byte             'ログ用小分類
    Dim strCorner1 As String
    Dim strCorner2 As String
    
    On Error Resume Next
    
    Call gsGetSettiCorner
    Call gsGetCornerName
    Call gsGetCornerType        'EG20 V30.1.0.1 ADD

    'タブ数を設置コーナ数とする
    SSTab1.Tab = 0

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
        Else
            SSTab1.TabVisible(intCount) = False
        End If
    
    Next intCount
    
    'Call sDisp_MasterData(SSTab1.Tab)  'EG20 V30.1.0.1 DEL
                                        '今回追加した外部マスタ入力画面から戻ってきたときに再表示できるようActivateイベントに移動
    
    
    
Exit Sub

'エラー処理
Err_LOG:

    'エラーログの出力
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, MASTER_INPUT_GAMEN_START, 0)
     
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : SSTab1_Click
'//  機能名称  : タブ押下時処理
'//  機能概要  : メール受信タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub SSTab1_Click(PreviousTab As Integer)
    '選択中のタブインデックスをセット（外部マスタ入力画面で必要のため）
    gintSelectedCornerTabIdx = SSTab1.Tab
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//
'//     概要      : 「メール受信用タイマ」がタイムアップした時のイベントプロシージャ
'//     説明      : メール受信処理を行う。
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 【プログレスバー表示機能見直し対応】
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014年度施策 【EG20_KANSI05_01】
'//     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  'メール受信エリア
    Dim lngLength As Long            '受信メールバイトサイズ
    Dim intStatus As Integer         '受信メールチェック結果
    Dim iResponse As Integer
    
    On Error Resume Next
    
    'メールを受信する。
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
    '受信メールがあれば、メールＩＤ毎の処理をする。
        Select Case udtReadMail.udtlHeader.dwId        'メールＩＤ
            Case ML_ID_PROEND_ORD
                '「プロセス終了指示」を受信した場合、
                '「プロセス終了指示受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                'プロセスの終了処理を行う
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示」を受信した場合
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '表示元画面（保守データ収集画面）をアクティブ表示する。
                AppActivate frmInputMstData.Caption, False
                pfFormActive (frmInputMstData.hwnd)           ' EG20 V8.1.0.1【EG20_KANSI05_01】ADD
            Case ML_ID_MASTER_UPDATE_RES
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加開始
                'プログレスバーを消去する
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1【プログレスバー表示機能見直し対応】追加終了
                
                '「マスタ更新完了通知」を受信した場合
                If fReadMailCheck(udtReadMail) = False Then
                    iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "マスタデータ更新結果")
                Else
                    iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "マスタデータ更新結果")
                    Call sDisp_MasterData(SSTab1.Tab)
                    Call sDisp_ParaData(SSTab1.Tab)     'EG20 V30.1.0.1 ADD
                End If
                Call sSetEnable(True)
            Case Else
                 'その他のメールを受信した場合
                 '「メールID不正」ログ出力
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sDisp_MasterData
'//  機能名称  : マスタデータ表示処理
'//  機能概要  : 現在選択中のコーナのマスタファイルデータを表示する。
'//
'//              型        名称      意味
'//  引数      : Integer   intTab    選択タブ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDisp_MasterData(ByVal intTab As Integer)

    Dim bySyoAssort         As Byte                 'ログ用小分類
    Dim intCorner           As Integer              'コーナ番号カウンタ
    Dim intCnt, intCnt2     As Integer              'カウンタ
    Dim cFso                As FileSystemObject
    Dim cFile               As File
    Dim dtUpdate            As Date                 '更新日時
    Dim strFilePath         As String               'ファイルパス
    Dim strFileName         As String               'ファイル名
    Dim intFileNo           As Integer              'ファイル番号
    Dim strNum              As String               'マスタ数
    Dim strNo()             As String               'マスタ番号
    Dim strMasterName()     As String               'マスタ名称
    Dim strVer              As String               'バージョン
    Dim intDataCnt          As Integer              'データカウンタ
    Dim intFileNumber       As Integer
    Dim intItemNum          As Integer
    Dim strDateTime         As String
    Dim byBuf()             As Byte
    Dim lngFileSize         As Long
    
    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    bySyoAssort = L3AN_FILE
    
    '未使用のファイル番号を取得
    intFileNumber = FreeFile

    '設定情報ファイルをオープンする
    Open MASTER_DATA_NAME_FILE For Input As #intFileNumber
    
    For intCnt = 0 To 1
        Input #intFileNumber, strNum, strMasterName

        'マスタ数を設定する
        If intCnt = 1 Then
            intItemNum = CInt(strNum)
        End If
    Next
    
    ReDim strNo(intItemNum - 1)
    ReDim strMasterName(intItemNum - 1)
    
    For intCnt = 0 To intItemNum - 1
        Input #intFileNumber, strNo(intCnt), strMasterName(intCnt)
    Next intCnt
    
    Close #intFileNumber
    
    grdData(intTab).Redraw = False      '自動再描画解除
    
    
    'コーナ分ループ
    For intCorner = 0 To 5
        '設置コーナのデータを取得する
        'If SSTab1.TabVisible(intCorner) = True Then    'EG20 V30.1.0.1 DEL
        'EG20 V30.1.0.1 ADD START
        '設置コーナかつ在来コーナのみデータを取得する
        If (SSTab1.TabVisible(intCorner) = True) And (gintCornerType(intCorner) = CORNER_TYPE_ZAIRAI) Then
        'EG20 V30.1.0.1 ADD END
        
            'フォルダを指定
            strFilePath = PATH_KANSI & "DESHU" & Format(intCorner + 1, "00") & DIR_MASTER_V
            intDataCnt = grdData(intCorner).FixedRows
            
            'グリッドを初期化
            For intCnt = grdData(intCorner).FixedRows To grdData(intCorner).Rows - 2
                Call grdData(intCorner).RemoveItem(1)
            Next
            
            For intCnt = 0 To grdData(intCorner).Cols - 1
                grdData(intCorner).TextMatrix(1, intCnt) = ""
            Next
    
            intDataCnt = 1
             grdData(intCorner).FormatString = GRID_TITLE
            Set cFile = Nothing
            Set cFso = New FileSystemObject
            
            For intCnt = 0 To intItemNum - 1
                strFileName = Dir(strFilePath & "MASTER" & Format(strNo(intCnt), "00") & ".dat")
    
                If strFileName = Empty Then
                    strVer = ""
                    strDateTime = ""
                Else
                    'ファイルの更新日時を取得
                    Set cFile = cFso.GetFile(strFilePath & strFileName)
                    dtUpdate = cFile.DateLastModified
                    strDateTime = Format(dtUpdate, "yyyy年m月d日h時nn分")
            
                    lngFileSize = cFile.Size
                    ReDim byBuf(lngFileSize - 1)
            
                    intFileNo = FreeFile
                    'ファイルオープン
                    
                    Open strFilePath & strFileName For Binary As intFileNo Len = lngFileSize
            
                    'データをファイルから読み込む
                    Get #intFileNo, , byBuf
                
                    'バージョンを取得
                    strVer = CStr(byBuf(3))
                    strVer = Format(strVer, "000")
                
                    Close #intFileNo
                End If
                    
                'データ表示
                If intDataCnt > 0 Then
                    grdData(intCorner).AddItem ""
                End If
                grdData(intCorner).TextMatrix(intDataCnt, 0) = strNo(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 1) = strMasterName(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 2) = strVer
                grdData(intCorner).TextMatrix(intDataCnt, 3) = strDateTime
                intDataCnt = intDataCnt + 1
    
            Next intCnt
                    
            Call sSetRowFill(intCorner)
            '表示位置を初期位置へ
            grdData(intCorner).TopRow = grdData(intCorner).FixedRows
        End If
        
        
    Next intCorner
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    Me.Refresh
    
Exit Sub

Err_LOG:

    'EG20 V30.1.0.1 ADD START
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    'EG20 V30.1.0.1 ADD END
    
    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    'エラーログの出力
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, MASTER_INPUT_DISP_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : sDisp_ParaData
'//  機能名称  : パラメータデータ表示処理
'//  機能概要  : 現在選択中のコーナのマスタファイルデータを表示する。
'//
'//              型        名称      意味
'//  引数      : Integer   intTab    選択タブ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sDisp_ParaData(ByVal intTab As Integer)

    Dim bySyoAssort         As Byte                 'ログ用小分類
    Dim intCorner           As Integer              'コーナ番号カウンタ
    Dim intCnt, intCnt2     As Integer              'カウンタ
    Dim cFso                As FileSystemObject
    Dim cFile               As File
    Dim dtUpdate            As Date                 '更新日時
    Dim strFilePath         As String               'ファイルパス
    Dim strFileName         As String               'ファイル名
    Dim intFileNo           As Integer              'ファイル番号
    Dim strNum              As String               'マスタ数
    Dim strNo()             As String               'マスタ番号
    Dim strParaName()       As String               'マスタ名称
    Dim strParaFile()       As String               'パラメータデータファイル名
    Dim strVer              As String               'バージョン
    Dim intDataCnt          As Integer              'データカウンタ
    Dim intFileNumber       As Integer
    Dim intItemNum          As Integer
    Dim strDateTime         As String
    Dim byBuf()             As Byte
    Dim lngFileSize         As Long
    Dim uParaFoot           As PARA_FOOT            'パラメータデータのフッタ部
    Dim i                   As Integer
    Dim intMuIdx            As Integer
    Dim strMutexFile        As String
    
    '内部ファイルエラーのトラップ
    On Error GoTo Err_LOG
    bySyoAssort = L3AN_FILE
    
    '未使用のファイル番号を取得
    intFileNumber = FreeFile

    '設定情報ファイルをオープンする
    Open PARA_DATA_NAME_FILE For Input As #intFileNumber
    
    For intCnt = 0 To 1
        Input #intFileNumber, strNum, strParaName, strParaFile

        'マスタ数を設定する
        If intCnt = 1 Then
            intItemNum = CInt(strNum)
        End If
    Next
    
    intMuIdx = 0
    Erase mlngHandle
    
    ReDim strNo(intItemNum - 1)
    ReDim strParaName(intItemNum - 1)
    ReDim strParaFile(intItemNum - 1)
    
    For intCnt = 0 To intItemNum - 1
        Input #intFileNumber, strNo(intCnt), strParaName(intCnt), strParaFile(intCnt)
    Next intCnt
    
    Close #intFileNumber
    
    grdData(intTab).Redraw = False      '自動再描画解除
    
    
    'コーナ分ループ
    For intCorner = 0 To 5
        '設置コーナかつ幹線コーナのデータを取得する
        intMuIdx = 0
        If (SSTab1.TabVisible(intCorner) = True) And (gintCornerType(intCorner) = CORNER_TYPE_KANSEN) Then
        
            'フォルダを指定
            strFilePath = PATH_KANSI & "N_GATE" & Format(intCorner + 1, "00") & DIR_NPARA_V
            intDataCnt = grdData(intCorner).FixedRows
            
            'グリッドを初期化
            For intCnt = grdData(intCorner).FixedRows To grdData(intCorner).Rows - 2
                Call grdData(intCorner).RemoveItem(1)
            Next
            
            For intCnt = 0 To grdData(intCorner).Cols - 1
                grdData(intCorner).TextMatrix(1, intCnt) = ""
            Next
    
            intDataCnt = 1
             grdData(intCorner).FormatString = GRID_TITLE
            Set cFile = Nothing
            Set cFso = New FileSystemObject
            
            For intCnt = 0 To intItemNum - 1
                strFileName = Dir(strFilePath & strParaFile(intCnt))
    
                If strFileName = Empty Then
                    strVer = ""
                    strDateTime = ""
                Else
                    '排他処理(OPEN)
                    strMutexFile = "MU_PARAMETER" & Format(intCorner + 1, "00")
                    mlngHandle(intMuIdx) = dllOpenMutex(strMutexFile)
                    If mlngHandle(intMuIdx) <> 0 Then
                        dllWaitForSingleObject (mlngHandle(intMuIdx))     '排他処理(GET)
                    End If
                    
                    'ファイルの更新日時を取得
                    Set cFile = cFso.GetFile(strFilePath & strFileName)
                    dtUpdate = cFile.DateLastModified
                    strDateTime = Format(dtUpdate, "yyyy年m月d日h時nn分")
            
                    lngFileSize = cFile.Size
                    ReDim byBuf(lngFileSize - 1)
            
                    intFileNo = FreeFile
                    'ファイルオープン
                    'Open strFilePath & strFileName For Binary As intFileNo Len = lngFileSize
                    'Binaryで開く場合はLen節は意味なし。（サイズは32,767 バイト以下である必要がある。しかしパラメータはそれ以上）
                    Open strFilePath & strFileName For Binary As intFileNo
            
                    'パラメータデータのフッタ情報を取得する
                    Get #intFileNo, lngFileSize - Len(uParaFoot) + 1, uParaFoot
            
                    'バージョンを取得
                    strVer = ""
                    For i = 0 To UBound(uParaFoot.byVersion)
                        strVer = strVer & Right$("0" & Hex(uParaFoot.byVersion(i)), 2)
                    Next i
                    strVer = Format(strVer, "000")

                    Close #intFileNo
                    
                    If mlngHandle(intMuIdx) <> 0 Then
                        '排他処理(FREE)
                        Call dllReleaseMutex(mlngHandle(intMuIdx))
                        '排他処理(CLOSE)
                        Call dllCloseHandle(mlngHandle(intMuIdx))
                    End If
                    intMuIdx = intMuIdx + 1
                End If
                    
                'データ表示
                If intDataCnt > 0 Then
                    grdData(intCorner).AddItem ""
                End If
                grdData(intCorner).TextMatrix(intDataCnt, 0) = strNo(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 1) = strParaName(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 2) = strVer
                grdData(intCorner).TextMatrix(intDataCnt, 3) = strDateTime
                intDataCnt = intDataCnt + 1
    
            Next intCnt
                    
            Call sSetRowFill(intCorner)
            '表示位置を初期位置へ
            grdData(intCorner).TopRow = grdData(intCorner).FixedRows
        End If
        
        
    Next intCorner
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    Me.Refresh
    
Exit Sub

Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    For intCnt = 0 To intMuIdx - 1
        '排他処理（FREE)
        Call dllReleaseMutex(mlngHandle(intCnt))
        '排他処理(CLOSE)
        Call dllCloseHandle(mlngHandle(intCnt))
    Next
  
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    'エラーログの出力
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, MASTER_INPUT_DISP_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sSetRowFill
'//  機能名称  : グリッド行埋め処理
'//  機能概要  : グリッドの行数と背景色を設定
'//
'//              型        名称      意味
'//  引数      : Integer   intTab    選択タブ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSetRowFill(ByVal intTab As Integer)

    Dim intCnt, intCnt2 As Integer
    
    '行数が１ページの表示件数になるように、空白行を作成する。
    If (grdData(intTab).Rows - grdData(intTab).FixedRows) Mod DispKensu > 0 Then
        grdData(intTab).Rows = grdData(intTab).Rows + (DispKensu - (grdData(intTab).Rows - grdData(intTab).FixedRows) Mod DispKensu)
    End If
    
    'グリッドの行背景色を設定
    grdData(intTab).RowHeight(0) = 232
    For intCnt = 1 To (grdData(intTab).Rows - 1)
        grdData(intTab).Row = intCnt
        grdData(intTab).RowHeight(intCnt) = 232
        For intCnt2 = 0 To grdData(intTab).Cols - 1
            grdData(intTab).Col = intCnt2
            If (intCnt Mod 2) = 0 Then
            '偶数行の背景色は「FFFFFF」
                grdData(intTab).CellBackColor = "&H00FFFFFF"
             Else
                '奇数行の背景色は「DFDDDE」
                grdData(intTab).CellBackColor = "&H00DFDDDE"
            End If
        Next intCnt2
    Next intCnt
        
    grdData(intTab).Redraw = True               '自動再描画再開
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : fCDATAMailSend
'//  機能名称  : マスタ更新要求送信処理
'//  機能概要  : 初期処理時：メールを送信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fCDATAMailSend() As Boolean

    Dim udtMail As MAIL_INFO_RES    'マスタ更新要求メール送信エリア
    Dim lngRet As Long              '関数戻り値
    Dim lngErrCode As Long          'エラーコード
    
    On Error Resume Next
 
    'マスタ更新要求を監マに送信する。
    udtMail.mlHeader.dwId = ML_ID_MASTER_UPDATE_CMD
    udtMail.mlHeader.dwSize = MlSize.MASTER_UPDATE_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequestType = ML_ID_MASTER_UPDATE_H       'データ種別
    udtMail.dwSts = SSTab1.Tab + 1                      'コーナ番号
    
    lngRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       '「マスタデータ入力画面：マスタ更新要求送信異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, MASTER_UPDATE_REQ_SEND, lngErrCode)
       fCDATAMailSend = False
       Exit Function
    Else
       '「マスタデータ入力画面：マスタ更新要求送信正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, MASTER_UPDATE_REQ_SEND, 0)
       fCDATAMailSend = True
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : fReadMailCheck
'//  機能名称  : マスタ更新完了通知メールチェック処理
'//  機能概要  : メール受信時：メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim lngErrCode As Long
    
    On Error Resume Next
    
    'データ種別チェック
    If udtReadMail.lngData(0) <> ML_ID_MASTER_UPDATE_H Then
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MRECEIVE + 1
        Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MASTER_UPDATE_REQ_RECV, lngErrCode)
        fReadMailCheck = False
        Exit Function
    End If
    
    'コーナチェック
    If udtReadMail.lngData(1) <> (SSTab1.Tab + 1) Then
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MRECEIVE + 2
        Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MASTER_UPDATE_REQ_RECV, lngErrCode)
        fReadMailCheck = False
        Exit Function
    End If
    
    '処理結果チェック
    If udtReadMail.lngData(2) > 0 Then
        fReadMailCheck = False
        Exit Function
    End If
    
    fReadMailCheck = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  関数名称  : sSetEnable
'//  機能名称  : 活性状態制御
'//  機能概要  : コマンドボタンの活性・非活性を制御する
'//
'//              型        名称      意味
'//  引数      : Boolean   blnEnable [IN]活性／不活性フラグ
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.2.0.1) 2014-06-25  REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応２
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sSetEnable(ByVal blnEnable As Boolean)

    Dim lngErrCode As Long
    
    On Error Resume Next
    
    cmdKoshin.Enabled = blnEnable
    cmdMasterInput.Enabled = blnEnable
    cmdUSBRemove.Enabled = blnEnable
    cmdModoru_Menu.Enabled = blnEnable
    cmdExtMstInput.Enabled = blnEnable      'EG20 V30.2.0.1 ADD
    SSTab1.Enabled = blnEnable
    
End Sub
