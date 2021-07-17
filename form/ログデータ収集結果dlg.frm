VERSION 5.00
Begin VB.Form dlgLogDataKekka 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   3195
   ClientLeft      =   3015
   ClientTop       =   4200
   ClientWidth     =   6030
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "ＯＫ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "○○終了しました。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "dlgLogDataKekka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'/
'/  関数名称     : cmdOK_Click
'/  機能名称     : 「OK」ボタン押下処理
'/  機能概要     : 「OK」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.1.1) 2006-10-10  CODED BY [TCC] K.Hayashi
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx  CODED BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    
'EG20 保守画面モックアップ作成 DEL START
'    'ログ出力
'    Call psPutLog(LOG_DLGLOGDATAKEKKA_CMDOKCLICK)
'EG20 保守画面モックアップ作成 DEL END

    '画面をUnload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'/
'/  関数名称     : Form_Load
'/  機能名称     : Form_Load時処理
'/  機能概要     : Form_Load時処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.1.1) 2006-10-10  CODED BY [TCC] K.Hayashi
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx  CODED BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
'EG20 保守画面モックアップ作成 DEL START
'    'ログ出力
'    Call psPutLog(LOG_DLGLOGDATAKEKKA_FORMLOAD)
'
'    '配置設定
'    Me.Top = DIALOGTOP
'    Me.Left = DIALOGLEFT
'    Me.Height = DIALOGHEIGHT
'    Me.Width = DIALOGWIDTH
'
'    lblTitle = "正常終了しました。"
'EG20 保守画面モックアップ作成 DEL END
'EG20 保守画面モックアップ作成 ADD START
    '配置設定
    Me.Top = 3495
    Me.Left = 2985
    Me.Height = 3375
    Me.Width = 6165
'EG20 保守画面モックアップ作成 ADD END

End Sub

