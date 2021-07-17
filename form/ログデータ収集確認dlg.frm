VERSION 5.00
Begin VB.Form dlgLogDataKakunin 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   2955
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
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル"
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
      Left            =   3240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
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
      Left            =   1200
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '透明
      Caption         =   " 未送の時間帯別データを全てクリアしますが    よろしいですか？"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5535
   End
End
Attribute VB_Name = "dlgLogDataKakunin"
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
'    Call psPutLog(LOG_DLGLOGDATAKAKUNIN_CMDOKCLICK)
'
'    frmICMLogKanri.blnOk = True
'EG20 保守画面モックアップ作成 DEL END

'EG20 保守画面モックアップ作成 ADD START
    blnOk = True
'EG20 保守画面モックアップ作成 ADD END

    '画面をUnload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'/
'/  関数名称     : cmdCancel_Click
'/  機能名称     : 「キャンセル」ボタン押下処理
'/  機能概要     : 「キャンセル」ボタン押下処理を行う
'/
'/                   型          名称            意味
'/  引数         :
'/  戻り値       :
'/
'/  ORIGINAL     : (1.0.1.1) 2006-10-10  CODED BY [TCC] K.Hayashi
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx  CODED BY [xxx]
'/ 備考:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
'EG20 保守画面モックアップ作成 DEL START
'    'ログ出力
'    Call psPutLog(LOG_DLGLOGDATAKAKUNIN_CMDCANCELCLICK)
'
'    frmICMLogKanri.blnOk = False
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
'    Call psPutLog(LOG_DLGLOGDATAKAKUNIN_FORMLOAD)
'
'    '配置設定
'    Me.Top = DIALOGTOP
'    Me.Left = DIALOGLEFT
'    Me.Height = DIALOGHEIGHT
'    Me.Width = DIALOGWIDTH
'
'    If DISPSTS_CPU = gintDispStatus Then
'        '投入口表示CPU
'        lblTitle(0).Caption = DEF_LOG_CPULABEL
'    Else
'        'EGX改札機
'        lblTitle(0).Caption = DEF_LOG_EGXLABEL
'    End If
'EG20 保守画面モックアップ作成 DEL END
'EG20 保守画面モックアップ作成 ADD START
    
    'OK釦押下フラグ初期化
    blnOk = False
    
    '配置設定
    Me.Top = 3495
    Me.Left = 2985
    Me.Height = 3375
    Me.Width = 6165
'EG20 保守画面モックアップ作成 ADD END


End Sub

