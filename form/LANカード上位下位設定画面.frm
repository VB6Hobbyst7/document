VERSION 5.00
Begin VB.Form frmLanSettei 
   BorderStyle     =   0  'なし
   Caption         =   "LANカード上位下位設定"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2445
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   479
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   10089
      Width           =   3135
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   10200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   23
      Top             =   6680
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   8520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   22
      Top             =   6680
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   10200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   21
      Top             =   5360
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   8520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   20
      Top             =   5360
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   10200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   19
      Top             =   4040
      Width           =   1575
   End
   Begin VB.Timer tmrMail 
      Left            =   3600
      Top             =   7800
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  機器情報設定    画面へ戻る"
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
      Left            =   9480
      TabIndex        =   18
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton SetteiUpdata 
      Caption         =   "設定更新"
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
      Left            =   6840
      TabIndex        =   16
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "LANカード設定(5)"
      Height          =   975
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   6120
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   8520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   4040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LANカード設定(4)"
      Height          =   975
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   10200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   10
      Top             =   2720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LANカード設定(3)"
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   8520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   2720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LANカード設定(2)"
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   10200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   1400
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   8520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   1400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LANカード設定(1)"
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   10200
      TabIndex        =   35
      Top             =   6280
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   8520
      TabIndex        =   34
      Top             =   6280
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   10200
      TabIndex        =   33
      Top             =   4950
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   8520
      TabIndex        =   32
      Top             =   4950
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   10200
      TabIndex        =   31
      Top             =   3650
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   8520
      TabIndex        =   30
      Top             =   3650
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   10200
      TabIndex        =   29
      Top             =   2325
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   8520
      TabIndex        =   28
      Top             =   2325
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   10200
      TabIndex        =   27
      Top             =   1005
      Width           =   1575
   End
   Begin VB.Label lblIpDisp 
      Alignment       =   2  '中央揃え
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   8520
      TabIndex        =   26
      Top             =   1005
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      Caption         =   "接続機器２"
      Height          =   375
      Left            =   10320
      TabIndex        =   25
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      Caption         =   "接続機器１"
      Height          =   375
      Left            =   8520
      TabIndex        =   24
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "LANカード設定"
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
      TabIndex        =   17
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmLanSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  ファイル名  ：frmLanSettei.frm
'//  パッケージ名：LANカード上位下位設定画面
'//
'//  概要：LANカード上位下位設定画面
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//                 フェーズ２対応
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応　上位LAN設定見直し
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(1.18.0.1) 2010-01-09  REVISED BY [TCC] S.Terao
'//                 不具合対応
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　LANカード設定仕様変更
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//                 ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Const ASRT_IDSERV = &H80000000

'------------------------------------------------------------------------------
'アダプタ情報取得API
'------------------------------------------------------------------------------
Private Const NO_ERROR = 0
Private Const ERROR_BUFFER_OVERFLOW = 111
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_NO_DATA = 232
Private Const ERROR_NOT_SUPPORTED = 50
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH = 128  '// arb.
Private Const MAX_ADAPTER_NAME_LENGTH = 256         '// arb.
Private Const MAX_ADAPTER_ADDRESS_LENGTH = 8        '// arb.
Private Type IP_ADDRESS_STRING
    addr(15)        As Byte
End Type
Private Type IP_ADDR_STRING
    pNext           As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    Context         As Long
End Type
Private Type IP_ADAPTER_INFO
    pNext                   As Long
    ComboIndex              As Long
    AdapterName(MAX_ADAPTER_NAME_LENGTH + 3)          As Byte
    Description(MAX_ADAPTER_DESCRIPTION_LENGTH + 3)   As Byte
    AddressLength           As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1)           As Byte
    dwIndex                 As Long
    uType                   As Long
    bDhcpEnabled            As Long
    pCurrentIpAddress       As Long
    IpAddressList           As IP_ADDR_STRING
    GatewayList             As IP_ADDR_STRING
    DhcpServer              As IP_ADDR_STRING
    bHaveWins               As Long
    PrimaryWinsServer       As IP_ADDR_STRING
    SecondaryWinsServer     As IP_ADDR_STRING
    LeaseObtained           As Long 'time_t
    LeaseExpires            As Long 'time_t
End Type
Private Declare Function GetAdaptersInfo Lib "IPHLPAPI.DLL" ( _
    pAdapterInfo As Byte, _
    pOutBufLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Private Type userIP_ADAPTER
    Description As String       'アダプタ名
    MACAddr As String           'アダプタアドレス
    IPAddr As String            'IPアドレス
    Subnet As String            'サブネットマスク
    Gateway As String           'ゲートウェイ
    DHCP As String              'DHCP
    WINS1 As String             'プライマリWINS
    WINS2 As String             'セカンダリWINS
    LeaseObtain As String       'DHCPリース取得日
    LeaseExpire As String       'DHCPリース期限
End Type

'V1.11.0.1 DEL START
'Private Type SetteiFile
'    sLanCardName(0 To 4) As String         'LANカード設定ファイル：LANカード名
'    sAdapterName(0 To 4) As String         'LANカード設定ファイル：LANアダプタ名    'V1.11.0.1 ADD
''    sCombBox(0 To 4) As Integer           'LANカード設定ファイル：LANカード値      'V1.11.0.1 DEL
'    sCombBox1(0 To 4) As Integer           'LANカード設定ファイル：LANカード値      'V1.11.0.1 ADD
'    sCombBox2(0 To 4) As Integer           'LANカード設定ファイル：LANカード値2     'V1.11.0.1 ADD
'End Type
'Private SetteiFile As SetteiFile
'V1.11.0.1 DEL END
Private sGetLanInfo(0 To 4) As String * MAX_PATH_SIZE 'LANカード情報iniファイル用

Private Const SETTEI_ARI = 0
Private Const SETTEI_NASI = 1
Private Const MN_MAIL_INTERVAL = 1000      'メールタイマのインターバル値
Private Const KANMA = ","
'Private Const DEFULT_KIKI_NAME = "上位"    'デフォルト値   'V1.11.0.1 DEL
Private iDispFlag(0 To 4) As Integer       'フラグ
Private Const LAN_MAX_SETTEI = 5           'LAN設定最大値
Private Const KIKI_MAX_SETTEI = 10         '接続機器最大値   'V1.11.0.1 ADD
Private iLan_Defult As Integer             'LAN設定最大値
'V1.11.0.1 ADD START
Private Const DEFULT_NASI = 0          'デフォルト値
Private Const DEFULT_JIKAI = 1         'デフォルト値
Private Const DEFULT_JOUI = 2          'デフォルト値
Private Const DEFULT_SHUNYU_JOUI = 2   'デフォルト値
Private Const DEFULT_IC_JOUI = 3       'デフォルト値
Private iLanInfo As Integer            '上位LAN設定情報
'V1.11.0.1 ADD END
Private Const DEFULT_ZEN_KIKI = 1      'デフォルト値    'V2.7.0.1 ADD

Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : LANカード上位下位設定画面(アクティブ時)
'//  機能概要  : メール受信用、タイマ起動
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
    'タイマを起動する
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : LANカード上位下位設定画面(ディアクティブ時)
'//  機能概要  : メール受信用、タイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    'タイマを停止する
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : LANカード上位下位設定画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応　上位LAN設定見直し
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　LANカード設定仕様変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim i               As Integer          'ループ用カウンタ
    Dim j               As Integer          'ループ用カウンタ
    Dim fso As New FileSystemObject
    Dim CreateFile As TextStream
    Dim bRet            As Boolean
    
    On Error Resume Next
   
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    bRet = True
    
    'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    SetteiUpdata.Enabled = False
    
    '初期は非表示とする。
    For i = 0 To LAN_MAX_SETTEI - 1
        Frame1(i).Visible = False
'        cmbLanSelect(i).Visible = False    'V1.11.0.1 DEL
        'V1.11.0.1 ADD START
        cmbLanSelect(i * 2).Visible = False
        cmbLanSelect((i * 2) + 1).Visible = False
        'V1.11.0.1 ADD DEL
        SetteiFile.sLanCardName(i) = ""
'        SetteiFile.sCombBox(i) = 0         'V1.11.0.1 DEL
        'V1.11.0.1 ADD START
        SetteiFile.sCombBox1(i) = 0
        SetteiFile.sCombBox2(i) = 0
        'V1.11.0.1 ADD END
        'V1.21.0.1 ADD START
        lblIpDisp(i * 2).Caption = ""
        lblIpDisp((i * 2) + 1).Caption = ""
        lblIpDisp(i * 2).Visible = False
        lblIpDisp((i * 2) + 1).Visible = False
        'V1.21.0.1 ADD END
    Next
        
    'V1.11.0.1 ADD START
    'HOSHU.INIから上位ＬＡＮ設定情報を取得する
    iLanInfo = GetPrivateProfileInt(HOSHU_LANINFO_SEC, HOSHU_LANINFO_KEY, LANINFO_DEFAULT, HOSHU_FILE)
    'V1.11.0.1 ADD END
    
    'LAN情報INIファイル取得処理
    bRet = pfGetLanInfo
    If bRet = True Then
        If Dir(LAN_CARD_SETTEI_CSVFILE, vbNormal) = "" Then
           'ファイル無し時→ファイル作成
           Set CreateFile = fso.CreateTextFile(LAN_CARD_SETTEI_CSVFILE, True)
           CreateFile.Close
        Else
          'LANカード設定ファイル情報取得処理
          bRet = pfGetLANSetteiFileInfo
        End If
        
        If bRet = True Then
           '表示のためのLANカード名取得処理を行う。
           bRet = psGetLANSettei
           If bRet = True Then
              SetteiUpdata.Enabled = True
           End If
        End If
    End If
    
    '「LANカード上位下位設定画面：表示｣
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LAN_CARD_SETTEI_GAMEN_START, 0)
   
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下処理
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    '「LANカード上位下位設定：消去」
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LAN_CARD_SETTEI_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psGetLANSettei
'//  機能名称  : OSよりLANカード情報取得処理
'//  機能概要  : OSよりLANカード情報の取得を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応　上位LAN設定見直し
'//     REVISIONS :(1.18.0.1) 2010-01-09  REVISED BY [TCC] S.Terao
'//                 不具合対応
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　LANカード設定仕様変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function psGetLANSettei() As Boolean
    
    Dim udtIPAdaptInfo() As IP_ADAPTER_INFO
    Dim bytAdaptInfo()  As Byte
    Dim lngBufLen       As Long
    Dim lngRet          As Long
    Dim lngPtr          As Long
    Dim i               As Integer
    Dim Fnc_GetIPAdapt As String
    Dim Name As String
    Dim Name2 As String     'V1.11.0.1 ADD
    Dim lngErrCode As Long
    'V1.18.0.1　ADD START
    Dim sSection As String            'レジストリ接続セクション名
    Dim iErrCnt   As Integer
    'V1.18.0.1　ADD END
    'V1.21.0.1 ADD START
    Dim sIPaddress1 As String         '接続機器１：IPアドレス(フロント)
    Dim sIPaddress2 As String         '接続機器２：IPアドレス(サブ)
    Dim sGetIPaddress As String       'IPアドレス情報
    Dim sSeachStr As String           '次文字有無
    Dim iStrCnt As Integer            'スペース位置
    Dim iByteSts  As Integer          'バイト数
    'V1.21.0.1 ADD END
    
    On Error Resume Next
    
    psGetLANSettei = True
    Fnc_GetIPAdapt = vbNullString
    ' データの取得バッファの初期化
    ReDim udtIPAdaptInfo(0)
    ReDim bytAdaptInfo(0)
    
    ' アダプタ情報を取得するために必要なバッファサイズを取得する
    lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
    
    Select Case lngRet
    
    Case ERROR_BUFFER_OVERFLOW  ' バッファサイズが小さい
        
        ' 取得したバッファサイズより必要なバッファを確保
        ReDim bytAdaptInfo(lngBufLen)
        
        ' アダプタ情報を取得する
        lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
        
        If lngRet = NO_ERROR Then
            ReDim pIpAdapter(0)
            ' 取得したアダプタ情報をユーザー定義型変数へコピー
            CopyMemory udtIPAdaptInfo(0), bytAdaptInfo(0), LenB(udtIPAdaptInfo(0))
            lngPtr = udtIPAdaptInfo(0).pNext
            i = 0
            Do While Not lngPtr = 0 ' 配列の次のアドレスへのポインタをチェックする
                i = i + 1
                ReDim Preserve udtIPAdaptInfo(i)
                ReDim Preserve pIpAdapter(i)
                CopyMemory udtIPAdaptInfo(i), ByVal lngPtr, LenB(udtIPAdaptInfo(0))
                lngPtr = udtIPAdaptInfo(i).pNext
            Loop
            
            '表示LANカード名称を取得する。
            For i = 0 To UBound(udtIPAdaptInfo)
              'アダプタ名
              Name = ByteToStr(udtIPAdaptInfo(i).Description)
              'V1.11.0.1 ADD START
              'LANサービス名
              Name2 = ByteToStr(udtIPAdaptInfo(i).AdapterName)
              
              'V1.21.0.1 ADD START
              sIPaddress1 = ""
              sIPaddress2 = ""
              ' IPアドレス
              sSection = ""
              sSection = "SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Interfaces\\"
              sSection = sSection + Name2 + "\\"
              sGetIPaddress = pfIpAddressGetReg(HKEY_LOCAL_MACHINE, sSection, "IpAddress")
                            
              iStrCnt = InStr(sGetIPaddress, Chr(0)) 'スペース位置の把握
              iByteSts = LenB(sGetIPaddress)         '全体の長さ把握
              sIPaddress1 = Left(sGetIPaddress, InStr(sGetIPaddress, Chr(0)) - 1)
              sSeachStr = Mid(sGetIPaddress, iStrCnt + 1, 1) '続きがあるかどうか判定
              If 0 = InStr(sSeachStr, Chr(0)) And sSeachStr <> "" Then
                 sGetIPaddress = Mid(sGetIPaddress, iStrCnt + 1, iByteSts - iStrCnt)
                 sIPaddress2 = Left(sGetIPaddress, InStr(sGetIPaddress, Chr(0)) - 1)
              End If
              'V1.21.0.1 ADD END
              
              'V1.18.0.1 ADD START
              '接続名取得
               sSection = ""
               sSection = "SYSTEM\\CurrentControlSet\\Control\\Network\\"
               sSection = sSection + "{4D36E972-E325-11CE-BFC1-08002BE10318}\\"
               sSection = sSection + Name2
               sSection = sSection + "\\Connection"
               SetteiFile.sLanName(i) = pfGetReg(HKEY_LOCAL_MACHINE, sSection, "Name")
               If SetteiFile.sLanName(i) = "" Then
                  Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
                  psGetLANSettei = False
        
                  '異常時は画面、ブランク表示とする⇒LANカード名が不明では「設定更新」釦処理も
                  '異常となることがわかっているため
                  For iErrCnt = 0 To LAN_MAX_SETTEI - 1
                      Frame1(iErrCnt).Visible = False
                      cmbLanSelect(iErrCnt * 2).Visible = False
                      cmbLanSelect((iErrCnt * 2) + 1).Visible = False
                      SetteiFile.sLanCardName(iErrCnt) = ""
                      SetteiFile.sCombBox1(iErrCnt) = 0
                      SetteiFile.sCombBox2(iErrCnt) = 0
                  Next
                  Exit Function
               End If
              'V1.18.0.1 ADD END
              'V1.11.0.1 ADD END
              If Name <> "" Then
'                 pfCreateFileDisp Name, i  'V1.11.0.1 DEL
                 'pfCreateFileDisp Name, Name2, i   'V1.11.0.1 ADD 'V1.21.0.1 DEL
                  pfCreateFileDisp Name, Name2, i, sIPaddress1, sIPaddress2 'V1.21.0.1 ADD
              End If
            Next
          Call sLogTraceReq(LTYP_NORMAL, L3AN_API, LAN_CARD_GET_INFO_OK, lngErrCode)
          psGetLANSettei = True
        Else
          Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
          psGetLANSettei = False
        End If
    ' エラー表示
    Case ERROR_INVALID_PARAMETER
       Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
       psGetLANSettei = False
    Case ERROR_NO_DATA
       Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
       psGetLANSettei = False
    Case ERROR_NOT_SUPPORTED
       Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
       psGetLANSettei = False
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : ByteToStr
'//  機能名称  : バイト配列よりNULL終端データを文字列に変換処理
'//  機能概要  : 文字列に文字変換処理を行う。
'//
'//              型        名称      意味
'//  引数      : Variant　bytDest　 [IN]変換対象配列
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function ByteToStr(bytDest As Variant) As String
    Dim strBuf  As String
  
    On Error Resume Next
    ' バイト配列よりNULL終端までのデータで文字列型を生成
    strBuf = StrConv(bytDest, vbUnicode)
    ByteToStr = Left$(strBuf, InStr(1, strBuf, vbNullChar) - 1)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetLanInfo
'//  機能名称  : LANカード情報INIファイルより外部機器選択部名称一覧取得処理
'//  機能概要  : INIファイルより情報を取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応　上位LAN設定見直し
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetLanInfo() As Boolean

  Dim iRet As Integer
  Dim sKey As String
  Dim i As Integer
  Dim iCnt As Integer
  Dim iCount As Integer
  Dim sLANInfo As String * MAX_PATH_SIZE
  Dim iSetteiCnt As Integer
  Dim iCntUp As Integer
  Dim sKikiName As String
    
  On Error Resume Next
  
  sKey = LAN_KEY
  
  pfGetLanInfo = False
  
' V1.11.0.1 DEL START
'  For iCnt = 0 To LAN_MAX_SETTEI - 1
'    i = 1
'    iCntUp = 0
'    iSetteiCnt = 0
'    Do While Not EOF(1)
'       iRet = GetPrivateProfileString(LAN_SEC, _
'                                       sKey & i, _
'                                       "", sLANInfo, Len(sLANInfo), _
'                                       PATH_LAN_CARD_FILE)
'        If 0 <> iRet Then
'           '1コンボボックス表示外部機器名称表示
'           sKikiName = Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
'           If sKikiName <> "" And Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 2, 1) = "," Then
'              cmbLanSelect(iCnt).AddItem Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
'              sGetLanInfo(i - 1) = Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))
'              If i = 1 Then
'                 'デフォルト値決定
'                 iLan_Defult = Mid(sGetLanInfo(i - 1), 1, 1)
'              End If
'           End If
'           iCntUp = iCntUp + 1
'        ElseIf i <> 1 And sGetLanInfo(i - 1) <> " " Then
'           'ループ終了
'           iCntUp = iCntUp + 1
'           pfGetLanInfo = True
'           Exit Do
'        Else
'           '設定なし
'           iSetteiCnt = iSetteiCnt + 1
'        End If
'        i = i + 1
'    Loop
'  Next
'
'  If iSetteiCnt = iCntUp Then
'     pfGetLanInfo = False
'  Else
'     pfGetLanInfo = True
'  End If
' V1.11.0.1 DEL END
  ' V1.11.0.1 ADD START
  For iCnt = 0 To LAN_MAX_SETTEI - 1
    i = 1
    iCntUp = 0
    iSetteiCnt = 0
    Do While Not EOF(1)
       iRet = GetPrivateProfileString(LAN_SEC, _
                                       sKey & i, _
                                       "", sLANInfo, Len(sLANInfo), _
                                       PATH_LAN_CARD_FILE)
        If 0 <> iRet Then
           '1コンボボックス表示外部機器名称表示
           sKikiName = Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
           If sKikiName <> "" And Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 2, 1) = "," Then
              cmbLanSelect(iCnt * 2).AddItem Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
              cmbLanSelect((iCnt * 2) + 1).AddItem Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
              sGetLanInfo(i - 1) = Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))
           End If
        ElseIf i <> 1 And sGetLanInfo(i - 1) <> " " Then
           'ループ終了
           pfGetLanInfo = True
           Exit Do
        End If
        i = i + 1
    Loop
  Next
  ' V1.11.0.1 ADD END
  
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfGetLANSetteiFileInfo
'//  機能名称  : LANカード設定ファイル内情報を読み込む
'//  機能概要  : LANカード設定ファイルより設定ファイル情報を取得する。
'//
'//              型        名称          意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応　上位LAN設定見直し
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20フェーズ２対応【操作卓 結合No.36関連】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetLANSetteiFileInfo() As Boolean

    Dim bRet            As Boolean  '戻り値
    Dim intFileNo       As Integer  'ファイル番号
    Dim strWork         As String   '作業エリア
    Dim intCnt          As Integer  'カウンター
    Dim lngErrCode      As Long     'エラーコード
    Dim sFData()        As String
    Dim iFCnt           As Integer
    Dim iFLoop          As Integer
    Dim iFLoop2         As Integer
    Dim iRet            As Integer
    Dim lLen            As Long     'ファイルサイズ
    
    pfGetLANSetteiFileInfo = True
    
On Error GoTo DispFileGetInfo_Error
   
   lLen = FileLen(LAN_CARD_SETTEI_CSVFILE)             'ファイルサイズの取得
   If lLen = 0 Then
      Exit Function
   End If
   
   'LANカード設定ファイルのファイル番号を取得する。
   intFileNo = FreeFile

   'LANカード設定ファイルオープン
   Open LAN_CARD_SETTEI_CSVFILE For Input As #intFileNo
    
   intCnt = 0
   
   'LANカード設定ファイル内情報を取得する。
'   Do While Not EOF(1)                                     ' EG20 V3.3.0.1削除
    Do While Not EOF(intFileNo)                             ' EG20 V3.3.0.1追加
     Line Input #intFileNo, strWork
           
     If Len(strWork) <> 0 Then
        intCnt = intCnt + 1
        'データの取得
'        ReDim sFData(2)    'V1.11.0.1 DEL
        ReDim sFData(4)     'V1.11.0.1 ADD
        iFCnt = 1
            
        For iFLoop = 1 To Len(strWork)
            If Mid(strWork, iFLoop, 1) <> "," Then
               iFLoop2 = iFLoop
               Do
                iFLoop2 = iFLoop2 + 1
                If iFLoop2 > Len(strWork) Then
                   sFData(iFCnt) = Mid(strWork, iFLoop, iFLoop2 - iFLoop)
                   iFCnt = iFCnt + 1
                   If iFCnt >= 16 Then
                      Exit For
                   End If
                  
                   iFLoop = iFLoop2
                   Exit Do
                End If
                       
                If Mid(strWork, iFLoop2, 1) = "," Then
                   sFData(iFCnt) = Mid(strWork, iFLoop, iFLoop2 - iFLoop)
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
        
        If intCnt <= LAN_MAX_SETTEI Then
           If Len(Trim(sFData(1))) > 0 Then
              'ファイルに記載されているLANカード名を、LANカード設定ファイルエリアに保持
              SetteiFile.sLanCardName(intCnt - 1) = Trim(sFData(1))
           End If
        
           If Len(Trim(sFData(2))) > 0 Then
              'ファイルに記載されているLANカード値を、LANカード設定ファイルエリアに保持
'              SetteiFile.sCombBox(intCnt - 1) = Trim(sFData(2))    'V1.11.0.1 DEL
              SetteiFile.sCombBox1(intCnt - 1) = Trim(sFData(2))     'V1.11.0.1 ADD
           End If
           'V1.11.0.1 ADD START
           If Len(Trim(sFData(3))) > 0 Then
              'ファイルに記載されているLANカード値を、LANカード設定ファイルエリアに保持
              SetteiFile.sCombBox2(intCnt - 1) = Trim(sFData(3))
           End If
           If Len(Trim(sFData(4))) > 0 Then
              'ファイルに記載されているLANアダプタ名を、LANカード設定ファイルエリアに保持
              SetteiFile.sAdapterName(intCnt - 1) = Trim(sFData(4))
           End If
           'V1.11.0.1 ADD END
        End If
     End If
   Loop
   
   'LANカード設定ファイルをクローズする。
    Close #intFileNo
 Exit Function
    
DispFileGetInfo_Error:
  'オープン異常時は処理終了
  Close #intFileNo 'V1.11.0.1 ADD
  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
  pfGetLANSetteiFileInfo = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfCreateFileDisp
'//  機能名称  : 設定比較を行い画面表示処理
'//  機能概要  : 設定比較を行い、画面に表示する。
'//
'//              型        名称          意味
'//  引数      : String   sLanCardName  [IN]LANカード名称
'//              Integer  iCnt          [IN]設定番号
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応　上位LAN設定見直し
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　LANカード設定仕様変更
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//                 ＪＲ東海　ＮＥＧ→ＥＧＲコンバート対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Function pfCreateFileDisp(sLanCardName As String, iCnt As Integer)     'V1.11.0.1 DEL
'Private Function pfCreateFileDisp(sLanCardName As String, sAdapterName As String, iCnt As Integer)      'V1.11.0.1 ADD
Private Function pfCreateFileDisp(sLanCardName As String, sAdapterName As String, iCnt As Integer, sIPaddress1 As String, sIPaddress2 As String)     'V1.21.0.1 ADD
      
      Dim bRet            As Boolean  '戻り値
    Dim strWork         As String   '作業エリア
    Dim intCnt          As Integer  'カウンター
    Dim lngErrCode      As Long     'エラーコード
    Dim i As Integer                'カウンター
    Dim iIniCnt As Integer          'カウンター１
    Dim bChkFlag1 As Boolean         'チェックフラグ1
    Dim bChkFlag2 As Boolean         'チェックフラグ2
    
   If iCnt < LAN_MAX_SETTEI Then
    '表示部をフレーム、ラベル、外部機器名称を表示
     Frame1(iCnt).Visible = True
'     cmbLanSelect(iCnt).Visible = True             'V1.11.0.1 DEL
     cmbLanSelect(iCnt * 2).Visible = True          'V1.11.0.1 ADD
     cmbLanSelect((iCnt * 2) + 1).Visible = True    'V1.11.0.1 ADD
     'V1.21.0.1 ADD START
     '接続機器１：値チェック。アドレスが「0.0.0.0」時はブランク
     If sIPaddress1 <> "" And sIPaddress1 <> "0.0.0.0" Then
        lblIpDisp(iCnt * 2).Caption = sIPaddress1
        lblIpDisp(iCnt * 2).Visible = True
     End If
     '接続機器２：値チェック。アドレスが「0.0.0.0」時はブランク
     If sIPaddress2 <> "" And sIPaddress2 <> "0.0.0.0" Then
        lblIpDisp((iCnt * 2) + 1).Caption = sIPaddress2
        lblIpDisp((iCnt * 2) + 1).Visible = True
     End If
     'V1.21.0.1 ADD END
     
     'OS取得LANカード名を表示する。
     Label2(iCnt).Caption = sLanCardName
                
     For i = 0 To LAN_MAX_SETTEI - 1
        'OS取得LANカード名と、ファイル内記載LANカード名を比較する。
        If SetteiFile.sLanCardName(i) = sLanCardName Then
           '一致：設定ファイルに設定有り
           iDispFlag(iCnt) = SETTEI_ARI
           Exit For
        End If
        If i = LAN_MAX_SETTEI - 1 Then
            '不一致：設定ファイルに設定無し
             iDispFlag(iCnt) = SETTEI_NASI
        End If
     Next
     
     If iDispFlag(iCnt) = SETTEI_NASI Then
        '画面表示LANカード名保持エリアに保持
        SetteiFile.sLanCardName(iCnt) = sLanCardName
        'ファイルに設定がなければデフォルト値をいれる
'        SetteiFile.sCombBox(iCnt) = iLan_Defult    'V1.11.0.1 DEL
        'V1.11.0.1 ADD START
        'LANカードアダプタ名保持エリアに保持
        SetteiFile.sAdapterName(iCnt) = sAdapterName
'        If iLanInfo = 1 Then                    'V2.7.0.1 DEL
        If iLanInfo = LAN_CARD_SET_TYPE1 Then    'V2.7.0.1 ADD
            'LAN上位設定＝方式１
            '接続機器を設定／保持
            If iCnt = 0 Then
                SetteiFile.sCombBox1(iCnt) = DEFULT_JIKAI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_NASI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            ElseIf iCnt = 1 Then
                SetteiFile.sCombBox1(iCnt) = DEFULT_JOUI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_NASI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            Else
                SetteiFile.sCombBox1(iCnt) = DEFULT_NASI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_NASI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            End If
'        ElseIf iLanInfo = 2 Then                  'V2.7.0.1 DEL
        ElseIf iLanInfo = LAN_CARD_SET_TYPE2 Then  'V2.7.0.1 ADD
            'LAN上位設定＝方式２
            '接続機器を設定／保持
            If iCnt = 0 Then
                SetteiFile.sCombBox1(iCnt) = DEFULT_JIKAI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_NASI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            ElseIf iCnt = 1 Then
                SetteiFile.sCombBox1(iCnt) = DEFULT_SHUNYU_JOUI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_IC_JOUI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            Else
                SetteiFile.sCombBox1(iCnt) = DEFULT_NASI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_NASI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            End If
        'V2.7.0.1 ADD START
        ElseIf iLanInfo = LAN_CARD_SET_TYPE3 Then
            'LAN上位設定＝方式３
            '接続機器を設定／保持
            If iCnt = 0 Then
                SetteiFile.sCombBox1(iCnt) = DEFULT_ZEN_KIKI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_NASI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            Else
                SetteiFile.sCombBox1(iCnt) = DEFULT_NASI
                cmbLanSelect(iCnt * 2).ListIndex = SetteiFile.sCombBox1(iCnt)
                SetteiFile.sCombBox2(iCnt) = DEFULT_NASI
                cmbLanSelect((iCnt * 2) + 1).ListIndex = SetteiFile.sCombBox2(iCnt)
            End If
        'V2.7.0.1 ADD END
        End If
        'V1.11.0.1 ADD END
     End If
     
     'V1.11.0.1 ADD START
     '設定ありの場合、設定ファイルから情報を取得
     If iDispFlag(iCnt) = SETTEI_ARI Then
        bChkFlag1 = False
        bChkFlag2 = False
        For iIniCnt = 0 To LAN_MAX_SETTEI - 1 'LANカード情報INIファイル分ループ
        '外部機器名称値を元に、LANカード情報INIファイルと、設定がファイルにあったかどうかを判断する。
            If SetteiFile.sCombBox1(iCnt) = Mid(sGetLanInfo(iIniCnt), 1, 1) Then
                '外部機器名称値に対応した文言を表示する。
                cmbLanSelect(iCnt * 2).ListIndex = iIniCnt
                bChkFlag1 = True
            End If
            If SetteiFile.sCombBox2(iCnt) = Mid(sGetLanInfo(iIniCnt), 1, 1) Then
                '外部機器名称値に対応した文言を表示する。
                cmbLanSelect((iCnt * 2) + 1).ListIndex = iIniCnt
                bChkFlag2 = True
            End If
        Next
        If bChkFlag1 = False Then
            '「－」を表示
            cmbLanSelect(iCnt * 2).ListIndex = DEFULT_NASI
        End If
        If bChkFlag2 = False Then
            '「－」を表示
            cmbLanSelect((iCnt * 2) + 1).ListIndex = DEFULT_NASI
        End If
     End If
     'V1.11.0.1 ADD END

'V1.11.0.1 DEL START
'     For intCnt = 0 To LAN_MAX_SETTEI - 1    '設定ファイル内記載外部機器名称値分ループ
'       For iIniCnt = 0 To LAN_MAX_SETTEI - 1 'LANカード情報INIファイル分ループ
'         '外部機器名称値を元に、LANカード情報INIファイルと、設定がファイルにあったかどうかを判断する。
'         If SetteiFile.sCombBox(iCnt) = Mid(sGetLanInfo(iIniCnt), 1, 1) And iDispFlag(iCnt) = SETTEI_ARI Then
'            '外部機器名称値に対応した文言を表示する。
'            cmbLanSelect(iCnt).ListIndex = iIniCnt
'            Exit Function
'         Else
'            '上記以外の場合
'            cmbLanSelect(iCnt).ListIndex = 0
'         End If
'       Next
'     Next
'V1.11.0.1 DEL END
  End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SetteiUpdata_Click
'//  機能名称  : 「設定更新」釦押下処理
'//  機能概要  : LANカード設定ファイルに画面値を取得する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-Rフェーズ3残件項目対応　上位LAN設定見直し
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 釦の押下可／不可処理追加
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　LANカード設定仕様変更
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 ファイルクローズ処理追加
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub SetteiUpdata_Click()
   Dim iResponse   As Integer   'メッセージボックスの戻り値
   Dim i           As Integer   'カウンタ－１
   Dim iCnt2       As Integer   'カウンター２
   Dim iCnt1       As Integer   'カウンター３
   Dim iKikiCnt    As Integer   '接続機器カウンター     'V1.11.0.1 ADD
   Dim intFileNo   As Integer   'ファイル番号
   Dim sLanName    As String    '書込みLANカード名
   Dim sAdapterName As String   '書込みLANアダプタ名    'V1.11.0.1 ADD
'   Dim sListCnt    As String    '書込み外部機器名称値  'V1.11.0.1 DEL
   Dim sListCnt1    As String    '書込み外部機器名称値  'V1.11.0.1 ADD
   Dim sListCnt2    As String    '書込み外部機器名称値  'V1.11.0.1 ADD
   Dim iLenCnt     As Integer   '文字列数
   Dim lngErrCode  As Long      'エラーコード
'   Dim iListIndex  As Integer   'インデックス値    'V1.11.0.1 DEL
   Dim iListIndex1  As Integer   'インデックス値     'V1.11.0.1 ADD
   Dim iListIndex2  As Integer   'インデックス値     'V1.11.0.1 ADD
   Dim bKikiChk     As Boolean   '機器チェックフラグ 'V1.11.0.1 ADD
   Dim bRet         As Boolean   '戻り値             'V1.11.0.1 ADD
   
On Error GoTo SetteiUpdata_Click

   '「LANカード上位下位設定画面：設定更新釦押下」
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LAN_CARD_SETTEICHANGE_BUTTOM, 0)

'V1.12.0.1 ADD START
   '各ボタンを押下不可とする
    sButtonEnabled (False)
'V1.12.0.1 ADD END
 
   '設定更新確認ポップアップ表示
   iResponse = MsgBox("設定値を反映しますがよろしいですか？", _
                       vbOKCancel + vbQuestion, "設定更新確認")
   If iResponse = vbOK Then
      'V1.11.0.1 ADD START
      '釦押下不可に更新
      sButtonEnabled (False)
      '機器選択重複チェックを行う
      bKikiChk = False
      For i = 0 To LAN_MAX_SETTEI - 1
        If cmbLanSelect(i).Visible = True And Frame1(i).Visible = True Then
            If cmbLanSelect(i * 2).ListIndex = DEFULT_NASI And cmbLanSelect((i * 2) + 1).ListIndex = DEFULT_NASI Then
                ' LANカードに機器が設定されていない
                iResponse = MsgBox("LANカード設定(" & i + 1 & ")に接続機器が設定されていません。", _
                                    vbOKOnly + vbCritical, "設定更新結果")
                '釦押下可に更新
                sButtonEnabled (True)
                Exit Sub
            End If
            For iKikiCnt = 0 To KIKI_MAX_SETTEI - 1
                '同じコマンドは再ループ
                If i <> iKikiCnt Then
                    '接続機器なしの場合はチェック対象外
                    If cmbLanSelect(i).ListIndex <> DEFULT_NASI Then
                        '接続機器の重複を比較
                        If cmbLanSelect(i).ListIndex = cmbLanSelect(iKikiCnt).ListIndex Then
                            '重複あり
                            bKikiChk = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        '機器チェックフラグをチェック
        If bKikiChk = True Then
            '設定重複ポップアップ表示
            iResponse = MsgBox("接続機器に同一の機器が設定されています。", _
                                vbOKOnly + vbCritical, "設定更新結果")
            '釦押下可に更新
            sButtonEnabled (True)
            Exit Sub
        End If
      Next
      
      'LANカードにIPアドレス／コンピュータ名を設定
      bRet = True
      bRet = LANSelectNetworkSet(Me)
      If bRet <> True Then
        'IPアドレス設定異常により処理終了　ポップアップは関数内で表示
        '釦押下可に更新
        sButtonEnabled (True)
        Exit Sub
      End If
      'V1.11.0.1 ADD END
      
      '未使用のファイル番号を取得する
      intFileNo = FreeFile
    
      'LANカード情報設定ファイルをオープンする。
      Open LAN_CARD_SETTEI_CSVFILE For Output Access Write As #intFileNo
     
      For i = 0 To LAN_MAX_SETTEI - 1
        If cmbLanSelect(i).Visible = True And Frame1(i).Visible = True Then
            'V1.11.0.1 ADD START
            For iCnt1 = 0 To cmbLanSelect(i).ListCount
                iListIndex1 = cmbLanSelect(i * 2).ListIndex
                iListIndex2 = cmbLanSelect((i * 2) + 1).ListIndex
                '表示LANカード名を取得
                sLanName = SetteiFile.sLanCardName(i)
                'LANアダプタ名を取得
                sAdapterName = SetteiFile.sAdapterName(i)
                '選択外部機器名称番号取得
                sListCnt1 = Mid(sGetLanInfo(iListIndex1), 1, 1)
                sListCnt2 = Mid(sGetLanInfo(iListIndex2), 1, 1)
                Print #intFileNo, sLanName; KANMA; sListCnt1; KANMA; sListCnt2; KANMA; sAdapterName
                Exit For
            Next
            'V1.11.0.1 ADD END
'V1.11.0.1 DEL START
'          '表示されているものを設定ファイルに書き込む。
'          For iCnt1 = 0 To cmbLanSelect(i).ListCount
'              iListIndex = cmbLanSelect(i).ListIndex
'              '表示LANカード名を取得
'               sLanName = SetteiFile.sLanCardName(i)
'               '選択外部機器名称番号取得
'               sListCnt = Mid(sGetLanInfo(iListIndex), 1, 1)
'               Print #intFileNo, sLanName; KANMA; sListCnt
'               Exit For
'            Next
'V1.11.0.1 DEL END
         End If
      Next
      'LANカード情報設定ファイルをクローズする。
      Close #intFileNo
      Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LAN_CARD_SETTEICHANGE_OK, 0)
      'V1.11.0.1 ADD START
       '正常終了
        'iResponse = MsgBox("正常終了しました。", vbOKOnly + vbInformation, "設定更新結果")  'V1.21.0.1 DEL
        iResponse = MsgBox("正常終了しました。" & Chr(vbKeyReturn) & "監視盤を再起動してください。", vbOKOnly + vbInformation, "設定更新結果")  'V1.21.0.1 ADD
        '釦押下可に更新
        sButtonEnabled (True)
      'V1.11.0.1 ADD END
'V1.12.0.1 ADD START
   Else
        sButtonEnabled (True)
'V1.12.0.1 ADD END
   End If
'V1.21.0.1 ADD START
 'LANカード設定ファイル情報取得処理
  sLanInfoUpData
'V1.21.0.1 ADD END
Exit Sub

SetteiUpdata_Click:
    'V1.21.0.1 ADD  START
    If intFileNo > 0 Then
        Close #intFileNo
    End If
    'V1.21.0.1 ADD  END
    
   '「LANカード上位下位設定画面：設定更新釦押下処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LAN_CARD_SETTEICHANGE_ERROR, lngErrCode)
   '設定更新結果(異常)ポップアップ表示
   iResponse = MsgBox("設定更新は異常終了しました。", _
                       vbOKOnly + vbCritical, "設定更新結果")
    '釦押下可に更新
    sButtonEnabled (True)
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
'//     ORIGINAL  :(1.4.0.1) 2009-04-03   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  'メールを受信する。
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
        AppActivate frmLanSettei.Caption, False
        pfFormActive (frmLanSettei.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sButtonEnabled
'//  機能名称  : 表示画面の釦コントロールを行う。
'//  機能概要  : 「設定更新」釦押下時処理：釦を押下可/押下不可にする。
'//
'//              型        名称      意味
'//  引数      : Boolean　bSet　　　[IN]釦のコントロール(TRUE：押下可,FALSE：押下不可)
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.11.0.1) 2009-10-26   CODED   BY [TCC] D.Yamashita
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sButtonEnabled(bSet As Boolean)

    Dim i           As Integer   'カウンタ－
    
    On Error Resume Next

    SetteiUpdata.Enabled = bSet         '設定更新ボタン
    cmdReturn.Enabled = bSet            'メニュー画面へ戻るボタン
    For i = 0 To LAN_MAX_SETTEI - 1
        If cmbLanSelect(i).Visible = True Then
            cmbLanSelect(i * 2).Enabled = bSet
            cmbLanSelect((i * 2) + 1).Enabled = bSet
        End If
    Next

End Sub

'V1.21.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : sLanInfoUpData
'//  機能名称  : LANカード設定画面の更新処理を行う。
'//  機能概要  : 画面更新処理
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.21.0.1) 2010-04-08   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub sLanInfoUpData()
    Dim bRet As Boolean
    Dim i As Integer
    
    SetteiUpdata.Enabled = False
    
    '初期は非表示とする。
    For i = 0 To LAN_MAX_SETTEI - 1
        Frame1(i).Visible = False
        cmbLanSelect(i * 2).Visible = False
        cmbLanSelect((i * 2) + 1).Visible = False
        SetteiFile.sLanCardName(i) = ""
        lblIpDisp(i * 2).Caption = ""
        lblIpDisp((i * 2) + 1).Caption = ""
        lblIpDisp(i * 2).Visible = False
        lblIpDisp((i * 2) + 1).Visible = False
    Next
    
    bRet = psGetLANDispUpData
    
    SetteiUpdata.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : psGetLANDispUpData
'//  機能名称  : OSよりLANカード情報取得処理
'//  機能概要  : OSよりLANカード情報の取得を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.21.0.1) 2010-04-08  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　LANカード設定仕様変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function psGetLANDispUpData() As Boolean
    
    Dim udtIPAdaptInfo() As IP_ADAPTER_INFO
    Dim bytAdaptInfo()  As Byte
    Dim lngBufLen       As Long
    Dim lngRet          As Long
    Dim lngPtr          As Long
    Dim i               As Integer
    Dim Fnc_GetIPAdapt As String
    Dim Name As String
    Dim Name2 As String
    Dim lngErrCode As Long
    Dim sSection As String            'レジストリ接続セクション名
    Dim iErrCnt   As Integer
    Dim sIPaddress1 As String         '接続機器１：IPアドレス(フロント)
    Dim sIPaddress2 As String         '接続機器２：IPアドレス(サブ)
    Dim sGetIPaddress As String       'IPアドレス情報
    Dim sSeachStr As String           '次文字有無
    Dim iStrCnt As Integer            'スペース位置
    Dim iByteSts  As Integer          'バイト数
    
    On Error Resume Next
    
    psGetLANDispUpData = True
    Fnc_GetIPAdapt = vbNullString
    ' データの取得バッファの初期化
    ReDim udtIPAdaptInfo(0)
    ReDim bytAdaptInfo(0)
    
    ' アダプタ情報を取得するために必要なバッファサイズを取得する
    lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
    
    Select Case lngRet
    
     Case ERROR_BUFFER_OVERFLOW  ' バッファサイズが小さい
        
        ' 取得したバッファサイズより必要なバッファを確保
        ReDim bytAdaptInfo(lngBufLen)
        
        ' アダプタ情報を取得する
        lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
        
        If lngRet = NO_ERROR Then
            ReDim pIpAdapter(0)
            ' 取得したアダプタ情報をユーザー定義型変数へコピー
            CopyMemory udtIPAdaptInfo(0), bytAdaptInfo(0), LenB(udtIPAdaptInfo(0))
            lngPtr = udtIPAdaptInfo(0).pNext
            i = 0
            Do While Not lngPtr = 0 ' 配列の次のアドレスへのポインタをチェックする
                i = i + 1
                ReDim Preserve udtIPAdaptInfo(i)
                ReDim Preserve pIpAdapter(i)
                CopyMemory udtIPAdaptInfo(i), ByVal lngPtr, LenB(udtIPAdaptInfo(0))
                lngPtr = udtIPAdaptInfo(i).pNext
            Loop
            
            '表示LANカード名称を取得する。
            For i = 0 To UBound(udtIPAdaptInfo)
              'アダプタ名
              Name = ByteToStr(udtIPAdaptInfo(i).Description)
              'LANサービス名
              Name2 = ByteToStr(udtIPAdaptInfo(i).AdapterName)
              
              sIPaddress1 = ""
              sIPaddress2 = ""
              ' IPアドレス
              sSection = ""
              sSection = "SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Interfaces\\"
              sSection = sSection + Name2 + "\\"
              sGetIPaddress = pfIpAddressGetReg(HKEY_LOCAL_MACHINE, sSection, "IpAddress")
                            
              iStrCnt = InStr(sGetIPaddress, Chr(0)) 'スペース位置の把握
              iByteSts = LenB(sGetIPaddress)         '全体の長さ把握
              sIPaddress1 = Left(sGetIPaddress, InStr(sGetIPaddress, Chr(0)) - 1)
              sSeachStr = Mid(sGetIPaddress, iStrCnt + 1, 1) '続きがあるかどうか判定
              If 0 = InStr(sSeachStr, Chr(0)) And sSeachStr <> "" Then
                 sGetIPaddress = Mid(sGetIPaddress, iStrCnt + 1, iByteSts - iStrCnt)
                 sIPaddress2 = Left(sGetIPaddress, InStr(sGetIPaddress, Chr(0)) - 1)
              End If
              If Name <> "" Then
                  pfCreateFileDispUpData Name, i, sIPaddress1, sIPaddress2
              End If
            Next
          Call sLogTraceReq(LTYP_NORMAL, L3AN_API, LAN_CARD_GET_INFO_OK, lngErrCode)
          psGetLANDispUpData = True
        Else
          Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
          psGetLANDispUpData = False
        End If
    ' エラー表示
    Case ERROR_INVALID_PARAMETER
       Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
       psGetLANDispUpData = False
    Case ERROR_NO_DATA
       Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
       psGetLANDispUpData = False
    Case ERROR_NOT_SUPPORTED
       Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
       psGetLANDispUpData = False
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfCreateFileDispUpData
'//  機能名称  : 画面再描画処理
'//  機能概要  : 画面再描画を行う。
'//
'//              型        名称          意味
'//  引数      : String   sLanCardName  [IN]LANカード名称
'//              Integer  iCnt          [IN]設定番号
'//              String   sIPaddress1   [IN]接続機器１：IPアドレス
'//              String   sIPaddress2　 [IN]接続機器２：IPアドレス
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL :(1.21.0.1) 2010-04-08  CODED BY [TCC] S.Terao
'//                 EG-R監視盤　２月対応　LANカード設定仕様変更
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfCreateFileDispUpData(sLanCardName As String, iCnt As Integer, sIPaddress1 As String, sIPaddress2 As String)
      
   On Error Resume Next
    
   If iCnt < LAN_MAX_SETTEI Then
    '表示部をフレーム、ラベル、外部機器名称を表示
     Frame1(iCnt).Visible = True
     cmbLanSelect(iCnt * 2).Visible = True
     cmbLanSelect((iCnt * 2) + 1).Visible = True
     '接続機器１：値チェック。アドレスが「0.0.0.0」時はブランク
     If sIPaddress1 <> "" And sIPaddress1 <> "0.0.0.0" Then
        lblIpDisp(iCnt * 2).Caption = sIPaddress1
        lblIpDisp(iCnt * 2).Visible = True
     End If
     '接続機器２：値チェック。アドレスが「0.0.0.0」時はブランク
     If sIPaddress2 <> "" And sIPaddress2 <> "0.0.0.0" Then
        lblIpDisp((iCnt * 2) + 1).Caption = sIPaddress2
        lblIpDisp((iCnt * 2) + 1).Visible = True
     End If
     
     'OS取得LANカード名を表示する。
     Label2(iCnt).Caption = sLanCardName
   End If
End Function
'V1.21.0.1 ADD END
