VERSION 5.00
Begin VB.Form frmLanSettei 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "LAN�J�[�h��ʉ��ʐݒ�"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2445
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
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
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   23
      Top             =   6680
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   22
      Top             =   6680
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   21
      Top             =   5360
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   20
      Top             =   5360
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   19
      Top             =   4040
      Width           =   1575
   End
   Begin VB.Timer tmrMail 
      Left            =   3600
      Top             =   7800
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  �@����ݒ�    ��ʂ֖߂�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "�ݒ�X�V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "LAN�J�[�h�ݒ�(5)"
      Height          =   975
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   6120
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   13
      Top             =   4040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LAN�J�[�h�ݒ�(4)"
      Height          =   975
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   10
      Top             =   2720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LAN�J�[�h�ݒ�(3)"
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   7
      Top             =   2720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LAN�J�[�h�ݒ�(2)"
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   4
      Top             =   1400
      Width           =   1575
   End
   Begin VB.ComboBox cmbLanSelect 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   1400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LAN�J�[�h�ݒ�(1)"
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8055
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "XXX.XXX.XXX.XXX"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Caption         =   "�ڑ��@��Q"
      Height          =   375
      Left            =   10320
      TabIndex        =   25
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      Caption         =   "�ڑ��@��P"
      Height          =   375
      Left            =   8520
      TabIndex        =   24
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "LAN�J�[�h�ݒ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
'//  �t�@�C����  �FfrmLanSettei.frm
'//  �p�b�P�[�W���FLAN�J�[�h��ʉ��ʐݒ���
'//
'//  �T�v�FLAN�J�[�h��ʉ��ʐݒ���
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ��@���LAN�ݒ茩����
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(1.18.0.1) 2010-01-09  REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@LAN�J�[�h�ݒ�d�l�ύX
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//                 �i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Const ASRT_IDSERV = &H80000000

'------------------------------------------------------------------------------
'�A�_�v�^���擾API
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
    Description As String       '�A�_�v�^��
    MACAddr As String           '�A�_�v�^�A�h���X
    IPAddr As String            'IP�A�h���X
    Subnet As String            '�T�u�l�b�g�}�X�N
    Gateway As String           '�Q�[�g�E�F�C
    DHCP As String              'DHCP
    WINS1 As String             '�v���C�}��WINS
    WINS2 As String             '�Z�J���_��WINS
    LeaseObtain As String       'DHCP���[�X�擾��
    LeaseExpire As String       'DHCP���[�X����
End Type

'V1.11.0.1 DEL START
'Private Type SetteiFile
'    sLanCardName(0 To 4) As String         'LAN�J�[�h�ݒ�t�@�C���FLAN�J�[�h��
'    sAdapterName(0 To 4) As String         'LAN�J�[�h�ݒ�t�@�C���FLAN�A�_�v�^��    'V1.11.0.1 ADD
''    sCombBox(0 To 4) As Integer           'LAN�J�[�h�ݒ�t�@�C���FLAN�J�[�h�l      'V1.11.0.1 DEL
'    sCombBox1(0 To 4) As Integer           'LAN�J�[�h�ݒ�t�@�C���FLAN�J�[�h�l      'V1.11.0.1 ADD
'    sCombBox2(0 To 4) As Integer           'LAN�J�[�h�ݒ�t�@�C���FLAN�J�[�h�l2     'V1.11.0.1 ADD
'End Type
'Private SetteiFile As SetteiFile
'V1.11.0.1 DEL END
Private sGetLanInfo(0 To 4) As String * MAX_PATH_SIZE 'LAN�J�[�h���ini�t�@�C���p

Private Const SETTEI_ARI = 0
Private Const SETTEI_NASI = 1
Private Const MN_MAIL_INTERVAL = 1000      '���[���^�C�}�̃C���^�[�o���l
Private Const KANMA = ","
'Private Const DEFULT_KIKI_NAME = "���"    '�f�t�H���g�l   'V1.11.0.1 DEL
Private iDispFlag(0 To 4) As Integer       '�t���O
Private Const LAN_MAX_SETTEI = 5           'LAN�ݒ�ő�l
Private Const KIKI_MAX_SETTEI = 10         '�ڑ��@��ő�l   'V1.11.0.1 ADD
Private iLan_Defult As Integer             'LAN�ݒ�ő�l
'V1.11.0.1 ADD START
Private Const DEFULT_NASI = 0          '�f�t�H���g�l
Private Const DEFULT_JIKAI = 1         '�f�t�H���g�l
Private Const DEFULT_JOUI = 2          '�f�t�H���g�l
Private Const DEFULT_SHUNYU_JOUI = 2   '�f�t�H���g�l
Private Const DEFULT_IC_JOUI = 3       '�f�t�H���g�l
Private iLanInfo As Integer            '���LAN�ݒ���
'V1.11.0.1 ADD END
Private Const DEFULT_ZEN_KIKI = 1      '�f�t�H���g�l    'V2.7.0.1 ADD

Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : LAN�J�[�h��ʉ��ʐݒ���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�A�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
    '�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : LAN�J�[�h��ʉ��ʐݒ���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�A�^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    '�^�C�}���~����
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : LAN�J�[�h��ʉ��ʐݒ���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ��@���LAN�ݒ茩����
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@LAN�J�[�h�ݒ�d�l�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim i               As Integer          '���[�v�p�J�E���^
    Dim j               As Integer          '���[�v�p�J�E���^
    Dim fso As New FileSystemObject
    Dim CreateFile As TextStream
    Dim bRet            As Boolean
    
    On Error Resume Next
   
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    bRet = True
    
    '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    SetteiUpdata.Enabled = False
    
    '�����͔�\���Ƃ���B
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
    'HOSHU.INI�����ʂk�`�m�ݒ�����擾����
    iLanInfo = GetPrivateProfileInt(HOSHU_LANINFO_SEC, HOSHU_LANINFO_KEY, LANINFO_DEFAULT, HOSHU_FILE)
    'V1.11.0.1 ADD END
    
    'LAN���INI�t�@�C���擾����
    bRet = pfGetLanInfo
    If bRet = True Then
        If Dir(LAN_CARD_SETTEI_CSVFILE, vbNormal) = "" Then
           '�t�@�C�����������t�@�C���쐬
           Set CreateFile = fso.CreateTextFile(LAN_CARD_SETTEI_CSVFILE, True)
           CreateFile.Close
        Else
          'LAN�J�[�h�ݒ�t�@�C�����擾����
          bRet = pfGetLANSetteiFileInfo
        End If
        
        If bRet = True Then
           '�\���̂��߂�LAN�J�[�h���擾�������s���B
           bRet = psGetLANSettei
           If bRet = True Then
              SetteiUpdata.Enabled = True
           End If
        End If
    End If
    
    '�uLAN�J�[�h��ʉ��ʐݒ��ʁF�\���
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LAN_CARD_SETTEI_GAMEN_START, 0)
   
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t��������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    '�uLAN�J�[�h��ʉ��ʐݒ�F�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LAN_CARD_SETTEI_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psGetLANSettei
'//  �@�\����  : OS���LAN�J�[�h���擾����
'//  �@�\�T�v  : OS���LAN�J�[�h���̎擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ��@���LAN�ݒ茩����
'//     REVISIONS :(1.18.0.1) 2010-01-09  REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@LAN�J�[�h�ݒ�d�l�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
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
    'V1.18.0.1�@ADD START
    Dim sSection As String            '���W�X�g���ڑ��Z�N�V������
    Dim iErrCnt   As Integer
    'V1.18.0.1�@ADD END
    'V1.21.0.1 ADD START
    Dim sIPaddress1 As String         '�ڑ��@��P�FIP�A�h���X(�t�����g)
    Dim sIPaddress2 As String         '�ڑ��@��Q�FIP�A�h���X(�T�u)
    Dim sGetIPaddress As String       'IP�A�h���X���
    Dim sSeachStr As String           '�������L��
    Dim iStrCnt As Integer            '�X�y�[�X�ʒu
    Dim iByteSts  As Integer          '�o�C�g��
    'V1.21.0.1 ADD END
    
    On Error Resume Next
    
    psGetLANSettei = True
    Fnc_GetIPAdapt = vbNullString
    ' �f�[�^�̎擾�o�b�t�@�̏�����
    ReDim udtIPAdaptInfo(0)
    ReDim bytAdaptInfo(0)
    
    ' �A�_�v�^�����擾���邽�߂ɕK�v�ȃo�b�t�@�T�C�Y���擾����
    lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
    
    Select Case lngRet
    
    Case ERROR_BUFFER_OVERFLOW  ' �o�b�t�@�T�C�Y��������
        
        ' �擾�����o�b�t�@�T�C�Y���K�v�ȃo�b�t�@���m��
        ReDim bytAdaptInfo(lngBufLen)
        
        ' �A�_�v�^�����擾����
        lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
        
        If lngRet = NO_ERROR Then
            ReDim pIpAdapter(0)
            ' �擾�����A�_�v�^�������[�U�[��`�^�ϐ��փR�s�[
            CopyMemory udtIPAdaptInfo(0), bytAdaptInfo(0), LenB(udtIPAdaptInfo(0))
            lngPtr = udtIPAdaptInfo(0).pNext
            i = 0
            Do While Not lngPtr = 0 ' �z��̎��̃A�h���X�ւ̃|�C���^���`�F�b�N����
                i = i + 1
                ReDim Preserve udtIPAdaptInfo(i)
                ReDim Preserve pIpAdapter(i)
                CopyMemory udtIPAdaptInfo(i), ByVal lngPtr, LenB(udtIPAdaptInfo(0))
                lngPtr = udtIPAdaptInfo(i).pNext
            Loop
            
            '�\��LAN�J�[�h���̂��擾����B
            For i = 0 To UBound(udtIPAdaptInfo)
              '�A�_�v�^��
              Name = ByteToStr(udtIPAdaptInfo(i).Description)
              'V1.11.0.1 ADD START
              'LAN�T�[�r�X��
              Name2 = ByteToStr(udtIPAdaptInfo(i).AdapterName)
              
              'V1.21.0.1 ADD START
              sIPaddress1 = ""
              sIPaddress2 = ""
              ' IP�A�h���X
              sSection = ""
              sSection = "SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Interfaces\\"
              sSection = sSection + Name2 + "\\"
              sGetIPaddress = pfIpAddressGetReg(HKEY_LOCAL_MACHINE, sSection, "IpAddress")
                            
              iStrCnt = InStr(sGetIPaddress, Chr(0)) '�X�y�[�X�ʒu�̔c��
              iByteSts = LenB(sGetIPaddress)         '�S�̂̒����c��
              sIPaddress1 = Left(sGetIPaddress, InStr(sGetIPaddress, Chr(0)) - 1)
              sSeachStr = Mid(sGetIPaddress, iStrCnt + 1, 1) '���������邩�ǂ�������
              If 0 = InStr(sSeachStr, Chr(0)) And sSeachStr <> "" Then
                 sGetIPaddress = Mid(sGetIPaddress, iStrCnt + 1, iByteSts - iStrCnt)
                 sIPaddress2 = Left(sGetIPaddress, InStr(sGetIPaddress, Chr(0)) - 1)
              End If
              'V1.21.0.1 ADD END
              
              'V1.18.0.1 ADD START
              '�ڑ����擾
               sSection = ""
               sSection = "SYSTEM\\CurrentControlSet\\Control\\Network\\"
               sSection = sSection + "{4D36E972-E325-11CE-BFC1-08002BE10318}\\"
               sSection = sSection + Name2
               sSection = sSection + "\\Connection"
               SetteiFile.sLanName(i) = pfGetReg(HKEY_LOCAL_MACHINE, sSection, "Name")
               If SetteiFile.sLanName(i) = "" Then
                  Call sLogTraceReq(LTYP_ERROR, L3AN_API, LAN_CARD_GET_INFO_ERROR, lngErrCode)
                  psGetLANSettei = False
        
                  '�ُ펞�͉�ʁA�u�����N�\���Ƃ����LAN�J�[�h�����s���ł́u�ݒ�X�V�v�t������
                  '�ُ�ƂȂ邱�Ƃ��킩���Ă��邽��
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
    ' �G���[�\��
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
'//  �֐�����  : ByteToStr
'//  �@�\����  : �o�C�g�z����NULL�I�[�f�[�^�𕶎���ɕϊ�����
'//  �@�\�T�v  : ������ɕ����ϊ��������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Variant�@bytDest�@ [IN]�ϊ��Ώ۔z��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function ByteToStr(bytDest As Variant) As String
    Dim strBuf  As String
  
    On Error Resume Next
    ' �o�C�g�z����NULL�I�[�܂ł̃f�[�^�ŕ�����^�𐶐�
    strBuf = StrConv(bytDest, vbUnicode)
    ByteToStr = Left$(strBuf, InStr(1, strBuf, vbNullChar) - 1)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGetLanInfo
'//  �@�\����  : LAN�J�[�h���INI�t�@�C�����O���@��I�𕔖��̈ꗗ�擾����
'//  �@�\�T�v  : INI�t�@�C���������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ��@���LAN�ݒ茩����
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
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
'           '1�R���{�{�b�N�X�\���O���@�햼�̕\��
'           sKikiName = Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
'           If sKikiName <> "" And Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 2, 1) = "," Then
'              cmbLanSelect(iCnt).AddItem Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
'              sGetLanInfo(i - 1) = Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))
'              If i = 1 Then
'                 '�f�t�H���g�l����
'                 iLan_Defult = Mid(sGetLanInfo(i - 1), 1, 1)
'              End If
'           End If
'           iCntUp = iCntUp + 1
'        ElseIf i <> 1 And sGetLanInfo(i - 1) <> " " Then
'           '���[�v�I��
'           iCntUp = iCntUp + 1
'           pfGetLanInfo = True
'           Exit Do
'        Else
'           '�ݒ�Ȃ�
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
           '1�R���{�{�b�N�X�\���O���@�햼�̕\��
           sKikiName = Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
           If sKikiName <> "" And Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 2, 1) = "," Then
              cmbLanSelect(iCnt * 2).AddItem Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
              cmbLanSelect((iCnt * 2) + 1).AddItem Mid((Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))), 3)
              sGetLanInfo(i - 1) = Left$(sLANInfo, (InStr(sLANInfo, vbNullChar) - 1))
           End If
        ElseIf i <> 1 And sGetLanInfo(i - 1) <> " " Then
           '���[�v�I��
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
'//  �֐�����  : pfGetLANSetteiFileInfo
'//  �@�\����  : LAN�J�[�h�ݒ�t�@�C��������ǂݍ���
'//  �@�\�T�v  : LAN�J�[�h�ݒ�t�@�C�����ݒ�t�@�C�������擾����B
'//
'//              �^        ����          �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ��@���LAN�ݒ茩����
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetLANSetteiFileInfo() As Boolean

    Dim bRet            As Boolean  '�߂�l
    Dim intFileNo       As Integer  '�t�@�C���ԍ�
    Dim strWork         As String   '��ƃG���A
    Dim intCnt          As Integer  '�J�E���^�[
    Dim lngErrCode      As Long     '�G���[�R�[�h
    Dim sFData()        As String
    Dim iFCnt           As Integer
    Dim iFLoop          As Integer
    Dim iFLoop2         As Integer
    Dim iRet            As Integer
    Dim lLen            As Long     '�t�@�C���T�C�Y
    
    pfGetLANSetteiFileInfo = True
    
On Error GoTo DispFileGetInfo_Error
   
   lLen = FileLen(LAN_CARD_SETTEI_CSVFILE)             '�t�@�C���T�C�Y�̎擾
   If lLen = 0 Then
      Exit Function
   End If
   
   'LAN�J�[�h�ݒ�t�@�C���̃t�@�C���ԍ����擾����B
   intFileNo = FreeFile

   'LAN�J�[�h�ݒ�t�@�C���I�[�v��
   Open LAN_CARD_SETTEI_CSVFILE For Input As #intFileNo
    
   intCnt = 0
   
   'LAN�J�[�h�ݒ�t�@�C���������擾����B
'   Do While Not EOF(1)                                     ' EG20 V3.3.0.1�폜
    Do While Not EOF(intFileNo)                             ' EG20 V3.3.0.1�ǉ�
     Line Input #intFileNo, strWork
           
     If Len(strWork) <> 0 Then
        intCnt = intCnt + 1
        '�f�[�^�̎擾
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
              '�t�@�C���ɋL�ڂ���Ă���LAN�J�[�h�����ALAN�J�[�h�ݒ�t�@�C���G���A�ɕێ�
              SetteiFile.sLanCardName(intCnt - 1) = Trim(sFData(1))
           End If
        
           If Len(Trim(sFData(2))) > 0 Then
              '�t�@�C���ɋL�ڂ���Ă���LAN�J�[�h�l���ALAN�J�[�h�ݒ�t�@�C���G���A�ɕێ�
'              SetteiFile.sCombBox(intCnt - 1) = Trim(sFData(2))    'V1.11.0.1 DEL
              SetteiFile.sCombBox1(intCnt - 1) = Trim(sFData(2))     'V1.11.0.1 ADD
           End If
           'V1.11.0.1 ADD START
           If Len(Trim(sFData(3))) > 0 Then
              '�t�@�C���ɋL�ڂ���Ă���LAN�J�[�h�l���ALAN�J�[�h�ݒ�t�@�C���G���A�ɕێ�
              SetteiFile.sCombBox2(intCnt - 1) = Trim(sFData(3))
           End If
           If Len(Trim(sFData(4))) > 0 Then
              '�t�@�C���ɋL�ڂ���Ă���LAN�A�_�v�^�����ALAN�J�[�h�ݒ�t�@�C���G���A�ɕێ�
              SetteiFile.sAdapterName(intCnt - 1) = Trim(sFData(4))
           End If
           'V1.11.0.1 ADD END
        End If
     End If
   Loop
   
   'LAN�J�[�h�ݒ�t�@�C�����N���[�Y����B
    Close #intFileNo
 Exit Function
    
DispFileGetInfo_Error:
  '�I�[�v���ُ펞�͏����I��
  Close #intFileNo 'V1.11.0.1 ADD
  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
  pfGetLANSetteiFileInfo = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfCreateFileDisp
'//  �@�\����  : �ݒ��r���s����ʕ\������
'//  �@�\�T�v  : �ݒ��r���s���A��ʂɕ\������B
'//
'//              �^        ����          �Ӗ�
'//  ����      : String   sLanCardName  [IN]LAN�J�[�h����
'//              Integer  iCnt          [IN]�ݒ�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ��@���LAN�ݒ茩����
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@LAN�J�[�h�ݒ�d�l�ύX
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//                 �i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Function pfCreateFileDisp(sLanCardName As String, iCnt As Integer)     'V1.11.0.1 DEL
'Private Function pfCreateFileDisp(sLanCardName As String, sAdapterName As String, iCnt As Integer)      'V1.11.0.1 ADD
Private Function pfCreateFileDisp(sLanCardName As String, sAdapterName As String, iCnt As Integer, sIPaddress1 As String, sIPaddress2 As String)     'V1.21.0.1 ADD
      
      Dim bRet            As Boolean  '�߂�l
    Dim strWork         As String   '��ƃG���A
    Dim intCnt          As Integer  '�J�E���^�[
    Dim lngErrCode      As Long     '�G���[�R�[�h
    Dim i As Integer                '�J�E���^�[
    Dim iIniCnt As Integer          '�J�E���^�[�P
    Dim bChkFlag1 As Boolean         '�`�F�b�N�t���O1
    Dim bChkFlag2 As Boolean         '�`�F�b�N�t���O2
    
   If iCnt < LAN_MAX_SETTEI Then
    '�\�������t���[���A���x���A�O���@�햼�̂�\��
     Frame1(iCnt).Visible = True
'     cmbLanSelect(iCnt).Visible = True             'V1.11.0.1 DEL
     cmbLanSelect(iCnt * 2).Visible = True          'V1.11.0.1 ADD
     cmbLanSelect((iCnt * 2) + 1).Visible = True    'V1.11.0.1 ADD
     'V1.21.0.1 ADD START
     '�ڑ��@��P�F�l�`�F�b�N�B�A�h���X���u0.0.0.0�v���̓u�����N
     If sIPaddress1 <> "" And sIPaddress1 <> "0.0.0.0" Then
        lblIpDisp(iCnt * 2).Caption = sIPaddress1
        lblIpDisp(iCnt * 2).Visible = True
     End If
     '�ڑ��@��Q�F�l�`�F�b�N�B�A�h���X���u0.0.0.0�v���̓u�����N
     If sIPaddress2 <> "" And sIPaddress2 <> "0.0.0.0" Then
        lblIpDisp((iCnt * 2) + 1).Caption = sIPaddress2
        lblIpDisp((iCnt * 2) + 1).Visible = True
     End If
     'V1.21.0.1 ADD END
     
     'OS�擾LAN�J�[�h����\������B
     Label2(iCnt).Caption = sLanCardName
                
     For i = 0 To LAN_MAX_SETTEI - 1
        'OS�擾LAN�J�[�h���ƁA�t�@�C�����L��LAN�J�[�h�����r����B
        If SetteiFile.sLanCardName(i) = sLanCardName Then
           '��v�F�ݒ�t�@�C���ɐݒ�L��
           iDispFlag(iCnt) = SETTEI_ARI
           Exit For
        End If
        If i = LAN_MAX_SETTEI - 1 Then
            '�s��v�F�ݒ�t�@�C���ɐݒ薳��
             iDispFlag(iCnt) = SETTEI_NASI
        End If
     Next
     
     If iDispFlag(iCnt) = SETTEI_NASI Then
        '��ʕ\��LAN�J�[�h���ێ��G���A�ɕێ�
        SetteiFile.sLanCardName(iCnt) = sLanCardName
        '�t�@�C���ɐݒ肪�Ȃ���΃f�t�H���g�l�������
'        SetteiFile.sCombBox(iCnt) = iLan_Defult    'V1.11.0.1 DEL
        'V1.11.0.1 ADD START
        'LAN�J�[�h�A�_�v�^���ێ��G���A�ɕێ�
        SetteiFile.sAdapterName(iCnt) = sAdapterName
'        If iLanInfo = 1 Then                    'V2.7.0.1 DEL
        If iLanInfo = LAN_CARD_SET_TYPE1 Then    'V2.7.0.1 ADD
            'LAN��ʐݒ聁�����P
            '�ڑ��@���ݒ�^�ێ�
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
            'LAN��ʐݒ聁�����Q
            '�ڑ��@���ݒ�^�ێ�
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
            'LAN��ʐݒ聁�����R
            '�ڑ��@���ݒ�^�ێ�
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
     '�ݒ肠��̏ꍇ�A�ݒ�t�@�C����������擾
     If iDispFlag(iCnt) = SETTEI_ARI Then
        bChkFlag1 = False
        bChkFlag2 = False
        For iIniCnt = 0 To LAN_MAX_SETTEI - 1 'LAN�J�[�h���INI�t�@�C�������[�v
        '�O���@�햼�̒l�����ɁALAN�J�[�h���INI�t�@�C���ƁA�ݒ肪�t�@�C���ɂ��������ǂ����𔻒f����B
            If SetteiFile.sCombBox1(iCnt) = Mid(sGetLanInfo(iIniCnt), 1, 1) Then
                '�O���@�햼�̒l�ɑΉ�����������\������B
                cmbLanSelect(iCnt * 2).ListIndex = iIniCnt
                bChkFlag1 = True
            End If
            If SetteiFile.sCombBox2(iCnt) = Mid(sGetLanInfo(iIniCnt), 1, 1) Then
                '�O���@�햼�̒l�ɑΉ�����������\������B
                cmbLanSelect((iCnt * 2) + 1).ListIndex = iIniCnt
                bChkFlag2 = True
            End If
        Next
        If bChkFlag1 = False Then
            '�u�|�v��\��
            cmbLanSelect(iCnt * 2).ListIndex = DEFULT_NASI
        End If
        If bChkFlag2 = False Then
            '�u�|�v��\��
            cmbLanSelect((iCnt * 2) + 1).ListIndex = DEFULT_NASI
        End If
     End If
     'V1.11.0.1 ADD END

'V1.11.0.1 DEL START
'     For intCnt = 0 To LAN_MAX_SETTEI - 1    '�ݒ�t�@�C�����L�ڊO���@�햼�̒l�����[�v
'       For iIniCnt = 0 To LAN_MAX_SETTEI - 1 'LAN�J�[�h���INI�t�@�C�������[�v
'         '�O���@�햼�̒l�����ɁALAN�J�[�h���INI�t�@�C���ƁA�ݒ肪�t�@�C���ɂ��������ǂ����𔻒f����B
'         If SetteiFile.sCombBox(iCnt) = Mid(sGetLanInfo(iIniCnt), 1, 1) And iDispFlag(iCnt) = SETTEI_ARI Then
'            '�O���@�햼�̒l�ɑΉ�����������\������B
'            cmbLanSelect(iCnt).ListIndex = iIniCnt
'            Exit Function
'         Else
'            '��L�ȊO�̏ꍇ
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
'//  �֐�����  : SetteiUpdata_Click
'//  �@�\����  : �u�ݒ�X�V�v�t��������
'//  �@�\�T�v  : LAN�J�[�h�ݒ�t�@�C���ɉ�ʒl���擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ��@���LAN�ݒ茩����
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(1.21.0.1) 2010-04-08  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@LAN�J�[�h�ݒ�d�l�ύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub SetteiUpdata_Click()
   Dim iResponse   As Integer   '���b�Z�[�W�{�b�N�X�̖߂�l
   Dim i           As Integer   '�J�E���^�|�P
   Dim iCnt2       As Integer   '�J�E���^�[�Q
   Dim iCnt1       As Integer   '�J�E���^�[�R
   Dim iKikiCnt    As Integer   '�ڑ��@��J�E���^�[     'V1.11.0.1 ADD
   Dim intFileNo   As Integer   '�t�@�C���ԍ�
   Dim sLanName    As String    '������LAN�J�[�h��
   Dim sAdapterName As String   '������LAN�A�_�v�^��    'V1.11.0.1 ADD
'   Dim sListCnt    As String    '�����݊O���@�햼�̒l  'V1.11.0.1 DEL
   Dim sListCnt1    As String    '�����݊O���@�햼�̒l  'V1.11.0.1 ADD
   Dim sListCnt2    As String    '�����݊O���@�햼�̒l  'V1.11.0.1 ADD
   Dim iLenCnt     As Integer   '������
   Dim lngErrCode  As Long      '�G���[�R�[�h
'   Dim iListIndex  As Integer   '�C���f�b�N�X�l    'V1.11.0.1 DEL
   Dim iListIndex1  As Integer   '�C���f�b�N�X�l     'V1.11.0.1 ADD
   Dim iListIndex2  As Integer   '�C���f�b�N�X�l     'V1.11.0.1 ADD
   Dim bKikiChk     As Boolean   '�@��`�F�b�N�t���O 'V1.11.0.1 ADD
   Dim bRet         As Boolean   '�߂�l             'V1.11.0.1 ADD
   
On Error GoTo SetteiUpdata_Click

   '�uLAN�J�[�h��ʉ��ʐݒ��ʁF�ݒ�X�V�t�����v
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LAN_CARD_SETTEICHANGE_BUTTOM, 0)

'V1.12.0.1 ADD START
   '�e�{�^���������s�Ƃ���
    sButtonEnabled (False)
'V1.12.0.1 ADD END
 
   '�ݒ�X�V�m�F�|�b�v�A�b�v�\��
   iResponse = MsgBox("�ݒ�l�𔽉f���܂�����낵���ł����H", _
                       vbOKCancel + vbQuestion, "�ݒ�X�V�m�F")
   If iResponse = vbOK Then
      'V1.11.0.1 ADD START
      '�t�����s�ɍX�V
      sButtonEnabled (False)
      '�@��I���d���`�F�b�N���s��
      bKikiChk = False
      For i = 0 To LAN_MAX_SETTEI - 1
        If cmbLanSelect(i).Visible = True And Frame1(i).Visible = True Then
            If cmbLanSelect(i * 2).ListIndex = DEFULT_NASI And cmbLanSelect((i * 2) + 1).ListIndex = DEFULT_NASI Then
                ' LAN�J�[�h�ɋ@�킪�ݒ肳��Ă��Ȃ�
                iResponse = MsgBox("LAN�J�[�h�ݒ�(" & i + 1 & ")�ɐڑ��@�킪�ݒ肳��Ă��܂���B", _
                                    vbOKOnly + vbCritical, "�ݒ�X�V����")
                '�t�����ɍX�V
                sButtonEnabled (True)
                Exit Sub
            End If
            For iKikiCnt = 0 To KIKI_MAX_SETTEI - 1
                '�����R�}���h�͍ă��[�v
                If i <> iKikiCnt Then
                    '�ڑ��@��Ȃ��̏ꍇ�̓`�F�b�N�ΏۊO
                    If cmbLanSelect(i).ListIndex <> DEFULT_NASI Then
                        '�ڑ��@��̏d�����r
                        If cmbLanSelect(i).ListIndex = cmbLanSelect(iKikiCnt).ListIndex Then
                            '�d������
                            bKikiChk = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        '�@��`�F�b�N�t���O���`�F�b�N
        If bKikiChk = True Then
            '�ݒ�d���|�b�v�A�b�v�\��
            iResponse = MsgBox("�ڑ��@��ɓ���̋@�킪�ݒ肳��Ă��܂��B", _
                                vbOKOnly + vbCritical, "�ݒ�X�V����")
            '�t�����ɍX�V
            sButtonEnabled (True)
            Exit Sub
        End If
      Next
      
      'LAN�J�[�h��IP�A�h���X�^�R���s���[�^����ݒ�
      bRet = True
      bRet = LANSelectNetworkSet(Me)
      If bRet <> True Then
        'IP�A�h���X�ݒ�ُ�ɂ�菈���I���@�|�b�v�A�b�v�͊֐����ŕ\��
        '�t�����ɍX�V
        sButtonEnabled (True)
        Exit Sub
      End If
      'V1.11.0.1 ADD END
      
      '���g�p�̃t�@�C���ԍ����擾����
      intFileNo = FreeFile
    
      'LAN�J�[�h���ݒ�t�@�C�����I�[�v������B
      Open LAN_CARD_SETTEI_CSVFILE For Output Access Write As #intFileNo
     
      For i = 0 To LAN_MAX_SETTEI - 1
        If cmbLanSelect(i).Visible = True And Frame1(i).Visible = True Then
            'V1.11.0.1 ADD START
            For iCnt1 = 0 To cmbLanSelect(i).ListCount
                iListIndex1 = cmbLanSelect(i * 2).ListIndex
                iListIndex2 = cmbLanSelect((i * 2) + 1).ListIndex
                '�\��LAN�J�[�h�����擾
                sLanName = SetteiFile.sLanCardName(i)
                'LAN�A�_�v�^�����擾
                sAdapterName = SetteiFile.sAdapterName(i)
                '�I���O���@�햼�̔ԍ��擾
                sListCnt1 = Mid(sGetLanInfo(iListIndex1), 1, 1)
                sListCnt2 = Mid(sGetLanInfo(iListIndex2), 1, 1)
                Print #intFileNo, sLanName; KANMA; sListCnt1; KANMA; sListCnt2; KANMA; sAdapterName
                Exit For
            Next
            'V1.11.0.1 ADD END
'V1.11.0.1 DEL START
'          '�\������Ă�����̂�ݒ�t�@�C���ɏ������ށB
'          For iCnt1 = 0 To cmbLanSelect(i).ListCount
'              iListIndex = cmbLanSelect(i).ListIndex
'              '�\��LAN�J�[�h�����擾
'               sLanName = SetteiFile.sLanCardName(i)
'               '�I���O���@�햼�̔ԍ��擾
'               sListCnt = Mid(sGetLanInfo(iListIndex), 1, 1)
'               Print #intFileNo, sLanName; KANMA; sListCnt
'               Exit For
'            Next
'V1.11.0.1 DEL END
         End If
      Next
      'LAN�J�[�h���ݒ�t�@�C�����N���[�Y����B
      Close #intFileNo
      Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LAN_CARD_SETTEICHANGE_OK, 0)
      'V1.11.0.1 ADD START
       '����I��
        'iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�ݒ�X�V����")  'V1.21.0.1 DEL
        iResponse = MsgBox("����I�����܂����B" & Chr(vbKeyReturn) & "�Ď��Ղ��ċN�����Ă��������B", vbOKOnly + vbInformation, "�ݒ�X�V����")  'V1.21.0.1 ADD
        '�t�����ɍX�V
        sButtonEnabled (True)
      'V1.11.0.1 ADD END
'V1.12.0.1 ADD START
   Else
        sButtonEnabled (True)
'V1.12.0.1 ADD END
   End If
'V1.21.0.1 ADD START
 'LAN�J�[�h�ݒ�t�@�C�����擾����
  sLanInfoUpData
'V1.21.0.1 ADD END
Exit Sub

SetteiUpdata_Click:
    'V1.21.0.1 ADD  START
    If intFileNo > 0 Then
        Close #intFileNo
    End If
    'V1.21.0.1 ADD  END
    
   '�uLAN�J�[�h��ʉ��ʐݒ��ʁF�ݒ�X�V�t���������ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LAN_CARD_SETTEICHANGE_ERROR, lngErrCode)
   '�ݒ�X�V����(�ُ�)�|�b�v�A�b�v�\��
   iResponse = MsgBox("�ݒ�X�V�ُ͈�I�����܂����B", _
                       vbOKOnly + vbCritical, "�ݒ�X�V����")
    '�t�����ɍX�V
    sButtonEnabled (True)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v  : ���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-03   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmLanSettei.Caption, False
        pfFormActive (frmLanSettei.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sButtonEnabled
'//  �@�\����  : �\����ʂ̖t�R���g���[�����s���B
'//  �@�\�T�v  : �u�ݒ�X�V�v�t�����������F�t��������/�����s�ɂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Boolean�@bSet�@�@�@[IN]�t�̃R���g���[��(TRUE�F������,FALSE�F�����s��)
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.11.0.1) 2009-10-26   CODED   BY [TCC] D.Yamashita
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sButtonEnabled(bSet As Boolean)

    Dim i           As Integer   '�J�E���^�|
    
    On Error Resume Next

    SetteiUpdata.Enabled = bSet         '�ݒ�X�V�{�^��
    cmdReturn.Enabled = bSet            '���j���[��ʂ֖߂�{�^��
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
'//  �֐�����  : sLanInfoUpData
'//  �@�\����  : LAN�J�[�h�ݒ��ʂ̍X�V�������s���B
'//  �@�\�T�v  : ��ʍX�V����
'//
'//              �^        ����      �Ӗ�
'//  ����      :
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.21.0.1) 2010-04-08   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sLanInfoUpData()
    Dim bRet As Boolean
    Dim i As Integer
    
    SetteiUpdata.Enabled = False
    
    '�����͔�\���Ƃ���B
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
'//  �֐�����  : psGetLANDispUpData
'//  �@�\����  : OS���LAN�J�[�h���擾����
'//  �@�\�T�v  : OS���LAN�J�[�h���̎擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.21.0.1) 2010-04-08  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@LAN�J�[�h�ݒ�d�l�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
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
    Dim sSection As String            '���W�X�g���ڑ��Z�N�V������
    Dim iErrCnt   As Integer
    Dim sIPaddress1 As String         '�ڑ��@��P�FIP�A�h���X(�t�����g)
    Dim sIPaddress2 As String         '�ڑ��@��Q�FIP�A�h���X(�T�u)
    Dim sGetIPaddress As String       'IP�A�h���X���
    Dim sSeachStr As String           '�������L��
    Dim iStrCnt As Integer            '�X�y�[�X�ʒu
    Dim iByteSts  As Integer          '�o�C�g��
    
    On Error Resume Next
    
    psGetLANDispUpData = True
    Fnc_GetIPAdapt = vbNullString
    ' �f�[�^�̎擾�o�b�t�@�̏�����
    ReDim udtIPAdaptInfo(0)
    ReDim bytAdaptInfo(0)
    
    ' �A�_�v�^�����擾���邽�߂ɕK�v�ȃo�b�t�@�T�C�Y���擾����
    lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
    
    Select Case lngRet
    
     Case ERROR_BUFFER_OVERFLOW  ' �o�b�t�@�T�C�Y��������
        
        ' �擾�����o�b�t�@�T�C�Y���K�v�ȃo�b�t�@���m��
        ReDim bytAdaptInfo(lngBufLen)
        
        ' �A�_�v�^�����擾����
        lngRet = GetAdaptersInfo(bytAdaptInfo(0), lngBufLen)
        
        If lngRet = NO_ERROR Then
            ReDim pIpAdapter(0)
            ' �擾�����A�_�v�^�������[�U�[��`�^�ϐ��փR�s�[
            CopyMemory udtIPAdaptInfo(0), bytAdaptInfo(0), LenB(udtIPAdaptInfo(0))
            lngPtr = udtIPAdaptInfo(0).pNext
            i = 0
            Do While Not lngPtr = 0 ' �z��̎��̃A�h���X�ւ̃|�C���^���`�F�b�N����
                i = i + 1
                ReDim Preserve udtIPAdaptInfo(i)
                ReDim Preserve pIpAdapter(i)
                CopyMemory udtIPAdaptInfo(i), ByVal lngPtr, LenB(udtIPAdaptInfo(0))
                lngPtr = udtIPAdaptInfo(i).pNext
            Loop
            
            '�\��LAN�J�[�h���̂��擾����B
            For i = 0 To UBound(udtIPAdaptInfo)
              '�A�_�v�^��
              Name = ByteToStr(udtIPAdaptInfo(i).Description)
              'LAN�T�[�r�X��
              Name2 = ByteToStr(udtIPAdaptInfo(i).AdapterName)
              
              sIPaddress1 = ""
              sIPaddress2 = ""
              ' IP�A�h���X
              sSection = ""
              sSection = "SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Interfaces\\"
              sSection = sSection + Name2 + "\\"
              sGetIPaddress = pfIpAddressGetReg(HKEY_LOCAL_MACHINE, sSection, "IpAddress")
                            
              iStrCnt = InStr(sGetIPaddress, Chr(0)) '�X�y�[�X�ʒu�̔c��
              iByteSts = LenB(sGetIPaddress)         '�S�̂̒����c��
              sIPaddress1 = Left(sGetIPaddress, InStr(sGetIPaddress, Chr(0)) - 1)
              sSeachStr = Mid(sGetIPaddress, iStrCnt + 1, 1) '���������邩�ǂ�������
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
    ' �G���[�\��
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
'//  �֐�����  : pfCreateFileDispUpData
'//  �@�\����  : ��ʍĕ`�揈��
'//  �@�\�T�v  : ��ʍĕ`����s���B
'//
'//              �^        ����          �Ӗ�
'//  ����      : String   sLanCardName  [IN]LAN�J�[�h����
'//              Integer  iCnt          [IN]�ݒ�ԍ�
'//              String   sIPaddress1   [IN]�ڑ��@��P�FIP�A�h���X
'//              String   sIPaddress2�@ [IN]�ڑ��@��Q�FIP�A�h���X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.21.0.1) 2010-04-08  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@LAN�J�[�h�ݒ�d�l�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfCreateFileDispUpData(sLanCardName As String, iCnt As Integer, sIPaddress1 As String, sIPaddress2 As String)
      
   On Error Resume Next
    
   If iCnt < LAN_MAX_SETTEI Then
    '�\�������t���[���A���x���A�O���@�햼�̂�\��
     Frame1(iCnt).Visible = True
     cmbLanSelect(iCnt * 2).Visible = True
     cmbLanSelect((iCnt * 2) + 1).Visible = True
     '�ڑ��@��P�F�l�`�F�b�N�B�A�h���X���u0.0.0.0�v���̓u�����N
     If sIPaddress1 <> "" And sIPaddress1 <> "0.0.0.0" Then
        lblIpDisp(iCnt * 2).Caption = sIPaddress1
        lblIpDisp(iCnt * 2).Visible = True
     End If
     '�ڑ��@��Q�F�l�`�F�b�N�B�A�h���X���u0.0.0.0�v���̓u�����N
     If sIPaddress2 <> "" And sIPaddress2 <> "0.0.0.0" Then
        lblIpDisp((iCnt * 2) + 1).Caption = sIPaddress2
        lblIpDisp((iCnt * 2) + 1).Visible = True
     End If
     
     'OS�擾LAN�J�[�h����\������B
     Label2(iCnt).Caption = sLanCardName
   End If
End Function
'V1.21.0.1 ADD END
