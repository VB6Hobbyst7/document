VERSION 5.00
Begin VB.Form frmJprPrint 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�W���[�i����"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkJprKind 
      Caption         =   "�ݒ�l�ꗗ"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   32
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame FraPrintKind 
      Caption         =   "�󎚍��ڎw��"
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   11655
      Begin VB.CheckBox chkJprKind 
         Caption         =   "���D�@�ێ�ݒ�f�[�^"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   39
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "�w�s�x�f�[�^�m�F(�ݺ��޺�ō��@����`)"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   38
         Top             =   1080
         Width           =   5055
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "�w���@��h�c"
         Height          =   255
         Index           =   7
         Left            =   8520
         TabIndex        =   37
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "���؃I�t���C���o��"
         Height          =   255
         Index           =   6
         Left            =   8520
         TabIndex        =   36
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "�ғ��o�[�W�����ꗗ"
         Height          =   255
         Index           =   5
         Left            =   8520
         TabIndex        =   35
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "���p���z�f�[�^"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "�ʉ߃f�[�^"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   33
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "�w�s�x�f�[�^�m�F(����)"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "�w�s�x�f�[�^�m�F(�w���)"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "�P�T���@"
      Height          =   375
      Index           =   14
      Left            =   10080
      TabIndex        =   27
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "�P�S���@"
      Height          =   375
      Index           =   13
      Left            =   10080
      TabIndex        =   26
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "�V���@"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame FraGouki 
      Caption         =   "���@�w��"
      Height          =   1935
      Left            =   4800
      TabIndex        =   12
      Top             =   600
      Width           =   6975
      Begin VB.CheckBox chkGouki 
         Caption         =   "�P�U���@"
         Height          =   375
         Index           =   15
         Left            =   5280
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�P�R���@"
         Height          =   375
         Index           =   12
         Left            =   5280
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�P�Q���@"
         Height          =   375
         Index           =   11
         Left            =   3600
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�P�P���@"
         Height          =   375
         Index           =   10
         Left            =   3600
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�P�O���@"
         Height          =   375
         Index           =   9
         Left            =   3600
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�X���@"
         Height          =   375
         Index           =   8
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�W���@"
         Height          =   375
         Index           =   7
         Left            =   2160
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�U���@"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�T���@"
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�S���@"
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�R���@"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�Q���@"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "�P���@"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FraCorner 
      Caption         =   "�R�[�i�w��"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4455
      Begin VB.CheckBox chkCorner 
         Caption         =   "�R�[�i�U"
         Height          =   225
         Index           =   5
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "�R�[�i�T"
         Height          =   225
         Index           =   4
         Left            =   2640
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "�R�[�i�S"
         Height          =   225
         Index           =   3
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "�R�[�i�R"
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "�R�[�i�Q"
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "�R�[�i�P"
         Height          =   225
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��"
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
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   9600
      Top             =   360
   End
   Begin VB.ListBox LstStatus 
      Height          =   2310
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   11655
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   15000
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�f�[�^���W�E�o��  ��ʂ֖߂�"
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�W���[�i����"
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
      TabIndex        =   3
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmJprPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmJprPrint.frm
'//  �p�b�P�[�W���F�W���[�i���󎚉��
'/
'//  �T�v�F�V�X�e��������(�Ď���)���
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-18   CODED   BY [TCC] T.Takajima
'//     REVISIONS :(EG20 V7.4.0.1) 2013-07-22   CODED   BY [TCC] T.Nakajima
'//                 ���܂�����o��t���[�ݒ��ʑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή�
'//                 �yHKRK_Kansi07_003_01�z SUB_GATE_KAN.INI�t�H�[�}�b�g�������Ή�
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-10  CODED   BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(EG30 V32.2.0.1) 2016-07-20  CODED   BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή� �v���O���X�o�[��\���Ή�
'//     REVISIONS :(EG30 V35.3.0.1) 2019-07-03   REVISED BY [TCC] H.Kondoh
'//                2019�N�x�{���Ή�
'//     REVISIONS :
'//
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'���������s�t���O
Private bSysFormat As Boolean

Private Const APL_INTERVAL = 390000     '�A�v���N���^�C�}�f�t�H���g�l
Public glbFilePath  As String             '�t�@�C���p�X     'V1.12.0.1 ADD
Dim lngMAX_Time As Long                    'INI�擾�ݒ�l
Dim lngtime     As Long                    '���݃^�C�}�l
Private iSendType As Integer            '�v����ʒl
'V1.5.0.1 ADD END
'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        '���O�N���^�C�}�f�t�H���g�l(30�b)
Dim lngLogMAX_Time As Long                'INI�擾�ݒ�l(���O�j
'V1.20.0.1 ADD END
Dim intJprFile        As Integer        'EG20 V30.1.0.1 ADD


' �W���[�i���o�͐ݒ���
Private Type JPR_PRINT_SETTING_INFO
    iCornerCount        As Integer          ' �`�F�b�N���ꂽ�R�[�i��
    iCorner(5)          As Integer          ' �`�F�b�N���ꂽ�R�[�i�ꗗ
    iGoukiCount         As Integer          ' �`�F�b�N���ꂽ���@��
    iGouki(15)          As Integer          ' �`�F�b�N���ꂽ���@�ꗗ
    iJprCount           As Integer          ' �`�F�b�N���ꂽ�W���[�i����ސ�
'    iJprKind(7)         As Integer          ' �`�F�b�N���ꂽ�W���[�i���ꗗ      'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
'    iJprKind(8)         As Integer          ' �`�F�b�N���ꂽ�W���[�i���ꗗ      'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD    'EG30 V32.1.0.1 DEL
    iJprKind(9)         As Integer          ' �`�F�b�N���ꂽ�W���[�i���ꗗ      'EG30 V32.1.0.1 ADD
End Type
Private Enum JPR_KIND
    JPR_KIND_EKI_INFO = 0           ' �w�s�x�f�[�^�m�F(�w���)
    JPR_KIND_JIKAI_INFO = 1         ' �w�s�x�f�[�^�m�F(����)
    JPR_KIND_SETTING_LST = 2        ' �ݒ�l�ꗗ
    JPR_KIND_TUKA_DATA = 3          ' �ʉ߃f�[�^
    JPR_KIND_RIYO_KINGAKU = 4       ' ���p���z�f�[�^
    JPR_KIND_KADO_VER = 5           ' �ғ��o�[�W�����ꗗ
    JPR_KIND_SIMEKIRI = 6           ' ���؃I�t���C���o��
    JPR_KIND_EKIMU_ID = 7           ' �w���@��ID
    JPR_KIND_SUBGATE_INFO = 8       ' �w�s�x�f�[�^�m�F(�G���R�[�h�R�[�i���@����`)    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
    JPR_KIND_GATE_CFG = 9          ' ���D�@�ێ�ݒ�f�[�^  'EG30 V32.1.0.1 ADD
End Enum
Dim udtJprPrintSetteingInfo    As JPR_PRINT_SETTING_INFO
Dim udtInitJprSetting           As JPR_PRINT_SETTING_INFO
Dim iJprIdx                     As Integer          '�������̃W���[�i��

'�@��\���f�[�^�i�w���j�C���[�W�t�@�C���ǎ�p�̍\����
Private Type EKIINFO_IMAGE_FILE
    sType       As String                '���
    sGoki       As String                '���@
    sNo         As String                '��ʖ��ʔ�
    sCorner     As String                '�R�[�i        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    sTuuban     As String                '�ʔ�
    sKoumoku    As String                '����
    sKubun      As String                '�敪
    sSettei     As String                '�ݒ�l
    sSyosai     As String                '�ݒ�l�ڍ�
End Type

'�w�s�x�f�[�^�i���D�@)�C���[�W�t�@�C���ǂݎ��p�̍\����
Private Type JIKAIINFO_IMAGE_FILE
    strBunrui_Dai  As String               '�啪��
    strBunrui_Tyu  As String               '������
    srtBunrui_Sho   As String               '������
    strCorner       As String               '�R�[�i
    strKomoku       As String               '����
    strKubun        As String               '�敪
    strData         As String               '�f�[�^
    strSetShosai    As String               '�ڍ�
    
End Type

'�w�s�x�f�[�^�m�F(����)�W���[�i���o�̓t�@�C���쐬�e�[�u��
Private Type JIKAI_JPREDIT_TBL
    strKomoku       As String               '���ږ��i�{�敪)
    strBunrui_Sho   As String               '���̍��ڂ��w�������ރR�[�h
    strKubun        As String               '���̍��ڂ��w���敪
End Type

'�ғ��o�[�W�����o�͋敪
Private Enum mintDispDiv
    KADOVER_FILE_DISP = 0
    KADOVER_FILE_OUTPUT
End Enum

'�}�̏o�̓t�@�C���ǂݎ��p�̍\����(�ʉ�/���p���z)
Private Type BAITAI_OUTPUT_IMAGE_FILE
    strKomokuName       As String          '���ږ�
    strGoukei           As String          '�ʉߍ��v
    srtGoukiValue(15)   As String          '���@�ʂ̒l(���g�p)
End Type

'EG20 V30.1.0.1 ADD START
'�}�̏o�̓t�@�C���ǂݎ��p�̍\����(�ʉ�/���p���z)�y�����p�z
Private Type BAITAI_OUTPUT_IMAGE_FILE_KAN
    strKomokuName       As String          '���ږ�
    strGoukei           As String          '�ʉߍ��v
    strNorikae          As String          '�ʉߏ抷(���g�p)
    strTukaChoku        As String          '�ʉߒ���(���g�p)
    srtGoukiValue(31)   As String          '���@�ʂ̒l(���g�p)
End Type
'EG20 V30.1.0.1 ADD END

'�ݒ�ꗗ�t�@�C���ǂݎ��p�\����(OPERATE_SET##.CSV�j
Private Type SETTEI_OUTPUT_IMAGE_FILE
    strDaiKomoku        As String           '�區�ږ�
    strKomoku           As String           '���ږ�
    strValue            As String           '�ݒ�l
    strChangeFlg        As String           '�ύX�t���O 'EG30 V32.1.0.1 ADD
End Type

'�ғ��o�[�W�����t�@�C���ǂݎ��p�\����(KadoVerDisp.csv)
Private Type KADO_VER_DISP_IMAGE_FILE
    strKishu            As String           '�@�핪�ށi�t�@�C���ǂݍ��ݗp�j
    strCorner           As String           '�R�[�i���ށi�t�@�C���ǂݍ��ݗp�j
    strGokiDiv          As String           '���@���ށi�t�@�C���ǂݍ��ݗp�j
    strName             As String           '�@�햼�i�t�@�C���ǂݍ��ݗp�j
    strMaker            As String           '���[�J���i�t�@�C���ǂݍ��ݗp�j
    strVer              As String           '�o�[�W�����i�t�@�C���ǂݍ��ݗp�j
    strDate             As String           '�쐬���t�i�t�@�C���ǂݍ��ݗp�j
End Type

'EG30 V32.1.0.1 ADD START
'���D�@�ێ�ݒ�f�[�^ �ǂݎ��p�̍\����(JP_CFG�R�[�i���@�ԍ�.csv)
Private Type GATE_CFG_DATA_FILE
    strInfoName         As String           '��񕔖�
    strBunrui_Dai       As String           '�區��
    strBunrui_Chu       As String           '������
    strBunrui_Syo       As String           '������
    strValue            As String           '�ݒ�l
    strChangeFlg        As String           '�ύX�L���t���O
End Type
'EG30 V32.1.0.1 ADD END


'�W���[�i���ҏW���ԃt�@�C��
Private Const EKIMU_DEFU = "APL\APL_WORK"
Private Const EDIT_DATA_EKIINFO = PATH_WORK & "EKI_DISP_EKIINFO.csv"    '�w�s�x�f�[�^�m�F(�w���)"
Private Const EDIT_DATA_JIKAIINFO = PATH_WORK & "EKI_DISP_GATE_JPR.csv" '�w�s�x�f�[�^�m�F(����)"
Private Const EDIT_DATA_SETTEI = PATH_WORK & "OPERATE_SET##.csv"        '�ݒ�l�ꗗ
Private Const EDIT_DATA_KADOVERSION = PATH_WORK & "KadoVerDisp####"     '�ғ��o�[�W�����ꗗ�iKadoVerDisp�R�[�i�ԍ��A���@�ԍ��j
Private Const EDIT_DATA_SIMEKIRI = PATH_WORK & "SIME##.txt"             '���؃I�t���C���o��
Private Const EDIT_DATA_EKIMUID = PATH_WORK & "MN_VERSI.txt"            '�w���@��ID
Private Const EDIT_DATA_TUKA = PATH_SHUKEI_SEND & "TUKA*.csv"           '�ʉ߃f�[�^
Private Const EDIT_DATA_RIYO = PATH_SHUKEI_SEND & "ICRIYO*.csv"         '���p���z�f�[�^
Private Const EDIT_DATA_GATECFG = PATH_WORK & "JP_CFG####"              '���D�@�ێ�ݒ�f�[�^  'EG30 V32.1.0.1 ADD
Private Const APL = "APL"
Private Const LOG = "LOG"
Private Const Data = "DATA"
Private Const BACKUP = "BACKUP"

Private Const MAX_KOMOKU_NUM_TUKA = 51                      '�ʉߊO���}�̍ő區�ڐ�
Private Const MAX_KOMOKU_NUM_KINGAKU = 16                   '���z�O���}�̍ő區�ڐ�
'EG20 V30.1.0.1 ADD START
Private Const MAX_TUKA_SHUKEI_KOUMOKU = 7                                 '�����ʉ߃f�[�^�̍ő�W�v���ڐ��i�u���b�N�P�ʁj
Private Const MAX_KOMOKU_NUM_TUKA_KAN = 51                                '�����ʉ߃f�[�^ �ő區�ڐ�
Private Const MAX_KOMOKU_NUM_UNKOU_FUNOU = 1                              '�����ʉ߃f�[�^ �^�s�s�\�f�[�^ �ő區�ڐ�
Private Const MAX_KOMOKU_NUM_NORIKAE_TUKA = 51                            '���� �抷 �ݗ����ʉ߃f�[�^ �ő區�ڐ�
Private Const MAX_KOMOKU_NUM_JIEKI_KYUSAI = 51                            '���� ���w����~�ϒʉ߃f�[�^ �ő區�ڐ�
Private Const MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI = 51                      '���� ���C��������~�ʉ߃f�[�^ �ő區�ڐ�

Private Const MAX_KINGAKU_SHUKEI_KOUMOKU = 11                             '�������z�f�[�^�̍ő�W�v���ڐ��i�u���b�N�P�ʁj
Private Const MAX_KOMOKU_NUM_SUICA_RIYO = 11                              '�������z�f�[�^ �X�C�J���p���z�@�ő區�ڐ�
Private Const MAX_KOMOKU_NUM_SUICA_SEISAN = 32                            '�������z�f�[�^ �X�C�J��ЊԐ��Z�f�[�^ �ő區�ڐ�
Private Const MAX_KOMOKU_NUM_AUTOCHARGE = 34                              '�������z�f�[�^ �I�[�g�`���[�W�f�[�^�@�ő區�ڐ�


'�W�v���ځi�ʉ߃f�[�^�j
Private Enum mintTukaShukeiKoumoku
    SHUKEI_KAISATU_KANSEN_TUKA = 0      '�y���D���@�V�����ʉ߃f�[�^�z
    SHUKEI_SHUSATU_KANSEN_TUKA          '�y�W�D���@�V�����ʉ߃f�[�^�z
    SHUKEI_IC_UNKO_FUNOU                '�y�^�s�s�\�f�[�^�z
    SHUKEI_KAN_ZAI_TUKA                 '�y��-�ݏ抷�ʉ߃f�[�^�z
    SHUKEI_ZAI_KAN_TUKA                 '�y��-���抷�ʉ߃f�[�^�z
    SHUKEI_JIEKI_KYUSAI                 '�y���w����~�ϒʉ߃f�[�^�z
    SHUKEI_KAISHU_CHUSHI                '�y���C��������~�ʉ߃f�[�^�z
End Enum

'�W�v���ځi���z�f�[�^�j
Private Enum mintKingakuShukeiKoumoku
    SHUKEI_KAI_OTONA_SUICA_RIYO         '�y���D���@��l�@�V�����X�C�J���p���v���z�z
    SHUKEI_SHU_OTONA_SUICA_RIYO         '�y�W�D���@��l�@�V�����X�C�J���p���v���z�z
    SHUKEI_KAI_SHONI_SUICA_RIYO         '�y���D���@�����@�V�����X�C�J���p���v���z�z
    SHUKEI_SHU_SHONI_SUICA_RIYO         '�y�W�D���@�����@�V�����X�C�J���p���v���z�z
    SHUKEI_SEISAN_SHIHARAI              '�y�X�C�J��ЊԐ��Z�f�[�^�@�^���x���z�z
    SHUKEI_KAI_AUTOCHARGE               '�y���D���@�I�[�g�`���[�W�f�[�^�z
    SHUKEI_SHU_AUTOCHARGE               '�y�W�D���@�I�[�g�`���[�W�f�[�^�z
    SHUKEI_KAN_OTONA_SUICA_RIYO         '�y�����^���@��l�@�X�C�J���p���v���z�z
    SHUKEI_KAN_SHONI_SUICA_RIYO         '�y�����^���@�����@�X�C�J���p���v���z�z
    SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO    '�y�抷�ݗ��^���@��l�@�X�C�J���p���v���z�z
    SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO    '�y�抷�ݗ��^���@�����@�X�C�J���p���v���z�z
End Enum

'GAIBU_OUTPUT.INI�̃L�[�ԍ�
Private Enum mintGaibuOutputKey
    GAIBU_INI_TUKA = 0                  '�ʉ߃f�[�^
    GAIBU_INI_ICSF_KIKAN                'ICSF���s���ԕʗ��p���z�f�[�^
    GAIBU_INI_IC_CARD_SHIHARAI          'IC�J�[�h��ЊԐ��Z�f�[�^�i�^���x���z�j
    GAIBU_INI_AUTO_CHARGE               '�I�[�g�`���[�W�f�[�^
    GAIBU_INI_IC_UNKOU_FUNOU            'IC�J�[�h�^�s�s�\�����f�[�^
    GAIBU_INI_TUKA_KAN_ZAI              '��-�ݏ抷�ʉ߃f�[�^
    GAIBU_INI_TUKA_ZAI_KAN              '��-���抷�ʉ߃f�[�^
    GAIBU_INI_IC_KIKAN_KANSEN           '�����^��IC���s�@�֕ʗ��p���z�f�[�^
    GAIBU_INI_IC_KIKAN_ZAIRAI           '�抷�ݗ��^��IC���s���ԕʗ��p���z�f�[�^
    GAIBU_INI_KYUSAI                    '���w����~�ϒʉ߃f�[�^
    GAIBU_INI_KAISHU_CHUSI              '���C��������~�ʉ߃f�[�^
End Enum

'Private ReadSetteiSubGate()             As JIKAIINFO_IMAGE_FILE     'EKI_DISP_SUBGATE.CSV��1�R�[�i���̃f�[�^    'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL
'EG20 V30.1.0.1 ADD END
Private ReadSetteiSubGate(0 To 191)           As JIKAIINFO_IMAGE_FILE     'EKI_DISP_SUBGATE.CSV 1�`32���@ �@�`�E�i�Œ�j     'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
Private Const SUBGATE_ITEM_NUM = 6      ' SUB_GATE_KAN.INI�̎��Е��̍��ڐ��F6                                   'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD

Private Const MAX_JPR_KETA_MAX = 30 'JPR1�s�ő�30�o�C�g(���p30����)

Private Const MAX_KADO_PG = 6       '���D�@1�䓖����ɓ���v���O�������i�v���O��������f�[�^)

Private Const FOOTER_STRING = "*************END**************"


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : cmdPrint_Click
'//  �@�\����  : �u�󎚁v�t����������
'//  �@�\�T�v  : ������������s����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17  CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdPrint_Click()
    Dim i       As Integer
    Dim bRet    As Boolean
    Dim intCount    As Integer
    
    '�u�W���[�i���󎚉�ʁF�󎚊J�n�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_BUTTON, 0)
    
    '�{�^���A�`�F�b�N�{�b�N�X���A�N�e�B�u�ɂ���
    Call JPRScreenEnable(False)
    
    ' �ݒ��Ԃ��擾����
    Call GetPrintSettings
    
    ' �R�[�i�`�F�b�N
    If udtJprPrintSetteingInfo.iCornerCount = 0 Then
        '�R�[�i�ɉ����`�F�b�N����Ă��Ȃ��̂ŏ������s
        LstStatus.AddItem "�R�[�i���`�F�b�N����Ă��܂���"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    '���@�`�F�b�N
    If udtJprPrintSetteingInfo.iGoukiCount = 0 Then
        '���@�ɉ����`�F�b�N����Ă��Ȃ��̂ŏ������s
        LstStatus.AddItem "���@���`�F�b�N����Ă��܂���"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    '�󎚍��ڎw��`�F�b�N
    If udtJprPrintSetteingInfo.iJprCount = 0 Then
        '�R�[�i�ɉ����`�F�b�N����Ă��Ȃ��̂ŏ������s
        LstStatus.AddItem "�󎚍��ڂ��w�肳��Ă��܂���"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    '�`�F�b�N���ꂽ�R�[�i�͐ݒu����Ă��邩���Ȃ����̏����Z�b�g���Ă���
    Erase glngTergetCorner
    For intCount = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        '���̃R�[�i���ݒu����Ă��邩�H
        If IsTaisyoCorner(udtJprPrintSetteingInfo.iCorner(intCount)) = True Then
            glngTergetCorner(udtJprPrintSetteingInfo.iCorner(intCount) - 1) = CMN_ONOFF.CMN_ON
        Else
            glngTergetCorner(udtJprPrintSetteingInfo.iCorner(intCount) - 1) = CMN_ONOFF.CMN_OFF
        End If
    Next intCount
    
    '�󎚍��ڃ`�F�b�N�ɉ����ĕҏW�������Ăяo���B
    iJprIdx = 0
    Call JprOutputProc
 
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JPREdit_EkiInfo
'//  �@�\����  : �w�s�x�f�[�^�m�F(�w���)�C���[�W�t�@�C���쐬
'//  �@�\�T�v  : �w�s�x�f�[�^�m�F(�w���)�̃W���[�i���C���[�W�t�@�C�����쐬����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-27   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-15  REVISED BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_EkiInfo() As Boolean

    Dim strFileName          As String          '�t�@�C����
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim lRetVal              As Long            '�߂�l
    Dim sCommand             As String          '�R�}���h������
'V1.12.0.1 ADD START
    Dim sWriteDir            As String              '�������ݐ�t�H���_��
    Dim intFileNumber        As Integer             '�t�@�C���|�C���^
    Dim strLineCount         As String              '�s���J�E���^
    Dim i                    As Integer             '���[�v�J�E���^�P
    Dim j                    As Integer             '���[�v�J�E���^�Q
    Dim k                    As Integer             '���[�v�J�E���^�R
    Dim l                    As Integer             '���[�v�J�E���^�S
    Dim ReadFileSettei()     As EKIINFO_IMAGE_FILE  '�t�@�C���Ǎ��p�\����
    Dim fso         As New FileSystemObject         '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim FsoTS As TextStream

    Dim bRet                 As Boolean         '�֐��߂�l
    Dim lErrCode             As Long            '�G���[�R�[�h
    
    Dim strNowType          As String           '�������啪��
    Dim strNowShoNo         As String           '���������ڔԍ�
    Dim strNowTuban         As String           '���������ڒʔ�
    Dim strNowCorner        As String           '�������R�[�i
    Dim strNowKubn          As String           '�������敪
    
    Dim intCount            As Integer          '�R�[�i�C���f�b�N�X �O�F�R�[�i1
    
    'EG30 V32.1.0.1 ADD START
    Dim strEkiSettiBefPath  As String           '���݉w�ݒ�f�[�^�i�ύX�O�ۑ��j
    Dim strGetValue         As String * 64      'DLL�ɂ���Đݒ肳��邽�߁A64�Œ蒷�ɂ��Ă���
    Dim strCompValue        As String           '�ݒ�l�i�ύX�O�ۑ��j
    Dim strChangeFlg        As String           '�ύX��
    Dim intValueLen         As Integer          '�擾�����ݒ�l�̒���
    'EG30 V32.1.0.1 ADD END
    
    On Error GoTo Err_handler
    
    
    '��ʂŎw�肳�ꂽ�R�[�i�͐ݒu����Ă��邩�H
    If pfSettingCheck(False) = False Then
        '���ׂĖ��ݒu�̃R�[�i�Ȃ̂ŃG���[�Ƃ���
        GoTo Err_handler
    End If
    
    '////////////////////////////////////////////////
    '// �R�[�i������ʂ�擾
    gsGetCornerName
   
    '�e�X�g�łƂ肠�����A�R�[�i1
    intCount = 0
    
    '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���쐬
    bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���폜
        Kill EKI_TUDO_CHK_EKI_INFO_FILE
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
        JPREdit_EkiInfo = False
        Exit Function
    End If
    
    'CSV�t�@�C���̌����擾
    intFileNumber = FreeFile
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber
    
    Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1�ǉ�
        Line Input #intFileNumber, strLineCount
        j = j + 1
    Loop
    'CSV�t�@�C���N���[�Y
    Close #intFileNumber
    
    '��L�������A��������ɕێ�
    '�Đݒ�
    ReDim ReadFileSettei(j) As EKIINFO_IMAGE_FILE   '�t�@�C���Ǎ��p�G���A
        
    'CSV�t�@�C���I�[�v��
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber

    '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
    For i = 0 To UBound(ReadFileSettei) - 1
        Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, _
        ReadFileSettei(i).sCorner, ReadFileSettei(i).sTuuban, ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, _
        ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
    Next i

    'CSV�t�@�C���N���[�Y
    Close #intFileNumber
    
    'EG30 V32.1.0.1 ADD START
    '���̃R�[�i�̕ύX�O�f�[�^�ۑ����ꂽ�f�[�^����������ɓW�J����
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", "0")
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    '/////////////////////////////////////
    '�W���[�i���C���[�W�t�@�C���쐬
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
   
    '�W���[�i���o�̓C���[�W�t�@�C�����쐬
    Open EKI_JPR_EKIINFO_TXTFILE For Output As #intFileNumber
    
    '�^�C�g���\��
    'PrintHeader intFileNumber, "�w�s�x�f�[�^�m�F�i�w���j"    'EG30 V32.1.0.1 DEL
    PrintHeader3 intFileNumber, "�w�s�x�f�[�^�m�F�i�w���j", pfGetSaveDate(0) 'EG30 V32.1.0.1 ADD
    Print #intFileNumber, "�ݒu�w�F" & Trim(pfGetEkiNameInfo(NotEkiVer))
    '�`�F�b�N���ꂽ�R�[�i�����Ń��[�v
    For k = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        intCount = udtJprPrintSetteingInfo.iCorner(k) - 1  '��ʂŎw�肳�ꂽ�R�[�i-1
        If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(k)) = True Then
            
            ' ���̃R�[�i�͐ݒu����Ă���̂ŃW���[�i���o�͂�
            '1�R�[�i�ڂ����ݒu�w�Ɛݒu�R�[�i�̊Ԃ͋�s���Ȃ�
            If k <> 0 Then
                Print #intFileNumber, ""
            End If
            Print #intFileNumber, "�ݒu�R�[�i�F" & gstrCornerName(intCount)

            '////////////////////////////////
            '// �e�ݒ���o��
            '////////////////////////////////
            strNowType = ""
            strNowShoNo = ""
            strNowKubn = ""
            
            For i = 0 To UBound(ReadFileSettei) - 1
            
                If strNowType <> ReadFileSettei(i).sType Then
                    '�V�����啪�ދ敪�ɂȂ����̂Ń^�C�g������
                    Print #intFileNumber, ""
                    Select Case ReadFileSettei(i).sType
                        Case "1"
                            'Print #intFileNumber, "�y�w���z"     'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "�@�y�w���z"    'EG30 V32.1.0.1 ADD
                        Case "2"
                            'Print #intFileNumber, "�y�Ď��z"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "�@�y�Ď��z"  'EG30 V32.1.0.1 ADD
                        Case "3"
                            'Print #intFileNumber, "�y�l�b�g���[�N�z"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "�@�y�l�b�g���[�N�z"  'EG30 V32.1.0.1 ADD
                        Case "7"
                            'Print #intFileNumber, "�y��ʁz"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "�@�y��ʁz"  'EG30 V32.1.0.1 ADD
                    End Select
                    strNowType = ReadFileSettei(i).sType
                End If
                
                '���ڔԍ����O��Ɠ����ꍇ�͏o�͂��Ȃ�
                'If strNowShoNo <> ReadFileSettei(i).sNo Then
                If strNowShoNo <> ReadFileSettei(i).sNo Or strNowKubn <> ReadFileSettei(i).sKubun Then
                    '���ږ�+�敪+�ݒ�l���o��
                    If (CInt(ReadFileSettei(i).sCorner) = intCount + 1) Or (CInt(ReadFileSettei(i).sCorner) = 0) Then
                        
                        'EG30 V32.1.0.1 ADD START
                        '�ύX�O�f�[�^�ۑ����ꂽ�ݒ�l�Ɣ�r����
                        bRet = dllGetEkiInfoValue(CInt(ReadFileSettei(i).sType), _
                                                    CInt(ReadFileSettei(i).sGoki), _
                                                    CInt(ReadFileSettei(i).sNo), _
                                                    CInt(ReadFileSettei(i).sCorner), _
                                                    strGetValue, _
                                                    intValueLen)
                        strCompValue = strGetValue
                        If (intValueLen <> 0) Then
                            strCompValue = MidByte(strGetValue, 1, intValueLen)
                            strCompValue = Trim(strCompValue)
                        ElseIf (intValueLen = 0) Then
                            strCompValue = ""
                        End If
                        
                        If (bRet = False) Or (ReadFileSettei(i).sSettei <> strCompValue) Then
                            strChangeFlg = DIFF_MARK_STRING_ON
                        Else
                            strChangeFlg = DIFF_MARK_STRING_OFF
                        End If
                        'EG30 V32.1.0.1 ADD END
                        
                        
                        '/////////////////////////////////////////
                        '//���L�̍��ڂ͉w�s�x�f�[�^�ƃW���[�i���̏o�̓f�[�^���قȂ�`���ɂȂ�̂ŋ����I�ɕϊ�����
                        '/////////////////////////////////////////
                        
                        '�啪�ށF�P �����ށF�O �����ށF�P�W�u���ށv�̒l�́u9 9 9 9 9 9�v�`�� ���p�X�y�[�X2������1�����ɕύX
                        If (ReadFileSettei(i).sType = "1") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "18") Then
                           ReadFileSettei(i).sSettei = Replace(ReadFileSettei(i).sSettei, "  ", " ")
                        End If
                        '�啪�ށF�Q �����ށF�O �����ށF�P�u�R�[�i�ԍ��i�΂h�c�T�[�o)�v�̒l�́u�h�c�v�͑S�p
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "1") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "�h�c")
                        End If
                        '�啪�ށF�Q �����ށF�O �����ށF�Q�u�i�΂h�c�T�[�o)�v�̒l�́u�h�c�v�͑S�p
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "2") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "�h�c")
                        End If
                        '�啪�ށF�Q �����ށF�O �����ށF�Q�u�i�΂h�c�T�[�o)�v�u�h�c�v�͑S�p
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "2") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "�h�c")
                        End If
                        '�啪�ށF�Q �����ށF�O �����ށF�R�u�i�΂h�c�T�[�o)�v�u�h�c�v�͑S�p
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "3") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "�h�c")
                        End If
                        '�啪�ށF�Q �����ށF�O �����ށF�W�u�i�΂h�c�T�[�o)�v�u�h�c�v�͑S�p
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "8") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "�h�c")
                        End If

                        '�啪�ށF�Q �����ށF�O �����ށF�X�u�i�΂h�c�T�[�o)�v�u�h�c�v�͑S�p
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "9") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "�h�c")
                        End If
                     
                        '�啪�ށF�V �����ށF�O �����ށF�Q�P�u�ێ烆�[�U�ݒ胁�j���[��� ���l���[�h����ݒ�t�v
                        '���ږ��Ƌ敪�̊ԂɃX�y�[�X���ЂƂ��������
                        If (ReadFileSettei(i).sType = "7") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "21") Then
                           ReadFileSettei(i).sKoumoku = ReadFileSettei(i).sKoumoku & Space(1)
                        End If
                     
                        'Print #intFileNumber, ReadFileSettei(i).sKoumoku & " " & ReadFileSettei(i).sKubun & " " & ReadFileSettei(i).sSettei    'EG30 V32.1.0.1 DEL
                        Print #intFileNumber, strChangeFlg & ReadFileSettei(i).sKoumoku & " " & ReadFileSettei(i).sKubun & " " & ReadFileSettei(i).sSettei  'EG30 V32.1.0.1 ADD
                        strNowShoNo = ReadFileSettei(i).sNo
                        strNowKubn = ReadFileSettei(i).sKubun
                    End If
                End If
            
            Next i
        Else
             '�ݒu����Ă��Ȃ��R�[�i�Ȃ̂Ŏ��̃R�[�i��
        End If
    Next k
    
    Print #intFileNumber, ""
    Print #intFileNumber, FOOTER_STRING
    
    Close #intFileNumber
    
    JPREdit_EkiInfo = True
    Exit Function
    
Err_handler:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    '�ُ�I��
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ�e�L�X�g�o�͌���")
    JPREdit_EkiInfo = False

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-14   CODED   BY [TCC] N.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    '�u�W���[�i���󎚉�ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �V�X�e��������(�Ď���)���(�A�N�e�B�u��)
'//  �@�\�T�v  : �őO�ʕ\�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    pfFormActive (hwnd)
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �V�X�e��������(�Ď���)���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �W���[�i���󎚉��(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.0.1.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer    '�J�E���^�[
   
    On Error Resume Next
    
    '�u�W���[�i���󎚉�ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_GAMEN_START, 0)
    
    ' �R�[�i�`�F�b�N�{�b�N�X
    For i = 0 To chkCorner.Count - 1
        chkCorner(i).Value = 1
    Next i

    ' ���@�`�F�b�N�{�b�N�X
    For i = 0 To chkGouki.Count - 1
        chkGouki(i).Value = 1
    Next i

    ' �f�[�^����
    For i = 0 To chkJprKind.Count - 1
        chkJprKind(i).Value = 0
    Next i
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   
   'INI�t�@�C�����A�v���N���^�C�}�l���擾
   lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      APL_INTERVAL, HOSHU_FILE)
   '�擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
   If lngMAX_Time = 0 Then
      lngMAX_Time = APL_INTERVAL
   End If
   
   'INI�t�@�C����胍�O�N���^�C�}�l���擾
   lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
   '�擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
   If lngLogMAX_Time = 0 Then
      lngLogMAX_Time = LOG_INTERVAL
   End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
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
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF           '���[����M�G���A
    Dim lngLength As Long                    '��M���[���o�C�g�T�C�Y
    Dim lngMlSts  As Long                    '��M���[���̃X�e�[�^�X
    Dim bRet  As Boolean
    Dim lngDataKind As Long                 '��ʏo�͗v��RES�̃f�[�^���
    
    On Error Resume Next

    '���[������M����B
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
   '��M���[��������΁A���[���h�c���̏���������B
        Select Case udtReadMail.udtlHeader.dwId        '���[���h�c
            Case ML_ID_JPR_PRINT_RES
                '�u�W���[�i������v��RES��M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, JPR_PRINT_RES_RECV, 0)
                lngMlSts = udtReadMail.lngData(0)
                If (udtJprPrintSetteingInfo.iJprKind(iJprIdx) = JPR_KIND.JPR_KIND_TUKA_DATA) Or _
                   (udtJprPrintSetteingInfo.iJprKind(iJprIdx) = JPR_KIND.JPR_KIND_RIYO_KINGAKU) Then
                    
                    '�ʉ߃f�[�^�܂��͗��p���z���o�͂��Ă���Ƃ��̓W���[�i���󎚗v��RES����M������A
                    '�W�v�ɉ�ʏo�͊����ʒm�𑗐M����B
                    If lngMlSts = 0 Then
                        bRet = SendMessageGamenOutComplete(ML_GAMEN_OUT_STS.ML_STS_OK)
                    Else
                        bRet = SendMessageGamenOutComplete(ML_GAMEN_OUT_STS.ML_STS_NG)
                    End If
                Else
                    bRet = True
                End If
                
                If (lngMlSts = 0) And (bRet = True) Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), True)
                Else
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                End If
            
            Case ML_ID_INFO_RES
                '�u���v��RES��M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
                ' �������ʂ��m�F
                lngMlSts = udtReadMail.lngData(1)
                If lngMlSts = 0 Then
                    '�ҏW�������s���B
                    bRet = JprEdit_EkimuId()
                    If bRet = True Then
                       '�W���[�i���󎚗v��CMD�𑗐M
                        bRet = SendMessageJprPrint(EKIMUKIKI_ID_TXTFILE, ML_CUT_ARI)
                        If bRet = False Then
                            Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                            Exit Sub
                        End If
                    Else
                        Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                    End If
                    
                Else
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                End If
            
            Case ML_ID_GAMEN_OUTPUT_RES
                '�u��ʏo�͗v��RES��M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
                '�������ʊm�F
                lngMlSts = udtReadMail.lngData(1)
                '�f�[�^���
                lngDataKind = udtReadMail.lngData(2)
                If lngMlSts = 0 Then
                    '�ҏW�������s��
                    bRet = JprEdit_TukaData(lngDataKind)
                    If bRet = True Then
                        '�W���[�i���󎚗v��CMD�𑗐M
                        If lngDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
                            bRet = SendMessageJprPrint(TUKA_TXTFILE, ML_CUT_ARI)
                        ElseIf lngDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then
                            bRet = SendMessageJprPrint(ICRIYO_TXTFILE, ML_CUT_ARI)
                        Else
                            bRet = False
                        End If
                            
                        If bRet = False Then
                            '�ҏW���������s�A��ʏo�͊����ʒm���ُ�ő��M
                            SendMessageGamenOutComplete (ML_GAMEN_OUT_STS.ML_STS_NG)
                            '�ُ�Ȃ̂ŁA��ʏo�͊����ʒm���b�Z�[�W�̑��M�Ɏ��s���悤���������ʂُ͈�
                            Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                            Exit Sub
                        End If
                    Else
                        '�ҏW���������s�A��ʏo�͊����ʒm���ُ�ő��M
                        SendMessageGamenOutComplete (ML_GAMEN_OUT_STS.ML_STS_NG)
                        '�ُ�Ȃ̂ŁA��ʏo�͊����ʒm���b�Z�[�W�̑��M�Ɏ��s���悤���������ʂُ͈�
                        Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                        Exit Sub
                    End If
                Else
                    'RES���ُ�̂��߁A�I���i��ʏo�͊����ʒm�͑��M���Ȃ�)
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                End If
                
            Case ML_ID_PROEND_ORD
                '�u�v���Z�X�I���w���v����M�����ꍇ�A
                '�u�v���Z�X�I���w����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                '�v���Z�X�̏I���������s��
                pfAbortProc
            
            Case ML_ID_HOSHU_ACTIVE_REQ
                '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
                AppActivate frmJprPrint.Caption, False
                pfFormActive (frmJprPrint.hwnd)
                
            Case Else
                '�u���[��ID�s���v���O�o��
                Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : SendMessageJprPrint
'//  �@�\����  : �W���[�i���󎚗v�����b�Z�[�W�𑗐M����
'//  �@�\�T�v  : �o�̓v���Z�X�ɃW���[�i���󎚗v���𑗐M����
'//
'//              �^        ����      �Ӗ�
'//  ����      : String    strFileName   �o�̓t�@�C����
'//              Byte      byCut         0:�J�b�g�Ȃ�   1�F�J�b�g����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function SendMessageJprPrint(strFileName As String, byCut As Byte) As Boolean

    Dim udtMail As MAIL_JPR_PRINT_CMD   '�W���[�i������v�����[�����M�G���A
    Dim lngRet As Long                  '�֐��߂�l
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim bTmpArray() As Byte
    Dim i       As Integer
    On Error Resume Next
    
    
    '�W���[�i���󎚗v�����o�̓v���Z�X�ɑ��M����B
    udtMail.mlHeader.dwId = ML_ID_JPR_PRINT_REQ
    udtMail.mlHeader.dwSize = MlSize.JPR_PRINT_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    bTmpArray = StrConv(strFileName, vbFromUnicode)
    For i = 0 To UBound(bTmpArray)
        'udtMail.byOutputFilePath(i) = Chr(bTmpArray(i))
        udtMail.byOutputFilePath(i) = bTmpArray(i)
    Next
    udtMail.dwCut = byCut                                   '�J�b�g�L��
    udtMail.dwOutputDataPoint = 0                           '�o�̓f�[�^�|�C���g
    
    lngRet = DssSendMail(MAIL_SLOT_OUTPUT, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       '�u�W���[�i���󎚉�ʁF�W���[�i������v�����M�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, JPR_PRINT_REQ_SEND, lngErrCode)
       SendMessageJprPrint = False
       Exit Function
    Else
       '�u�W���[�i���󎚉�ʁF�W���[�i������v�����M����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, JPR_PRINT_REQ_SEND, 0)
       SendMessageJprPrint = True
    End If
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : SendMessageInfoReq
'//  �@�\����  : ���v��CMD���b�Z�[�W�𑗐M����
'//  �@�\�T�v  : ID���ɏ��v���v��CMD�𑗐M����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function SendMessageInfoReq() As Boolean
    
    Dim bRet As Boolean                 '�߂�l
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim udtMail As MAIL_INFO_CMD        '��ʕ\���v��
    Dim uMail As ML_KYOTU_INF           '���[��
 
   '�o�b�t�@�t���b�V���v�������O�v���Z�X�ɑ��M����
   '���v��CMD(�w���@��ID=0)��ID����ɑ��M����
   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
   iSendType = MailCmdType.ML_DT_EKIMU_ID
   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '�u�w���@��ID�m�F�F���v��CMD���M�ُ�v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageInfoReq = False
      Exit Function
   Else
      '�u�w���@��ID�m�F�F���v��CMD���M����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If
   
   SendMessageInfoReq = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : SendMessageGamenOutReq
'//  �@�\����  : ��ʏo�͗v��CMD���M
'//  �@�\�T�v  : �W�v�ɉ�ʏo�͗v��CMD�𑗐M����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function SendMessageGamenOutReq(dwDataKind As Long) As Boolean
    
    Dim bRet As Boolean                     '�߂�l
    Dim lngErrCode As Long                  '�G���[�R�[�h
    Dim udtMail As MAIL_GAMEN_OUTPUT_CMD    '��ʏo�͗v��
    Dim uMail As ML_KYOTU_INF               '���[��
 
   '�o�b�t�@�t���b�V���v�������O�v���Z�X�ɑ��M����
   '��ʏo�͗v��CMD���W�v�ɑ��M����
   udtMail.mlHeader.dwId = ML_ID_GAMEN_OUTPUT_REQ
   udtMail.mlHeader.dwSize = MlSize.GAMEN_OUT_REQ
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwSeqence = 0                ' �V�[�P���X�ԍ�0�Œ�
   udtMail.dwDataKind = dwDataKind
   bRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '�u��ʏo�͗v��CMD���M�ُ�v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageGamenOutReq = False
      Exit Function
   Else
      '�u��ʏo�͗v��CMD���M����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If

   SendMessageGamenOutReq = True

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : SendMessageGamenOutComplete
'//  �@�\����  : ��ʏo�͗v�������ʒm���M
'//  �@�\�T�v  : �W�v�ɉ�ʏo�͊����ʒm�𑗐M����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Long     dwStatus   ���b�Z�[�W�ɃZ�b�g����X�e�[�^�X
'//
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function SendMessageGamenOutComplete(dwStatus As Long) As Boolean
    
    Dim bRet As Boolean                     '�߂�l
    Dim lngErrCode As Long                  '�G���[�R�[�h
    Dim udtMail As MAIL_GAMEN_OUTPUT_COMP   '��ʏo�͗v�������ʒm
    Dim uMail As ML_KYOTU_INF               '���[��
 
   '�o�b�t�@�t���b�V���v�������O�v���Z�X�ɑ��M����
   '��ʏo�͗v��CMD���W�v�ɑ��M����
   udtMail.mlHeader.dwId = ML_ID_GAMEN_OUTPUT_COMP
   udtMail.mlHeader.dwSize = MlSize.GAMEN_OUT_COMP
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwSequence = 0                ' �V�[�P���X�ԍ�0�Œ�
   udtMail.dwStatus = dwStatus
   bRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '�u��ʏo�͗v�������ʒm���M�ُ�v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageGamenOutComplete = False
      Exit Function
   Else
      '�u��ʏo�͗v�������ʒm���M����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If
   
   SendMessageGamenOutComplete = True

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : ResultDisp
'//  �@�\����  : �W���[�i��������ʕ\��
'//  �@�\�T�v  : �W���[�i���̈�����ʂ�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer    iJprKind    �W���[�i�����
'//              Boolean    bResult     ����(true/false)
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  REVISED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_003_01�z
'//     REVISIONS :(32.1.0.1) 2016-06-10  REVISED BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub ResultDisp(iJprKind As Integer, bResult As Boolean)
    Dim strStatus   As String
    Dim strJprName  As String

    '�������ʕ����쐬
    Select Case iJprKind
        Case JPR_KIND.JPR_KIND_EKI_INFO
            strJprName = "�w�s�x�f�[�^�m�F�i�w���j"
            
        Case JPR_KIND.JPR_KIND_JIKAI_INFO
            strJprName = "�w�s�x�f�[�^�m�F�i�����j"
            
        Case JPR_KIND.JPR_KIND_SETTING_LST
            strJprName = "�ݒ�l�ꗗ"
            
        Case JPR_KIND.JPR_KIND_TUKA_DATA
            strJprName = "�ʉ߃f�[�^"
            
        Case JPR_KIND.JPR_KIND_RIYO_KINGAKU
            strJprName = "���p���z�f�[�^"
            
        Case JPR_KIND.JPR_KIND_KADO_VER
            strJprName = "�ғ��o�[�W�����ꗗ"
            
        Case JPR_KIND.JPR_KIND_SIMEKIRI
            strJprName = "���؃I�t���C���o��"
            
        Case JPR_KIND.JPR_KIND_EKIMU_ID
            strJprName = "�w���@��h�c"
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        Case JPR_KIND.JPR_KIND_SUBGATE_INFO
            strJprName = "�w�s�x�f�[�^�m�F�i�ݺ��޺�ō��@����`�j"
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
        'EG30 V32.1.0.1 ADD START
        Case JPR_KIND.JPR_KIND_GATE_CFG
            strJprName = "���D�@�ێ�ݒ�f�[�^"
        'EG30 V32.1.0.1 ADD END
    End Select
    
    If bResult = True Then
        '����
        LstStatus.AddItem strJprName & "    " & "����I�����܂���"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        'Call JPRScreenEnable(True)
    Else
        '�ُ�
        LstStatus.AddItem strJprName & "    " & "�ُ�I�����܂���"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        'Call JPRScreenEnable(True)
    End If

    iJprIdx = iJprIdx + 1
    If iJprIdx < udtJprPrintSetteingInfo.iJprCount Then
        '2��ޖڈȍ~�̃W���[�i���o��
        JprOutputProc
    Else
        '�S�W���[�i���o�͊����Ȃ�΁A�t�A�`�F�b�N�{�b�N�X����\�ɂ���B
        Call JPRScreenEnable(True)
        iJprIdx = 0
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JPRScreenEnable
'//  �@�\����  : �W���[�i�������ʂ̐ݒ�ύX�ې���
'//  �@�\�T�v  : �W���[�i���󎚉�ʂ̓��e��ύX�̉ۂ𐧌䂷��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Boolean   bEnable    true:�ύX�\  false:�ύX�s��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.2.0.1) 2016-07-20  CODED   BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή� �v���O���X�o�[��\���Ή�
'//                 ���[���������ȏ���g�p����W���[�i�������邽�߁A�v���O���X�o�[�\�����ɊԂɍ���Ȃ��Ȃ邽��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub JPRScreenEnable(bEnable As Boolean)
    Dim i   As Integer
    
    ' �R�[�i�`�F�b�N�{�b�N�X
    For i = 0 To chkCorner.Count - 1
        chkCorner(i).Enabled = bEnable
    Next i
    
    ' ���@�`�F�b�N�{�b�N�X
    For i = 0 To chkGouki.Count - 1
        chkGouki(i).Enabled = bEnable
    Next i
 
    ' �f�[�^����
    For i = 0 To chkJprKind.Count - 1
        chkJprKind(i).Enabled = bEnable
    Next i
    
    '�󎚃{�^��
    cmdPrint.Enabled = bEnable
    
    '�߂�{�^��
    cmdReturn.Enabled = bEnable
    
    If bEnable = False Then
        '�X�e�[�^�X�\�������N���A����i�󎚃{�^�������őO��̏������ʂ��N���A)
         LstStatus.Clear
        '�v���O���X�o�[��\������
        ' EG30 V32.2.0.1 DEL START
        'Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_JPR_OUT)
        ' EG30 V32.2.0.1 DEL END
    Else
        '�v���O���X�o�[����������
        ' EG30 V32.2.0.1 DEL START
        'Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ, PRG_JPR_OUT)
        ' EG30 V32.2.0.1 DEL END
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : GetPrintSettings
'//  �@�\����  : ��ʂ̃`�F�b�N��Ԃ��擾
'//  �@�\�T�v  : �w�肳�ꂽ�R�[�i�����擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Boolean   bEnable    true:�ύX�\  false:�ύX�s��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub GetPrintSettings()
    Dim i               As Integer
    Dim k               As Integer
    Dim iCornerCount    As Integer
    Dim iGoukiCount     As Integer
    Dim iJprCount       As Integer
    
    ' �W���[�i���ݒ�����N���A����
    udtJprPrintSetteingInfo = udtInitJprSetting
    
    
    ' �R�[�i�̃`�F�b�N���
    k = 0
    For i = 0 To chkCorner.Count - 1
        If chkCorner(i).Value = 1 Then
            iCornerCount = iCornerCount + 1
            udtJprPrintSetteingInfo.iCorner(k) = i + 1
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iCornerCount = iCornerCount
    
    '���@�̃`�F�b�N���
    k = 0
    For i = 0 To chkGouki.Count - 1
        If chkGouki(i).Value = 1 Then
            iGoukiCount = iGoukiCount + 1
            udtJprPrintSetteingInfo.iGouki(k) = i + 1
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iGoukiCount = iGoukiCount
    
    '�W���[�i����ʂ̃`�F�b�N���
    k = 0
    For i = 0 To chkJprKind.Count - 1
        If chkJprKind(i).Value = 1 Then
            iJprCount = iJprCount + 1
            udtJprPrintSetteingInfo.iJprKind(k) = i
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iJprCount = iJprCount
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JprOutputProc
'//  �@�\����  : �W���[�i���o�͏���
'//  �@�\�T�v  : �o�̓t�@�C���쐬�Əo�̓v���Z�X�ɗv���𑗐M
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  REVISED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_003_01�z
'//     REVISIONS :(32.1.0.1) 2016-06-10  REVISED BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub JprOutputProc()
    Dim bRet        As Boolean
    
    'EG30 V32.1.0.1 ADD START
    Dim i, j            As Integer  '�R�[�i�A���@�J�E���^
    Dim intComSts       As Integer  '�ʐM���
    Dim blnSkipFlg      As Boolean  '�ێ�ݒ�f�[�^�Ȃ�
    Dim intGateNo       As Integer  '���@�ԍ��i1�`32�j
    'EG30 V32.1.0.1 ADD END
    Select Case udtJprPrintSetteingInfo.iJprKind(iJprIdx)
        Case JPR_KIND.JPR_KIND_EKI_INFO          ' �w�s�x�f�[�^�m�F(�w���)
            bRet = JPREdit_EkiInfo
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_EKI_INFO, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(EKI_JPR_EKIINFO_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
            
        Case JPR_KIND.JPR_KIND_JIKAI_INFO            ' �w�s�x�f�[�^�m�F(����)
            bRet = JPREdit_JikaiInfo
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_JIKAI_INFO, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(EKI_JPR_GATE_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If

        Case JPR_KIND.JPR_KIND_SETTING_LST           ' �ݒ�l�ꗗ
            bRet = JprEdit_SetteiList
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_SETTING_LST, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(SETTI_TXTFLE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        
        Case JPR_KIND.JPR_KIND_TUKA_DATA             ' �ʉ߃f�[�^
            ' ���b�Z�[�W�𑗐M���ĕҏW���t�@�C���̍쐬���˗�����̂ŁA�����ł͕ҏW�����͌Ă΂Ȃ��B
            ' �ҏW�������ĂԂ̂�RES���[������M�����Ƃ��B�W���[�i���󎚗v���͕ҏW�������I����Ă���ĂԁB
            ' �w�肳�ꂽ�R�[�i�����ݒu�Ȃ�Ώ����͂��Ȃ�
            If pfSettingCheck(False) = True Then
                bRet = SendMessageGamenOutReq(Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI)
                If bRet = False Then
                    Call ResultDisp(JPR_KIND.JPR_KIND_TUKA_DATA, bRet)
                    Exit Sub
                End If
            Else
                Call ResultDisp(JPR_KIND.JPR_KIND_TUKA_DATA, False)
                Exit Sub
            End If
        
        Case JPR_KIND.JPR_KIND_RIYO_KINGAKU          ' ���p���z�f�[�^
            ' ���b�Z�[�W�𑗐M���ĕҏW���t�@�C���̍쐬���˗�����̂ŁA�����ł͕ҏW�����͌Ă΂Ȃ��B
            ' �ҏW�������ĂԂ̂�RES���[������M�����Ƃ��B�W���[�i���󎚗v���͕ҏW�������I����Ă���ĂԁB
            ' �w�肳�ꂽ�R�[�i�����ݒu�Ȃ�Ώ����͂��Ȃ�
            If pfSettingCheck(False) = True Then
                bRet = SendMessageGamenOutReq(Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI)
                If bRet = False Then
                    Call ResultDisp(JPR_KIND.JPR_KIND_RIYO_KINGAKU, bRet)
                    Exit Sub
                End If
            Else
                Call ResultDisp(JPR_KIND.JPR_KIND_RIYO_KINGAKU, False)
                Exit Sub
            End If
        
        Case JPR_KIND.JPR_KIND_KADO_VER              ' �ғ��o�[�W�����ꗗ
            bRet = JprEdit_KadoVersion
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_KADO_VER, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(KADOVER_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        
        Case JPR_KIND.JPR_KIND_SIMEKIRI              ' ���؃I�t���C���o��
            bRet = JprEdit_SimekiriOffline
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_SIMEKIRI, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(SIMEKIRI_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        
        Case JPR_KIND.JPR_KIND_EKIMU_ID              ' �w���@��ID
            ' ���b�Z�[�W�𑗐M���ĕҏW���t�@�C���̍쐬���˗�����̂ŁA�����ł͕ҏW�����͌Ă΂Ȃ��B
            ' �ҏW�������ĂԂ̂�RES���[������M�����Ƃ��B�W���[�i���󎚗v���͕ҏW�������I����Ă���ĂԁB
            bRet = SendMessageInfoReq
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_EKIMU_ID, bRet)
                Exit Sub
            End If
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        Case JPR_KIND.JPR_KIND_SUBGATE_INFO         ' �w�s�x�f�[�^�m�F(�G���R�[�h�R�[�i���@����`)
            bRet = JPREdit_SubGateInfo
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_SUBGATE_INFO, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(EKI_JPR_SUBGATE_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
        'EG30 V32.1.0.1 ADD START
        Case JPR_KIND.JPR_KIND_GATE_CFG             ' ���D�@�ێ�ݒ�f�[�^
            '�`�F�b�N����Ă�����D�@�̒ʐM��Ԃ��擾����
            For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    '���̃R�[�i�A���@�͐ݒu����Ă��邩�H
                    If pfCornerGokiToGateNo(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j), intGateNo) = True Then
                    
                        '�Ď��ՋN���L���`�F�b�N
                        If CheckAppStart(PROC_KANRI) <> 0 Then
                            gpfGetjikaiConectSts intComSts, intGateNo
                            If intComSts <> CONECTSTS_NORMAL Then
                                Exit For
                            End If
                        End If
                    End If
                Next j
                '1��ł��ʐM�ُ�̉��D�@������΁A�x����\������̂ŁA�R�[�i�P�ʂ̃��[�v�𔲂���
                If intComSts <> CONECTSTS_NORMAL Then
                    Exit For
                End If
            Next i
            
            '�X�e�[�^�X�\�����ɒʐM�ُ���D�@�����邱�Ƃ�\������
            If intComSts <> CONECTSTS_NORMAL Then
                LstStatus.AddItem "�I�������R�[�i�ɒʐM�ُ�̉��D�@������܂�"
                LstStatus.AddItem "�ʐM�ُ퍆�@�̉��D�@�ێ�ݒ�f�[�^�͍ŐV�Ŗ����\��������܂�"
                LstStatus.Selected(LstStatus.ListCount - 1) = True
            End If
            
            bRet = JprEdit_GateCfg(blnSkipFlg)
            '���D�@�ێ�ݒ�f�[�^����M�̉��D�@�����������߁A�W���[�i���󎚂ł��Ȃ��������Ƃ�\������B
            If blnSkipFlg = True Then
                LstStatus.AddItem "���D�@�ێ�ݒ�f�[�^���󎚂ł��Ȃ��������D�@������܂�"
                LstStatus.Selected(LstStatus.ListCount - 1) = True
            End If
            
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_GATE_CFG, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(GATE_CFG_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        'EG30 V32.1.0.1 ADD END
    End Select

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JPREdit_JikaiInfo
'//  �@�\����  : �u�󎚁v�t����������
'//  �@�\�T�v  : ���݉w�ݒ�t�@�C���i����)���e�L�X�g�\������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.1.0.1) 2014-05-01  CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_003_01�z
'//                 �E���D�@�ݒu�����̈󎚂͕ʃW���[�i���֓Ɨ�������
'//     REVISIONS :(32.1.0.1) 2016-06-16  CODED   BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_JikaiInfo() As Boolean

    Dim strFileName             As String                   '�t�@�C����
    Dim bRet                    As Boolean                  '�֐��߂�l
    Dim lErrCode                As Long                     '�G���[�R�[�h
    Dim strLineCount            As String                   '�s���J�E���^
    
    Dim sWriteDir               As String                   '�������ݐ�t�H���_��
    Dim intFileNumber           As Integer                  '�t�@�C���|�C���^
    
    Dim ReadFileSettei()        As JIKAIINFO_IMAGE_FILE     '�������C���[�W�t�@�C��
    Dim i                       As Integer
    Dim j                       As Integer
    Dim iCornerIdx              As Integer                  '�R�[�i�C���f�b�N�X(���Ԗڂ̃R�[�i)
    
    Dim fso                     As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim FsoTS                   As TextStream

    Dim strNowShobunrui         As String                   '���ݕҏW���̏����ރR�[�h
    Dim strNowKubun             As String                   '���ݕҏW���̋敪
    Dim strNowCorner            As String                   '���ݕҏW���̃R�[�i
    
    'EG30 V32.1.0.1 ADD START
    Dim strEkiSettiBefPath      As String           '���݉w�ݒ�f�[�^�i�ύX�O�ۑ��j
    Dim strGetValue             As String * 64      'DLL�ɂ���Đݒ肳��邽�߁A64�Œ蒷�ɂ��Ă���
    Dim strCompValue            As String           '�ݒ�l�i�ύX�O�ۑ��j
    Dim strChangeFlg            As String           '�ύX��
    Dim intValueLen             As Integer          '�擾�����ݒ�l�̒���
    Dim intGateNo               As Integer          '1�`32���@
    'EG30 V32.1.0.1 ADD END
    
    '�G���[���[�`����錾
    On Error GoTo OUTPUT_ERROR
    '��ʂŎw�肳�ꂽ�R�[�i�͐ݒu����Ă��邩�H
    If pfSettingCheck(True) = False Then
        '���ׂĖ��ݒu�̃R�[�i�A���@�Ȃ̂ŃG���[�Ƃ���
        GoTo OUTPUT_ERROR
    End If
    
    '�C���[�W�t�@�C���̏o�͐�
    sWriteDir = EKI_JPR_GATE_TXTFILE

    '�w�s�x�f�[�^�m�F�i�����j�C���[�W�t�@�C���쐬
    bRet = dllGetEkiIniDataJpr(1, EKI_TUDO_CHK_GATE_FILE_JPR, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '�w�s�x�f�[�^�m�F�i�����j�C���[�W�t�@�C���폜
        Kill EKI_TUDO_CHK_GATE_FILE_JPR
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
        JPREdit_JikaiInfo = False
        Exit Function
    End If
    
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'    'EG20 V30.1.0.1 ADD START
'    '�����⏕CSV�t�@�C���쐬
'    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
'    If bRet = False Then
'        '�����⏕CSV�t�@�C���폜
'        Kill EKI_TUDO_CHK_SUBGATE_FILE
'        '�ُ탍�O�o��
'        Call pfOutPutErrLog(lErrCode)
'        JPREdit_JikaiInfo = False
'        Exit Function
'    End If
'    'EG20 V30.1.0.1 ADD END
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
    
    
    ' �R�[�i���̐ݒ菈��
    Call gsGetCornerName

    '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '�ُ�I��
        JPREdit_JikaiInfo = False
        Exit Function
        
    End If

    '�w�s�x�f�[�^(����)�C���[�W�t�@�C���̌������擾
    '�t�@�C���ԍ��擾
    intFileNumber = FreeFile
    Open EKI_TUDO_CHK_GATE_FILE_JPR For Input As #intFileNumber
    
    'CSV�t�@�C���s���J�E���g�i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
        Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1�ǉ�
            Line Input #intFileNumber, strLineCount
            j = j + 1
        Loop
    
    'CSV�t�@�C���N���[�Y
    Close #intFileNumber

    '�t�@�C���ԍ��擾
    intFileNumber = FreeFile

    '�Đݒ�
    ReDim ReadFileSettei(j) As JIKAIINFO_IMAGE_FILE        '�t�@�C���Ǎ��p�G���A
        
    'CSV�t�@�C���I�[�v��
    Open EKI_TUDO_CHK_GATE_FILE_JPR For Input As #intFileNumber

    '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
        For i = 0 To j - 1
            Input #intFileNumber, ReadFileSettei(i).strBunrui_Dai, ReadFileSettei(i).strBunrui_Tyu, _
            ReadFileSettei(i).srtBunrui_Sho, ReadFileSettei(i).strCorner, ReadFileSettei(i).strKomoku, _
            ReadFileSettei(i).strKubun, ReadFileSettei(i).strData, ReadFileSettei(i).strSetShosai
        Next

    'CSV�t�@�C���N���[�Y
    Close #intFileNumber
    
    'EG30 V32.1.0.1 ADD START
    '���̃R�[�i�̕ύX�O�f�[�^�ۑ����ꂽ�f�[�^����������ɓW�J����(�R�[�i0�j
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", "0")
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    '///////////////////////////////////////
    '// �W���[�i���o�̓C���[�W�t�@�C�����쐬
    '///////////////////////////////////////
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '�W���[�i���o�̓C���[�W�t�@�C�����쐬
    Open sWriteDir For Output As #intFileNumber
    
    '�^�C�g���\��
    'PrintHeader intFileNumber, "�w�s�x�f�[�^�m�F"  'EG30 V32.1.0.1 DEL
    PrintHeader3 intFileNumber, "�w�s�x�f�[�^�m�F", pfGetSaveDate(0)
    Print #intFileNumber, "�ݒu�w�F" & Trim(pfGetEkiNameInfo(NotEkiVer))
    
    strNowShobunrui = ""
    strNowKubun = ""
    strNowCorner = ""
    
    For i = 0 To UBound(ReadFileSettei) - 1
        '�R�[�i���؂�ւ�������H
        If (ReadFileSettei(i).strCorner <> strNowCorner) Then
            'iCornerIdx = iCornerIdx + 1
            'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'            'EG20 V30.1.0.1 ADD START
'            '�����R�[�i�o�͂�2�R�[�i�ڈȍ~������ꍇ�A2�R�[�i�ڂɓ���O�Ɏ����⏕���o��
'            If strNowCorner <> "" Then
'                pfOutPutSubGate CInt(strNowCorner), intFileNumber
'            End If
'            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
            '�ݒu�R�[�i���o��
            If IsTaisyoCorner(CInt(ReadFileSettei(i).strCorner)) = True Then
                
                '�ΏۃR�[�i�ł����Ă��Ώۍ��@���Ȃ���������Ȃ�
                For j = 0 To 15
                    If IsTaisyoGoki(CInt(ReadFileSettei(i).strCorner), j + 1) = True Then
                        If i <> 0 Then
                            Print #intFileNumber, ""
                        End If
                        Print #intFileNumber, "�ݒu�R�[�i�F" & gstrCornerName(CInt(ReadFileSettei(i).strCorner) - 1)
                        Exit For
                    End If
                Next j
            End If
            strNowCorner = ReadFileSettei(i).strCorner
        End If
    
        '���̍��@�͏o�͑Ώۂ��H
        If IsTaisyoGoki(CInt(ReadFileSettei(i).strCorner), CInt(ReadFileSettei(i).strBunrui_Tyu)) = True Then
            '�����ނƋ敪����v���Ȃ���΃^�C�g�����o�͂���
            If (ReadFileSettei(i).srtBunrui_Sho <> strNowShobunrui) Or (ReadFileSettei(i).strKubun <> strNowKubun) Then
                '�^�C�g�����o��
                Print #intFileNumber, ""
                'Print #intFileNumber, "�y" & ReadFileSettei(i).strKomoku & "�z" & ReadFileSettei(i).strKubun   'EG30 V32.1.0.1 DEL
                Print #intFileNumber, "�@�y" & ReadFileSettei(i).strKomoku & "�z" & ReadFileSettei(i).strKubun  'EG30 V32.1.0.1 ADD
                strNowShobunrui = ReadFileSettei(i).srtBunrui_Sho
                strNowKubun = ReadFileSettei(i).strKubun
            End If
            '�e���@�̐ݒ���o��
            
            'EG30 V32.1.0.1 ADD START
            '�W���[�i���ҏW�f�[�^�t�@�C���̒����ނ̓R�[�i���@�ԍ����Z�b�g����A����ɃR�[�i�ԍ����Z�b�g����Ă��邪�A
            '��r����ƂȂ�EKI_SETTI.CSV�͒����ނ͂P�`�R�Q�ŃR�[�i�ԍ��͂O�ƂȂ��Ă��邽�߁A�R�[�i���@�ԍ����P�`�R�Q�ɕϊ�����
            If pfCornerGokiToGateNo(CInt(ReadFileSettei(i).strCorner), CInt(ReadFileSettei(i).strBunrui_Tyu), intGateNo) = True Then
            
                '�ύX�O�f�[�^�ۑ����ꂽ�ݒ�l�Ɣ�r����(�啪�ނ����D�@�̏ꍇ�́A�R�[�i��0�Œ�Ō����j
                bRet = dllGetEkiInfoValue(CInt(ReadFileSettei(i).strBunrui_Dai), _
                                            intGateNo, _
                                            CInt(ReadFileSettei(i).srtBunrui_Sho), _
                                            0, _
                                            strGetValue, _
                                            intValueLen)
                strCompValue = strGetValue
                If (intValueLen <> 0) Then
                    strCompValue = MidByte(strGetValue, 1, intValueLen)
                    strCompValue = Trim(strCompValue)
                ElseIf (intValueLen = 0) Then
                    strCompValue = ""
                End If
                
                If (bRet = False) Or (ReadFileSettei(i).strData <> strCompValue) Then
                    strChangeFlg = DIFF_MARK_STRING_ON
                Else
                    strChangeFlg = DIFF_MARK_STRING_OFF
                End If
            '��r����̍��@�����Ȃ�������u���v
            Else
                strChangeFlg = DIFF_MARK_STRING_ON
            End If
            'EG30 V32.1.0.1 ADD END
            
            '���@�ԍ���\�����鍀�ڂ͍��@�ԍ���99�`���ɕϊ�����
            If (ReadFileSettei(i).strBunrui_Dai = 4 And ReadFileSettei(i).srtBunrui_Sho = 5) Or _
               (ReadFileSettei(i).strBunrui_Dai = 4 And ReadFileSettei(i).srtBunrui_Sho = 7) Then
                ReadFileSettei(i).strData = Format(ReadFileSettei(i).strData, "0#")
            End If
                
            'Print #intFileNumber, ReadFileSettei(i).strBunrui_Tyu & "���@ " & ReadFileSettei(i).strData    'EG30 V32.1.0.1 DEL
            Print #intFileNumber, strChangeFlg & ReadFileSettei(i).strBunrui_Tyu & "���@ " & ReadFileSettei(i).strData    'EG30 V32.1.0.1 ADD
            'bJprFlg = True
        End If
    Next i
    Print #intFileNumber, ""
    
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'    'EG20 V30.1.0.1 ADD START
'    '�����⏕���o��
'    pfOutPutSubGate CInt(strNowCorner), intFileNumber
'    Print #intFileNumber, ""
'    'EG20 V30.1.0.1 ADD END
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
    
    Print #intFileNumber, FOOTER_STRING
    '�t�@�C�����N���[�Y����B
    Close #intFileNumber
    Set fso = Nothing
    JPREdit_JikaiInfo = True
    Exit Function

OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    JPREdit_JikaiInfo = False
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : JPREdit_SubGateInfo
'//  �@�\����  : �u�󎚁v�t����������
'//  �@�\�T�v  : ���݉w�ݒ�t�@�C���i�G���R�[�h�R�[�i���@����`)���e�L�X�g�\������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(30.3.0.1) 2014-10-01  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_SubGateInfo() As Boolean

    Dim strFileName             As String                   '�t�@�C����
    Dim bRet                    As Boolean                  '�֐��߂�l
    Dim lErrCode                As Long                     '�G���[�R�[�h
    Dim strLineCount            As String                   '�s���J�E���^
    
    Dim sWriteDir               As String                   '�������ݐ�t�H���_��
    Dim intFileNumber           As Integer                  '�t�@�C���|�C���^
    
    Dim ReadFileSettei()        As JIKAIINFO_IMAGE_FILE     '�������C���[�W�t�@�C��
    Dim i                       As Integer
    Dim j                       As Integer
    Dim iCornerIdx              As Integer                  '�R�[�i�C���f�b�N�X(���Ԗڂ̃R�[�i)
    
    Dim fso                     As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim FsoTS                   As TextStream

    Dim strNowShobunrui         As String                   '���ݕҏW���̏����ރR�[�h
    Dim strNowKubun             As String                   '���ݕҏW���̋敪
    Dim strNowCorner            As String                   '���ݕҏW���̃R�[�i
    
    '�G���[���[�`����錾
    On Error GoTo OUTPUT_ERROR
    
    '�C���[�W�t�@�C���̏o�͐�
    sWriteDir = EKI_JPR_SUBGATE_TXTFILE

    '�����⏕CSV�t�@�C���쐬
    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '�����⏕CSV�t�@�C���폜
        Kill EKI_TUDO_CHK_SUBGATE_FILE
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
        JPREdit_SubGateInfo = False
        Exit Function
    End If
    
    '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '�ُ�I��
        JPREdit_SubGateInfo = False
        Exit Function
        
    End If

    '///////////////////////////////////////
    '// �W���[�i���o�̓C���[�W�t�@�C�����쐬
    '///////////////////////////////////////
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '�W���[�i���o�̓C���[�W�t�@�C�����쐬
    Open sWriteDir For Output As #intFileNumber
    
    '�^�C�g���\��
    PrintHeader2 intFileNumber, "�w�s�x�f�[�^�m�F", "(�G���R�[�h�R�[�i���@����`)"
    Print #intFileNumber, "�ݒu�w�F" & Trim(pfGetEkiNameInfo(NotEkiVer))
    
    strNowShobunrui = ""
    strNowKubun = ""
    strNowCorner = ""
    
    If pfOutPutSubGate(0, intFileNumber) = False Then
        GoTo OUTPUT_ERROR
    End If
    Print #intFileNumber, ""
    
    Print #intFileNumber, FOOTER_STRING
    '�t�@�C�����N���[�Y����B
    Close #intFileNumber
    Set fso = Nothing
    JPREdit_SubGateInfo = True
    Exit Function

OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    JPREdit_SubGateInfo = False
End Function
'EG30 V32.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  �֐�����  : JprEdit_GateCfg
'//  �@�\����  : ���D�@�ێ�ݒ�f�[�^�W���[�i���C���[�W�t�@�C���쐬
'//  �@�\�T�v  : ���D�@�ێ�ݒ�f�[�^�W���[�i���C���[�W�t�@�C�����쐬����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Boolean   bSkipFlg  ���D�@�ێ�f�[�^���������߁A�W���[�i���ҏW���X�L�b�v�������@������B
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-10   CODED   BY [TCC] T.Nakajima
'//             2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_GateCfg(ByRef bSkipFlg As Boolean) As Boolean

    Dim strOutputFile As String         '�o�̓t�@�C��
    Dim lngRet As Long                  '�֐��Ԃ�l
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim iOutFile    As Integer          '�t�@�C���ԍ�
    Dim ReadFileGateCfg()    As GATE_CFG_DATA_FILE  '���D�@�ێ�ݒ�f�[�^
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim l           As Integer
    Dim strJpCfgPath            As String  '���@�ʐݒ�R���t�B�O�t�@�C��(JP�p)
    Dim strSetteiBefFolder      As String  '�����Ď��ՕύX�O�ۑ��̈�
    Dim strJpCfgPathBef         As String  '�ύX�O�ۑ��������@�ʐݒ�R���t�B�O�t�@�C����
    Dim strDispImageFileName    As String  '�ҏW�f�[�^�t�@�C����
    Dim objFs                   As New FileSystemObject
    Dim intFileNo               As Integer
    
    Dim blnRet                  As Boolean  '�ҏW�f�[�^�쐬�֐��߂�l
    
    Dim strMutexName    As String       '�~���[�e�b�N�X��
    
    Dim strNowInfoName  As String       '���ݏo�͒��̏�񕔖�
    Dim strNowDai       As String       '���ݏo�͒��̑區��
    Dim strNowChu       As String       '���ݏo�͒��̒�����
    
    Dim iKoumokuByte          As Integer '���ږ��̃o�C�g��
    Dim iValueByte          As Integer '�ݒ�l�̃o�C�g��
    Dim iSpaceByte          As Integer '���Ԃɑ}������X�y�[�X�̃o�C�g��
    Dim strChangeFlg        As String  '�ύX�t���O
    Dim strSyoName          As String  '�����ږ�
    Dim strValue            As String  '�ݒ�l
    Dim blnInfoNameFlg      As Boolean '���s�t���O�i��񕔖�����̑區�ځ|�����ږ��̒��O�͉��s�Ȃ�)
    Dim intOutCount         As Integer  '�o�͉\���@��
    Dim intOutCountbyCorner(0 To 5) As Integer '�o�͉\���@���i�R�[�i���j
    Dim intGateNo           As Integer  '1�`32���@
    Dim bResult(0 To 5, 0 To 15) As Boolean  '�W���[�i���ҏW�f�[�^�t�@�C���o�͌���
    Dim strExistsCheckFileName As String     '0101.CSV�`0616.CSV�܂ł̃t�@�C���̑��݂��`�F�b�N
    Dim bDelFlg                 As Boolean  '�폜�t���O
    Const COLON_LEN = 2                 '�u�F�v�̃o�C�g��
    
    On Error GoTo Err_handler
    bDelFlg = False
    intOutCount = 0
    
    '�ݒu�w
    gsGetStationName
    '�������
    gsGetGateInfo
    '�R�[�i��
    gsGetCornerName
    '�R�[�i�^�C�v
    gsGetCornerType
    
    '��ʂŎw�肳�ꂽ�R�[�i�͐ݒu����Ă��邩�H
    If pfSettingCheck(True) = False Then
        '���ׂĖ��ݒu�̃R�[�i�A���@�Ȃ̂ŃG���[�Ƃ���
        GoTo Err_handler
    End If
    
    '�o�̓t�@�C�����ҏW
    strOutputFile = GATE_CFG_TXTFILE
    
    '�W���[�i���ҏW�f�[�^�t�@�C�������ׂč폜����
    '�폜�t�@�C�������݂��Ȃ��ꍇ��Err_Handler�ɂ����Ă��܂����߁A���݃`�F�b�N���s���B
    '�t�@�C������ł�������΁A���C���h�J�[�h�ɂ��t�@�C���폜���ł���̂ŁA���[�v�𔲂���
    strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", "*") & ".csv"
    For i = 1 To 6
        For j = 1 To 16
            strExistsCheckFileName = Replace(EDIT_DATA_GATECFG, "####", Format(i, "0#") & Format(j, "0#")) & ".csv"
            If objFs.FileExists(strExistsCheckFileName) Then
                objFs.DeleteFile strDispImageFileName
                bDelFlg = True
                Exit For
            End If
        Next j
        'JP_CFG*.CSV�ō폜�ς�
        If bDelFlg = True Then
            Exit For
        End If
    Next i
    
    bSkipFlg = False
    
    '�t�@�C���o�͊֐���Call
    '�`�F�b�N����Ă���R�[�i�A���@���ɂ��ď���
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        intOutCountbyCorner(i) = 0
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            '���̃R�[�i�A���@�͐ݒu����Ă��邩�H
            If pfCornerGokiToGateNo(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j), intGateNo) = True Then
                
                strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                '�~���[�e�b�N�X�����쐬
                strMutexName = Replace(MU_N_CFG, "##", Format(intGateNo, "0#"))

                strJpCfgPath = PATH_DATA & Replace(JP_CFG, "##", Format(intGateNo, "0#"))
                strSetteiBefFolder = PATH_OPERATE & "CORNER" & udtJprPrintSetteingInfo.iCorner(i) & "\\SETTEI_BEF\\"
                strJpCfgPathBef = strSetteiBefFolder & Replace(JP_CFG, "##", Format(intGateNo, "0#"))
                
                '���f�[�^(JP_CFGnn.GAT)�����݂����ꍇ�́A�W���[�i���f�[�^�t�@�C�����쐬
                If objFs.FileExists(strJpCfgPath) = True Then
                    bResult(i, j) = dllCreateGateCfgData(gintCornerType(udtJprPrintSetteingInfo.iCorner(i) - 1), _
                                                strDispImageFileName, strJpCfgPath, strJpCfgPathBef, strMutexName, lngErrCode)
                    If bResult(i, j) <> False Then
                        '�W���[�i���ҏW�f�[�^�t�@�C������ȏ���Ă���΁A���̍��@�Ŏ��s���Ă�����\�̂��߁B
                        intOutCount = intOutCount + 1
                        intOutCountbyCorner(i) = intOutCountbyCorner(i) + 1
                    Else
                        '�e�L�X�g�쐬�������s�ɂ��X�L�b�v
                        bSkipFlg = True
                    End If
                Else
                    ' �u���D�@�ێ�ݒ�f�[�^���󎚂ł��Ȃ��������D�@������܂��v��\�����邽�߂�ON
                    bSkipFlg = True
                End If
                
            End If
        Next j
    Next i
    
    '�W���[�i���ҏW�f�[�^�t�@�C�������Ă���΁A�W���[�i���o�͉\
    If intOutCount > 0 Then
        '���D�@�ێ�ݒ�f�[�^ �W���[�i���C���[�W�t�@�C�����쐬
        iOutFile = FreeFile
        Open strOutputFile For Output As #iOutFile
        
        '�w�b�_�[��
        PrintHeader iOutFile, "���D�@�ێ�ݒ�f�[�^�m�F"
        
        '�ݒu�w
        Print #iOutFile, "�ݒu�w�F" & gstrStationName(0)
        
        '�`�F�b�N���ꂽ�R�[�i�������[�v
        For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
            Erase ReadFileGateCfg
            
            '���̃R�[�i�̉��D�@�����ׂĉ��D�@�ێ�ݒ�f�[�^�������Ă��Ȃ��ꍇ�͈󎚂��Ȃ�
            If intOutCountbyCorner(i) > 0 Then
                '�R�[�i��
                Print #iOutFile, "�ݒu�R�[�i�F" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                '�ۑ�����
                Print #iOutFile, "�ۑ������F" & pfGetSaveDate(udtJprPrintSetteingInfo.iCorner(i))
    
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    ' �W���[�i���ҏW�f�[�^�t�@�C���쐬����������̏ꍇ��
                    If bResult(i, j) <> False Then
                        '���̍��@���ݒu����Ă��邩�H
                        intFileNo = FreeFile
                        strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", _
                            Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                    
                        '�W���[�i���ҏW�f�[�^�t�@�C�����I�[�v���i���̃t�@�C�������݂��Ȃ��ꍇ�͂����ɂ͗��Ȃ��j
                        Open strDispImageFileName For Input As #intFileNo
                
                        '��ʕ\���p�f�[�^(csv)���G���A�ɓǂݍ���
                        k = 0
                        Do While Not EOF(intFileNo)
                            ReDim Preserve ReadFileGateCfg(k)
                            Input #intFileNo, _
                                    ReadFileGateCfg(k).strInfoName, _
                                    ReadFileGateCfg(k).strBunrui_Dai, ReadFileGateCfg(k).strBunrui_Chu, _
                                    ReadFileGateCfg(k).strBunrui_Syo, ReadFileGateCfg(k).strValue, ReadFileGateCfg(k).strChangeFlg
                            k = k + 1
                        Loop
                        '�t�@�C���N���[�Y
                        Close #intFileNo
                        
                        '���@�ԍ�
                        Print #iOutFile, "���@�ԍ��F" & Format(udtJprPrintSetteingInfo.iGouki(j), "00") & "���@"
                        
                        '��������1�䕪�̉��D�@�ێ�ݒ�f�[�^�̓��e(�{��)���󎚂��郋�[�v
                        strNowInfoName = ""
                        strNowDai = ""
                        strNowChu = ""
                        blnInfoNameFlg = False
                        For l = 0 To UBound(ReadFileGateCfg)
                            '��񕔖����قȂ�΁A��؂�^�C�g�����o�͂��� ���񕔖���
                            If strNowInfoName <> ReadFileGateCfg(l).strInfoName Then
                                Print #iOutFile, ""
                                Print #iOutFile, "�@" & ReadFileGateCfg(l).strInfoName
                                strNowInfoName = ReadFileGateCfg(l).strInfoName
                                blnInfoNameFlg = True
                            End If
                            '�啪�ށA�����ނ��قȂ�ꍇ�́A��؂�^�C�g�����o�͂��� �y�啪��-�����ށz
                            If strNowDai <> ReadFileGateCfg(l).strBunrui_Dai Or strNowChu <> ReadFileGateCfg(l).strBunrui_Chu Then
                                '��؂�^�C�g��(�啪�ށ|������)�̒��O��1�s���s�����邪�A��񕔖��̒���͉��s�Ȃ�
                                If blnInfoNameFlg = False Then
                                    Print #iOutFile, ""
                                Else
                                    blnInfoNameFlg = False
                                End If
                                Print #iOutFile, "�@�y" & ReadFileGateCfg(l).strBunrui_Dai & "�|" & ReadFileGateCfg(l).strBunrui_Chu & "�z"
                                strNowDai = ReadFileGateCfg(l).strBunrui_Dai
                                strNowChu = ReadFileGateCfg(l).strBunrui_Chu
                            End If
                            
                            '�ύX�t���O �{ ���ږ� �{ ":" �{ �ݒ�l���o��
                            '�������X�y�[�X�𒆊Ԃɓ���邩�H
                            strSyoName = RTrim(ReadFileGateCfg(l).strBunrui_Syo)
                            strValue = RTrim(ReadFileGateCfg(l).strValue)
                            iKoumokuByte = LenB(StrConv(strSyoName, vbFromUnicode))
                            iValueByte = LenB(StrConv(strValue, vbFromUnicode))
                            '�W���[�i��1�s���ő�30�o�C�g
                            iSpaceByte = MAX_JPR_KETA_MAX - DIFF_MARK_LEN - iKoumokuByte - COLON_LEN - iValueByte
                            If iSpaceByte <= 0 Then
                                iSpaceByte = 0
                            End If
                            If ReadFileGateCfg(l).strChangeFlg = "" Then
                                strChangeFlg = DIFF_MARK_STRING_OFF
                            Else
                                strChangeFlg = DIFF_MARK_STRING_ON
                            End If
                            
                            Print #iOutFile, strChangeFlg & strSyoName & Space(iSpaceByte) & "�F" & strValue
                        
                        Next l
                        
                        Print #iOutFile, ""
                    End If
                Next j
            End If
        Next i
        
        Print #iOutFile, FOOTER_STRING
        Close #iOutFile
      
        JprEdit_GateCfg = True
    Else
        '�o�͑ΏۂƂȂ���D�@��1������݂��Ȃ��̂ŁA�X�L�b�v�͖����Ƃ���B
        bSkipFlg = False
        JprEdit_GateCfg = False
    End If
    Exit Function

Err_handler:

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    If iOutFile > 0 Then
        Close #iOutFile
    End If
    
    Set objFs = Nothing

    'MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
    '�u�W���[�i���󎚉�ʁi���D�@�ێ�ݒ�f�[�^�j�F�W���[�i���C���[�W�t�@�C���쐬�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_GateCfg = False

End Function
'EG30 V32.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : IsTaisyoGoki
'//  �@�\����  : �w�荆�@�m�F����
'//  �@�\�T�v  : �C���[�W�t�@�C���ɏo�͂��鍀�ڂ͏o�͑Ώۂ��m�F����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   iCorner   �R�[�i�ԍ�
'//              Integer   iGouki    ���@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function IsTaisyoGoki(iCorner As Integer, iGouki As Integer) As Boolean
    Dim bRet        As Boolean
    Dim i           As Integer
    Dim j           As Integer
    
    bRet = False
    
    If pfCornerGokiCheck(iCorner, iGouki) = False Then
        '���ݒu�̍��@�Ȃ̂ŁA��ʏ�`�F�b�N����Ă��Ă�false
        IsTaisyoGoki = False
        Exit Function
    End If
    
    
    For j = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        For i = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            If udtJprPrintSetteingInfo.iCorner(j) = iCorner Then
                If udtJprPrintSetteingInfo.iGouki(i) = iGouki Then
                    bRet = True
                    Exit For
                End If
            End If
        Next i
    Next j
    
    IsTaisyoGoki = bRet
   
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : IsTaisyoCorner
'//  �@�\����  : �w��R�[�i�m�F����
'//  �@�\�T�v  : �C���[�W�t�@�C���ɏo�͂��鍀�ڂ͏o�͑Ώۂ��m�F����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   iCorner   �R�[�i�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function IsTaisyoCorner(iCorner As Integer) As Boolean
    Dim bRet        As Boolean
    Dim i           As Integer
    
    bRet = False
    '���̃R�[�i�͐ݒu����Ă��邩�H
    If pfCornerGokiCheck(iCorner) = True Then
        
        For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
            If udtJprPrintSetteingInfo.iCorner(i) = iCorner Then
                bRet = True
                Exit For
            End If
        Next i
    End If
    
    IsTaisyoCorner = bRet
   
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JprEdit_SetteiList
'//  �@�\����  : �ݒ�l�ꗗ�o��
'//  �@�\�T�v  : �ݒ�l�ꗗ�W���[�i���̃C���[�W�t�@�C����ҏW����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   iCorner   �R�[�i�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(7.4.0.1) 2013-07-22   REVISED BY [TCC] T.Nakajima
'//                ���܂�����o��t���[�ݒ��ʑΉ�
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-17   REVISED BY [TCC] T.Nakajima
'//                2016�N�x�{���Ή�
'//     REVISIONS :(EG30 V35.3.0.1) 2019-07-03   REVISED BY [TCC] H.Kondoh
'//                2019�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_SetteiList() As Boolean
    Dim strFilePath As String           '�o�̓t�@�C���p�X
    Dim intCount As Integer             '�J�E���^
    Dim intOutFile As Integer           '�o�̓t�@�C���ԍ�
    Dim intTgtFileNo As Integer         '�o�͑Ώېݒ�t�@�C���ԍ�
    Dim strTgtFileName As String        '�o�͑Ώېݒ�t�@�C��
    Dim strTargetFile() As String       '�o�͑Ώۃt�@�C��
    Dim intFileNum As Integer
    Dim objFileObj As FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim ReadFileSettei()    As SETTEI_OUTPUT_IMAGE_FILE   '�t�@�C���Ǎ��p�\����
    Dim strCsvBuffer        As String
    Dim strCammaArray()     As String
    Dim i As Integer
    Dim strNowDaikomoku     As String
    Dim strNowKomoku        As String
    Dim FsoTS   As TextStream
    Dim iKomkuByte          As Integer '���ږ��̃o�C�g��
    Dim iValueByte          As Integer '�ݒ�l�̃o�C�g��
    Dim iSpaceByte          As Integer '���Ԃɑ}������X�y�[�X�̃o�C�g��
    Dim intJprFile            As Integer
    Dim strNyujoFree(3)       As String
    Dim iSeparatePos          As Integer    '��ʖ��̂��u:�v�ŋ�؂��Ă����ꍇ�̋�؂�ʒu
    'EG30 V32.1.0.1 ADD START
    Dim strChangeFlg        As String  '�ύX�t���O
    'EG30 V32.1.0.1 ADD END

    Set objFileObj = New FileSystemObject
    
    On Error GoTo Err_handler
    
    'EG20 V30.1.0.1 ADD START
    '�ݒu�w
    gsGetStationName
    '�������
    gsGetGateInfo
    '�R�[�i��
    gsGetCornerName
    '�R�[�i�^�C�v
    gsGetCornerType
    'EG20 V30.1.0.1 ADD END
    '��ʂŎw�肳�ꂽ�R�[�i�͐ݒu����Ă��邩�H
    If pfSettingCheck(False) = False Then
        '���ׂĖ��ݒu�Ȃ̂ŃG���[
        GoTo Err_handler
    End If
    
    '�o�͑Ώېݒ�t�@�C�����I�[�v������B
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE
    
    '�o�͑Ώېݒ�t�@�C�������݂��Ȃ��ꍇ�ُ͈�I��
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '�o�͑Ώۃt�@�C�������擾
    Input #intTgtFileNo, intFileNum
    
    '�o�͑Ώۃt�@�C�����擾
    ReDim strTargetFile(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFile)
        Input #intTgtFileNo, strTargetFile(intCount)
    Next
    
    Close #intTgtFileNo
    
    'EG20 V30.1.0.1 ADD START
    '�����R�[�i�[�ɑ΂���o�͑Ώۃt�@�C���̓��e���m�ۂ���
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE_KAN
    
    '�o�͑Ώېݒ�t�@�C�������݂��Ȃ��ꍇ�ُ͈�I��
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '�o�͑Ώۃt�@�C�������擾
    Input #intTgtFileNo, intFileNum
    
    '�o�͑Ώۃt�@�C�����擾
    ReDim strTargetFileKan(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFileKan)
        Input #intTgtFileNo, strTargetFileKan(intCount)
    Next
    
    Close #intTgtFileNo
    'EG20 V30.1.0.1 ADD END
    
    '////////////////////////////////
    '�W���[�i���C���[�W�t�@�C���쐬
    '�O��̏o�͍ς݂̃W���[�i���C���[�W�t�@�C���͏����Ă���(�R�[�i�P�ʂŒǋL���Ă�������)
    If Dir(SETTI_TXTFLE) <> "" Then
        Kill SETTI_TXTFLE
    End If
    
    '�R�[�i�P�ʂŐݒ�l�ꗗ��CSV�t�@�C�����쐬����
    '�w�b�_���쐬
    intJprFile = FreeFile
    Open SETTI_TXTFLE For Output As #intJprFile
    PrintHeader intJprFile, "�ݒ�l�ꗗ"
    
    '�ݒu�w
    'gsGetStationName   'EG20 V30.1.0.1 DEL
    Print #intJprFile, "�ݒu�w�F" & gstrStationName(0)
    '�R�[�i��
    'gsGetCornerName    'EG20 V30.1.0.1 DEL

    For intCount = 0 To UBound(glngTergetCorner)
        
        If glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON Then
            '�R�[�i�P�ʂŐݒ�t�@�C���ꗗ(�ҏW�pCSV)�쐬 OPERATE_SETTI99.csv
            strFilePath = Replace(EDIT_DATA_SETTEI, "##", Format(intCount + 1, "0#"))
            
            '---- �ݒ�ꗗ�e�L�X�g�쐬 �J�n
            '�t�@�C���쐬
            If objFileObj.FileExists(strFilePath) = True Then
                objFileObj.DeleteFile (strFilePath)
            End If
            Call objFileObj.CreateTextFile(strFilePath)
            
            '�o�̓t�@�C�����I�[�v������B
            intOutFile = FreeFile
            Open strFilePath For Output As #intOutFile
    
            'ID�ݒ�l���o��
            'If gsubOutput_Id(intCount + 1, intOutFile, True) = False Then      'EG30 V32.1.0.1 DEL
            If gsubOutput_Id_JPR(intCount + 1, intOutFile, True) = False Then   'EG30 V32.1.0.1 ADD
                GoTo Err_handler
            End If
            
            'EG20 V30.1.0.1 DEL START
            '���o��t���[�t�@�C�����o��
'            If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
'
'            '�j�Փ��t�@�C�����o��
'            If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
            'EG20 V30.1.0.1 DEL END
            
            'EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                '�����R�[�i�[�̏ꍇ
                '�V�����s���p�����[�^���o��
                'If gsubOutput_ParaKan(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                '�ݗ���IC����p�����[�^���o��
                'If gsubOutput_ParaKan(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                
                '�ݗ���IC�ʉߏ����p�����[�^���o��
                'If gsubOutput_ParaKan(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
            Else
                '�ݗ��R�[�i�[�̏ꍇ
                '���o��t���[�t�@�C�����o��
                'If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 DEL
                If gsubOutput_Free_InOut_JPR(intCount + 1, intOutFile) = False Then     'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                
                '�j�Փ��t�@�C�����o��
                'If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then  'EG30 V32.1.0.1 V32.1.0.1 DEL
                If gsubOutput_Shukusai_JPR(intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
            End If
            'EG20 V30.1.0.1 ADD END

            Close #intOutFile
            '---- �ݒ�ꗗ�e�L�X�g�쐬 �I��
            
            '�o�͂����ҏW���f�[�^���G���A�ɃZ�b�g����
            Set FsoTS = objFileObj.OpenTextFile(strFilePath, ForReading)
            i = 0
            Do Until FsoTS.AtEndOfStream = True
                ReDim Preserve ReadFileSettei(i)
                strCsvBuffer = FsoTS.ReadLine
                '�J���}���L�[���[�h�Ɋe���ڂ�؂�o���B
                strCammaArray = Split(strCsvBuffer, ",")
                ReadFileSettei(i).strDaiKomoku = strCammaArray(0)   '�區��
                ReadFileSettei(i).strKomoku = strCammaArray(1)      '���ږ�
                ReadFileSettei(i).strValue = strCammaArray(2)       '�ݒ�l
                ReadFileSettei(i).strChangeFlg = strCammaArray(3)   '�ύX�t���O
                
                i = i + 1
            Loop
            FsoTS.Close
            
            '�ǂݍ��񂾃G���A����W���[�i���C���[�W�t�@�C�����쐬����
            
            '�R�[�i��
            Print #intJprFile, "�ݒu�R�[�i�F" & gstrCornerName(intCount)
            '�ۑ�����
            Print #intJprFile, "�ۑ������F" & pfGetSaveDate(intCount + 1)

            strNowDaikomoku = ""
            strNowKomoku = ""
            
            For i = 0 To UBound(ReadFileSettei)
                '�區�ڂ��o�͂��邩�H
                If strNowDaikomoku <> ReadFileSettei(i).strDaiKomoku Then
                    '�������ANULL�̏ꍇ�͏����@�ȍ~�Ȃ̂Ōp��
                    If ReadFileSettei(i).strDaiKomoku <> "" Then
                        'Print #intJprFile, ""  'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            '�����ڃ��x���̐ؑւ͉��s���Ȃ�
                            iSeparatePos = InStr(ReadFileSettei(i).strDaiKomoku, ":")
                            If iSeparatePos > 0 Then
                                If Left(strNowDaikomoku, iSeparatePos - 1) <> Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) Then
                                    '�啪�ރ��x���ňقȂ��Ă���̂ŉ��s
                                    Print #intJprFile, ""
                                Else
                                End If
                            Else
                                Print #intJprFile, ""
                            End If
                        Else
                            Print #intJprFile, ""
                        End If
                        
                        'Print #intJprFile, "�y" & ReadFileSettei(i).strDaiKomoku & "�z"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            '�����R�[�i�̏ꍇ
                            '�區��(��ʖ��̂�":"�ŋ�؂��Ă����炻���ŕ�����)
                            iSeparatePos = InStr(ReadFileSettei(i).strDaiKomoku, ":")
                            If iSeparatePos > 0 Then
                                '�啪�ނ�������������o�͂��Ȃ�
                                If Left(strNowDaikomoku, iSeparatePos - 1) <> Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) Then
                                    'Print #intJprFile, "�y" & Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) & "�z"    'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "�@�y" & Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) & "�z"   'EG30 V32.1.0.1 ADD
                                    '�����ڂ��o�͂���
                                    'Print #intJprFile, Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1)   'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "�@" & Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1) 'EG30 V32.1.0.1 ADD
                                Else
                                    '�啪�ނ܂ł͓����Ȃ̂ŁA�����ڂ������o�͂���B
                                    'Print #intJprFile, Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1)   'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "�@" & Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1) 'EG30 V32.1.0.1 ADD
                                End If
                            Else
                                'Print #intJprFile, "�y" & ReadFileSettei(i).strDaiKomoku & "�z"    'EG30 V32.1.0.1 DEL
                                Print #intJprFile, "�@�y" & ReadFileSettei(i).strDaiKomoku & "�z"   'EG30 V32.1.0.1 ADD
                            End If
                        Else
                            'Print #intJprFile, "�y" & ReadFileSettei(i).strDaiKomoku & "�z"    'EG30 V32.1.0.1 DEL
                            Print #intJprFile, "�@�y" & ReadFileSettei(i).strDaiKomoku & "�z"   'EG30 V32.1.0.1 ADD
                        End If
                        'EG20 V30.1.0.1 ADD END
                        strNowDaikomoku = ReadFileSettei(i).strDaiKomoku
                    End If
                End If
                
                '����t���[�ݒ��ʂ͐ݒ�l�����s������K�v������B
                If ReadFileSettei(i).strDaiKomoku = "����t���[�ݒ���" Then
                    '����t���[1�`6�̐����͑S�p�ɕύX����B(�d�l�ɂ��킹�邽��)
                    Select Case ReadFileSettei(i).strKomoku
                        Case "����t���[1"
                            'strNyujoFree(0) = "����t���[�P"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�P"    'EG30 V32.1.0.1 ADD
                        Case "����t���[2"
                            'strNyujoFree(0) = "����t���[�Q"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�Q"    'EG30 V32.1.0.1 ADD
                        Case "����t���[3"
                            'strNyujoFree(0) = "����t���[�R"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�R"    'EG30 V32.1.0.1 ADD
                        Case "����t���[4"
                            'strNyujoFree(0) = "����t���[�S"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�S"    'EG30 V32.1.0.1 ADD
                        Case "����t���[5"
                            'strNyujoFree(0) = "����t���[�T"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�T"    'EG30 V32.1.0.1 ADD
                        Case "����t���[6"
                            'strNyujoFree(0) = "����t���[�U"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�U"    'EG30 V32.1.0.1 ADD
                        Case "����t���[7"
                            'strNyujoFree(0) = "����t���[�V"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�V"    'EG30 V32.1.0.1 ADD
                        Case "����t���[8"
                            'strNyujoFree(0) = "����t���[�W"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�W"    'EG30 V32.1.0.1 ADD
                        Case "����t���[9"
                            'strNyujoFree(0) = "����t���[�X"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@����t���[�X"    'EG30 V32.1.0.1 ADD
                    End Select
                    'EG30 V32.1.0.1 DEL START
'                    strNyujoFree(1) = MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 11, 16) '�J�n����
'                    strNyujoFree(2) = MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 38, 16) '�I������
'                    strNyujoFree(3) = MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(22) & MidByte(ReadFileSettei(i).strValue, 61, 4) '����
                    'EG30 V32.1.0.1 DEL END
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "�@" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    strNyujoFree(1) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 11, 16)  '�J�n����
                    strNyujoFree(2) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 38, 16) '�I������
                    strNyujoFree(3) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(20) & MidByte(ReadFileSettei(i).strValue, 61, 4) '����
                    'EG30 V32.1.0.1 ADD END
                    '1�s�o��
                    Print #intJprFile, strNyujoFree(0)
                    Print #intJprFile, strNyujoFree(1)
                    Print #intJprFile, strNyujoFree(2)
                    Print #intJprFile, strNyujoFree(3)
                'EG20 V7.4.0.1 ADD START
                '���܂�����o��t���[�ݒ��ʂ͐ݒ�����s������K�v������
                'ElseIf ReadFileSettei(i).strDaiKomoku = "���܂�����o��t���[�ݒ���" Then    'EG30 V32.1.0.1 DEL
                ElseIf ReadFileSettei(i).strDaiKomoku = "���ׂ�o��t���[�ݒ���" Then         'EG30 V32.1.0.1 ADD
                    '����t���[1�`6�̐����͑S�p�ɕύX����B(�d�l�ɂ��킹�邽��)
                    Select Case ReadFileSettei(i).strKomoku
                        'Case "���܂�����o��t���[1"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[1"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�P" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�P"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[2"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[2"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�Q" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�Q"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[3"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[3"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�R" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�R"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[4"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[4"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�S" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�S"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[5"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[5"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�T" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�T"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[6"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[6"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�U" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�U"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[7"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[7"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�V" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�V"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[8"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[8"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�W" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�W"  'EG30 V32.1.0.1 ADD
                        'Case "���܂�����o��t���[9"   'EG30 V32.1.0.1 DEL
                        Case "���ׂ�o��t���[9"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "���܂�����o��t���[�X" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "�@���ׂ�o��t���[�X"  'EG30 V32.1.0.1 ADD
                    End Select
                    'EG30 V32.1.0.1 DEL START
                    'strNyujoFree(1) = MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 11, 16) '�J�n����
                    'strNyujoFree(2) = MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 38, 16) '�I������
                    'strNyujoFree(3) = MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(22) & MidByte(ReadFileSettei(i).strValue, 61, 4) '����
                    'EG30 V32.1.0.1 DEL END
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "�@" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    strNyujoFree(1) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 11, 16)  '�J�n����
                    strNyujoFree(2) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 38, 16)  '�I������
                    strNyujoFree(3) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(20) & MidByte(ReadFileSettei(i).strValue, 61, 4) '����
                    'EG30 V32.1.0.1 ADD END
                    '1�s�o��
                    Print #intJprFile, strNyujoFree(0)
                    Print #intJprFile, strNyujoFree(1)
                    Print #intJprFile, strNyujoFree(2)
                    Print #intJprFile, strNyujoFree(3)
                'EG20 V7.4.0.1 ADD END
                Else
                
                    If ReadFileSettei(i).strKomoku = "" Then
                        '�O��̍��ږ����g��
                        ReadFileSettei(i).strKomoku = strNowKomoku
                    End If
                    strNowKomoku = ReadFileSettei(i).strKomoku
                    '��O����
                    '��{�̓e�L�X�g�o�͂�CSV�̕������g�����A���L�̍��ڂ����̓W���[�i���d�l�ɂ��킹��K�v������B
                    Select Case ReadFileSettei(i).strKomoku
                        Case "���ԑъJ�n����", _
                             "���ԑяI������", _
                             "�L�����ԑъJ�n����", _
                             "�L�����ԑяI������"
                            '�u�S���@�F99��99���v�� �u �S���@ 99��99���v�ɕϊ�����
                            ReadFileSettei(i).strValue = Space(1) & Replace(ReadFileSettei(i).strValue, "�F", " ")
                        
                        'Case "�ʉ߃T�[�r�X�s���ۗ�"    'EG30 V32.1.0.1 DEL
                        'EG30 V32.1.0.1 ADD START
'EG30 V35.3.0.1 DEL Start
'                        Case "�ʉ߃T�[�r�X�s���ۗ�", _
'                             "IC��ЊԌo�H�A����", _
'                             "�I�[�g�`���[�W�@�\", _
'                             "������t�F�[���Z�[�t", _
'                             "���ʌ��t�F�[���Z�[�t", _
'                             "IC�J�[�h�����O�\��", _
'                             "IC�J�[�h������ē�", _
'                             "���l���[�h�����ē�", _
'                             "IC�ē��\�����"
'                        'EG30 V32.1.0.1 ADD END
'EG30 V35.3.0.1 ADD End
'EG30 V35.3.0.1 ADD Start
                        Case "�ʉ߃T�[�r�X�s���ۗ�", _
                             "IC��ЊԌo�H�A����", _
                             "�I�[�g�`���[�W�@�\", _
                             "������t�F�[���Z�[�t", _
                             "���ʌ��t�F�[���Z�[�t", _
                             "IC�J�[�h�����O�\��", _
                             "IC�J�[�h������ē�", _
                             "���l���[�h�����ē�", _
                             "IC�ē��\�����", _
                             "�J�n�N����", _
                             "�I���N����"
'EG30 V35.3.0.1 ADD End
                            '�u�ʉ߃T�[�r�X�s���ۗ��S���@�Fxx�v�� �u�ʉ߃T�[�r�X�s���ۗ� �S���@�Fxx�v�ɕϊ�����
                            ReadFileSettei(i).strValue = Space(1) & ReadFileSettei(i).strValue
                        
                        Case "���l���[�h����ݒ�"
                            '�ݒ�l�����l�ɂ���
                            ReadFileSettei(i).strValue = ReadFileSettei(i).strValue & Space(8 - LenB(ReadFileSettei(i).strValue))
                    End Select
                
                    '���ږ��o��
                    '�������X�y�[�X�𒆊Ԃɓ���邩�H
                    iKomkuByte = LenB(StrConv(ReadFileSettei(i).strKomoku, vbFromUnicode))
                    iValueByte = LenB(StrConv(ReadFileSettei(i).strValue, vbFromUnicode))
                    '�W���[�i��1�s���ő�30�o�C�g
                    'iSpaceByte = MAX_JPR_KETA_MAX - iKomkuByte - iValueByte    'EG30 V32.1.0.1 DEL
                    iSpaceByte = MAX_JPR_KETA_MAX - DIFF_MARK_LEN - iKomkuByte - iValueByte  'EG30 V32.1.0.1 ADD
                    If iSpaceByte < 0 Then
                        'iSpaceByte = 0    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            iSpaceByte = 1
                        Else
                            iSpaceByte = 0
                        End If
                        'EG20 V30.1.0.1 ADD END
                    ElseIf iSpaceByte = 0 Then
                        'iSpaceByte = 0     'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            iSpaceByte = 1
                        Else
                            iSpaceByte = 0
                        End If
                        'EG20 V30.1.0.1 DEL END
                    End If
                    
                    Space (iSpaceByte)
                    '1�s�o��
                    'Print #intJprFile, ReadFileSettei(i).strKomoku & Space(iSpaceByte) & ReadFileSettei(i).strValue    'EG30 V32.1.0.1 DEL
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "�@" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    Print #intJprFile, strChangeFlg & ReadFileSettei(i).strKomoku & Space(iSpaceByte) & ReadFileSettei(i).strValue
                    'EG30 V32.1.0.1 ADD END
                End If
              
            Next i
            Print #intJprFile, ""
        End If
    Next intCount
    
    Print #intJprFile, FOOTER_STRING
    
    Close #intJprFile
    Set objFileObj = Nothing
    
    JprEdit_SetteiList = True
    Exit Function
    
'�G���[����
Err_handler:

    If intTgtFileNo > 0 Then
        Close #intTgtFileNo
    End If
    If intOutFile > 0 Then
        Close #intOutFile
    End If
    If intJprFile > 0 Then
        Close #intJprFile
    End If
    Set objFileObj = Nothing
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
    JprEdit_SetteiList = False
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JprEdit_SimekiriOffline
'//  �@�\����  : ���؃I�t���C���o�̓W���[�i���ҏW����
'//  �@�\�T�v  : ���؃I�t���C���o�͂̃C���[�W�t�@�C����ҏW����
'//
'//              �^        ����      �Ӗ�
'//  ����      :
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_SimekiriOffline() As Boolean
    
    Dim objFso As New FileSystemObject                  ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim objTs   As TextStream
    Dim bProceed As Boolean                             ' ���؏����J�n�t���O
    Dim nListCnt As Integer                             ' �t�@�C���i�[��
    Dim szSaveFolder As String                          ' �ۑ���t�H���_
    Dim szFileName As String                            ' �t�@�C����
    Dim iResponse As Integer
    Dim Index       As Integer                          '�C���f�b�N�X
    Dim iOutFile    As Integer
    
    On Error GoTo ErrorHandler                          ' �G���[�n���h���̓o�^
    
    'EG20 V30.1.0.1 ADD START
    ' �R�[�i���擾
    gsGetCornerName
    ' �R�[�i�^�C�v�擾
    gsGetCornerType
    
    ' �w���擾
    gsGetStationName
    ' EG20 V30.1.0.1 ADD END
    
    '�`�F�b�N���ꂽ�R�[�i�͐ݒu����Ă��邩�H�i�ǂꂩ�ЂƂł������OK)
    If pfSettingCheck(False) = False Then
        '���ׂĖ��ݒu�̃R�[�i�Ȃ̂ŃG���[
        GoTo ErrorHandler
    End If
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // ������
    Index = 0
    Erase gOfflineFileList

    ReDim Preserve gOfflineFileList(0)
    bProceed = False
    nListCnt = 0
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // �W���[�i���C���[�W�t�@�C���쐬
    'EG20 V30.1.0.1 DEL START
    ' �R�[�i���擾
    'gsGetCornerName
    
    ' �w���擾
    'gsGetStationName
    'EG20 V30.1.0.1 DELEND
    
    '�W���[�i���C���[�W�t�@�C�����I�[�v��
    iOutFile = FreeFile
    Open SIMEKIRI_TXTFILE For Output As #iOutFile
    
    '�w�b�_�����o��
    PrintHeader iOutFile, "���؃I�t���C���o��"
    
    '�ݒu�w/�ݒu�R�[�i
    Print #iOutFile, "�ݒu�w�F" & gstrStationName(0)
    
    For Index = 0 To UBound(glngTergetCorner)
    
        If glngTergetCorner(Index) = CMN_ONOFF.CMN_ON Then
            
            ' /////////////////////////////////////////////////////////////////////////
            ' // ���؏o�̓f�[�^�͑��݂��邩�H�iD:\KANSI\SHUKEI\SEND_DATA\SIMEKIRI##.DAT�j
            szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(Index + 1, "0#"))
            If objFso.FileExists(szFileName) = True Then              ' �t�@�C�����̎擾�`�F�b�N
                nListCnt = nListCnt + 1                             ' �t�@�C�����̃J�E���^���A�b�v����
                ReDim Preserve gOfflineFileList(nListCnt)           ' �t�@�C�����i�[�G���A���g������
                gOfflineFileList(nListCnt - 1) = szFileName         ' �t�@�C���p�X���i�[
                bProceed = True
            End If
            
                
            If bProceed = False Then
                GoTo ErrorHandler
            End If
            
            ' /////////////////////////////////////////////////////////////////////////
            ' // �ҏW�f�[�^�t�@�C�����쐬
            ' // �R�[�i���Ƃ̒��؃e�L�X�g�t�@�C�����쐬
            bProceed = sOutPutOfflineData(Index)
            If bProceed = False Then
                GoTo ErrorHandler
            End If
            
            Print #iOutFile, "�ݒu�R�[�i�F" & gstrCornerName(Index)
            Print #iOutFile, ""
            
            '1�R�[�i���̒��؃f�[�^��ǂݍ���
            szFileName = Replace(EDIT_DATA_SIMEKIRI, "##", Format(Index + 1, "0#"))
            Set objTs = objFso.OpenTextFile(szFileName, ForReading)
            Print #iOutFile, objTs.ReadAll
            objTs.Close
            Set objFso = Nothing
        End If
    Next Index
    
    '�t�b�^���o��
    Print #iOutFile, FOOTER_STRING
    
    Close #iOutFile
    Set objFso = Nothing

    JprEdit_SimekiriOffline = True
    Exit Function

' /////////////////////////////////////////////////////////
' // �G���[����
ErrorHandler:
    'Call MsgBox("�ُ�I�����܂����B", vbOKOnly, "�I�t���C���o�͌���")
    If iOutFile > 0 Then
        Close #iOutFile
    End If

    Set objFso = Nothing

    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 ALL Rights Reserved
'//
'//  �֐�����  : sOutPutOfflineData
'//  �@�\����  : �I�t���C���f�[�^�}�̏o�͏���
'//  �@�\�T�v  : �R�[�i���Ƃɒ��؃t�@�C��(�e�L�X�g�t�@�C��)���쐬����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sOutPutOfflineData(dwCornerIdx As Integer) As Boolean
            
    Dim szFileName As String                            ' �t�@�C����
    Dim lResult As Long                                 ' ��������
    Dim dwSequense As Long                              ' �V�[�P���X�ԍ�

    ' //////////////////////////////////////////////////////////////
    ' // �t�@�C���쐬����
    ' // �Q�ƌ��t�@�C��SIMEKIRI##.DAT�̃t�@�C�������쐬
    szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(dwCornerIdx + 1, "0#"))
    
    ' //////////////////////////////////////////////////////////////
    ' // �R�[�i���Ƃ̒��؃f�[�^(�e�L�X�g)���쐬
    dwSequense = 0                              ' �V�[�P���X�ԍ�:0�Œ�
    'EG20 V30.1.0.1 DEL START
'    lResult = dllCreateShimekiriFileJpr(dwCornerIdx + 1, dwSequense, _
'                                        PATH_WORK, _
'                                        szFileName)
    'EG20 V30.1.0.1 ADD START
    If gintCornerType(dwCornerIdx) = CORNER_TYPE_KANSEN Then
        '�����R�[�i�Ȃ�Ί����R�[�i�p�̊֐����Ăяo��
        lResult = dllCreateShimekiriFileJprKan(dwCornerIdx + 1, dwSequense, _
                                                PATH_WORK, _
                                                szFileName)
    Else
        '�ݗ��R�[�i�Ȃ�΍ݗ��R�[�i�p�̊֐����Ăяo��
        lResult = dllCreateShimekiriFileJpr(dwCornerIdx + 1, dwSequense, _
                                            PATH_WORK, _
                                            szFileName)
    End If
    'EG20 V30.1.0.1 ADD END
    If lResult = False Then
        sOutPutOfflineData = False
        Exit Function
    End If

    sOutPutOfflineData = True
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JprEdit_KadoVersion
'//  �@�\����  : �ғ��o�[�W�����W���[�i���C���[�W�t�@�C���쐬
'//  �@�\�T�v  : �ғ��o�[�W�����W���[�i���C���[�W�t�@�C�����쐬����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-05-07   CODED   BY [TCC] T.Nakajima
'//             �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_KadoVersion() As Boolean

    Dim strOutputFile As String         '�o�̓t�@�C��
    Dim lngRet As Long                  '�֐��Ԃ�l
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim iOutFile    As Integer          '�t�@�C���ԍ�
    Dim ReadFileKado()    As KADO_VER_DISP_IMAGE_FILE '�ғ��o�[�W�����ꗗ���f�[�^
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim l           As Integer
    Dim strDispImageFileName As String
    Dim objFs       As New FileSystemObject
    Dim intFileNo   As Integer
    Dim iHeadFlg    As Integer
    
    
    On Error GoTo Err_handler
    
    '��ʂŎw�肳�ꂽ�R�[�i�͐ݒu����Ă��邩�H
    If pfSettingCheck(True) = False Then
        '���ׂĖ��ݒu�̃R�[�i�A���@�Ȃ̂ŃG���[�Ƃ���
        GoTo Err_handler
    End If
    
    '�o�̓t�@�C�����ҏW
    strOutputFile = KADOVER_TXTFILE
    
    '// �R�[�i������ʂ�擾 �擾���ʂ�gstrCornerName(0 to 5)�ɓ����Ă���
    gsGetCornerName
    'EG20 V30.0.1.1 ADD START
    ' �R�[�i�^�C�v�擾
    gsGetCornerType
    'EG20 V30.0.1.1 ADD END

    
    '�w�����擾   �擾���ʂ�gstrStationName(0 to 5)�ɓ����Ă���
    gsGetStationName
    
    iHeadFlg = 0
    
    '�t�@�C���o�͊֐���Call
    '�`�F�b�N����Ă���R�[�i�A���@���̃o�[�W�����t�@�C������������o��
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            '���̃R�[�i�A���@�͐ݒu����Ă��邩�H
            If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j)) = True Then
        
                strDispImageFileName = Replace(EDIT_DATA_KADOVERSION, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                'EG20 V30.1.0.1 DEL START
'                lngRet = dllCreateKadoVersionFile(mintDispDiv.KADOVER_FILE_DISP, udtJprPrintSetteingInfo.iCorner(i), _
'                                                  udtJprPrintSetteingInfo.iGouki(j), strDispImageFileName, PATH_IDU_APP, PATH_LDU_APP)
                'EG20 V30.1.0.1 DEL END
                'EG20 V30.1.0.1 ADD START
                If gintCornerType(udtJprPrintSetteingInfo.iCorner(i) - 1) = CORNER_TYPE_KANSEN Then
                
                    lngRet = dllCreateKadoVersionFileKan(mintDispDiv.KADOVER_FILE_DISP, udtJprPrintSetteingInfo.iCorner(i), _
                                                      udtJprPrintSetteingInfo.iGouki(j), strDispImageFileName, PATH_IDU_APP, PATH_LDU_APP)
                
                Else
                    lngRet = dllCreateKadoVersionFile(mintDispDiv.KADOVER_FILE_DISP, udtJprPrintSetteingInfo.iCorner(i), _
                                                      udtJprPrintSetteingInfo.iGouki(j), strDispImageFileName, PATH_IDU_APP, PATH_LDU_APP)
                End If
                'V30.1.0.1 ADD END
                
                '�ُ�I�����̓G���[������
                If lngRet = 0 Then
                    GoTo Err_handler
                    Exit Function
                End If
                
                '�t�@�C�������݂��Ȃ��ꍇ�̓G���[������
                If objFs.FileExists(strDispImageFileName) = False Then
                    GoTo Err_handler
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    '�ғ��o�[�W�����ꗗ �W���[�i���C���[�W�t�@�C�����쐬
    iOutFile = FreeFile
    Open strOutputFile For Output As #iOutFile
    
    '�w�b�_�[��
    PrintHeader iOutFile, "�ғ��o�[�W�����ꗗ"
    
    '�ݒu�w
    Print #iOutFile, "�ݒu�w�F" & gstrStationName(0)
    Print #iOutFile, ""
    
    '��ʕ\���p�t�@�C�����I�[�v��
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        Erase ReadFileKado
        If i > 0 Then
            '1�R�[�i�ڂ͑S�̃o�[�W������\�����Ă���R�[�i�����o��
            '���̃R�[�i�͐ݒu����Ă��邩�H
            'If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i)) = True Then
            If IsTaisyoCorner(udtJprPrintSetteingInfo.iCorner(i)) = True Then
                '�ΏۃR�[�i�ł����Ă��Ώۍ��@���Ȃ���������Ȃ�
                For j = 0 To 15
                    If IsTaisyoGoki(udtJprPrintSetteingInfo.iCorner(i), j + 1) = True Then
                        Print #iOutFile, "�R�[�i���F" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                        Exit For
                    End If
                Next j
                        
            End If
        End If
    
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            '���̍��@���ݒu����Ă��邩�H
            If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j)) = True Then
    
                intFileNo = FreeFile
                strDispImageFileName = Replace(EDIT_DATA_KADOVERSION, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                Open strDispImageFileName For Input As #intFileNo
        
                '��ʕ\���p�f�[�^(csv)���G���A�ɓǂݍ���
                k = 0
                Do While Not EOF(intFileNo)
                    ReDim Preserve ReadFileKado(k)
                    'intKishu, intCorner, intGokiDiv, strName, strMaker, strVer, strDate
                    Input #intFileNo, _
                            ReadFileKado(k).strKishu, ReadFileKado(k).strCorner, ReadFileKado(k).strGokiDiv, _
                            ReadFileKado(k).strName, ReadFileKado(k).strMaker, ReadFileKado(k).strVer, ReadFileKado(k).strDate
                    k = k + 1
                Loop
                '�t�@�C���N���[�Y
                Close #intFileNo
                
                '�ŏ��̃��[�v�����S�̏���\��
                'If i = 0 And j = 0 Then
                If iHeadFlg = 0 Then
                    
                    '�����Ď��ՑS�̃o�[�W����
                    Print #iOutFile, "�����Ď��ՑS�̃o�[�W����"
                    Print #iOutFile, ReadFileKado(0).strVer
                    
                    '�����Ď���
                    Print #iOutFile, "�����Ď���"
                    Print #iOutFile, ReadFileKado(1).strVer
                    
                    '�h�c�t�o�[�W����
                    Print #iOutFile, "�h�c�t"
                    Print #iOutFile, ReadFileKado(2).strVer
                    
                    '�k�c�t�o�[�W����
                    Print #iOutFile, "�k�c�t"
                    Print #iOutFile, ReadFileKado(3).strVer
                    Print #iOutFile, ""
                    
                    '�����
                    Print #iOutFile, "�����"
                    Print #iOutFile, ReadFileKado(4).strVer
                    Print #iOutFile, ""
                    
                    '�R�[�i��
                    Print #iOutFile, "�R�[�i���F" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                    
                    iHeadFlg = 1
                End If
    
                '���@�ԍ�
                Print #iOutFile, "���@�ԍ��F" & Format(udtJprPrintSetteingInfo.iGouki(j), "00") & "���@"
                '�e�v���O�����o�[�W����(6�s�ڂ���e�v���O�����o�[�W����)
                For l = 0 To k - 1
                    If ReadFileKado(l).strKishu = "06" Then
                        '�\���̏ꍇ�̓o�[�W�������o���Ȃ�
                        If ReadFileKado(l).strName = "�\���P" Or ReadFileKado(l).strName = "�\���Q" Then
                            Print #iOutFile, ReadFileKado(l).strName
                        Else
                            Print #iOutFile, ReadFileKado(l).strName & Space(11 - LenB(StrConv(ReadFileKado(l).strName, vbFromUnicode))) & ReadFileKado(l).strVer
                        End If
                    End If
                Next l
                Print #iOutFile, ""
            End If
        Next j
    Next i
    
    Print #iOutFile, FOOTER_STRING
    Close #iOutFile
    
  
    JprEdit_KadoVersion = True
    Exit Function

Err_handler:

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    If iOutFile > 0 Then
        Close #iOutFile
    End If
    
    Set objFs = Nothing

    'MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
    '�u�ғ��o�[�W�����Ǘ���ʁF�ғ��o�[�W�������}�̏o�͏����ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_KadoVersion = False

End Function


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2013 All Right Reserved
'//
'//  �֐����� : JprEdit_EkimuId
'//  �T�v     : �w���@��ID�W���[�i���C���[�W�t�@�C���쐬����
'//  ����     : �w���@��ID�W���[�i���C���[�W�t�@�C�����쐬����
'//  ���Ұ�   :
'//           :
'//
'//  ORIGINAL  �F(EG20 V7.2.0.1) 2013-06-26  CODED BY  [TCC] T.Nakajima
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_EkimuId() As Boolean
    
    Dim sEkimuIDFile    As String   '�w���@��ID�t�@�C���p�X
    Dim iRet            As Integer  'INI�擾�߂�l
    Dim sFolder         As String * MAX_PATH_SIZE  '�t�H���_��
    Dim sFile           As String   '�t�@�C����
    Dim MyName          As String   '�t�@�C����������
    Dim bRet            As Boolean  '�߂�l
    Dim lngErrCode      As Long     '�G���[�R�[�h
    Dim intFileNo       As Integer  '�t�@�C���ԍ�
    Dim strWork         As String   '��ƃG���A
    Dim dwErrsts        As Long
    Dim sFolderName     As String
    Dim objFso          As New FileSystemObject
    Dim objTs           As TextStream
    
        
    On Error GoTo Err_handler
    sFolder = ""
    
    '�������ʁF���펞�͉�ʕ\������
    iRet = GetPrivateProfileString(IDU_SECTION_NAME, _
                                   IDU_EKIMUID_KEY, _
                                   EKIMU_DEFU, sFolder, Len(sFolder), _
                                   PATH_IDU_INI_FILE)
    If iRet = 0 Then
      sFolder = EKIMU_DEFU
    End If
    sEkimuIDFile = ""
    '�v����ʒl���t�@�C�����쐬
    sFile = Replace(EKIMU_ID_FILE, "##", Format(iSendType, "0#"))
    If iRet = 0 Then
       sFolderName = RTrim(sFolder)
    Else
       sFolderName = Mid(sFolder, 1, iRet)
    End If
    '�p�X�ϊ�����
    sFolderName = pfChangeFolderName(sFolderName)
    '�w���@��ID�t�@�C���p�X�쐬
    sEkimuIDFile = sFolderName & "\" & sFile
    '�t�@�C���L���`�F�b�N
    If Dir(sEkimuIDFile, vbNormal) = "" Then
       Exit Function
    End If
    
    '/////////////////////////////////////////////////////////////////////
    '//�ێ��p�֐��F�w���@��ID��ʕ\���p�t�@�C���쐬����
    '////////////////////////////////////////////////////////////////////
    bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE, PATH_IDU_APP, 1) 'V1.8.0.1 ADD
    
    If bRet = False Then
        GoTo Err_handler
        Exit Function
    End If
    
    
    '/////////////////////////////////////////////////////////////////////
    '//�W���[�i���C���[�W�t�@�C�����쐬
    '////////////////////////////////////////////////////////////////////
    intFileNo = FreeFile
    Open EKIMUKIKI_ID_TXTFILE For Output As #intFileNo
    
    '�w�b�_���o��
    PrintHeader intFileNo, "�w���@��h�c�o��"
    
    '�ݒu�w��
    gsGetStationName
    Print #intFileNo, "�ݒu�w�F" & gstrStationName(0)
    Print #intFileNo, ""
    
    '�f�[�^�����Ȃ���
    Set objTs = objFso.OpenTextFile(MN_VERSI_FILE, ForReading)
    Print #intFileNo, objTs.ReadAll
    objTs.Close
    Set objFso = Nothing
    
    '�t�b�^���쐬
    Print #intFileNo, FOOTER_STRING
    
    Close #intFileNo
    
    JprEdit_EkimuId = True
    
    Exit Function

Err_handler:

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    
    Set objFso = Nothing

    'MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
    '�u�ғ��o�[�W�����Ǘ���ʁF�ғ��o�[�W�������}�̏o�͏����ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_EkimuId = False
    
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfChangeFolderName
'//  �@�\����  : �t�H���_�p�X�ϊ�����
'//  �@�\�T�v  : INI�t�@�C�����擾�����t�H���_��`�̕ϊ����s���B
'//
'//              �^        ����         �Ӗ�
'//  ����      : String sFolderName    [IN]INI��`
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Function pfChangeFolderName(sFolderName As String) As String
   Dim iPath As Integer
   Dim sRootPath As String
   Dim sFolder As String
      
   '�u���v�ʒu���擾
   iPath = InStr(sFolderName, "\")
   If iPath = 0 Then
     sRootPath = Mid(sFolderName, 1)
   Else
     '�u���v�O��������擾
     sRootPath = Mid(sFolderName, 1, iPath - 1)
     '�u���v�㕶������擾
     sFolder = Mid(sFolderName, iPath + 1)
   End If
   Select Case sRootPath
      Case APL
        '�A�v�����[�g
        sRootPath = PATH_IDU_APP
      Case LOG
        '���O���[�g
        sRootPath = PATH_IDU_LOG
      Case Data
        'DB���[�g
        sRootPath = PATH_IDU_DB
      Case BACKUP
        '�o�b�N�A�b�v���[�g
        sRootPath = PATH_BUC
   End Select
    '�p�X�A��
    pfChangeFolderName = sRootPath + "\" + sFolder
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : JprEdit_TukaData
'//  �@�\����  : �ʉ߃f�[�^/���p���z�W���[�i���C���[�W�t�@�C���쐬
'//  �@�\�T�v  : �ʉ߃f�[�^/���p���z�W���[�i���C���[�W�t�@�C�����쐬����
'//
'//              �^        ����      �Ӗ�
'//  ����      : long      dwDataKind �f�[�^���    �ʉߔ}�́F306010
'//                                                 ���p�}�́F306020
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-01   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_TukaData(dwDataKind As Long) As Boolean
    
    Dim strFilePath As String           '�o�̓t�@�C���p�X
    Dim intCount As Integer             '�J�E���^
    'EG20 V30.1.0.1 DEL START
'    Dim intOutFile As Integer           '�o�̓t�@�C���ԍ�
'    Dim strBaitaiFileName As String     ' �}�̏o�̓t�@�C�� TUKA�R�[�i��YYYYMMDDhhmmss.csv ICRIYO�R�[�i��YYYYMMDDhhmmss.csv
'    Dim ReadFileBaitai()  As BAITAI_OUTPUT_IMAGE_FILE '�}�̏o�̓t�@�C��
'    Dim strLineCount()  As String
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim l As Integer
'    Dim strCammaArray() As String   '�J���}��؂��1���ڂ����o�����f�[�^

'    Dim fso As New FileSystemObject
'    Dim FsoTS As TextStream
    
'    Dim iKomokuMaxCnt       As Integer      ' �W�v�f�[�^���ڂ̍ő吔
'    Dim iStartLineKaisatu   As Integer      ' ���D���f�[�^�̊J�n�s�i�b�r�u�t�@�C���́j
'    Dim iStartLineShusatu   As Integer      ' �W�D���f�[�^�̊J�n�s�i�b�r�u�t�@�C���́j
    'EG20 V30.1.0.1 DEL END
    
    'Dim intJprFile        As Integer
    
    On Error GoTo Err_handler
    '��ʂŎw�肳�ꂽ�R�[�i�͐ݒu����Ă��邩�H
    If pfSettingCheck(False) = False Then
        '���ׂĖ��ݒu�Ȃ̂ŃG���[
        GoTo Err_handler
    End If
  
    '////////////////////////////////////////////////
    '// �ݒu�w�E�R�[�i������ʂ�擾
    gsGetStationName
    gsGetCornerName
    gsGetCornerType
    gsGetShukeiKoumoku     '�W�v���ڂ̏o�͗L�����擾    EG20 V30.1.0.1 ADD

   
    '�R�[�i�P�ʂŏ���
    
    '/////////////////////////////////////////////
    '// �W���[�i���C���[�W�t�@�C���쐬
    
    '�o�̓t�@�C�����I�[�v������B
    intJprFile = FreeFile
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        Open TUKA_TXTFILE For Output As #intJprFile
    ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then
        Open ICRIYO_TXTFILE For Output As #intJprFile
    Else
        JprEdit_TukaData = False
        Exit Function
    End If

   '�^�C�g���\��
   If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        'EG20 V30.1.0.1 DEL START �i�ݗ��Ɗ����ɂ���Đݒ�l���قȂ�̂ŃC���[�W�t�@�C�������ֈړ��j
'        iKomokuMaxCnt = MAX_KOMOKU_NUM_TUKA
'        iStartLineKaisatu = 6   '���D���̖��ׂ͌��t�@�C��(CSV)�z���(6)����
'        iStartLineShusatu = 60  '�W�D���̖��ׂ͌��t�@�C��(CSV)�z���(60)����
        'EG20 V30.1.0.1 DEL END
        PrintHeader intJprFile, "�ʉ߃f�[�^�o��"
    Else
        'EG20 V30.1.0.1 DEL START �i�ݗ��Ɗ����ɂ���Đݒ�l���قȂ�̂ŃC���[�W�t�@�C�������ֈړ��j
'        iKomokuMaxCnt = MAX_KOMOKU_NUM_KINGAKU
'        iStartLineKaisatu = 6   '���D���̖��ׂ͌��t�@�C��(CSV)�z���(6)����
'        iStartLineShusatu = 25  '�W�D���̖��ׂ͌��t�@�C��(CSV)�z���(60)����
        'EG20 V30.1.0.1 DEL END
        PrintHeader intJprFile, "���p���z�f�[�^�o��"
    End If

    '�ݒu�w�E�R�[�i���o��
    Print #intJprFile, "�ݒu�w�F" & gstrStationName(0)
    
    For intCount = 0 To UBound(glngTergetCorner)
    
        If glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON Then
            'EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
                    psMakeTukaImageFileKan intCount
                Else
                    psMakeRiyoImageFileKan intCount
                End If
            Else
                psMakeTukaRiyoImageFile intCount, dwDataKind
            End If
            
        
            'EG20 V30.1.0.1 DEL START
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then       '�ʉ߃f�[�^
'                strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(intCount) & gstrCornerName(intCount) & "*.csv")
'            ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then    '���p���z�f�[�^
'                strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(intCount) & gstrCornerName(intCount) & "*.csv")
'            Else
'                JprEdit_TukaData = False
'                Exit Function
'            End If
'
'            '////////////////////////////////////////////////
'            '// �ʉ߃f�[�^/���p���z�̔}�̏o�̓t�@�C�����擾
'            '�t�@�C���ԍ��擾
'            '�w���́{�R�[�i����yyyymmddhhmmss.csv
'            Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
'            j = FsoTS.Line
'            FsoTS.Close
'
'            ReDim strLineCount(j) As String         'CSV�t�@�C����1�s������Ă���
'
'            i = 0
'            Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
'            Do Until FsoTS.AtEndOfStream = True
'                strLineCount(i) = FsoTS.ReadLine
'                i = i + 1
'            Loop
'            FsoTS.Close
'            Set fso = Nothing
'
'            '�}�̏o�̓t�@�C���C���[�W�\���̂ɃZ�b�g����
'            ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE         '�t�@�C���Ǎ��p�G���A
'            l = 0
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
'
'                For i = 0 To j - 1
'                    Select Case i
'                        Case 0, 1, 2, 3, 4, 57, 58    'TUKAxxxx.csv��1�`4�s�ڂ܂ł̓^�C�g���Ȃ̂ŁA���ږ��ɃZ�b�g
'                            ReadFileBaitai(i).strKomokuName = strLineCount(i)
'                        Case Else
'                            '�J���}��؂��1���ڂ����o���B
'                            strCammaArray = Split(strLineCount(i), ",")
'                            For k = 0 To UBound(strCammaArray())
'                                If k = 0 Then
'                                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
'                                ElseIf k = 1 Then
'                                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
'                                Else
'                                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
'                                    l = l + 1
'                                End If
'                            Next k
'                    End Select
'                    l = 0
'                Next i
'            Else
'                For i = 0 To j - 1
'                    Select Case i
'                        Case 0, 1, 2, 3, 4, 22, 23    'ICRIYOxxxx.csv��1�`4�s�ڂ܂ł̓^�C�g���Ȃ̂ŁA���ږ��ɃZ�b�g
'                            ReadFileBaitai(i).strKomokuName = strLineCount(i)
'                        Case Else
'                            '�J���}��؂��1���ڂ����o���B
'                            strCammaArray = Split(strLineCount(i), ",")
'                            For k = 0 To UBound(strCammaArray())
'                                If k = 0 Then
'                                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
'                                ElseIf k = 1 Then
'                                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
'                                Else
'                                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
'                                    l = l + 1
'                                End If
'                            Next k
'                    End Select
'                    l = 0
'                Next i
'            End If
'
'            Print #intJprFile, "�ݒu�R�[�i�F" & gstrCornerName(intCount)
'            Print #intJprFile, ""
'
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
'                Print #intJprFile, "�y�ʉ߃f�[�^�z"
'            Else
'                Print #intJprFile, "�y�h�b�J�[�h���p���z�f�[�^�z"
'            End If
'            '/////////////////////
'            '���D���f�[�^�̏o��
'            Print #intJprFile, "���D���ʉߍ��v"
'
'            For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
'                If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "0" Then
'                    '���ږ���0���Z�b�g����Ă�����ȍ~�͏o�͂��Ȃ�
'                    Exit For
'                Else
'                    '���؃I�t���C���W���[�i���Ƃ��킹�邽�߂̗�O����
'                    If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "���̑�IC (��)" Then
'                        ReadFileBaitai(i + iStartLineKaisatu).strKomokuName = "���̑�IC(��)" & Space(38)   '�X�y�[�X������
'                    End If
'
'                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
'                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineKaisatu).strGoukei), "#,0"), 10)
'                End If
'            Next i
'            Print #intJprFile, ""
'
'            '/////////////////////
'            '�W�D���f�[�^�̏o��
'            Print #intJprFile, "�W�D���ʉߍ��v"
'
'            For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
'                If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "0" Then
'                    '���ږ���0���Z�b�g����Ă�����ȍ~�͏o�͂��Ȃ�
'                    Exit For
'                Else
'                    '���؃I�t���C���W���[�i���Ƃ��킹�邽�߂̗�O����
'                    If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "���̑�IC (��)" Then
'                        ReadFileBaitai(i + iStartLineShusatu).strKomokuName = "���̑�IC(��)" & Space(38)    '�X�y�[�X������
'                    End If
'
'                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineShusatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
'                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineShusatu).strGoukei), "#,0"), 10)
'                End If
'            Next i
'            Print #intJprFile, ""
        'EG20 V30.1.0.1 DEL END
            
        End If
    Next intCount
    
    Print #intJprFile, FOOTER_STRING
    Close #intJprFile
    
    JprEdit_TukaData = True
    Exit Function
    
'�G���[����
Err_handler:

    'EG20 V30.1.0.1 DEL START
'    If intOutFile > 0 Then
'        Close #intOutFile
'    End If
    'EG20 V30.1.0.1 DEL END
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

'    Set fso = Nothing      'EG20 V30.1.0.1 DEL
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
    JprEdit_TukaData = False
                                      
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : PadLeft
'//  �@�\����  : �E��
'//  �@�\�T�v  : �w��̕������ɂȂ�܂Ő擪�𕶎��Ŗ��߂�B
'//
'//              �^        ����         �Ӗ�
'//  ����      : string    strTarget    �����Ώە�����
'//              Integer   iLength      �����̒���
'//              string    chOne        ���߂镶��(�ȗ����͔��p�X�y�[�X)
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : string    �E�񂹂��ꂽ������
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function PadLeft(ByVal strTarget As String, ByVal iLength As Integer, Optional ByVal chOne As String = " ") As String
    
    Do While (Len(strTarget) < iLength)
        strTarget = chOne & strTarget
    Loop

    PadLeft = Right$(strTarget, iLength)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : PadRight
'//  �@�\����  : ���񂹁i�������X�y�[�X�Ŗ��߂�)
'//  �@�\�T�v  : �w��̕������ɂȂ�܂Ő擪�𕶎��Ŗ��߂�B
'//
'//              �^        ����         �Ӗ�
'//  ����      : string    strTarget    �����Ώە�����
'//              Integer   iLength      �����̒���
'//              string    chOne        ���߂镶��(�ȗ����͔��p�X�y�[�X)
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : string    ���񂹂��ꂽ������
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function PadRight(ByVal strTarget As String, ByVal iLength As Integer, Optional ByVal chOne As String = " ") As String
    Do While (Len(strTarget) < iLength)
        strTarget = strTarget & chOne
    Loop

    PadRight = Left$(strTarget, iLength)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : PrintHeader
'//  �@�\����  : �w�b�_���쐬
'//  �@�\�T�v  : �w�b�_�����쐬����B�i�W���[�i���̂P�`�S�s��)
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iFileNum     �t�@�C���ԍ�
'//              string    strTitle     �W���[�i���^�C�g��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.1.0.1 2016-06-14   REVISED BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader(iFileNum As Integer, strTitle As String)
    Dim lpSystemTime            As SYSTEMTIME               '���[�J���������擾
    'EG30 V32.1.0.1 DEL START
    'Print #iFileNum, "*************EG20*************"
    'EG30 V32.1.0.1 DEL END
    Print #iFileNum, strTitle
    
    '���[�J���������擾
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "�󎚓����F" & lpSystemTime.wYear & "�N" & Format(lpSystemTime.wMonth, "00") & "��" & Format(lpSystemTime.wDay, "00") & "��" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, ""
End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : PrintHeader2
'//  �@�\����  : �w�b�_���쐬
'//  �@�\�T�v  : �w�b�_�����쐬����B�i�W���[�i���̂P�`�S�s��)
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iFileNum     �t�@�C���ԍ�
'//              string    strTitle     �W���[�i���^�C�g��
'//              string    strTitle2    �W���[�i���^�C�g���Q�s��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V30.3.0.1) 2014-10-01   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_003_01�z
'//     REVISIONS :(EG30 V32.1.0.1 2016-06-14   REVISED BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader2(iFileNum As Integer, strTitle As String, strTitle2 As String)
    Dim lpSystemTime            As SYSTEMTIME               '���[�J���������擾
    
    'EG30 V32.1.0.1 DEL START
    'Print #iFileNum, "*************EG20*************"
    'EG30 V32.1.0.1 DEL END
    Print #iFileNum, strTitle
    Print #iFileNum, strTitle2
    
    '���[�J���������擾
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "�󎚓����F" & lpSystemTime.wYear & "�N" & Format(lpSystemTime.wMonth, "00") & "��" & Format(lpSystemTime.wDay, "00") & "��" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, "�ۑ������F" & pfGetSaveDate(0)    '�R�[�i0�̕ۑ�����  'EG30 V32.1.0.1 ADD
    Print #iFileNum, ""
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  �֐�����  : PrintHeader3
'//  �@�\����  : �w�b�_���쐬
'//  �@�\�T�v  : �w�b�_�����쐬����B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iFileNum     �t�@�C���ԍ�
'//              string    strTitle     �W���[�i���^�C�g��
'//              string    strSaveDate  �ۑ�����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-22   CODED   BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader3(iFileNum As Integer, strTitle As String, strSaveDate As String)
    Dim lpSystemTime            As SYSTEMTIME               '���[�J���������擾
    
    Print #iFileNum, strTitle
    '���[�J���������擾
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "�󎚓����F" & lpSystemTime.wYear & "�N" & Format(lpSystemTime.wMonth, "00") & "��" & Format(lpSystemTime.wDay, "00") & "��" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, "�ۑ������F" & strSaveDate
    Print #iFileNum, ""
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfCornerGokiCheck
'//  �@�\����  : �R�[�i���@�`�F�b�N
'//  �@�\�T�v  : ��ʂŃ`�F�b�N���ꂽ�R�[�i���@�����݂��邩�m�F����
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iCorner      �R�[�i(1�`6)
'//              Integer   iGoki        ���@�i�ȗ��\�F�ȗ����͍��@�̓`�F�b�N���Ȃ�) 1�`16
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : Boolean   true/false   true:�ݒu����Ă���   false:�ݒu����Ă��Ȃ�
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfCornerGokiCheck(iCornerNo As Integer, Optional iGoki As Integer = 0) As Boolean
    Dim i       As Integer
    Dim bRet    As Boolean
    bRet = False
    
    If gudtSettiCorner(iCornerNo - 1).intGokiNum > 0 Then
        ' �w�肵���R�[�i�͐ݒu����Ă���
        ' �p�����[�^�Ŏw�肳�ꂽ���@�͐ݒu����Ă��邩�H
        If iGoki <> 0 Then
            For i = 0 To 15
                If iGoki = gudtSettiCorner(iCornerNo - 1).intGokiNo(i) Then
                    bRet = True
                    Exit For
                End If
            Next i
        Else
            bRet = True
        End If
    Else
        '�w�肳�ꂽ�R�[�i�͐ݒu����Ă��Ȃ�
    End If
    
    pfCornerGokiCheck = bRet
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  �֐�����  : pfCornerGokiToGateNo
'//  �@�\����  : �R�[�i���@���_�����@�ԍ��ɕϊ�
'//  �@�\�T�v  : ��ʂŃ`�F�b�N���ꂽ�R�[�i���@�����݂��邩�m�F���A�_�����@�ԍ���Ԃ��B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iCorner      �R�[�i(1�`6)
'//              Integer   iGoki        ���@�i�ȗ��\�F�ȗ����͍��@�̓`�F�b�N���Ȃ�) 1�`16
'//              Integer   iGateNo      �_�����@(1�`32)
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : Boolean   true/false   true:�ݒu����Ă���   false:�ݒu����Ă��Ȃ�
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-28   CODED   BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfCornerGokiToGateNo(iCornerNo As Integer, iGoki As Integer, ByRef iGateNo As Integer) As Boolean
    Dim i       As Integer
    Dim bRet    As Boolean
    bRet = False
    iGateNo = 0
    If gudtSettiCorner(iCornerNo - 1).intGokiNum > 0 Then
        ' �w�肵���R�[�i�͐ݒu����Ă���
        ' �p�����[�^�Ŏw�肳�ꂽ���@�͐ݒu����Ă��邩�H
        If iGoki <> 0 Then
            For i = 0 To 15
                If iGoki = gudtSettiCorner(iCornerNo - 1).intGokiNo(i) Then
                    iGateNo = gudtSettiCorner(iCornerNo - 1).intGateNo(i)
                    bRet = True
                    Exit For
                End If
            Next i
        Else
            bRet = True
        End If
    Else
        '�w�肳�ꂽ�R�[�i�͐ݒu����Ă��Ȃ�
    End If
    
    pfCornerGokiToGateNo = bRet
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfSettingaCheck
'//  �@�\����  : �R�[�i���@�̐ݒu�m�F
'//  �@�\�T�v  : �W���[�i���ɏo�͂���R�[�i���@���ݒu����Ă��邩�m�F����B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Boolean   bGokiCheck   ���@�`�F�b�N�L��
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : Boolean   true/false   true:�ݒu����Ă���   false:�ݒu����Ă��Ȃ�
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfSettingCheck(Optional bGokiCheck As Boolean = True) As Boolean
    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer
    
    '��ʂŐݒ肳�ꂽ�����ǂꂩ�ЂƂł��ݒu����Ă���R�[�i���@�������OK�Ƃ���
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        If gudtSettiCorner(udtJprPrintSetteingInfo.iCorner(i) - 1).intGokiNum > 0 Then
            '���̃R�[�i�͐ݒu����Ă���
            ' �`�F�b�N���ꂽ���@�͂��̃R�[�i�ɑ��݂��Ă��邩�H(���@�`�F�b�N����̏ꍇ)
            If bGokiCheck = True Then
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    For k = 0 To 15
                        If udtJprPrintSetteingInfo.iGouki(j) = gudtSettiCorner(udtJprPrintSetteingInfo.iCorner(i) - 1).intGokiNo(k) Then
                            pfSettingCheck = True
                            Exit Function
                        End If
                    Next k
                Next j
            Else
                pfSettingCheck = True
                Exit Function
            End If
        End If
    Next i
                
    pfSettingCheck = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : MidByte
'//  �@�\����  : �R�[�i���@�̐ݒu�m�F
'//  �@�\�T�v  : �W���[�i���ɏo�͂���R�[�i���@���ݒu����Ă��邩�m�F����B
'//
'//              �^        ����         �Ӗ�
'//  ����      : String    strTarget     �Ώە�����
'//              long      iStart       �J�n�ʒu(1�o�C�g�`)
'//              Variant   ibyteCount   ����
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    :String                  ���o���ꂽ������
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function MidByte(ByVal strTarget As String, ByVal iStart As Long, Optional ByVal iByteCount As Variant) As String
    If IsMissing(iByteCount) = False Then
        MidByte = StrConv(MidB$(StrConv(strTarget, vbFromUnicode), iStart, iByteCount), vbUnicode)
    Else
        MidByte = StrConv(MidB$(StrConv(strTarget, vbFromUnicode), iStart), vbUnicode)
    End If
End Function


'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : psMakeTukaRiyoImageFile
'//  �@�\����  : �ʉ߃f�[�^/���p���z�f�[�^�W���[�i���̃C���[�W�t�@�C���쐬�i�ݗ��p�j
'//  �@�\�T�v  : �ʉ߃f�[�^����ї��p���z�f�[�^�W���[�i���̃C���[�W�t�@�C�����쐬����B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iCornerIdx   �R�[�i�C���f�b�N�X
'//              Long      dwDataKind   �f�[�^��ʁi�ʉ߃f�[�^�A���p���z�f�[�^�j
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : ����
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F      �����R�[�i�p�̃C���[�W�t�@�C���쐬�������ʓr�K�v�ƂȂ������߁A
'//              JprEdit_TukaData()����T�u���[�`����
'///////////////////////////////////////////////////////////////////
Private Sub psMakeTukaRiyoImageFile(iCornerIdx As Integer, dwDataKind As Long)
    
    Dim strBaitaiFileName   As String                       '�}�̏o�̓t�@�C�� TUKA�R�[�i��YYYYMMDDhhmmss.csv ICRIYO�R�[�i��YYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE     '�}�̏o�̓t�@�C��
    Dim intOutFile          As Integer                      '�o�̓t�@�C���ԍ�
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       '�J���}��؂��1���ڂ����o�����f�[�^
    Dim iKomokuMaxCnt       As Integer                      ' �W�v�f�[�^���ڂ̍ő吔
    Dim iStartLineKaisatu   As Integer                      ' ���D���f�[�^�̊J�n�s�i�b�r�u�t�@�C���́j
    Dim iStartLineShusatu   As Integer                      ' �W�D���f�[�^�̊J�n�s�i�b�r�u�t�@�C���́j
            
    On Error GoTo Err_handler
    
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then       '�ʉ߃f�[�^
        
        iKomokuMaxCnt = MAX_KOMOKU_NUM_TUKA
        iStartLineKaisatu = 6   '���D���̖��ׂ͌��t�@�C��(CSV)�z���(6)����
        iStartLineShusatu = 60  '�W�D���̖��ׂ͌��t�@�C��(CSV)�z���(60)����
        
        strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
    ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then    '���p���z�f�[�^
        
        iKomokuMaxCnt = MAX_KOMOKU_NUM_KINGAKU
        iStartLineKaisatu = 6   '���D���̖��ׂ͌��t�@�C��(CSV)�z���(6)����
        iStartLineShusatu = 25  '�W�D���̖��ׂ͌��t�@�C��(CSV)�z���(60)����
        
        strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
    End If
           
    '////////////////////////////////////////////////
    '// �ʉ߃f�[�^/���p���z�̔}�̏o�̓t�@�C�����擾
    '�t�@�C���ԍ��擾
    '�w���́{�R�[�i����yyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSV�t�@�C����1�s������Ă���
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '�}�̏o�̓t�@�C���C���[�W�\���̂ɃZ�b�g����
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE         '�t�@�C���Ǎ��p�G���A
    l = 0
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
    
        For i = 0 To j - 1
            Select Case i
                Case 0, 1, 2, 3, 4, 57, 58    'TUKAxxxx.csv��1�`4�s�ڂ܂ł̓^�C�g���Ȃ̂ŁA���ږ��ɃZ�b�g
                    ReadFileBaitai(i).strKomokuName = strLineCount(i)
                Case Else
                    '�J���}��؂��1���ڂ����o���B
                    strCammaArray = Split(strLineCount(i), ",")
                    For k = 0 To UBound(strCammaArray())
                        If k = 0 Then
                            ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                        ElseIf k = 1 Then
                            ReadFileBaitai(i).strGoukei = strCammaArray(k)
                        Else
                            ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                            l = l + 1
                        End If
                    Next k
            End Select
            l = 0
        Next i
    Else
        For i = 0 To j - 1
            Select Case i
                Case 0, 1, 2, 3, 4, 22, 23    'ICRIYOxxxx.csv��1�`4�s�ڂ܂ł̓^�C�g���Ȃ̂ŁA���ږ��ɃZ�b�g
                    ReadFileBaitai(i).strKomokuName = strLineCount(i)
                Case Else
                    '�J���}��؂��1���ڂ����o���B
                    strCammaArray = Split(strLineCount(i), ",")
                    For k = 0 To UBound(strCammaArray())
                        If k = 0 Then
                            ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                        ElseIf k = 1 Then
                            ReadFileBaitai(i).strGoukei = strCammaArray(k)
                        Else
                            ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                            l = l + 1
                        End If
                    Next k
            End Select
            l = 0
        Next i
    End If

    Print #intJprFile, "�ݒu�R�[�i�F" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        Print #intJprFile, "�y�ʉ߃f�[�^�z"
    Else
        Print #intJprFile, "�y�h�b�J�[�h���p���z�f�[�^�z"
    End If
    '/////////////////////
    '���D���f�[�^�̏o��
    Print #intJprFile, "���D���ʉߍ��v"
    
    For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
        If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "0" Then
            '���ږ���0���Z�b�g����Ă�����ȍ~�͏o�͂��Ȃ�
            Exit For
        Else
            '���؃I�t���C���W���[�i���Ƃ��킹�邽�߂̗�O����
            If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "���̑�IC (��)" Then
                ReadFileBaitai(i + iStartLineKaisatu).strKomokuName = "���̑�IC(��)" & Space(38)   '�X�y�[�X������
            End If
        
            Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
            & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineKaisatu).strGoukei), "#,0"), 10)
        End If
    Next i
    Print #intJprFile, ""
    
    '/////////////////////
    '�W�D���f�[�^�̏o��
    Print #intJprFile, "�W�D���ʉߍ��v"
    
    For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
        If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "0" Then
            '���ږ���0���Z�b�g����Ă�����ȍ~�͏o�͂��Ȃ�
            Exit For
        Else
            '���؃I�t���C���W���[�i���Ƃ��킹�邽�߂̗�O����
            If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "���̑�IC (��)" Then
                ReadFileBaitai(i + iStartLineShusatu).strKomokuName = "���̑�IC(��)" & Space(38)    '�X�y�[�X������
            End If
        
            Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineShusatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
            & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineShusatu).strGoukei), "#,0"), 10)
        End If
    Next i
    Print #intJprFile, ""
        
    
    'Print #intJprFile, FOOTER_STRING
    'Close #intJprFile
    
    Exit Sub
    
'�G���[����
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : psMakeTukaImageFileKan
'//  �@�\����  : �ʉ߃f�[�^�W���[�i���̃C���[�W�t�@�C���쐬�i�����p�j
'//  �@�\�T�v  : �ʉ߃f�[�^�W���[�i���̃C���[�W�t�@�C�����쐬����B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iCornerIdx   �R�[�i�C���f�b�N�X
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : ����
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psMakeTukaImageFileKan(iCornerIdx As Integer)
    
    Dim strBaitaiFileName   As String                       '�}�̏o�̓t�@�C�� TUKA�R�[�i��YYYYMMDDhhmmss.csv ICRIYO�R�[�i��YYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE_KAN '�}�̏o�̓t�@�C��
    Dim intOutFile          As Integer                      '�o�̓t�@�C���ԍ�
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       '�J���}��؂��1���ڂ����o�����f�[�^
    Dim iKomokuMaxCnt       As Integer                      ' �W�v�f�[�^���ڂ̍ő吔
    Dim iStartLine          As Integer                      '�e�W�v�u���b�N�̊J�n�s
                                                                
    On Error GoTo Err_handler
    
    '�e�W�v���ڂ̏o�͊J�n�ʒu���擾�iINI�t�@�C���ɂ��o�͗L�����w��ł��邽�߁A�J�n�ʒu�͉ςɂȂ�j
    
    strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
           
    '////////////////////////////////////////////////
    '// �ʉ߃f�[�^/���p���z�̔}�̏o�̓t�@�C�����擾
    '�t�@�C���ԍ��擾
    '�w���́{�R�[�i����yyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSV�t�@�C����1�s������Ă���
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '�}�̏o�̓t�@�C���C���[�W�\���̂ɃZ�b�g����
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE_KAN     '�t�@�C���Ǎ��p�G���A
    l = 0
    
    For i = 0 To j - 1
        If InStr(strLineCount(i), ",") = 0 Then
            '�J���}��؂�ɂȂ��Ă��Ȃ��s�͍��ږ��ɂƂ肠�����f�[�^���Z�b�g
            ReadFileBaitai(i).strKomokuName = strLineCount(i)
        Else
            '�J���}��؂��1���ڂ����o���B
            strCammaArray = Split(strLineCount(i), ",")
            For k = 0 To UBound(strCammaArray())
                If k = 0 Then
                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                ElseIf k = 1 Then
                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
                ElseIf k = 2 Then
                    ReadFileBaitai(i).strNorikae = strCammaArray(k)
                ElseIf k = 3 Then
                    ReadFileBaitai(i).strTukaChoku = strCammaArray(k)
                Else
                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                    l = l + 1
                End If
            Next k
        End If
        l = 0
    Next i

    Print #intJprFile, "�ݒu�R�[�i�F" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    Print #intJprFile, "�y�i�q���V�����ʉ߃f�[�^�z"
    
    '//////////////////////////////////////////////////////////
    '���D�� �V�����ʉ߃f�[�^�̏o��
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAISATU_KANSEN_TUKA)
        Print #intJprFile, "���D���@�V�����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_TUKA_KAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�W�D���@�V�����ʉ߃f�[�^�̏o��
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_SHUSATU_KANSEN_TUKA)
        
        Print #intJprFile, "�W�D���@�V�����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_TUKA_KAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�^�s�s�\�ʉ߃f�[�^�̏o��
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_UNKOU_FUNOU) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_IC_UNKO_FUNOU)
    
        Print #intJprFile, "�^�s�s�\�ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_UNKOU_FUNOU - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '���\�ݏ抷�@�ݗ����ʉ߃f�[�^�̏o��
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_KAN_ZAI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAN_ZAI_TUKA)

        Print #intJprFile, "���|�ݏ抷�@�ݗ����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_NORIKAE_TUKA - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�݁\���抷�@�ݗ����ʉ߃f�[�^�̏o��
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_ZAI_KAN) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_ZAI_KAN_TUKA)
        
        Print #intJprFile, "�݁|���抷�@�ݗ����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_NORIKAE_TUKA - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '���w����~�σf�[�^�̏o��
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KYUSAI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_JIEKI_KYUSAI)
    
        Print #intJprFile, "���w����~�ϒʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_JIEKI_KYUSAI - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '���C��������~�ʉ߃f�[�^�̏o��
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KAISHU_CHUSI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAISHU_CHUSHI)
    
        Print #intJprFile, "���C��������~�ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    'Print #intJprFile, FOOTER_STRING
    'Close #intJprFile
    
    Exit Sub
    
'�G���[����
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : psMakeRiyoImageFileKan
'//  �@�\����  : ���p���z�f�[�^�W���[�i���̃C���[�W�t�@�C���쐬�i�����p�j
'//  �@�\�T�v  : ���p���z�f�[�^�W���[�i���̃C���[�W�t�@�C�����쐬����B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iCornerIdx   �R�[�i�C���f�b�N�X
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : ����
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psMakeRiyoImageFileKan(iCornerIdx As Integer)
    
    Dim strBaitaiFileName   As String                       '�}�̏o�̓t�@�C�� TUKA�R�[�i��YYYYMMDDhhmmss.csv ICRIYO�R�[�i��YYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE_KAN '�}�̏o�̓t�@�C��
    Dim intOutFile          As Integer                      '�o�̓t�@�C���ԍ�
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       '�J���}��؂��1���ڂ����o�����f�[�^
    Dim iKomokuMaxCnt       As Integer                      ' �W�v�f�[�^���ڂ̍ő吔
    Dim iStartLine          As Integer                      '�e�W�v�u���b�N�̊J�n�s
                                                                
    On Error GoTo Err_handler
    
    '�e�W�v���ڂ̏o�͊J�n�ʒu���擾�iINI�t�@�C���ɂ��o�͗L�����w��ł��邽�߁A�J�n�ʒu�͉ςɂȂ�j
    
    strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
           
    '////////////////////////////////////////////////
    '// �ʉ߃f�[�^/���p���z�̔}�̏o�̓t�@�C�����擾
    '�t�@�C���ԍ��擾
    '�w���́{�R�[�i����yyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSV�t�@�C����1�s������Ă���
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '�}�̏o�̓t�@�C���C���[�W�\���̂ɃZ�b�g����
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE_KAN     '�t�@�C���Ǎ��p�G���A
    l = 0
    
    For i = 0 To j - 1
        If InStr(strLineCount(i), ",") = 0 Then
            '�J���}��؂�ɂȂ��Ă��Ȃ��s�͍��ږ��ɂƂ肠�����f�[�^���Z�b�g
            ReadFileBaitai(i).strKomokuName = strLineCount(i)
        Else
            '�J���}��؂��1���ڂ����o���B
            strCammaArray = Split(strLineCount(i), ",")
            For k = 0 To UBound(strCammaArray())
                If k = 0 Then
                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                ElseIf k = 1 Then
                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
                ElseIf k = 2 Then
                    ReadFileBaitai(i).strNorikae = strCammaArray(k)
                ElseIf k = 3 Then
                    ReadFileBaitai(i).strTukaChoku = strCammaArray(k)
                Else
                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                    l = l + 1
                End If
            Next k
        End If
        l = 0
    Next i

    Print #intJprFile, "�ݒu�R�[�i�F" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    Print #intJprFile, "�y�i�q���V�������z�f�[�^�z"
    
    '//////////////////////////////////////////////////////////
    '���D�� ��l �V�����X�C�J�ʉߍ��v�̏o��
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_OTONA_SUICA_RIYO)
        Print #intJprFile, "���D�� ��l �V��������ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�W�D�� ��l �V�����X�C�J�ʉߍ��v�̏o��
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_OTONA_SUICA_RIYO)
        Print #intJprFile, "�W�D�� ��l �V��������ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '���D�� ���� �V�����X�C�J�ʉߍ��v�̏o��
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_SHONI_SUICA_RIYO)
        Print #intJprFile, "���D�� ���� �V��������ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�W�D�� ���� �V�����X�C�J�ʉߍ��v�̏o��
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_SHONI_SUICA_RIYO)
        Print #intJprFile, "�W�D�� ���� �V��������ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    
    '//////////////////////////////////////////////////////////
    '�X�C�J��ЊԐ��Z�^���x�����ʉߍ��v
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_CARD_SHIHARAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SEISAN_SHIHARAI)
        Print #intJprFile, "�����ЊԐ��Z�^���x���ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_SEISAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If

    '//////////////////////////////////////////////////////////
    '���D���I�[�g�`���[�W�ʉߍ��v
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_AUTOCHARGE)
        Print #intJprFile, "���D�� �I�[�g�`���[�W�ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_AUTOCHARGE - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�W�D���I�[�g�`���[�W�ʉߍ��v
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_AUTOCHARGE)
        Print #intJprFile, "�W�D�� �I�[�g�`���[�W�ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_AUTOCHARGE - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�V�����^�� ��l�@�X�C�J�ʉߍ��v
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_KANSEN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAN_OTONA_SUICA_RIYO)
        Print #intJprFile, "�����^�� ��l ����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�V�����^�� �����@�X�C�J�ʉߍ��v
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_KANSEN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAN_SHONI_SUICA_RIYO)
        Print #intJprFile, "�����^�� ���� ����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�抷�ݗ��^�� ��l�@�X�C�J�ʉߍ��v
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_ZAIRAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO)
        Print #intJprFile, "�抷�ݗ��^�� ��l ����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '�抷�ݗ��^�� �����@�X�C�J�ʉߍ��v
    '//////////////////////////////////////////////////////////
    'INI�ŏo�͗L�ɐݒ肳��Ă���Ώo�͂���
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_ZAIRAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO)
        Print #intJprFile, "�抷�ݗ��^�� ���� ����ʉߍ��v"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '���ږ���0���Z�b�g����Ă�����o�͂��Ȃ�
            Else
                '���ږ���20���Ɏ��܂�Ȃ��ꍇ�͔��p�X�y�[�X�����Đ��l���o�́i�ʒu�͂��낦�Ȃ�)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    'Print #intJprFile, FOOTER_STRING
    'Close #intJprFile
    
    Exit Sub
    
'�G���[����
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : pfGetStartLineTuka
'//  �@�\����  : �w�肵���W�v���ڂ̈󎚊J�n�ʒu�擾
'//  �@�\�T�v  : �w�肵���W�v���ڂ̈󎚊J�n�ʒu��GAIBU_OUTPUT.INI�ɏ]���ċ��߂�B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   intShukeiKoumoku     �W�v����
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : Integer                �J�n�ʒu
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetStartLineTuka(intShukeiKoumoku As Integer) As Integer

    Dim intCount            As Integer
    Dim intIniIdx           As Integer  'GAIBU_OUTPUT.INI�̃L�[�ɑ΂���C���f�b�N�X
    
    Dim intStartLine        As Integer  '�ʉ߃f�[�^�̊J�n�s���iCSV��j
    
    Dim intNextBlockLine    As Integer  '���̏W�v�u���b�N�̃f�[�^������ʒu�iCSV��j
    
    Dim intNowLine           As Integer  'INI�t�@�C���̏o�͗L���ɏ]���āACSV�t�@�C�����ォ�珇�Ɍ��Ă������Ƃ��̌��ݍs
    
    intNowLine = 0
    intNextBlockLine = 6
    intIniIdx = 0
    
    For intCount = 0 To MAX_TUKA_SHUKEI_KOUMOKU - 1
    
        Select Case intCount
            Case mintTukaShukeiKoumoku.SHUKEI_KAISATU_KANSEN_TUKA              '�y���D�� �V�����ʉ߃f�[�^�z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_TUKA_KAN + 2
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_SHUSATU_KANSEN_TUKA              '�y�W�D���@�V�����ʉ߃f�[�^�z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_TUKA_KAN + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_IC_UNKO_FUNOU                    '�y�^�s�s�\�f�[�^�z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_UNKOU_FUNOU) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_UNKOU_FUNOU + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_KAN_ZAI_TUKA                    '�y��-�� �抷�ʉ߃f�[�^�z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_KAN_ZAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_NORIKAE_TUKA + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_ZAI_KAN_TUKA                    '�y��-�� �抷�ʉ߃f�[�^�z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_ZAI_KAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_NORIKAE_TUKA + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_JIEKI_KYUSAI                    '�y���w����~�ϒʉ߃f�[�^�z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KYUSAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_JIEKI_KYUSAI + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_KAISHU_CHUSHI                  '�y���C��������~�ʉ߃f�[�^�z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KAISHU_CHUSI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI + 3
                End If
            Case Else   '��L�ȊO�͋��z�f�[�^�Ɋւ���ݒ�̂��߃X�L�b�v
                
        End Select
        If intCount <> 0 Then
            intIniIdx = intIniIdx + 1
        End If
        
        '���߂����J�n�ʒu��������A���̍s����Ԃ�
        If intShukeiKoumoku = intCount Then
            pfGetStartLineTuka = intNowLine
            Exit Function
        End If
    Next
    
    '��L��For�����ő�񐔂܂ŉ���ďI�������Ƃ������Ƃ́A���߂����J�n�ʒu�����߂��Ȃ������B
    pfGetStartLineTuka = intNowLine

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : pfGetStartLineKingaku
'//  �@�\����  : �w�肵���W�v���ڂ̈󎚊J�n�ʒu�擾
'//  �@�\�T�v  : �w�肵���W�v���ڂ̈󎚊J�n�ʒu��GAIBU_OUTPUT.INI�ɏ]���ċ��߂�B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   intShukeiKoumoku     �W�v����
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : Integer                �J�n�ʒu
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetStartLineKingaku(intShukeiKoumoku As Integer) As Integer

    Dim intCount            As Integer
    Dim intIniIdx           As Integer  'GAIBU_OUTPUT.INI�̃L�[�ɑ΂���C���f�b�N�X
    
    Dim intStartLine        As Integer  '�ʉ߃f�[�^�̊J�n�s���iCSV��j
    
    Dim intNextBlockLine    As Integer  '���̏W�v�u���b�N�̃f�[�^������ʒu�iCSV��j
    
    Dim intNowLine           As Integer  'INI�t�@�C���̏o�͗L���ɏ]���āACSV�t�@�C�����ォ�珇�Ɍ��Ă������Ƃ��̌��ݍs
    
    intNowLine = 0
    intNextBlockLine = 6
    intIniIdx = 0
    
    For intCount = 0 To MAX_KINGAKU_SHUKEI_KOUMOKU - 1
    
        Select Case intCount
            '�y���D���@��l�@�V�����X�C�J���p���v���z�z
            '�y�W�D���@��l�@�V�����X�C�J���p���v���z�z
            '�y���D���@�����@�V�����X�C�J���p���v���z�z
            
            '�y�����^���@��l�@�X�C�J���p���v���z�z
            '�y�抷�ݗ��^���@��l�@�X�C�J���p���v���z�z
            Case mintKingakuShukeiKoumoku.SHUKEI_KAI_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_SHU_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAI_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAN_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO
                
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_RIYO + 2
                End If
            '�y�W�D���@�����@�V�����X�C�J���p���v���z�z
            '�y�����^���@�����@�X�C�J���p���v���z�z
            '�y�抷�ݗ��^���@�����@�X�C�J���p���v���z�z
            Case mintKingakuShukeiKoumoku.SHUKEI_SHU_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAN_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_RIYO + 3
                End If
            '�y�X�C�J��ЊԐ��Z�f�[�^�@�^���x���z�z
            Case mintKingakuShukeiKoumoku.SHUKEI_SEISAN_SHIHARAI
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_CARD_SHIHARAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_SEISAN + 3
                End If
            '�y���D���@�I�[�g�`���[�W�f�[�^�z
            Case mintKingakuShukeiKoumoku.SHUKEI_KAI_AUTOCHARGE
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_AUTOCHARGE + 2
                End If
            '�y�W�D���@�I�[�g�`���[�W�f�[�^�z
            Case mintKingakuShukeiKoumoku.SHUKEI_SHU_AUTOCHARGE
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_AUTOCHARGE + 3
                End If
            Case Else   '��L�ȊO�͋��z�f�[�^�Ɋւ���ݒ�̂��߃X�L�b�v
                
        End Select
        If intCount <> 0 Then
            intIniIdx = intIniIdx + 1
        End If
        
        '���߂����J�n�ʒu��������A���̍s����Ԃ�
        If intShukeiKoumoku = intCount Then
            pfGetStartLineKingaku = intNowLine
            Exit Function
        End If
    Next
    
    '��L��For�����ő�񐔂܂ŉ���ďI�������Ƃ������Ƃ́A���߂����J�n�ʒu�����߂��Ȃ������B
    pfGetStartLineKingaku = intNowLine

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : pfGetSubGateCsv
'//  �@�\����  : �����⏕���擾
'//  �@�\�T�v  : �w�肵���R�[�i�̎����⏕CSV�t�@�C�����擾����B
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   intCornerNo   �R�[�i�ԍ�
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : Integer                �擾���R�[�h��
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(EG20 V30.3.0.1) 2014-09-19   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_008_01�z
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Function pfGetSubGateCsv(intCornerNo As Integer) As Integer                                            ' EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL
Private Function pfGetSubGateCsv(intCornerNo As Integer, intGokiNo As Integer, intKomoku As Integer) As Integer 'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD

    Dim intFileNumber            As Integer
    Dim i                        As Integer
    Dim ReadBuf                  As JIKAIINFO_IMAGE_FILE    '�ǂݍ��݃o�b�t�@
        
    'Erase ReadSetteiSubGate        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL
    
    '�G���[���[�`����錾
    On Error GoTo Err_handler      'EG20 V30.3.0.1 ADD
    
    '�t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    'CSV�t�@�C���I�[�v��
    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
    
    '��v����R�[�i�ԍ��̃��R�[�h���G���A�ɕۑ����Ă���
    i = 0
    Do While Not EOF(intFileNumber)
                
        Input #intFileNumber, ReadBuf.strBunrui_Dai, ReadBuf.strBunrui_Tyu, _
            ReadBuf.srtBunrui_Sho, ReadBuf.strCorner, ReadBuf.strKomoku, _
            ReadBuf.strKubun, ReadBuf.strData, ReadBuf.strSetShosai
        
        If CInt(ReadBuf.strCorner) = intCornerNo Then
            If CInt(ReadBuf.strBunrui_Tyu) = intGokiNo Then     'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
                If CInt(ReadBuf.srtBunrui_Sho) = intKomoku Then     'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
                    'ReDim Preserve ReadSetteiSubGate(i) As JIKAIINFO_IMAGE_FILE    'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL
                    'ReadSetteiSubGate(i) = ReadBuf                                 'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
                    ReadSetteiSubGate((intGokiNo - 1) * SUBGATE_ITEM_NUM + (intKomoku - 1)) = ReadBuf
                    i = i + 1
                    Exit Do     '���@�A���ڔԍ��ōi�荞�ނ悤�ɂ����̂ŁA�߂�l�ƂȂ郌�R�[�h���͂O���P�ǂ��炩�ɂȂ�B EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
                End If      'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
            End If      'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
        End If
    Loop
    
    'CSV�t�@�C���N���[�Y
    Close #intFileNumber
    pfGetSubGateCsv = i
    
'EG20 V30.3.0.1 ADD START
    Exit Function
Err_handler:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    pfGetSubGateCsv = 0
'EG20 V30.3.0.1 ADD END


End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : pfOutPutSubGate
'//  �@�\����  : �����⏕���o��
'//  �@�\�T�v  : �w�肵���R�[�i�̎����⏕���e���W���[�i���`���ŏo�͂���
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   intCornerNo   �R�[�i�ԍ�
'//              Integer   intFileNumber �t�@�C���ԍ�
'//
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : Integer                �擾���R�[�h��
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(EG20 V30.3.0.1) 2014-09-19   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_003_01�z�A�yHKRK_Kansi07_008_01�z
'//  REVISIONS :(EG30 V32.1.0.1) 2016-06-16   CODED   BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub pfOutPutSubGate(intCornerNo As Integer, intFileNumber As Integer)  'EG20 V30.3.0.1 DEL
Private Function pfOutPutSubGate(intCornerNo As Integer, intFileNumber As Integer) As Boolean   'EG20 V30.3.0.1 ADD
    Dim intTitleFlg             As Integer                  '�����⏕�̑匩�o���̏o�̓t���O
    Dim intSubGateCnt           As Integer                  '�����⏕1�R�[�i���̃��R�[�h��
    Dim i                       As Integer
    Dim intGokiLoop             As Integer                  '���@1�`32 EG20 V30.3.0.1       �yHKRK_Kansi07_008_01�z ADD
    Dim intKomokuLoop           As Integer                  '�����ڇ@�`�E EG20 V30.3.0.1    �yHKRK_Kansi07_008_01�z ADD
    Dim intRet                  As Integer                  ' EG20 V30.3.0.1 ADD
    
    'EG30 V32.1.0.1 ADD START
    Dim bRet                    As Boolean
    Dim lErrCode                As Long
    Dim strEkiSettiBefPath      As String           '���݉w�ݒ�f�[�^�i�ύX�O�ۑ��j
    Dim strGetValue             As String * 64
    Dim strCompValue            As String           '�ݒ�l�i�ύX�O�ۑ��j
    Dim strChangeFlg            As String           '�ύX��
    Dim intValueLen             As Integer          '�擾�����ݒ�l�̒���
    'EG30 V32.1.0.1 ADD END


    '���̃R�[�i�̎����⏕�f�[�^���擾
    intTitleFlg = 0
    'intSubGateCnt = pfGetSubGateCsv(intCornerNo)    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zDEL
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
    'SUB_GATE_KAN.INI����R�[�i���Ȃ��Ȃ������߁A�R�[�i��0�Œ�ō��@�A���ڇ@�`�E�̏���EKI_DATA.CSV���猟��
    intSubGateCnt = 0                       'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
    For intGokiLoop = 0 To 31
        For intKomokuLoop = 0 To 5
            intRet = pfGetSubGateCsv(0, intGokiLoop + 1, intKomokuLoop + 1)
            If intRet = 0 Then
                ' CSV����̎擾������0���̏ꍇ�̓G���[�Ƃ���B
                pfOutPutSubGate = False
                Exit Function
            Else
                intSubGateCnt = intSubGateCnt + intRet
            End If
        Next
    Next
    
    'EG30 V32.1.0.1 ADD START
    '�R�[�i�O�̕ύX�O�ۑ����ꂽ�w�s�x�f�[�^�Ɣ�r����B
    '���̃R�[�i�̕ύX�O�f�[�^�ۑ����ꂽ�f�[�^����������ɓW�J����
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", 0)
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
    For i = 0 To intSubGateCnt - 1
        ' EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
        ' �w�肵���R�[�i�A���@�ɑΉ����郌�R�[�h���o�͂���K�v���Ȃ��Ȃ�A1�`32���@�Œ�ɂȂ�������If�����폜
        'If IsTaisyoGoki(CInt(ReadSetteiSubGate(i).strCorner), CInt(ReadSetteiSubGate(i).strBunrui_Tyu)) = True Then
        ' EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
        If intTitleFlg = 0 Then
            Print #intFileNumber, ""
            'Print #intFileNumber, "�y���D�@�@�ݒu�����@���Ёz����" 'EG30 V32.1.0.1 DEL
            Print #intFileNumber, "�@�y���D�@�@�ݒu�����@���Ёz����"    'EG30 V32.1.0.1 ADD
            intTitleFlg = 1
        End If
        
        'EG30 V32.1.0.1 ADD START
        '�ύX�O�f�[�^�ۑ����ꂽ�ݒ�l�Ɣ�r����
        bRet = dllGetEkiInfoValue(CInt(ReadSetteiSubGate(i).strBunrui_Dai), _
                                    CInt(ReadSetteiSubGate(i).strBunrui_Tyu), _
                                    CInt(ReadSetteiSubGate(i).srtBunrui_Sho), _
                                    0, _
                                    strGetValue, _
                                    intValueLen)
        strCompValue = strGetValue
        If (intValueLen <> 0) Then
            strCompValue = MidByte(strGetValue, 1, intValueLen)
            strCompValue = Trim(strCompValue)
        ElseIf (intValueLen = 0) Then
            strCompValue = "0"
        End If
        
        If (bRet = False) Or (CInt(ReadSetteiSubGate(i).strData) <> CInt(strCompValue)) Then
            strChangeFlg = DIFF_MARK_STRING_ON
        Else
            strChangeFlg = DIFF_MARK_STRING_OFF
        End If
        'EG30 V32.1.0.1 ADD END
        
        'ReadSetteiSubGate(i).strData = Format(ReadSetteiSubGate(i).strData, "0#")      'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
        ReadSetteiSubGate(i).strData = Format(ReadSetteiSubGate(i).strData, "00#")      'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
        Select Case CInt(ReadSetteiSubGate(i).srtBunrui_Sho)
            Case 1
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "FM�� ��Ű�ԍ�" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "FM�� ��Ű�ԍ�" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 2
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "FM�� ���@�ԍ�" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "FM�� ���@�ԍ�" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 3
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "�V����IC ��Ű�ԍ�" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "�V����IC ��Ű�ԍ�" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 4
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "�V����IC ���@�ԍ�" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "�V����IC ���@�ԍ�" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 5
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "NRZ�� ��Ű�ԍ�" & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "NRZ�� ��Ű�ԍ�" & " " & ReadSetteiSubGate(i).strData   'EG30 V32.1.0.1 ADD
            Case 6
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "NRZ�� ���@�ԍ�" & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "���@ " & "NRZ�� ���@�ԍ�" & " " & ReadSetteiSubGate(i).strData   'EG30 V32.1.0.1 ADD
            Case Else
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & ReadSetteiSubGate(i).strKomoku & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & ReadSetteiSubGate(i).strKomoku & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 ADD
        End Select
            
    Next i
    pfOutPutSubGate = True
End Function
'EG20 V30.1.0.1 ADD END
'EG30 V32.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  �֐�����  : pfGetSaveDate
'//  �@�\����  : �ύX�O�f�[�^�ۑ����t�擾����
'//  �@�\�T�v  : �R�[�i���Ƃɕۑ�����Ă���SaveDate.dat�̍X�V���t���擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer    intCorner   �擾����R�[�i�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String     �X�V���t    YYYY�NMM��DD��HH:MM
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-14   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetSaveDate(intCorner As Integer) As String
    Dim strFileName(0 To 1)     As String           '�쐬����
    Dim intCnt                  As Integer          '�J�E���^
    Dim lngHandle               As Long             '�n���h��

    Dim lpCreatTime             As FILETIME         '�쐬����
    Dim lpAccessTime            As FILETIME         '�ŏI�A�N�Z�X����
    Dim lpLastwTime             As FILETIME         '�X�V����
    Dim lpLocalTime             As FILETIME         '���[�J������
    Dim lpSystemTime            As SYSTEMTIME       '�V�X�e������
    Dim bRet                    As Boolean          '�߂�l
    
    Dim strSaveFile             As String
    
    On Error Resume Next

           
    '�ۑ��t�@�C���̓��t���擾
    strSaveFile = PATH_OPERATE & "CORNER" & CStr(intCorner) & "\\SETTEI_BEF\\" & SET_BEF_DATE_FILE
    If Dir(strSaveFile) = "" Then
        pfGetSaveDate = "    �N  ��  ��  :  "
        Exit Function
    Else
        '�t�@�C�����I�[�v��
        lngHandle = CreateFile(strSaveFile, GENERIC_READ, FILE_SHARE_READ, _
                                    0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)

        '�t�@�C���I�[�v��������ɍs��ꂽ���H
        If lngHandle = INVALID_HANDLE_VALUE Then GoTo ErrorHandler
            '�t�@�C���^�C����GET
            bRet = GetFileTime(lngHandle, lpCreatTime, lpAccessTime, lpLastwTime)
            If bRet = False Then GoTo APIError                          '�擾������ɍs��ꂽ���H
        
            '�t�@�C���^�C�������[�J���^�C���ɕϊ�
            bRet = FileTimeToLocalFileTime(lpLastwTime, lpLocalTime)    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
            If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H
        
            '���[�J���^�C�����V�X�e���^�C���ɕϊ�
            bRet = FileTimeToSystemTime(lpLocalTime, lpSystemTime)
            If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H
                
            '�n���h���̃N���[�Y
            Call CloseHandle(lngHandle)
        
            '�쐬���t��\������ (YYYY�NMM��DD��hh:mm)
            pfGetSaveDate = lpSystemTime.wYear & "�N" & _
                                Format(lpSystemTime.wMonth, "00") & "��" & _
                                Format(lpSystemTime.wDay, "00") & "��" & _
                                Format(lpSystemTime.wHour, "00") & ":" & _
                                Format(lpSystemTime.wMinute, "00")
    End If
            
    Exit Function

APIError:
    Call CloseHandle(lngHandle)             '�n���h���̃N���[�Y

ErrorHandler:
    pfGetSaveDate = "    �N  ��  ��  :  "
    
End Function
'EG30 V32.1.0.1 ADD END
