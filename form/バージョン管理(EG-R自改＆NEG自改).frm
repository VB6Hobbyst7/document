VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJVer 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�������D�@�o�[�W�����Ǘ�"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   -210
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   9.75
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
   StartUpPosition =   3  'Windows �̊���l
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "   ���j���[     ��ʂ֖߂�"
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
      Left            =   9360
      TabIndex        =   30
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "�}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   9720
      TabIndex        =   29
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "�e�L�X�g�\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   9720
      TabIndex        =   28
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "�}�̎�O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   27
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdLzhFileCopy 
      Caption         =   " ���k�t�@�C��      �����[�N �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   26
      Top             =   7200
      Width           =   2295
   End
   Begin VB.ListBox lstKan 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Frame fraResource 
      Caption         =   "�\�����\�[�X�w��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   7560
      TabIndex        =   17
      Top             =   720
      Width           =   4215
      Begin VB.OptionButton optSyubetu 
         Caption         =   "����CPU-Pro"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "���C��CPU-Pro"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "�T�uCPU-Pro"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "���C��CPU-OS"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "�\��1"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "�\��2"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�������s �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   5400
      TabIndex        =   14
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "���[�N�����s �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   5400
      TabIndex        =   13
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�}�́����[�N �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2760
      TabIndex        =   12
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "���[�N �N���A"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "�\���X�V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   9720
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "�����؂藣��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   9720
      TabIndex        =   10
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Frame fraVersion 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   7560
      TabIndex        =   16
      Top             =   2520
      Width           =   2055
      Begin VB.CheckBox chkFolder 
         Caption         =   "W ���[�N"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   1  '����
         Width           =   1815
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "O ��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Value           =   1  '����
         Width           =   1815
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "N ���s"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Value           =   1  '����
         Width           =   1815
      End
   End
   Begin VB.Timer tmrMail 
      Left            =   8160
      Top             =   8040
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "NEG�������D�@�o�[�W�����Ǘ�"
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
      TabIndex        =   31
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�R�����g"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   25
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "Ver"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   24
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�쐬����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5190
      TabIndex        =   23
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�t�@�C��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3930
      TabIndex        =   22
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�@�햼"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   21
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "̫���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2040
      TabIndex        =   20
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�t�@�C����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�^�C�v"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "frmJVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmJVer.frm
'//  �p�b�P�[�W���F�o�[�W�����Ǘ�(EG-R����/NEG����)���
'//
'//  �T�v�F�o�[�W�����Ǘ�(EG-R����/NEG����)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�v�����������`�F�b�N�����ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή�
'//                     �E�@�퐳�����`�F�b�N�����ǉ�/�u���[�N�����s�R�s�[�v��
'//                     �E�t�F�[�Y�Q�s��C��
'//                     �E�t�F�[�Y�P�s��C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.11.0.1) 2009-10-23  CODED   BY [TCC] D.Yamashita
'//                 �E�t�F�[�Y�R�c�����ڑΉ�
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 ���̓t�@�C���i�[�f�B���N�g���ʒu�ύX
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(1.16.0.1) 2009-12-21   REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                �@ �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                �A�u���j���[��ʂ֖߂�v�t�����ɂāA
'//                 �@�o�[�W�����Ǘ���ʂ̃o�[�W�����\���X�V���s��
'//                �B�\�����\�[�X���W�I�t�I���Ń��X�g�̕\���X�V
'//                �C���[�N�����s�R�s�[�ł̋@�퐳�����`�F�b�N�ύX
'//                �D���[�N�����s�R�s�[�ł̐������`�F�b�Nini�t�@�C����
'//                �EDir�֐���FileSystemObject�ɒu������
'//                �F�t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(V2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 EG-R(KK)�@������Ή��@KUK�������`�F�b�N�ύX
'//                 �}�̎�O�s��C��
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 �t�@�C�����`�F�b�N�s��C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Dim FolderSyubetu As Integer                 '�I�����\�[�X���

Dim FolderName(0 To 2, 0 To 7) As String     '�t�H���_��
Dim TitleBox(0 To 7) As String               '�^�C�g����
Dim LogBox(0 To 7) As String                 '���O�o�͗p�^�C�g����
Dim FileList() As String                     '�t�@�C�������X�g�ꗗ�i�[�G���A
Dim FileListType() As String                 '�t�@�C�����X�g�ꗗ�i�[�G���A�i�����㎩���^�C�v���܂ށj
Dim uVersion() As MN_VERSION_JIKAI           '�o�[�W�������i�[�G���A

'�I�𒆃��\�[�X��� =0=MN_RSOC_PRO�F�v���O�����A=1=MN_RSOC_HAN:����f�[�^
Dim iSelResource As Integer

Private Const MN_MAIL_INTERVAL = 1000       '���[���^�C�}�̃C���^�[�o���l

Private Const MN_FOLD_WRK = 0               '�u���[�N�v�t�H���_
Private Const MN_FOLD_NOW = 1               '�u���s�v�t�H���_
Private Const MN_FOLD_OLD = 2               '�u���v�t�H���_

'�o�[�W�����f�[�^�t�@�C���p�̍\����
Private Type MN_VERSION_FILE
    sFileName As String * 12                '�t�@�C����
    uFooter As MN_FOOT_BYTE                 '�t�b�^���
End Type

Private Type MN_VERSION_DAT
    strFolder(0 To 5) As String * 8         '�t�H���_��
    intFileNum(0 To 5) As Integer           '�t�@�C����
End Type
'�o�[�W�����f�[�^�t�@�C�����(�o�[�W����2)
Private Type MN_FILE_INFO_V2
    udtInfo As MN_VERSION_DAT               '�t�H���_���ƃt�@�C����
    uFileInfo() As MN_VERSION_FILE          '�t�@�C�����ƃt�b�^���
End Type

Dim uVerdataFile As MN_FILE_INFO_V2

Private Const HANKUKA_KUK = "HAN_KUKA.KUK"
Private Const INI_MAX = 5
Dim HAN_KUKA_DATA As HANTEI_DATA
Private Type HANTEI_DATA
    sHederKisyu(0 To 4) As String
    sHederFile(0 To 4) As String
    sFotterKisyu(0 To 4) As String
    sFotterFile(0 To 4) As String
End Type

'V1.4.0.1�@ADD�@START
Private Const FILE_NAME_MAX_SIZE = 12
Private Const FILE_NAME_SIZE = 19
'�y�^���f�[�^�������`�F�b�N�ُ�X�e�[�^�X��`�z
Private sNGSts As String        'NG�ʒu
Private sNGKoumoku As String    'NG����
'�yNG�ʒu�z
Private Const ERROR_HEDER = "�w�b�_"  '�w�b�_
Private Const ERROR_FOTTER = "�t�b�^" '�t�b�^
'�yNG���ځz
Private Const KISHU_NAME_ERROR = "�@�햼"       '�@�햼
Private Const FILE_NAME_ERRORE = "�t�@�C����"   '�t�@�C����
Private Const CREATE_DATA_ERROR = "�쐬���t"    '�쐬���t
Private Const VERSION_ERROR = "�o�[�W����"      '�o�[�W����
Private sJverName As String                     '�\�����b�Z�[�W�{�b�N�X�^�C�g��
Private Const EGR_JIKAI = "EG-R"                'EG-R
Private Const NEG_JIKAI = "NEG"                 'NEG
'V1.4.0.1�@ADD�@END
'V1.6.0.1 ADD START
Private Const EGR_JIKAI_KISHU = "EG5000"        'EG-R�����@�햼
Private Const NEG_JIKAI_KISHU = "EG2000"        'NEG�����@�햼
'V1.20.0.1 DEL START
'EG-R����
'Private Const EHANTEI_CPU_CHK_FILE = "ko_gateh.vef"
'Private Const EMAIN_CPU_CHK_FILE = "ko_gatep.vef"
'Private Const ESUB_CPU_CHK_FILE = "ko_gatef.vef"
'Private Const EMAIN_OS_CHK_FILE = "ko_gateo.vef"
''NEG����
'Private Const NHANTEI_CPU_CHK_FILE = "KO_GATEH.VEF"
'Private Const NMAIN_CPU_CHK_FILE = "KO_GATEP.VEF"
'Private Const NSUB_CPU_CHK_FILE = "KO_GATEF.VEF"
'Private Const NMAIN_OS_CHK_FILE = "KO_GATEO.VEF"
'V1.20.0.1 DEL END
'V1.20.0.1 ADD START
'EG-R����
Private EHANTEI_CPU_CHK_FILE As String
Private EMAIN_CPU_CHK_FILE As String
Private ESUB_CPU_CHK_FILE As String
Private EMAIN_OS_CHK_FILE As String
'NEG����
Private NHANTEI_CPU_CHK_FILE As String
Private NMAIN_CPU_CHK_FILE As String
Private NSUB_CPU_CHK_FILE As String
Private NMAIN_OS_CHK_FILE As String
'V1.20.0.1 ADD END
'V1.6.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdInstall_Click
'//  �@�\����  : �u�}�̎�O�v�t����������
'//  �@�\�T�v  : �}�̂̎�O�����s��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdInstall_Click()
   On Error Resume Next
   
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����Ǘ�(EG-R/NEG����)���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
   On Error Resume Next
    
    '���[����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����Ǘ�(EG-R/NEG����)���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    
    If blnCabfrmOpenFlg = True Then
        Call fnTsbCabCallDiverge
        Exit Sub
    End If

    '���[����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����Ǘ�(EG-R/NEG����)���(���[�h��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-10   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
   On Error Resume Next
 
    If gStrCurrentForm = sFormName_EJVer Then
       sJverName = EGR_JIKAI                        'V1.4.0.1 ADD
       Label1.Caption = "EG-R�������D�@�o�[�W�����Ǘ�"
       '�uEG-R�������D�@�ް�ޮ݉�ʁF�\���v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EJIKAI_VERASION_KANRI_GAMEN_START, 0)
    Else
       sJverName = NEG_JIKAI                        'V1.4.0.1 ADD
       '�uNEG�������D�@�ް�ޮ݉�ʁF�\���v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, NJIKAI_VERASION_KANRI_GAMEN_START, 0)
    End If
  
    '�f�[�^�W�J
    sSetFolderName

    '�ϐ��̏�����
    FolderSyubetu = 0

    '�o�[�W�������̃��X�g�{�b�N�X���쐬����
    fMakeListbox

    '���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdUpdate_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �e�t�����ɂ�鏈�����s���B
'//             �u���[�N�N���A�v�u�}�́����[�N�R�s�[�v�u���[�N�����s�R�s�[�v
'//             �u�������s�R�s�[�v
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   Index     [IN]�����t�C���f�b�N�X�l
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                  �E�t�F�[�Y�P�s��Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdUpdate_Click(Index As Integer)
   Dim iResponse As Integer         'MsgBox�{�^���R�[�h
   Dim lngErrCode As Long           '�G���[�R�[�h

   On Error Resume Next

' �����ꂽ�{�^���𔻒肷��B
Select Case Index
   Case 0
        '�u���[�N�N���A�v�{�^���̏ꍇ�B
        '�u�����ް�ޮ݊Ǘ��F���[�N�N���A�t�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_CREA_BUTTOM, 0)

        '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
        iResponse = MsgBox("�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "���A" _
               & Chr(vbKeyReturn) & "�S�č폜���܂��B    ��낵���ł����H", _
               vbYesNo + vbExclamation, _
               TitleBox(FolderSyubetu) & "  ���[�N �N���A")
        If iResponse = vbYes Then
            '[�͂�] �{�^����I�������ꍇ
            '���[�N�t�H���_���̃t�@�C�����폜����
           If sWrkFolderRemove <> True Then
              '�u�����ް�ޮ݊Ǘ��F���[�N�N���A�����ُ�v���O�o��
              lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
              Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_CREA_ERROR, lngErrCode)
              Exit Sub
           End If
           '�u�����ް�ޮ݊Ǘ��F���[�N�N���A��������v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_CREA_OK, 0)
           '�o�[�W������񃊃X�g�{�b�N�X���쐬����
           fMakeListbox
        End If
        
   Case 1
        '�u�}�́����[�N�R�s�[�v�{�^���̏ꍇ�B
        '�u�����ް�ޮ݁F�}�́����[�N�R�s�[�t�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_LZH_COPY_WRK_BUTTOM, 0)
        '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����
        sFDInstall "STD"
        
   Case 2
        '�u���[�N�����s�R�s�[�v�{�^���̏ꍇ�B
        '�u�����ް�ޮ݁F���[�N�����s�R�s�[�t�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_COPY_NOW_BUTTOM, 0)
 
        '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
        iResponse = MsgBox("�u���[�N�v�t�H���_�̓��e���A�u���s�v�t�H���_�ɓo�^���邱�Ƃɂ��A" _
                & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & " �̍ŐV�̃o�[�W�������A���s�o�[�W�����Ƃ��܂��B" _
                & Chr(vbKeyReturn) & "��낵���ł����H", _
               vbYesNo + vbExclamation, _
               TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�[")
        If iResponse = vbYes Then
        '[�͂�] �{�^����I�������ꍇ
            'V1.6.0.1�@DEL START
            'If CheckAppStart(PROC_KANRI) <> 0 Then
            '   '�ُ�
            '   If gStrCurrentForm = sFormName_EJVer Then
            '      MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
            '             vbOKOnly + vbExclamation, _
            '             "EG-R�������D�@ �o�[�W�����Ǘ�"
            '      Exit Sub
            '   Else
            '      MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
            '             vbOKOnly + vbExclamation, _
            '             "NEG�������D�@ �o�[�W�����Ǘ�"
            '      Exit Sub
            '   End If
            'End If
            'V1.6.0.1�@DEL END
            '�ŐV�o�[�W���������s�o�[�W�����Ƃ��ēo�^����
            If fNewVersion <> True Then
               '�u�����ް�ޮ݁F���[�N�����s�R�s�[�����ُ�v���O�o��
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
               Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_ERROR, lngErrCode)
               Exit Sub
            End If
            '�u�����ް�ޮ݁F���[�N�����s�R�s�[��������v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_OK, 0)
            '�o�[�W������񃊃X�g�{�b�N�X���쐬����
            fMakeListbox
        End If
        
   Case Else
        '�u�������s�R�s�[�v�{�^���̏ꍇ�B
        '�u�����ް�ޮ݁F�������s�R�s�[�t�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OLD_COPY_NOW_BUTTOM, 0)
       '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
       iResponse = MsgBox("�u���v�t�H���_�̓��e���A�u���s�v�t�H���_�ɖ߂����Ƃɂ��A" _
                & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "�̈ꐢ��O�̃o�[�W�������A" _
                & Chr(vbKeyReturn) & "���s�o�[�W�����Ƃ��܂��B  ��낵���ł����H", _
               vbYesNo + vbExclamation, _
               TitleBox(FolderSyubetu) & "  �������s �R�s�[")
       If iResponse = vbYes Then
       '[�͂�] �{�^����I�������ꍇ
            '�ꐢ��O�̃o�[�W���������s�o�[�W�����ɖ߂�
           If fOldVersion <> True Then
              '�u�����ް�ޮ݁F���[�N�����s�R�s�[�����ُ�v���O�o��
              lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
              Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_ERROR, lngErrCode)
              Exit Sub
            End If
            '�u�����ް�ޮ݁F�������s�R�s�[��������v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_OK, 0)
            '�o�[�W������񃊃X�g�{�b�N�X���쐬����
            fMakeListbox
       End If
       
  End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdLzhFileCopy_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �e�t�����ɂ�鏈�����s���B
'//             �u���k�t�@�C�������[�N�R�s�[�v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdLzhFileCopy_Click()
   
   On Error Resume Next
    
    '�u�����ް�ޮ݁F���ķ�ف�ܰ���߰�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_CAB_COPY_WRK_BUTTOM, 0)
 
    '���k�t�@�C������C���X�g�[������B
    sFDInstall "LZH"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdVer_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �e�t�����ɂ�鏈�����s���B
'//             �u�\���X�V�v�u�e�L�X�g�\���v�u�}�̏o�́v
'//             �u�����؂藣���v
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   Index     [IN]�����t�C���f�b�N�X�l
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdVer_Click(Index As Integer)
    Dim i As Integer                '�J�E���^
    Dim bFlag As Boolean            '�t���O
    Dim lRetVal As Long             '�߂�l
    Dim sCommand As String          '�R�}���h������
    Dim sWriteDir As String
    
    On Error GoTo ErrorHandle

    Select Case Index
        Case 0  '�u�\���X�V�v�t
            '�u�����ް�ޮ݁F�\���X�V�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)

           '�t�H���_�I�𕔂Ɏw��L���`�F�b�N
            bFlag = False                                 '�t���O���u�U�v�ɂ���
            For i = 0 To 2                                '�t�H���_�����J��Ԃ�
               If chkFolder(i).Value = CHECKBOX_ON Then   '�u�H�H�v�t�H���_���w�肳��Ă���
                  bFlag = True                            '�t���O���u�^�v�ɂ���
                  Exit For                                '���[�v�𔲂���
                End If
             Next
              
              If bFlag = False Then                       '�t�H���_�w�薳��
                If gStrCurrentForm = sFormName_EJVer Then
                  '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
                   MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                          vbOKOnly + vbExclamation, _
                          "EG-R�������D�@ �o�[�W�����Ǘ�"
                Else
                  '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
                   MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                          vbOKOnly + vbExclamation, _
                          "NEG�������D�@ �o�[�W�����Ǘ�"
                End If
                '�����𔲂���
                Exit Sub
              End If
              '�o�[�W������񃊃X�g�{�b�N�X���쐬����
              fMakeListbox
              
        Case 1 '�u�e�L�X�g�\���v�t
            '�u�����ް�ޮ݁F�e�L�X�g�\���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_TXT_BUTTOM, 0)

            '�t�H���_�I�𕔂Ɏw��L���`�F�b�N
            bFlag = False                                 '�t���O���u�U�v�ɂ���
            For i = 0 To 2                                '�t�H���_�����J��Ԃ�
               If chkFolder(i).Value = CHECKBOX_ON Then   '�u�H�H�v�t�H���_���w�肳��Ă���
                  bFlag = True                            '�t���O���u�^�v�ɂ���
                  Exit For                                '���[�v�𔲂���
               End If
            Next
                        
            If bFlag = False Then                       '�t�H���_�w�薳��
               If gStrCurrentForm = sFormName_EJVer Then
                 '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
                   MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                           vbOKOnly + vbExclamation, _
                           "EG-R�������D�@ �o�[�W�����Ǘ�"
               Else
                 '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
                   MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                          vbOKOnly + vbExclamation, _
                          "NEG�������D�@ �o�[�W�����Ǘ�"
               End If
                   '�����𔲂���
               Exit Sub
             End If


            '���X�g�{�b�N�X�̓��e���t�@�C���ɏ�������
            sWriteListbox
            sCommand = MN_EXE_MEMO & MN_VERSI_FILE '���������s�R�}���h���쐬����
            '���������N������
            lRetVal = Shell(sCommand, vbMaximizedFocus)
            '���������A�N�e�B�u�i�O�ʕ\���j�ɂ���
            AppActivate lRetVal, True
            SendKeys "{LEFT}", True
           
        Case 2 '�u�}�̏o�́v�t
            '�u�����ް�ޮ݁F�}�̏o�͖t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OUTPUT_BUTTOM, 0)
 
            '�}�̏o�͏���
             fMakeOutPutFile
           
        Case 3  '�u�����؂藣���v�t
            '�u�����ް�ޮ݁F�����؂藣���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_KIRIHANASI_BUTTOM, 0)

            '�ʐM�ڑ��E�ؒf��ʂ�\������B
            Load frmConectSts
            frmConectSts.Show 1
        Case Else
   End Select
ErrorHandle:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optSyubetu_Click
'//  �@�\����  : ���\�[�X�t����������
'//  �@�\�T�v  : �Ώۃ��\�[�X��ʂ�ύX����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                �\�����\�[�X���W�I�t�I���Ń��X�g�̕\���X�V
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub optSyubetu_Click(Index As Integer)
    'V1.20.0.1 ADD START
    Dim i As Integer                '�J�E���^
    Dim bFlag As Boolean            '�t���O
    'V1.20.0.1 ADD END

    '���\�[�X��ʂ�ύX����B'
    FolderSyubetu = Index
    
    'V1.20.0.1 ADD START
    '�t�H���_�I�𕔂Ɏw��L���`�F�b�N
    bFlag = False                                 '�t���O���u�U�v�ɂ���
    For i = 0 To 2                                '�t�H���_�����J��Ԃ�
        If chkFolder(i).Value = CHECKBOX_ON Then   '�u�H�H�v�t�H���_���w�肳��Ă���
            bFlag = True                            '�t���O���u�^�v�ɂ���
            Exit For                                '���[�v�𔲂���
        End If
    Next
    
    If bFlag = False Then                       '�t�H���_�w�薳��
        If gStrCurrentForm = sFormName_EJVer Then
            '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
            MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                    vbOKOnly + vbExclamation, _
                    "EG-R�������D�@ �o�[�W�����Ǘ�"
        Else
            '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
            MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                    vbOKOnly + vbExclamation, _
                    "NEG�������D�@ �o�[�W�����Ǘ�"
        End If
        '�����𔲂���
        Exit Sub
    End If
    
    '�o�[�W������񃊃X�g�{�b�N�X���쐬����
    fMakeListbox
    'V1.20.0.1 ADD END

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                �A�u���j���[��ʂ֖߂�v�t�����ɂāA
'//                 �@�o�[�W�����Ǘ���ʂ̃o�[�W�����\���X�V���s��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
   
   On Error Resume Next
    
   If gStrCurrentForm = sFormName_EJVer Then
      '�uEG-R�������D�@�ް�ޮ݉�ʁF�����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EJIKAI_VERASION_KANRI_GAMEN_END, 0)
   Else
      '�uNEG�������D�@�ް�ޮ݉�ʁF�����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, NJIKAI_VERASION_KANRI_GAMEN_END, 0)
   End If
   
   'V1.20.0.1 ADD START
   '�o�[�W�����Ǘ���ʂ̃o�[�W�����\���X�V�������s���B
   frmVersion.psGetVersion
   'V1.20.0.1 ADD END
   
   'NEG/EG-R�������D�@�o�[�W�����Ǘ���ʂ����
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeListbox
'//  �@�\����  : �o�[�W������񃊃X�g�{�b�N�X�쐬
'//  �@�\�T�v  : �e�t�H���_����o�[�W�����擾���s���A���X�g�{�b�N�X�쐬
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeListbox() As Boolean
    Dim bRet As Boolean                        '�߂�l

    On Error Resume Next

    '***********************************************
    '* �����㎩���t�H���_����S�Ẵo�[�W���������擾���� *
    '***********************************************
    ReDim uVersion(0)

    '����[�N��t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = True Then
        '�t�@�C�����X�g����o�[�W���������擾����
        sVersionInfo FolderName(0, FolderSyubetu), MN_FLDWRK
    End If

    '����s��t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(FolderName(1, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = True Then
        '�t�@�C�����X�g����o�[�W���������擾����
        sVersionInfo FolderName(1, FolderSyubetu), MN_FLDNOW
    End If

    '�����t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(FolderName(2, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = True Then
        '�t�@�C�����X�g����o�[�W���������擾����
        sVersionInfo FolderName(2, FolderSyubetu), MN_FLDOLD
    End If

    '�o�[�W���������t�@�C�������Ƀ\�[�g����
    sListboxSort

    '�o�[�W�����������X�g�{�b�N�X�ɃZ�b�g����
    sVerListDisp
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sVerListDisp
'//  �@�\����  : �o�[�W������񃊃X�g�{�b�N�X�ݒ�
'//  �@�\�T�v  : �擾�����o�[�W���������A���X�g�{�b�N�X�ɐݒ�
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sVerListDisp()
    Dim i As Integer                        '�J�E���^
    Dim uVerData(2) As MN_VERSION_JIKAI     '�o�[�W�������i�e�t�H���_�j
    Dim lDataNum As Long                    '�o�[�W�������

    On Error Resume Next

    '���X�g�{�b�N�X������������
    lstKan.Clear

    lDataNum = UBound(uVersion)             '�o�[�W������񐔎擾
    For i = 1 To lDataNum

        uVerData(0).sFileName = ""          '�t�@�C�������N���A����
        uVerData(1).sFileName = ""          '�t�@�C�������N���A����
        uVerData(2).sFileName = ""          '�t�@�C�������N���A����

        Select Case uVersion(i).iFolder     '�t�H���_����ΏۂƂ���
        Case MN_FLDWRK                      '�u���[�N�v�t�H���_�̏ꍇ
            uVerData(0) = uVersion(i)       '�u���[�N�v�t�H���_���Ɋi�[����
            If i + 1 <= lDataNum Then       '���̃f�[�^������?
                If uVersion(i).sFileName = uVersion(i + 1).sFileName Then
                                                        '�t�@�C����������?
                    Select Case uVersion(i + 1).iFolder '�t�H���_����ΏۂƂ���
                    Case MN_FLDNOW                      '�u���s�v�t�H���_�̏ꍇ
                        uVerData(1) = uVersion(i + 1)   '�u���s�v�t�H���_���Ɋi�[����
                        If i + 2 <= lDataNum Then       '���̃f�[�^������?
                            If uVersion(i + 1).sFileName = uVersion(i + 2).sFileName Then
                                                        '�t�@�C����������?
                                uVerData(2) = uVersion(i + 2)
                                                        '�u���v�t�H���_���Ɋi�[����
                                i = i + 2               '�J�E���^�����X�ɂ���
                            Else
                                i = i + 1               '�J�E���^�����ɂ���
                            End If
                        Else
                            i = i + 1                   '�J�E���^�����ɂ���
                        End If
                    Case MN_FLDOLD                      '�u���v�t�H���_�̏ꍇ
                        uVerData(2) = uVersion(i + 1)   '�u���v�t�H���_���Ɋi�[����
                        i = i + 1                       '�J�E���^�����ɂ���
                    End Select
                End If
            End If
        Case MN_FLDNOW                      '�u���s�v�t�H���_�̏ꍇ
            uVerData(1) = uVersion(i)       '�u���s�v�t�H���_���Ɋi�[����
            If i + 1 <= lDataNum Then       '���̃f�[�^������
                If uVersion(i).sFileName = uVersion(i + 1).sFileName Then
                                                    '�t�@�C����������?
                    uVerData(2) = uVersion(i + 1)   '�u���v�t�H���_���Ɋi�[����
                    i = i + 1                       '�J�E���^�����ɂ���
                End If
            End If
        Case MN_FLDOLD                      '�u���v�t�H���_�̏ꍇ
            uVerData(2) = uVersion(i)       '�u���v�t�H���_���Ɋi�[����
        End Select
        '�t�@�C�������܂Ƃ߂ă��X�g�{�b�N�X�ɐݒ�
        sVersionDisp uVerData()
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sListboxSort
'//  �@�\����  : �o�[�W�������̃\�[�g
'//  �@�\�T�v  : �o�[�W���������t�@�C�������Ƀ\�[�g����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sListboxSort()
    Dim i As Integer                '�J�E���^
    Dim j As Integer                '�J�E���^
    Dim uBuff As MN_VERSION_JIKAI   '�o�[�W�������i�[�o�b�t�@

    On Error Resume Next
   
    For i = 1 To UBound(uVersion) - 1
        For j = i + 1 To UBound(uVersion)
            '�t�@�C�����̔�r���s��
            If uVersion(j).sFileName < uVersion(i).sFileName Then
                '�t�@�C��������������Έڂ��ւ���
                uBuff = uVersion(i)
                uVersion(i) = uVersion(j)
                uVersion(j) = uBuff
            ElseIf uVersion(j).sFileName = uVersion(i).sFileName Then
                '�t�H���_�̔�r���s��
                If uVersion(j).iFolder = MN_FLDWRK And uVersion(i).iFolder = MN_FLDNOW Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                ElseIf uVersion(j).iFolder = MN_FLDNOW And uVersion(i).iFolder = MN_FLDOLD Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                ElseIf uVersion(j).iFolder = MN_FLDWRK And uVersion(i).iFolder = MN_FLDOLD Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                End If
            End If
        Next
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sVersionDisp
'//  �@�\����  : �o�[�W������񃊃X�g�{�b�N�X�ǉ�
'//  �@�\�T�v  : �o�[�W���������t�@�C�����P�ʂŃ��X�g�{�b�N�X�ɒǉ�����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sVersionDisp(uVerData() As MN_VERSION_JIKAI)
    Dim sFileName As String         '�t�@�C����������i�����㎩���^�C�v���܂ށj
    Dim sFileSize As String         '�t�@�C���T�C�Y������
    Dim sFileInfo(2) As String      '�o�[�W������񕶎���
    Dim sComment1(2) As String      '�R�����g������
    Dim sComment2(2) As String      '�R�����g������

   On Error Resume Next
    
    If uVerData(0).sFileName <> "" Then     '�u���[�N�v�t�H���_�Ƀt�@�C��������
        '�t�@�C�����i�[
        sFileName = StrConv(MidB(StrConv(uVerData(0).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    ElseIf uVerData(1).sFileName <> "" Then '�u���s�v�t�H���_�Ƀt�@�C��������
        '�t�@�C�����i�[
        sFileName = StrConv(MidB(StrConv(uVerData(1).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    Else                                    '�u���v�t�H���_�Ƀt�@�C��������
        '�t�@�C�����i�[
        sFileName = StrConv(MidB(StrConv(uVerData(2).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    End If
    sFileName = sFileName & " "

    If uVerData(0).sFileName <> "" Then     '�u���[�N�v�t�H���_�Ƀt�@�C��������
        '�o�[�W�������i�[
        sFileInfo(0) = " " & StrConv(MidB(StrConv(uVerData(0).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
        sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
        sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
        sFileInfo(0) = sFileInfo(0) & uVerData(0).sVersion
        sComment1(0) = " " & StrConv(MidB(StrConv(uVerData(0).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(0) = " " & StrConv(MidB(StrConv(uVerData(0).sComment, vbFromUnicode), 33, 32), vbUnicode)
    End If
    If uVerData(1).sFileName <> "" Then     '�u���s�t�H���_�Ƀt�@�C��������
        '�o�[�W�������i�[
        sFileInfo(1) = " " & StrConv(MidB(StrConv(uVerData(1).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
        sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
        sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
        sFileInfo(1) = sFileInfo(1) & uVerData(1).sVersion
        sComment1(1) = " " & StrConv(MidB(StrConv(uVerData(1).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(1) = " " & StrConv(MidB(StrConv(uVerData(1).sComment, vbFromUnicode), 33, 32), vbUnicode)
    End If
    If uVerData(2).sFileName <> "" Then     '�u���v�t�H���_�Ƀt�@�C��������
        '�o�[�W�������i�[
        sFileInfo(2) = " " & StrConv(MidB(StrConv(uVerData(2).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
        sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
        sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
        sFileInfo(2) = sFileInfo(2) & uVerData(2).sVersion
        sComment1(2) = " " & StrConv(MidB(StrConv(uVerData(2).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(2) = " " & StrConv(MidB(StrConv(uVerData(2).sComment, vbFromUnicode), 33, 32), vbUnicode)
    End If


    If chkFolder(0).Value = CHECKBOX_ON Then               '����[�N��t�H���_�\��
        If uVerData(0).sFileName <> "" Then         '����[�N��t�H���_�Ƀt�@�C���͂���
            If chkFolder(1).Value = CHECKBOX_ON Then       '����s��t�H���_�\��
                If uVerData(1).sFileName <> "" Then '����s��t�H���_�Ƀt�@�C���͂���
                    '����[�N��t�H���_�Ƣ���s��t�H���_���r����
                    If sFileInfo(0) = sFileInfo(1) Then
                        If chkFolder(2).Value = CHECKBOX_ON Then   '�����t�H���_�\��
                            If uVerData(2).sFileName <> "" Then
                                '����s��t�H���_�Ƣ����t�H���_���r����
                                If sFileInfo(1) = sFileInfo(2) Then
                                    lstKan.AddItem sFileName & "W N O" & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(0)
                                    End If
                                Else
                                    lstKan.AddItem sFileName & "W N  " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(0)
                                    End If
                                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(0)
                                    End If
                                    lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(2)
                                    End If
                                    If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '�����t�H���_�Ƀt�@�C�����Ȃ�
                                lstKan.AddItem sFileName & "W N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                     lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                            End If
                        Else                                '�����t�H���_��A�N�e�B�u�\��
                            lstKan.AddItem sFileName & "W N  " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                        End If
                    Else                            '����[�N��t�H���_�Ƣ���s��t�H���_�̃o�[�W�������Ⴄ
                        lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(0)
                        End If
                        If chkFolder(2).Value = CHECKBOX_ON Then   '�����t�H���_�\��
                            If uVerData(2).sFileName <> "" Then
                                '����s��t�H���_�Ƣ����t�H���_���r����
                                If sFileInfo(1) = sFileInfo(2) Then
                                    lstKan.AddItem Space(17) & "  N O" & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(1)
                                    End If
                                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(1)
                                    End If
                                Else
                                    lstKan.AddItem Space(17) & "  N  " & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(1)
                                    End If
                                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(1)
                                    End If
                                    lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment1(2)
                                    End If
                                    If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                        lstKan.AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '�����t�H���_�Ƀt�@�C�����Ȃ�
                                lstKan.AddItem Space(17) & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                            End If
                        Else
                            lstKan.AddItem Space(17) & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                        End If
                    End If
                Else                                    '����s��t�H���_�Ƀt�@�C�����Ȃ�
                    If chkFolder(2).Value = CHECKBOX_ON Then   '�����t�H���_�\��
                        If uVerData(2).sFileName <> "" Then
                            If sFileInfo(0) = sFileInfo(2) Then
                                lstKan.AddItem sFileName & "W   O" & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(0)
                                End If
                                If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(0)
                                End If
                                lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                            Else
                                lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(0)
                                End If
                                If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(0)
                                End If
                                lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(2)
                                End If
                                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(2)
                                End If
                                lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                            End If
                        Else                            '�����t�H���_�Ƀt�@�C�����Ȃ�
                            lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                            lstKan.AddItem Space(17) & "  N O" & " -------- --------  -------- ----"
                        End If
                    Else                                '�����t�H���_��A�N�e�B�u�\��
                        lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(0)
                        End If
                        lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                    End If
                End If
            Else                                        '����s��t�H���_��A�N�e�B�u�\��
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then
                        If sFileInfo(0) = sFileInfo(2) Then
                            lstKan.AddItem sFileName & "W   O" & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                        Else
                            lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(0)
                            End If
                            If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(0)
                            End If
                            lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                    '�����t�H���_�Ƀt�@�C�����Ȃ�
                        lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(0)
                        End If
                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(0)
                        End If
                        lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
                    lstKan.AddItem sFileName & "W    " & sFileInfo(0)
                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                        lstKan.AddItem Space(22) & sComment1(0)
                    End If
                    If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                        lstKan.AddItem Space(22) & sComment2(0)
                    End If
                End If
            End If
        Else                                '����[�N��t�H���_�Ƀt�@�C�����Ȃ�
            If chkFolder(1).Value = CHECKBOX_ON Then               '����s��t�H���_�\��
                If uVerData(1).sFileName <> "" Then         '����s��t�H���_�Ƀt�@�C���͂���
                    If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                        If uVerData(2).sFileName <> "" Then '�����t�H���_�Ƀt�@�C���͂���
                            '����s��t�H���_�Ƣ����t�H���_���r����
                            If sFileInfo(1) = sFileInfo(2) Then
                                lstKan.AddItem sFileName & "  N O" & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                            Else                            '����s��t�H���_�Ƣ����t�H���_�̃o�[�W�������Ⴄ
                                lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(1)
                                End If
                                If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(1)
                                End If
                                lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment1(2)
                                End If
                                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                    lstKan.AddItem Space(22) & sComment2(2)
                                End If
                                lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                            End If
                        Else                                '�����t�H���_�Ƀt�@�C���͂Ȃ�
                            lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                            lstKan.AddItem Space(17) & "W   O" & " -------- --------  -------- ----"
                        End If
                    Else                                    '�����t�H���_��A�N�e�B�u�\��
                        lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(1)
                        End If
                        If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(1)
                        End If
                        lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                    End If
                Else                                        '����s��t�H���_�Ƀt�@�C�����Ȃ�
                    If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                        If uVerData(2).sFileName <> "" Then
                            lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(2)
                            End If
                            lstKan.AddItem Space(17) & "W N  " & " -------- --------  -------- ----"
                        Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
                            lstKan.AddItem sFileName & "W N O" & " -------- --------  -------- ----"
                        End If
                    Else                                    '�����t�H���_��A�N�e�B�u�\��
                        lstKan.AddItem sFileName & "W N  " & " -------- --------  -------- ----"
                    End If
                End If
            Else                                        '����s��t�H���_��A�N�e�B�u�\��
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then
                        lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(2)
                        End If
                        If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(2)
                        End If
                        lstKan.AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                    Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
                        lstKan.AddItem sFileName & "W   O" & " -------- --------  -------- ----"
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
                    lstKan.AddItem sFileName & "W    " & " -------- --------  -------- ----"
                End If
            End If
        End If
    Else                                                '����[�N��t�H���_��A�N�e�B�u�\��
        If chkFolder(1).Value = CHECKBOX_ON Then               '����s��t�H���_�\��
            If uVerData(1).sFileName <> "" Then         '����s��t�H���_�Ƀt�@�C���͂���
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then '�����t�H���_�Ƀt�@�C���͂���
                        '����s��t�H���_�Ƣ����t�H���_���r����
                        If sFileInfo(1) = sFileInfo(2) Then
                            lstKan.AddItem sFileName & "  N O" & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                        Else
                            lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(1)
                            End If
                            If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(1)
                            End If
                            lstKan.AddItem Space(17) & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment1(2)
                            End If
                            If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                lstKan.AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                '�����t�H���_�Ƀt�@�C���͂Ȃ�
                        lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(1)
                        End If
                        If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(1)
                        End If
                        lstKan.AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
                    lstKan.AddItem sFileName & "  N  " & sFileInfo(1)
                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                        lstKan.AddItem Space(22) & sComment1(1)
                    End If
                    If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                        lstKan.AddItem Space(22) & sComment2(1)
                    End If
                End If
            Else                                        '����s��t�H���_�Ƀt�@�C�����Ȃ�
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then
                        lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment1(2)
                        End If
                        If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                            lstKan.AddItem Space(22) & sComment2(2)
                        End If
                        lstKan.AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                    Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
                        lstKan.AddItem sFileName & "  N O" & " -------- --------  -------- ----"
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
                    lstKan.AddItem sFileName & "  N  " & " -------- --------  -------- ----"
                End If
            End If
        Else                                    '����s��t�H���_��A�N�e�B�u�\��
            If uVerData(2).sFileName <> "" Then '�����t�H���_�Ƀt�@�C���͂���
                lstKan.AddItem sFileName & "    O" & sFileInfo(2)
                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
                    lstKan.AddItem Space(22) & sComment1(2)
                End If
                If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                    lstKan.AddItem Space(22) & sComment2(2)
                End If
            Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
                lstKan.AddItem sFileName & "    O" & " -------- --------  -------- ----"
            End If
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sVersionInfo
'//  �@�\����  : �o�[�W�������̎擾
'//  �@�\�T�v  : �t�@�C�����X�g�ꗗ����o�[�W���������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sPath
'//  �@�@�@    : Integer�@ iFolder
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sVersionInfo(sPath As String, iFolder As Integer)
    Dim i As Integer                    '�J�E���^
    Dim j As Integer                    '�J�E���^
    Dim sMyName As String               '�t�@�C����
    Dim iFileNumber As Integer          '�t�@�C���ԍ�
    Dim lLen As Long                    '�t�@�C���T�C�Y
    Dim uFooter As MN_FOOT              '�t�b�^���i�[�G���A
    Dim lPos As Long                    '�o�[�W�������i�[�ʒu
    Dim sDateTime As String
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD

   On Error Resume Next

    For i = 0 To UBound(FileList) - 1   '�t�@�C�����X�g��

        sMyName = sPath & "\" & FileList(i)     '�t�@�C���t���p�X���̍쐬

        'If Dir(sMyName) <> "" Then              '�t�@�C�������݂���?    'V1.20.0.1 DEL
        If objFso.FileExists(sMyName) = True Then  '�t�@�C�������݂���?    'V1.20.0.1 ADD
            lLen = FileLen(sMyName)             '�t�@�C���T�C�Y�̎擾

            iFileNumber = FreeFile              '���g�p�̃t�@�C���ԍ����擾����

            Open sMyName For Binary Access Read As #iFileNumber
                                                '�t�@�C���̃I�[�v��
            Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
                                                '�t�b�^���̎擾
            ReDim Preserve uVersion(UBound(uVersion) + 1)
                                                '�o�[�W�������i�[�G���A�̊g��
            lPos = UBound(uVersion)             '�o�[�W�������i�[�ʒu�̎擾
            uVersion(lPos).sFileName = UCase(FileListType(i))       '�t�@�C������啶���ɂ��ăZ�b�g
            uVersion(lPos).iFolder = iFolder                    '�t�H���_���Z�b�g
            uVersion(lPos).sMachineName = uFooter.sKisyu        '�@�햼�Z�b�g
            uVersion(lPos).sFooterFile = uFooter.sFileName      '�t�@�C�����Z�b�g

            sDateTime = ""
            For j = 0 To 3
                sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
            Next
            sDateTime = sDateTime & " "
            For j = 4 To 5
                sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
            Next
            uVersion(lPos).sFileDate = sDateTime
            uVersion(lPos).sVersion = uFooter.sVersion          '�o�[�W�������Z�b�g
            uVersion(lPos).sComment = uFooter.sHyoji            '�\��������Z�b�g

            Close #iFileNumber                  '�t�@�C������܂�
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fNewVersion
'//  �@�\����  : �ŐV�o�[�W��������
'//  �@�\�T�v  : �ŐV(���[�N)�o�[�W�������A���s(���s)�o�[�W�����ɓo�^
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sPath
'//  �@�@�@    : Integer�@ iFolder
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�v�����������`�F�b�N�����ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�Ή��s��C��
'//                 �t�F�[�Y�R�Ή��@�@�퐳�����`�F�b�N�����ǉ�
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fNewVersion() As Boolean
    Dim bRet As Boolean                      '�߂�l
    Dim lngCnt                  As Long      '�J�E���^�[
    Dim sSrcFileName            As String    '���[�N�t�H���_���t�@�C�����X�g
    Dim sFileName As String
    Dim lngErrCode As Long                   '�G���[�R�[�h
    'V1.4.0.1 ADD START
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim myLen As Long                        '������̒���
    Dim SysCodeTxt As String                 '�o�C�g�ϊ���(�S�p�����p)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    'V1.4.0.1 ADD END
    Dim iKansiAplChk As Integer              '�A�v���N���`�F�b�N�߂�l�@'V1.6.0.1 ADD
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error Resume Next
    
    '����[�N��t�H���_�̃t�@�C�����X�g����������
    '���[�N�t�H���_���t�@�C�������쐬
    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
    '�t�@�C���̌���������
    'If Dir(sSrcFileName) <> "" Then     'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
      Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else
      '�t�@�C�������݂��Ȃ�
      MsgBox "�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
             & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
             vbOKOnly + vbExclamation, _
             TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
     '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      fNewVersion = False
      Set objFso = Nothing    'V1.20.0.1 ADD
      Exit Function
    End If
  
    '����[�N��t�H���_����t�@�C�����X�g���擾����
    'bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)�@'V1.8.0.1 DEL
    
    bRet = pfSeitouseiChck    'V1.4.0.1�@ADD
    '�����v���O��������f�[�^�������`�F�b�N���s��(�Ώۃt�@�C���FHAN_KUKA.KUK)
'    bRet = fDataFileCheck(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST) 'V1.4.0.1�@DEL
'V1.8.0.1 ADD START
    '����[�N��t�H���_����t�@�C�����X�g���A�o�^�t�@�C�������J�E���g����
    If bRet = True Then
       bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    End If
'V1.8.0.1 ADD END

  If bRet = True Then
    '�����t�H���_���̃t�@�C����S�č폜����
     If sOldFolderRemove <> True Then
         fNewVersion = False
         Exit Function
     End If

    '�����t�H���_���̃t�@�C������s��t�H���_�̓��e�ɒu������
    If sCopyNOWtoOLD <> True Then
        fNewVersion = False
        Exit Function
    End If

    '����s��t�H���_���̃t�@�C���𢃏�[�N��t�H���_�̓��e�ɒu������
    If sCopyWRKtoNOW <> True Then
        fNewVersion = False
        Exit Function
    End If
    
 
    '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
    'V1.6.0.1�@ADD�@START
    '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
    'V1.6.0.1 ADD END
      If gStrCurrentForm = sFormName_EJVer Then
         psVersionUpdateReqest (ML_REQUEST_EGATE)
      Else
         psVersionUpdateReqest (ML_REQUEST_NGATE)
      End If
    'V1.6.0.1 ADD START
    Else
        '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
        gintGateVerInfUpdRes = MailSts.stsNormal
    End If
    'V1.6.0.1 ADD END
    
    '���D�@�o�[�W�����X�V��������
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
        '����
        MsgBox "�u���[�N�v�t�H���_�̓��e��,�u���s�v�t�H���_�ɓo�^���āA" _
                & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & " �̍ŐV�̃o�[�W�����Ƃ��܂����B", _
                vbOKOnly + vbExclamation, _
                TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
        fNewVersion = True
    Else
        '�ُ�
        If gStrCurrentForm = sFormName_EJVer Then
           MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
                  vbOKOnly + vbExclamation, _
                  "EG-R�������D�@ �o�[�W�����Ǘ�"
        Else
         MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
                 vbOKOnly + vbExclamation, _
                 "NEG�������D�@ �o�[�W�����Ǘ�"
        End If
        
        fNewVersion = False
    End If
  
    fNewVersion = True
  Else
    fNewVersion = False
  End If
End Function
'V1.4.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfSeitouseiChck
'//  �@�\����  : �v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �v���O��������f�[�^�������`�F�b�N�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�v�����������`�F�b�N����
'//     REVISIONS :(1.6.0.1) 2009-06-16  REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��s��C��
'//                 �t�F�[�Y�R�Ή��@�@�퐳�����`�F�b�N�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfSeitouseiChck() As Boolean
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim myLen As Long                        '������̒���
    Dim SysCodeTxt As String                 '�o�C�g�ϊ���(�S�p�����p)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    Dim bRet As Boolean
    Dim lngCnt As Long
    Dim iRet   As Integer                    '�o�[�W�����`�F�b�NDLL�߂�l
    Dim iGouki As Integer                    '���@�ԍ�
    Dim sVersionInfoPath As String           '�o�[�W�������t�@�C��(���@��)
    Dim iCnt             As Integer          '���@�J�E���^�[�@V1.6.0.1�@ADD
    On Error Resume Next
    
    pfSeitouseiChck = True
    
    '********************************
    '*�v�����������`�F�b�N
    '********************************
    '�����v���O��������f�[�^�������`�F�b�N���s��(�Ώۃt�@�C���FHAN_KUKA.KUK)
    bRet = fDataFileCheck(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    If bRet = False Then
       If sNGSts <> "" And sNGKoumoku <> "" Then
          MsgBox "�^���f�[�^�������`�F�b�N�ُ�(" & sNGSts & "�F" & sNGKoumoku & "�j", _
                 vbOKOnly + vbExclamation, _
                 sJverName & "�������D�@ �o�[�W�����Ǘ�"
       Else
          MsgBox "�ُ�I�����܂����B", _
                 vbOKOnly + vbExclamation, _
                 TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
       End If
       pfSeitouseiChck = False
       Exit Function
    End If

'V1.6.0.1 DEL START
'    '�T���l�`�F�b�N
'    For lngCnt = 0 To UBound(FileList) - 1
'        If pfFileSumChk(FolderName(0, FolderSyubetu) & "\" & FileList(lngCnt), lngSumRet) <> True Then
'            '�T���l�ُ�
'            If lngSumRet = SUM_CHK.SumErr Then
'               MsgBox "�T���l���ُ�ł��B" _
'                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                      vbOKOnly + vbExclamation, _
'                      sJverName & "�������D�@ �o�[�W�����Ǘ�"
'            '�T���l�ُ�ȊO�ُ�
'            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
'               MsgBox "�ُ�I�����܂����B", _
'                     vbOKOnly + vbExclamation, _
'                     TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
'            End If
'            pfSeitouseiChck = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
'            Exit Function
'        End If
'    Next
'
'    '�t�@�C�����ő�`�F�b�N
'    If UBound(FileList) > FILECNT_MAX Then
'       MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
'              & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'              vbOKOnly + vbExclamation, _
'              sJverName & "�������D�@ �o�[�W�����Ǘ�"
'      pfSeitouseiChck = False
'
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
'
'      Exit Function
'    End If
'
'    '�t�@�C�����T�C�Y�`�F�b�N
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
'    '�t�@�C�����X�g���I�[�v���B
'    Open FolderName(0, FolderSyubetu) & "\" & MN_FILELIST For Input As #iFileNumber
'    For i = 0 To lngFileListCnt
'       If i = lngFileListCnt Then
'          Exit For
'       End If
'       '�t�@�C�������擾����B
'       Input #iFileNumber, strWork
'       '�t�@�C������`�Ȃ�
'       If strWork = "" Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'          bRet = False
'          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'          Exit For
'       '�t�H�[�}�b�g�ُ�
'       ElseIf " " <> Mid(strWork, 2, 1) Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       ElseIf (InStr(strWork, ".") - 1) = -1 Then
'           MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       Else
'          '�t�@�C�����݂̂𒊏o
'          sGetFileListName = Mid(strWork, 3, 16)
'          '�擾�t�@�C�����̃T�C�Y���擾
'          myLen = LenB(StrConv(sGetFileListName, vbFromUnicode))                                              '���p���Z�̃o�C�g�����擾
'          If FILE_NAME_MAX_SIZE < myLen Then
'            '13�o�C�g�ȏ�̏ꍇ
'            MsgBox "�t�@�C�������ُ�ł��B" _
'                   & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                   vbOKOnly + vbExclamation, _
'                   sJverName & "�������D�@ �o�[�W�����Ǘ�"
'            bRet = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'            Exit For
'           End If
'        End If
'     Next
'    '�t�@�C�����X�g���N���[�Y�B
'    Close #iFileNumber
'V1.6.0.1 DEL END
'V1.11.0.1 DEL START
'    If gStrCurrentForm = sFormName_EJVer Then
''V1.6.0.1 ADD�@START
'   For iCnt = 1 To MAX_GATE_NO
'      'EG-R�����̂݁F�����o�[�W�����`�F�b�NDLL����
'      iGouki = pfGetGoukiNo(iCnt)
'      If iGouki <> 0 Then
''V1.6.0.1 ADD�@END
'       'iGouki = pfGetGoukiNo 'V1.6.0.1 DEL
'       sVersionInfoPath = Replace(GATE_VERSION_INFO_FILE, "##", Format(iGouki, "0#"))
'
'       'iRet = dllVerChk(E_EPRO1WRK & "\\" & GATE_VERSION_KANRI_FILE, PATH_GATE & sVersionInfoPath, PATH_HOSHU_LOG & GATE_VERSION_NGLIST_FILE)�@�@�@�@�@�@�@�@�@'V1.6.0.1�@DEL
'       iRet = dllVerChk(FolderName(0, FolderSyubetu) & "\" & GATE_VERSION_KANRI_FILE, PATH_GATE & sVersionInfoPath, PATH_HOSHU_LOG & GATE_VERSION_NGLIST_FILE)  'V1.6.0.1�@ADD
'       If iRet = 1 Then
'          bRet = True
'       Else
'          bRet = False
'          MsgBox "�ُ�I�����܂����B", _
'                 vbOKOnly + vbExclamation, _
'                 TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
'          'V1.6.0.1 ADD START
'           pfSeitouseiChck = False
'           Exit Function
'          'V1.6.0.1 ADD END
'       End If
'       End If 'V1.6.0.1 ADD
'      Next 'V1.6.0.1 ADD
'    End If
''V1.6.0.1 ADD START
'V1.11.0.1 DEL END
    '�@�퐳�����`�F�b�N(�Ώۃt�@�C���FXX_GATEY.VEF�@XX:���[�U�[���@Y�F�f�[�^���)
    bRet = fKishuCheck(FolderName(0, FolderSyubetu) & "\")
    If bRet = False Then
       MsgBox "�ُ�I�����܂����B", _
                  vbOKOnly + vbExclamation, _
                  TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
       pfSeitouseiChck = False
       Exit Function
    End If
'V1.6.0.1 ADD END

    pfSeitouseiChck = bRet
Exit Function

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pfSeitouseiChck = False
End Function
'V1.4.0.1 ADD END
'V1.6.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fKishuCheck
'//  �@�\����  : �����v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �ΏۂƂȂ�f�[�^�̋@�퐳�����`�F�b�N���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                ���[�N�����s�R�s�[�ł̋@�퐳�����`�F�b�N�ύX
'//                Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fKishuCheck(sFileList As String) As Boolean
    Dim sKisyu       As String * 8     '�擾�@�햼
    Dim sMyName      As String         '�@�퐳�����`�F�b�N���X�g�t�@�C����
    Dim sFileName    As String         '�t�@�C�����X�g�L�ڃt�@�C����
    Dim sChkFileName As String         '�@�퐳�����`�F�b�N�t�@�C���p�X
    Dim sVerChkFile  As String         '�o�[�W�����`�F�b�N�t�@�C����
    
    Dim lLen         As Long           '�t�@�C���T�C�Y
    Dim lPos         As Long           '�o�[�W�������i�[�ʒu
           
    Dim i            As Integer        '�J�E���^�[
    Dim iCnt         As Integer        '�o�^���R�[�h��
    Dim iListCnt     As Integer        '�t�@�C���i�[��
    Dim iFileNumber  As Integer        '�t�@�C���ԍ�

    Dim bRet         As Boolean        '�@�퐳�����`�F�b�N����

    Dim uHeder       As MN_HEDER       '�w�b�_���i�[�G���A
    Dim uFotter      As MN_FOOT        '�t�b�^���i�[�G���A
    
    Dim sChkData As String             '��r�������o    'V1.20.0.1 ADD
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�
     
    '������
    iCnt = 0
    iListCnt = 0
    iFileNumber = 0
    fKishuCheck = False
        
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)
    
    '�o�[�W�����f�[�^(�@�퐳�����`�F�b�N���X�g�t�@�C���p�X)�쐬
    sVerChkFile = fSelectFile
    
    '�t�@�C�����擾�s��=�@�퐳�����`�F�b�N�t�@�C���Ȃ�
    If sVerChkFile = "" Then
       '�������`�F�b�N���s���K�v�Ȃ����߁A�����Ԃ��B
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    
    sMyName = sFileList & sVerChkFile
    
    'If Dir(sMyName) <> "" Then              '�t�@�C�������݂���?     'V1.20.0.1 DEL
    If objFso.FileExists(sMyName) = True Then    '�t�@�C�������݂���?  'V1.20.0.1 ADD
       
       iFileNumber = FreeFile               '���g�p�̃t�@�C���ԍ����擾����
       
       Open sMyName For Input Access Read As #iFileNumber     '�o�[�W�����f�[�^�̃I�[�v��
       
       '�f�[�^�ǂݍ���
       Line Input #iFileNumber, sFileName
          
       '�ǂݍ��݃f�[�^���A�w�b�_���������B
       sFileName = Mid(sFileName, Len(uHeder) - 3)
       
       '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
       Do While Not EOF(iFileNumber)
          
          '�ǂݍ��݁B
          Line Input #iFileNumber, sFileName
           
           '�擾��񂪁u/�v�ȍ~�̃R�����g�Ȃ�ΏۊO�B
           '�f�[�^���{���ȊO�Ȃ�ΏۊO
           '�f�[�^���{���݂̂̏ꍇ�̂݁A�t�@�C�����擾���s���B
           If sFileName <> "" And Left$(sFileName, 1) <> "/" _
                              And " " = Mid(sFileName, 2, 1) Then   '�t�@�C���������݂���
              iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
              ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
              ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������
              '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
              FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
              FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 12)
              '�o�^���R�[�h�����J�E���g
              iCnt = iCnt + 1
            End If
       Loop
       
       Close #iFileNumber                                     '�t�@�C������܂��B
       iFileNumber = 0
    Else
       '�t�@�C�������݂��Ȃ��ꍇ�F�������`�F�b�N���s��Ȃ��B
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    
    'V1.20.0.1 ADD  START
    If iCnt = 0 Then
       '�t�@�C�����X�g�R�[�h�����݂��Ȃ��ꍇ�F�������`�F�b�N���s��Ȃ��B
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    'V1.20.0.1 ADD  END
    
    '�t�@�C���@�퐳�����`�F�b�N���s���B
    For i = 0 To iCnt - 1
         '�`�F�b�N�Ώۃt�@�C���p�X�쐬
        sChkFileName = sFileList & FileList(i)
    
        'If Dir(sChkFileName) <> "" Then              '�t�@�C�������݂���?  'V1.20.0.1 DEL
        If objFso.FileExists(sChkFileName) = True Then  '�t�@�C�������݂���?   'V1.20.0.1 ADD
            
            lLen = FileLen(sChkFileName)             '�t�@�C���T�C�Y�̎擾

            iFileNumber = FreeFile                   '���g�p�̃t�@�C���ԍ����擾����
            '�t�@�C���̃I�[�v�����s���B
            Open sChkFileName For Binary Access Read As #iFileNumber
            '�t�b�^���̎擾
            Get #iFileNumber, lLen - Len(uFotter) + 1, uFotter
            
            Close #iFileNumber                       '�t�@�C������܂�
            iFileNumber = 0
            
            '�@�햼�Z�b�g
            sKisyu = uFotter.sKisyu
            
            sChkData = "" '�������@'V1.20.0.1 ADD
            
            '�����`�F�b�N
            If gStrCurrentForm = sFormName_EJVer Then
               'EG-R������
               'If EGR_JIKAI_KISHU = Trim(sKisyu) Then  'V1.20.0.1 DEL
               'V1.20.0.1 ADD START
               '�������o
               sChkData = Left(sKisyu, Len(EGR_JIKAI_KISHU))
               If EGR_JIKAI_KISHU = sChkData Then
               'V1.20.0.1 ADD END
                   bRet = True  '�@�퐳�����F����
               Else
                   bRet = False '�@�퐳�����F�ُ�
                   fKishuCheck = bRet
                   Set objFso = Nothing    'V1.20.0.1 ADD
                   Exit Function
               End If
            Else
               'NEG������
               'If NEG_JIKAI_KISHU = Trim(sKisyu) Then    'V1.20.0.1 DEL
               'V1.20.0.1 ADD START
               '�������o
               sChkData = Left(sKisyu, Len(NEG_JIKAI_KISHU))
               If NEG_JIKAI_KISHU = sChkData Then
               'V1.20.0.1 ADD END
                   bRet = True  '�@�퐳�����F����
               Else
                   bRet = False '�@�퐳�����F�ُ�
                   fKishuCheck = bRet
                   Set objFso = Nothing    'V1.20.0.1 ADD
                   Exit Function
               End If
            End If

        End If
    Next

  fKishuCheck = bRet
  
  Set objFso = Nothing    'V1.20.0.1 ADD
  
 Exit Function

ErrorHandler:
   If iFileNumber <> 0 Then
       Close #iFileNumber                                     '�t�@�C������܂��B
   End If
    
   '�߂�l���ُ�Ƃ���
   fKishuCheck = False
       
   Set objFso = Nothing    'V1.20.0.1 ADD

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fSelectFile
'//  �@�\����  : �o�[�W�����`�F�b�N�t�@�C����
'//  �@�\�T�v  : �Ώۃo�[�W�����`�F�b�N�t�@�C�������擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSelectFile() As String
 
 If gStrCurrentForm = sFormName_EJVer Then
    '�o�[�W�����`�F�b�N�t�@�C������ݒ肷��B
    Select Case FolderSyubetu
       Case 0 '����CPU-Pro
            fSelectFile = EHANTEI_CPU_CHK_FILE
       
       Case 1 '���C��CPU-Pro
            fSelectFile = EMAIN_CPU_CHK_FILE
       
       Case 2 '�T�uCPU-Pro
            fSelectFile = ESUB_CPU_CHK_FILE
       
       Case 3 '���C��CPU-OS
            fSelectFile = EMAIN_OS_CHK_FILE
     
     End Select
  Else
    '�o�[�W�����`�F�b�N�t�@�C������ݒ肷��B
    Select Case FolderSyubetu
       Case 0 '����CPU-Pro
             fSelectFile = NHANTEI_CPU_CHK_FILE
      
       Case 1 '���C��CPU-Pro
            fSelectFile = NMAIN_CPU_CHK_FILE
       
       Case 2 '�T�uCPU-Pro
            fSelectFile = NSUB_CPU_CHK_FILE
       
       Case 3 '���C��CPU-OS
            fSelectFile = NMAIN_OS_CHK_FILE
     
    End Select
   End If

End Function
'V.1.6.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fOldVersion
'//  �@�\����  : ���o�[�W��������
'//  �@�\�T�v  : �ꐢ��O�̃o�[�W���������s(���s)�o�[�W�����ɕԂ��B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-29   REVISED BY [TCC] S.Terao
'//                �t�F�[�Y�R�Ή��@�Ǘ��ւ̃��[�����M�������u���[�N�����s�R�s�[�v���ɂ��킹��
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                Dir�֐���FileSystemObject�ɒu������
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fOldVersion() As Boolean
    Dim bRet As Boolean                     '�߂�l
    Dim lngCnt                  As Long     '�J�E���^�[
    Dim sSrcFileName            As String   '���t�H���_���t�@�C�����X�g
    Dim lngSumRet               As Long
    Dim lngErrCode              As Long     '�G���[�R�[�h
    Dim iKansiAplChk As Integer              '�A�v���N���`�F�b�N�߂�l�@'V1.6.0.1 ADD

    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error Resume Next
 
   '���t�H���_���̃t�@�C�����X�g����������B
    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '�u���v�t�H���_���t�@�C�������쐬����
    'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������  'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else                                '�t�@�C�������݂��Ȃ�
        MsgBox "�u���v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
                   & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)
 
        fOldVersion = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '�������I������
    End If
    
    '�����t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST)
  
    '����s��t�H���_���̃t�@�C����S�č폜����
    If sNowFolderRemove <> True Then
        fOldVersion = False
        Exit Function
    End If
    
    '�����t�H���_���̃t�@�C������s��t�H���_�̓��e�ɒu������
    If sCopyOLDtoNOW <> True Then
        fOldVersion = False
        Exit Function
    End If
'V1.6.0.1 DEL START
'   '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
'     If gStrCurrentForm = sFormName_EJVer Then
'        psVersionUpdateReqest (ML_REQUEST_EGATE)
'     Else
'        psVersionUpdateReqest (ML_REQUEST_NGATE)
'     End If
'V1.6.0.1 DEL END
'V1.6.0.1 ADD START
    '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
    '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
      If gStrCurrentForm = sFormName_EJVer Then
         psVersionUpdateReqest (ML_REQUEST_EGATE)
      Else
         psVersionUpdateReqest (ML_REQUEST_NGATE)
      End If
    Else
        '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
        gintGateVerInfUpdRes = MailSts.stsNormal
    End If
'V1.6.0.1 ADD END
     
     '���D�@�o�[�W�����X�V�����ُ�
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
        '����
        MsgBox "�u���v�t�H���_�̓��e���A�u���s�v�t�H���_�ɖ߂��āA" _
                    & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "�̈ꐢ��O�̃o�[�W�������A" _
                    & Chr(vbKeyReturn) & "���s�o�[�W�����Ƃ��܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        fOldVersion = True
    Else
        '�ُ�
        If gStrCurrentForm = sFormName_EJVer Then
          MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
                  vbOKOnly + vbExclamation, _
                  "EG-R�������D�@ �o�[�W�����Ǘ�"
        Else
           MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
                   vbOKOnly + vbExclamation, _
                   "NEG�������D�@ �o�[�W�����Ǘ�"
        End If
        fOldVersion = False
    End If

    fOldVersion = True
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetFolderName
'//  �@�\����  : �f�[�^�W�J
'//  �@�\�T�v  : �t�H���_���Ȃǂ̃f�[�^���O���[�o���G���A�ɓW�J����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                ���[�N�����s�R�s�[�ł̐������`�F�b�NINI�Ǎ���
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSetFolderName()

        TitleBox(0) = "����CPU-Pro"
        TitleBox(1) = "���C��CPU-Pro"
        TitleBox(2) = "�T�uCPU-Pro"
        TitleBox(3) = "���C��CPU-OS"
        TitleBox(4) = "�\��1"
        TitleBox(5) = "�\��2"
    
        LogBox(0) = "����"
        LogBox(1) = "���C��"
        LogBox(2) = "�T�u"
        LogBox(3) = "OS"
        LogBox(4) = "�\��1"
        LogBox(5) = "�\��2"
        
  If gStrCurrentForm = sFormName_EJVer Then
        '�t�H���_���ɐݒ���s��
        FolderName(0, 0) = E_EHAN1WRK
        FolderName(1, 0) = E_EHAN1NOW
        FolderName(2, 0) = E_EHAN1OLD
        FolderName(0, 1) = E_EPRO1WRK
        FolderName(1, 1) = E_EPRO1NOW
        FolderName(2, 1) = E_EPRO1OLD
        FolderName(0, 2) = E_ESCPUWRK
        FolderName(1, 2) = E_ESCPUNOW
        FolderName(2, 2) = E_ESCPUOLD
        FolderName(0, 3) = E_EOSWRK
        FolderName(1, 3) = E_EOSNOW
        FolderName(2, 3) = E_EOSOLD
        FolderName(0, 4) = E_EYOBI1WRK
        FolderName(1, 4) = E_EYOBI1NOW
        FolderName(2, 4) = E_EYOBI1OLD
        FolderName(0, 5) = E_EYOBI2WRK
        FolderName(1, 5) = E_EYOBI2NOW
        FolderName(2, 5) = E_EYOBI2OLD
    Else
        '�t�H���_���ɐݒ���s��
        FolderName(0, 0) = N_NHAN1WRK
        FolderName(1, 0) = N_NHAN1NOW
        FolderName(2, 0) = N_NHAN1OLD
        FolderName(0, 1) = N_NPRO1WRK
        FolderName(1, 1) = N_NPRO1NOW
        FolderName(2, 1) = N_NPRO1OLD
        FolderName(0, 2) = N_NSCPUWRK
        FolderName(1, 2) = N_NSCPUNOW
        FolderName(2, 2) = N_NSCPUOLD
        FolderName(0, 3) = N_NOSWRK
        FolderName(1, 3) = N_NOSNOW
        FolderName(2, 3) = N_NOSOLD
        FolderName(0, 4) = N_NYOBI1WRK
        FolderName(1, 4) = N_NYOBI1NOW
        FolderName(2, 4) = N_NYOBI1OLD
        FolderName(0, 5) = N_NYOBI2WRK
        FolderName(1, 5) = N_NYOBI2NOW
        FolderName(2, 5) = N_NYOBI2OLD
    End If

'V1.20.0.1 ADD START
'-------EG-R����-------
    ' �L�[��:����CPU-PRO��\
    EHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    
    ' �L�[��:���C��CPU-PRO��\
    EMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_PRO, PATH_GATEVER_FILE)
        
    ' �L�[���F�T�uCPU-PRO��\
    ESUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_SUB_PRO, PATH_GATEVER_FILE)
    
    ' �L�[��:���C��CPU-OS��\
    EMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_OS, PATH_GATEVER_FILE)
    
'-------NEG����-------
    ' �L�[��:����CPU-PRO��\
    NHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    
    ' �L�[��:���C��CPU-PRO��\
    NMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_PRO, PATH_GATEVER_FILE)
    
    ' �L�[���F�T�uCPU-PRO��\
    NSUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_SUB_PRO, PATH_GATEVER_FILE)
    
    ' �L�[��:���C��CPU-OS��\
    NMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_OS, PATH_GATEVER_FILE)
'V1.20.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fReadFileList
'//  �@�\����  : �t�@�C�����X�g�̎擾
'//  �@�\�T�v  : �t�@�C�����X�g���A�t�@�C�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fReadFileList(sFileList As String) As Boolean
    Dim iFileNumber As Integer      '�t�@�C���ԍ�
    Dim sFileName As String         '�t�@�C����
    Dim iListCnt As Integer         '�t�@�C���i�[��

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����

    Open sFileList For Input Access Read As #iFileNumber    '�t�@�C�����X�g�̃I�[�v��
    Do While Not EOF(iFileNumber)                           '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        Line Input #iFileNumber, sFileName                  '�f�[�^��ǂݍ��݂܂��B
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                '�t�@�C���������݂���
            iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
            ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
            ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������
            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
                                                            '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
        End If
    Loop
    Close #iFileNumber      '�t�@�C������܂��B

    fReadFileList = True    '�߂�l�𐳏�Ƃ���

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    'V1.21.0.1 ADD  START
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    'V1.21.0.1 ADD  END
    fReadFileList = False   '�߂�l���G���[�Ƃ���
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fDataFileCheck
'//  �@�\����  : �����v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �ΏۂƂȂ�HAN_KUKA.KUK�L���`�F�b�N���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-23  CODED   BY [TCC] D.Yamashita
'//                 �E�t�F�[�Y�R�c�����ڑΉ��@�ُ펞�N���[�Y�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fDataFileCheck(sFileList As String) As Boolean
    Dim iFileNumber As Integer      '�t�@�C���ԍ�
    Dim sFileName As String         '�t�@�C����
    Dim iListCnt As Integer         '�t�@�C���i�[��
    Dim sFolderPath As String       'HAN_KUKA.KUK�t�H���_�p�X�p
    Dim sHANKUKAPath As String      'HAN_KUKA.KUK�t���p�X�p
     
    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����

    Open sFileList For Input Access Read As #iFileNumber    '�t�@�C�����X�g�̃I�[�v��
    Do While Not EOF(iFileNumber)                           '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        Line Input #iFileNumber, sFileName                  '�f�[�^��ǂݍ��݂܂��B
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                '�t�@�C���������݂���
            iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
            ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
            ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������
            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
                                                            '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
            If HANKUKA_KUK = FileList(iListCnt - 1) Then
               'HAN_KUKA.KUK�t�@�C�����L�����ꍇ�A�f�[�^�������`�F�b�N���s���B
               psFolderPathGet sFileList, sFolderPath
               sHANKUKAPath = sFolderPath & HANKUKA_KUK
               If fHankukaChck(sHANKUKAPath) = False Then
                 '�f�[�^�������`�F�b�N�ُ펞�́A�߂�l��False��ݒ肷��B
                  fDataFileCheck = False
                  Close #iFileNumber        '�t�@�C������܂��B   'V1.11.0.1 ADD
                  Exit Function
               End If
            End If
        End If
  Loop
  
  Close #iFileNumber        '�t�@�C������܂��B

  fDataFileCheck = True     '�߂�l�𐳏�Ƃ���

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:               ' �G���[�������[�`���B
    fDataFileCheck = False  '�߂�l���G���[�Ƃ���
    Close #iFileNumber      '�t�@�C������܂��B
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fHankukaChck
'//  �@�\����  : HAN_KUKA.KUK�������`�F�b�N����
'//  �@�\�T�v  : �ΏۂƂȂ�HAN_KUKA.KUK�̓��e���`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-06   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//     REVISIONS :(1.16.0.1) 2009-12-21   REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(V2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 EG-R(KK)�@������Ή��@KUK�������`�F�b�N�ύX
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fHankukaChck(sFilePath As String) As Boolean
    Dim iFileNumber As Integer           '�t�@�C���ԍ�
    Dim i As Integer
    Dim lSts As Long
    Dim sKeyName As String
    Dim lPos As Long                     '�o�[�W�������i�[�ʒu
    Dim lLen As Long                     '�t�@�C���T�C�Y
    Dim uFooter As MN_FOOT          '�t�b�^���i�[�G���A
'    Dim uHeder As MN_FOOT           '�w�b�_���i�[�G���A     'V1.4.0.1 DEL
    Dim sDateTime As String
    Dim j As Integer
    Dim lngErrCode As Long          '�G���[�R�[�h
    'V1.4.0.1 ADD START
    Dim uHeder As HAN_KUKA_KUK_HEADER       '�w�b�_���i�[�G���A
    Dim sGetInfo As String * MAX_PATH_SIZE  'INI�t�@�C���擾�p
    Dim sChkFileData As String
    Dim iMojisu As Integer
    
    'V1.16.0.1 ADD Start
    Dim bChkSts As Boolean              '�`�F�b�N���ʃt���O
    Dim sChkData As String              '��r�������o
    'V1.16.0.1 ADD End
    
   '�������F����(�u�����N�j
    sNGSts = ""
    sNGKoumoku = ""
    'V1.4.0.1 ADD END
    Dim oFs As New FileSystemObject 'V2.5.0.1 ADD
    
    fHankukaChck = False
    
'V2.5.0.1 ADD START
 '�t�@�C���L���`�F�b�N���s���B
 If oFs.FileExists(sFilePath) = False Then
    '�t�@�C����������ΐ������`�F�b�N���s��Ȃ��B
    fHankukaChck = True
    Set oFs = Nothing
    Exit Function
 End If
'V2.5.0.1 ADD END

 'V1.4.0.1 DEL START
'   For i = 0 To INI_MAX
'      '�w�b�_�F���Ғl�@�햼�擾
'      sKeyName = Format(HEDER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sHederKisyu(i), _
'                                     Len(HAN_KUKA_DATA.sHederKisyu(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      '�w�b�_�F���Ғl�t�@�C�����擾
'      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sHederFile(i), _
'                                     Len(HAN_KUKA_DATA.sHederFile(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      '�t�b�^�F���Ғl�@�햼�擾
'      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sFotterKisyu(i), _
'                                     Len(HAN_KUKA_DATA.sFotterKisyu(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      '�t�b�^�F���Ғl�t�@�C�����擾
'      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sFotterFile(i), _
'                                     Len(HAN_KUKA_DATA.sFotterFile(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'    Next i
    'V1.4.0.1 DEL END
    'V1.4.0.1 ADD START
    '������
    For i = 0 To INI_MAX - 1
        HAN_KUKA_DATA.sHederKisyu(i) = ""
        HAN_KUKA_DATA.sHederFile(i) = ""
        HAN_KUKA_DATA.sFotterKisyu(i) = ""
        HAN_KUKA_DATA.sFotterFile(i) = ""
    Next
    For i = 0 To INI_MAX - 1
      '�w�b�_�F���Ғl�@�햼�擾
      sKeyName = Format(HEDER_KISHU_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
       
      Else
        HAN_KUKA_DATA.sHederKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      '�w�b�_�F���Ғl�t�@�C�����擾
      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
        
      Else
         HAN_KUKA_DATA.sHederFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      '�t�b�^�F���Ғl�@�햼�擾
      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
        
      Else
        HAN_KUKA_DATA.sFotterKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      '�t�b�^�F���Ғl�t�@�C�����擾
      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
      lSts = GetPrivateProfileString(HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      If lSts = False Then
        
      Else
        HAN_KUKA_DATA.sFotterFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
    Next i
    'V1.4.0.1 ADD END

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�
    
    'HAN_KUKA.KUK�t�@�C���T�C�Y�擾
    lLen = FileLen(sFilePath)
    
    '���g�p�̃t�@�C���ԍ����擾����
    iFileNumber = FreeFile
    
    'V1.4.0.1 DEL START
'    'HAN_KUKA.KUK�t�@�C�����I�[�v������B
'    Open sFilePath For Input Access Read As #iFileNumber
'
'    'HAN_KUKA.KUK�t�@�C���̃w�b�_�����擾����B
''    Get #iFileNumber, 1, uHeder
    'V1.4.0.1 DEL END

    'V1.4.0.1 ADD START
    'HAN_KUKA.KUK�t�@�C�����I�[�v������B
    Open sFilePath For Binary Access Read As #iFileNumber
            
    'HAN_KUKA.KUK�t�@�C���̃w�b�_�����擾����B
    Get #iFileNumber, 1, uHeder
    'V1.4.0.1 ADD END

   'HAN_KUKA.KUK�t�@�C���̃t�b�^�����擾����B
    Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter

    'HAN_KUKA.KUK�t�@�C�����N���[�Y����B
    Close #iFileNumber
    
    iFileNumber = 0                          'V1.4.0.1 ADD
'V1.4.0.1 DEL START
    '�@�햼/�t�@�C�����`�F�b�N
'    For i = 0 To 5
'       '�w�b�_���F�@�햼�`�F�b�N
'       If uHeder.sKisyu <> HAN_KUKA_DATA.sHederKisyu(i) Then
'          Exit Function
'       End If
'       '�w�b�_���F�t�@�C�����`�F�b�N
'       If uHeder.sFileName <> HAN_KUKA_DATA.sHederFile(i) Then
'          Exit Function
'       End If
'       '�t�b�^���F�@�햼�`�F�b�N
'       If uFooter.sKisyu <> HAN_KUKA_DATA.sFotterKisyu(i) Then
'          Exit Function
'       End If
'       '�t�b�^���F�t�@�C�����`�F�b�N
'       If uFooter.sFileName <> HAN_KUKA_DATA.sFotterFile(i) Then
'          Exit Function
'       End If
'     Next
'V1.4.0.1 DEL END
   'V1.4.0.1 ADD START
   '�w�b�_���F�@�햼�`�F�b�N
   iMojisu = InStr(uHeder.sKisyuName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sKisyuName, 1)
   Else
     sChkFileData = Mid(uHeder.sKisyuName, 1, iMojisu)
   End If
'V1.16.0.1 DEL START
'    For i = 0 To INI_MAX - 1
'      If sChkFileData <> HAN_KUKA_DATA.sHederKisyu(i) Then
'         If i = INI_MAX - 1 Then
'            '�@�햼���Ғl�S�s��v�F
'            sNGSts = ERROR_HEDER
'            sNGKoumoku = KISHU_NAME_ERROR
'            GoTo ErrorHandler
'         End If
'      Else
'        Exit For
'      End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
    bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sHederKisyu(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sHederKisyu(i)))
          If sChkData = HAN_KUKA_DATA.sHederKisyu(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    '�`�F�b�N���ʃt���O����
    If bChkSts = False Then
       '�@�햼���Ғl�S�s��v�F
        sNGSts = ERROR_HEDER
        sNGKoumoku = KISHU_NAME_ERROR
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END

   '�w�b�_���F�t�@�C�����`�F�b�N
   iMojisu = InStr(uHeder.sProgrumName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sProgrumName, 1)
   Else
     sChkFileData = Mid(uHeder.sProgrumName, 1, iMojisu)
   End If

'V1.16.0.1 DEL START
'   For i = 0 To INI_MAX - 1
'       If sChkFileData <> HAN_KUKA_DATA.sHederFile(i) Then
'         If i = INI_MAX - 1 Then
'            '�t�@�C�������Ғl�S�s��v�F
'            sNGSts = ERROR_HEDER
'            sNGKoumoku = FILE_NAME_ERRORE
'            GoTo ErrorHandler
'         End If
'      Else
'         Exit For
'      End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
    bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sHederFile(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sHederFile(i)))
          If sChkData = HAN_KUKA_DATA.sHederFile(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    '�`�F�b�N���ʃt���O����
    If bChkSts = False Then
       '�@�햼���Ғl�S�s��v�F
        sNGSts = ERROR_HEDER
        sNGKoumoku = FILE_NAME_ERRORE
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END
    
   '�쐬���t�`�F�b�N
   '�w�b�_���F�쐬���t�����l���ǂ���
    sDateTime = ""
    For j = 0 To 3
        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
    Next
    For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
    Next
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_HEDER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    
    '�o�[�W�������l�`�F�b�N
    If IsNumeric(uHeder.sVersion) = False Then
       sNGSts = ERROR_HEDER
       sNGKoumoku = VERSION_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    
   '�t�b�^���F�@�햼�`�F�b�N
   iMojisu = InStr(uFooter.sKisyu, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uFooter.sKisyu, 1)
   Else
     sChkFileData = Mid(uFooter.sKisyu, 1, iMojisu)
   End If
'V1.16.0.1 DEL START
'    For i = 0 To INI_MAX - 1
'      If sChkFileData <> HAN_KUKA_DATA.sFotterKisyu(i) Then
'         If i = INI_MAX - 1 Then
'             '�@�햼���Ғl�S�s��v�F
'             sNGSts = ERROR_FOTTER
'             sNGKoumoku = KISHU_NAME_ERROR
'             GoTo ErrorHandler
'          End If
'       Else
'         Exit For
'       End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
   bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sFotterKisyu(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterKisyu(i)))
          If sChkData = HAN_KUKA_DATA.sFotterKisyu(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    '�`�F�b�N���ʃt���O����
    If bChkSts = False Then
       '�@�햼���Ғl�S�s��v�F
        sNGSts = ERROR_FOTTER
        sNGKoumoku = KISHU_NAME_ERROR
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END

   '�t�b�^���F�t�@�C�����`�F�b�N
   iMojisu = InStr(uFooter.sFileName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uFooter.sFileName, 1)
   Else
     sChkFileData = Mid(uFooter.sFileName, 1, iMojisu)
   End If
'V1.16.0.1 DEL START
'    For i = 0 To INI_MAX - 1
'       If sChkFileData <> HAN_KUKA_DATA.sFotterFile(i) Then
'          If i = INI_MAX - 1 Then
'             '�@�햼���Ғl�S�s��v�F
'             sNGSts = ERROR_FOTTER
'             sNGKoumoku = FILE_NAME_ERRORE
'             GoTo ErrorHandler
'          End If
'       Else
'         Exit For
'       End If
'    Next
'   'V1.4.0.1 ADD END
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
   bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sFotterFile(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterFile(i)))
          If sChkData = HAN_KUKA_DATA.sFotterFile(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    '�`�F�b�N���ʃt���O����
    If bChkSts = False Then
       '�@�햼���Ғl�S�s��v�F
        sNGSts = ERROR_FOTTER
        sNGKoumoku = FILE_NAME_ERRORE
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END

'V1.4.0.1 DEL START
'   '�쐬���t�`�F�b�N
'   '�w�b�_���F�쐬���t�����l���ǂ���
'    sDateTime = ""
'    For j = 0 To 3
'        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
'    Next
'    sDateTime = sDateTime & " "
'    For j = 4 To 5
'        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
'    Next
'    If (Chr(sDateTime) >= "A" And Chr(sDateTime) <= "Z") And _
'        (Chr(sDateTime) >= "a" And Chr(sDateTime) <= "z") Then
'         Exit Function
'    End If
'V1.4.0.1 DEL END
      
    '�t�b�^���F�쐬���t�����l���ǂ���
     sDateTime = ""
     For j = 0 To 3
         sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
     Next
    'sDateTime = sDateTime & " " 'V1.4.0.1 DEL
     For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    'V1.4.0.1 DEL START
'    If (Chr(sDateTime) >= "A" And Chr(sDateTime) <= "Z") And _
'       (Chr(sDateTime) >= "a" And Chr(sDateTime) <= "z") Then
'        Exit Function
'    End If
    'V1.4.0.1 DEL END
    
    'V1.4.0.1 ADD START
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_FOTTER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    'V1.4.0.1 ADD END
    'V1.4.0.1 DEL START
'      '�o�[�W�����l�`�F�b�N
'    '�w�b�_���F�o�[�W�����l�����l���ǂ���
'    If (Chr(uHeder.sVersion) >= "A" And Chr(uHeder.sVersion) <= "Z") And _
'        (Chr(uHeder.sVersion) >= "a" And Chr(uHeder.sVersion) <= "z") Then
'        Exit Function
'    End If
'
'    '�t�b�^���F�o�[�W�����l�����l���ǂ���
'    If (Chr(uFooter.sVersion) >= "A" And Chr(uFooter.sVersion) <= "Z") And _
'       (Chr(uFooter.sVersion) >= "a" And Chr(uFooter.sVersion) <= "z") Then
'        Exit Function
'    End If
    'V1.4.0.1 DEL END
    
    'V1.4.0.1 ADD START
    '�o�[�W�����l�`�F�b�N
    '�t�b�^���F�o�[�W�����l�����l���ǂ���
    If IsNumeric(uFooter.sVersion) = False Then
       sNGSts = ERROR_FOTTER
       sNGKoumoku = VERSION_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    'V1.4.0.1 ADD END
    
    '�u�����ް�ޮ݁F�����`�F�b�N����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_OK, 0)
    
    '���ׂ�OK�̏ꍇ�ATRUE�ł�����B
    fHankukaChck = True

Exit Function 'V1.4.0.1 ADD
'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    'V1.4.0.1 ADD START
    If iFileNumber > 0 Then
       'HAN_KUKA.KUK�t�@�C�����N���[�Y����B
       Close #iFileNumber
    End If
    iFileNumber = 0
    'V1.4.0.1 ADD END
    
    '�u�����ް�ޮ݁F�����`�F�b�N�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   ' Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_OK, 0) 'V1.4.0.1 DEL
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_ERROR, lngErrCode)  'V1.4.0.1 ADD
    fHankukaChck = False   '�߂�l���G���[�Ƃ���
    'HAN_KUKA.KUK�t�@�C�����N���[�Y����B
    'Close #iFileNumber                        'V1.4.0.1 DEL
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sWriteListbox
'//  �@�\����  : �o�[�W�����e�L�X�g�t�@�C�������݁B
'//  �@�\�T�v  : ���X�g�{�b�N�X�̓��e���A�o�[�W�����e�L�X�g�t�@�C���ɏ������ށB
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sWriteListbox()
    Dim iFileNumber As Integer
    Dim i As Integer

    On Error Resume Next
    
    iFileNumber = FreeFile              '���g�p�̃t�@�C���ԍ����擾����

    Open MN_VERSI_FILE For Output Access Write As #iFileNumber
                                        '�t�@�C�������쐬���܂��B

    For i = 0 To lstKan.ListCount - 1

        Print #iFileNumber, lstKan.List(i) & Chr(vbKeyReturn)
                                        '�f�[�^����������
    Next

    Close #iFileNumber                  '�t�@�C������܂��B
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFDInstall
'//  �@�\����  : �}�̃C���X�g�[������
'//  �@�\�T�v  : �C���X�g�[���}�̃t�@�C�����A���[�N�t�H���_�ɃR�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�s��C��
'//                 �t�F�[�Y�R�Ή�
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ���̓t�@�C���i�[�f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//                 Dir�֐���FileSystemObject�ɒu������
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall(sFlag As String)
    Dim MyName As String            '�t�@�C���t���p�X��
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim iResponse As Integer        'MsgBox�{�^���R�[�h
    Dim sInputPass As String        '�C���X�g�[�����f�B���N�g����(STD)or�t�@�C����(LZH)
    Dim sInputFolder As String      '�C���X�g�[�����t�H���_���BLZH�̎��A�𓀐�t�H���_�B
    Dim lngErrCode As Long          '�G���[�R�[�h
    'V1.6.0.1 ADD START
    Dim bRet As Boolean             '�������`�F�b�N�߂�l
    Dim sChkName As String          '�`�F�b�N�t�@�C��
    'V1.6.0.1 ADD END
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                    '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^

    If sFlag = "STD" Then
    '�W���i�񈳏k�j�t�@�C���w��̎�:
    '�f�B���N�g���I����ʂ�\�������A���̓t�@�C���i�[�f�B���N�g�����𓾂�B
'       sInputPass = pfDirSelection("a:", "�C���X�g�[���}�̂̃f�B���N�g���I��")     'V1.12.0.1 DEL
        'sInputPass = pfDirSelection("H:", "�C���X�g�[���}�̂̃f�B���N�g���I��")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
        sInputPass = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
        If sInputPass = "" Then
        '�f�B���N�g�����w��Ȃ����͏����I��
            'V1.20.0.1 ADD START
            Set objFso = Nothing
            Set objFi = Nothing
            'V1.20.0.1 ADD END
            Exit Sub
        End If
        sInputFolder = sInputPass
    Else
    '���k�t�@�C���w��̎�:
    '���k�t�@�C���I����ʂ�\�������ALZH�t�@�C���t���p�X���𓾂�i�f�t�H���g�͂e�c��\���B�j�B
'       sInputPass = pfCabFileSelection("a:")     'V1.12.0.1 DEL
        'V1.20.0.1 DEL START
       'sInputPass = pfCabFileSelection("H:")      'V1.12.0.1 ADD
        'If sInputPass = "" Then Exit Sub '�t�@�C�����I������Ȃ���Ζ߂�B
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
        '�擾�t�@�C������������
        CommonDialog1.FileName = ""
        '�����f�B���N�g����ݒ�
        If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    '�t�H���_�I����ʃf�t�H���g�p�X�P�����݂��邩
            '���݂��邽�߁A�f�t�H���g�p�X�P�iH:�j��ݒ�
            CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
        Else
            '���݂��Ȃ����߁A�f�t�H���g�p�X�Q�iC:�j��ݒ�
            CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
        End If
        '�g���q��ݒ�
        CommonDialog1.Filter = "���k�t�@�C���i*.cab�j|*.cab|"
        '�t�@�C���I����ʂ��J��
        CommonDialog1.ShowOpen
        '�I�������t�@�C�������擾
        sInputPass = CommonDialog1.FileName
        If sInputPass = "" Then '�t�@�C�����I��
            Set objFso = Nothing
            Set objFi = Nothing
            Exit Sub    '�t�@�C�����I������Ȃ���Ώ������f
        End If
        
        Call ChDrive("D")  'V2.5.0.1 ADD
        
        'V1.20.0.1 ADD END
       '�𓀗p�ꎞ�t�H���_���쐬����B
       psMakeFolder MELTED_FOLDER_FULLPASS
       '���k�t�@�C�����A�𓀗p�ꎞ�t�H���_�ɉ𓀁E�i�[������B
        Call psCabReqest(CABREQEST.CAB_THAW, sInputPass, MELTED_FOLDER_FULLPASS)
        If glngCabErrCd <> 0 Then
            psDeleteFolder MELTED_FOLDER_FULLPASS
            'V1.20.0.1 ADD START
            Set objFso = Nothing
            Set objFi = Nothing
            'V1.20.0.1 ADD END
            Exit Sub
        End If
        sInputFolder = MELTED_FOLDER_FULLPASS
    End If
    
    '�u���[�N�R�s�[�m�F�v�|�b�v�A�b�v��ʕ\��
    iResponse = MsgBox(sInputPass & " �̑S�Ẵt�@�C�����A" _
                       & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                       & TitleBox(FolderSyubetu) & "�́u���[�N�v�t�H���_�ɃR�s�[���܂��B " _
                       & "��낵���ł����H", _
                       vbYesNo + vbExclamation, _
                       TitleBox(FolderSyubetu) & "  �}�́����[�N �R�s�[")
    If iResponse = vbNo Then
    '[������] �{�^����I��:�������Ȃ��B
    '�A���A���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B
       If sFlag = "LZH" Then
           psDeleteFolder MELTED_FOLDER_FULLPASS
       End If
        'V1.20.0.1 ADD START
        Set objFso = Nothing
        Set objFi = Nothing
        'V1.20.0.1 ADD END
        Exit Sub
    End If
    
    'V1.6.0.1 ADD START
    '�O�����̓v�����������`�F�b�N
    If sFlag = "STD" Then
       '�}�́����[�N �R�s�[��
       bRet = pfInstallSeitouseiChck(sInputPass)
    Else
       '���k�t�@�C�������[�N �R�s�[��
       bRet = pfInstallSeitouseiChck(MELTED_FOLDER_FULLPASS & "\")
    End If
    If bRet = False Then
       'V1.20.0.1 ADD START
       Set objFso = Nothing
       Set objFi = Nothing
       'V1.20.0.1 ADD END
       Exit Sub
    End If
    
    '�o�[�W�����`�F�b�N�t�@�C���L���`�F�b�N���s���B
    sChkName = fSelectFile
    'V1.20.0.1 DEL START
'    sChkName = Dir(FolderName(0, FolderSyubetu) & "\" & sChkName)
'    If sChkName <> "" Then
'      Kill FolderName(0, FolderSyubetu) & "\" & sChkName
'    End If
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    If objFso.FileExists(FolderName(0, FolderSyubetu) & "\" & sChkName) = True Then
        '�w��t�@�C�������݂���
        sChkName = objFso.GetFileName(FolderName(0, FolderSyubetu) & "\" & sChkName)
        Kill FolderName(0, FolderSyubetu) & "\" & sChkName
    Else
        sChkName = ""
    End If
    'V1.20.0.1 ADD END
    'V1.6.0.1 ADD START
    
    '�w��t�H���_���̃t�@�C�����A�S�āu���[�N�v�t�H���_�ɃR�s�[����B
    'V1.20.0.1 DEL START
'    MyName = Dir(sInputFolder & "\*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            '�}�̓��t�@�C�������쐬����
'            sSrcFileName = sInputFolder & "\" & MyName
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
'                '���[�N�t�H���_���t�@�C�������쐬����
'                sDstFileName = FolderName(0, FolderSyubetu) & "\" & MyName
'                '�}�̓��̃t�@�C�������[�N�t�H���_�ɃR�s�[����
'                FileCopy sSrcFileName, sDstFileName
'            End If
'        End If
'        MyName = Dir                    ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(sInputFolder).files   '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then  '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            '�}�̓��t�@�C�������쐬
            sSrcFileName = sInputFolder & "\" & MyName
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
                '���[�N�t�H���_���t�@�C�������쐬����
                sDstFileName = FolderName(0, FolderSyubetu) & "\" & MyName
                '�}�̓��̃t�@�C�������[�N�t�H���_�ɃR�s�[����
                FileCopy sSrcFileName, sDstFileName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    '���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B(�g�p�ς݂̂���)
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
    
    '�u���[�N�R�s�[����I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�C���X�g�[���}�̂̑S�Ẵt�@�C�����A" _
            & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "�́u���[�N�v�t�H���_��" _
            & Chr(vbKeyReturn) & "�R�s�[���܂����B", _
            vbOKOnly + vbExclamation, _
            TitleBox(FolderSyubetu) & "  �}�́����[�N �R�s�["
    
    '�u�����ް�ޮ݁F�}�́�ܰ���߰��������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_OK, 0)
    
    '�o�[�W������񃊃X�g�{�b�N�X���쐬����
    fMakeListbox
    Exit Sub    '�������I������

ErrorHandler:   ' �G���[�����B
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    Select Case Err.Number
        Case 53 ' �u�w���ʃt�@�C���Ȃ��v�|�b�v�A�b�v��ʕ\��
            MsgBox "�C���X�g�[���}�̂� " & TitleBox(FolderSyubetu) & "�́A" _
                   & Chr(vbKeyReturn) & "�ЂƂ����݂��܂���B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �����[�N �R�s�["
            Exit Sub
        Case 71 '�u�}�̂Ȃ��v�|�b�v�A�b�v��ʕ\��
            iResponse = MsgBox("�}�̂���������Ă��܂���B", _
                    vbRetryCancel + vbExclamation, _
                    TitleBox(FolderSyubetu) & "  �����[�N �R�s�[")
            If iResponse = vbRetry Then    '�u��蒼���v�{�^����I�������ꍇ
                Resume      ' �G���[�����������s���珈���ĊJ
            Else                            '�u�L�����Z���v�{�^����I�������ꍇ
                Exit Sub    '�������I������
            End If
        Case Else  '�u���[�N�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
           MsgBox "�C���X�g�[���}�̂���̃R�s�[�G���[���������܂����B" _
                   & Chr(vbKeyReturn) & "�G���[�R�[�h��" _
                   & str$(Err.Number), _
                   vbOKOnly + vbExclamation, _
                   "�����[�N �R�s�["
    End Select
    
    '�u�����ް�ޮ݁F�}�́�ܰ���߰�����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)
End Sub

'V1.6.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfInstallSeitouseiChck
'//  �@�\����  : �O�����̓v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �O�����̓v���O��������f�[�^�������`�F�b�N�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 �t�@�C�����`�F�b�N�s��C��
'//     REVISIONS :(X.X.X.X)----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfInstallSeitouseiChck(sInputPass As String) As Boolean
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim myLen As Long                        '������̒���
    Dim SysCodeTxt As String                 '�o�C�g�ϊ���(�S�p�����p)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    Dim bRet As Boolean
    Dim lngCnt As Long
    Dim iRet   As Integer                    '�o�[�W�����`�F�b�NDLL�߂�l
    Dim iGouki As Integer                    '���@�ԍ�
    Dim sVersionInfoPath As String           '�o�[�W�������t�@�C��(���@��)
    Dim sSrcFileName As String               '�t�@�C�����X�g��
    Dim lngErrCode   As Long
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error Resume Next
    
    pfInstallSeitouseiChck = True
    
    '********************************
    '*�v�����������`�F�b�N
    '********************************
    '�O���}�̃t�H���_���t�@�C�������쐬
    sSrcFileName = sInputPass & MN_FILELIST
    '�O���}�̂̌���������
    'If Dir(sSrcFileName) <> "" Then     'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else
      '�t�@�C�������݂��Ȃ�
      MsgBox "�}�̓��ɁA�t�@�C�����X�g�����݂��܂���B", _
             vbOKOnly + vbExclamation, _
             "�����[�N �R�s�["
     '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      pfInstallSeitouseiChck = False
      Set objFso = Nothing    'V1.20.0.1 ADD
      Exit Function
    End If

   '����[�N��t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(sInputPass & MN_FILELIST)

    '�T���l�`�F�b�N
    For lngCnt = 0 To UBound(FileList) - 1
        If pfFileSumChk(sInputPass & FileList(lngCnt), lngSumRet) <> True Then
            '�T���l�ُ�
            If lngSumRet = SUM_CHK.SumErr Then
               MsgBox "�T���l���ُ�ł��B" _
                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
                      vbOKOnly + vbExclamation, _
                      sJverName & "�������D�@ �o�[�W�����Ǘ�"
            '�T���l�ُ�ȊO�ُ�
            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
                   '�u���[�N�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
               MsgBox "�C���X�g�[���}�̂���̃R�s�[�G���[���������܂����B" _
                     & Chr(vbKeyReturn) & "�G���[�R�[�h��" _
                     & str$(Err.Number), _
                     vbOKOnly + vbExclamation, _
                    "�����[�N �R�s�["
            End If
            pfInstallSeitouseiChck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    '�t�@�C�����ő�`�F�b�N
    If UBound(FileList) > FILECNT_MAX Then
       MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
              & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
              vbOKOnly + vbExclamation, _
              sJverName & "�������D�@ �o�[�W�����Ǘ�"
      pfInstallSeitouseiChck = False

      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)

      Exit Function
    End If
'V2.6.0.1 DEL START
'    '�t�@�C�����T�C�Y�`�F�b�N
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
'
'    bRet = True
'
'    '�t�@�C�����X�g���I�[�v���B
'    Open sInputPass & MN_FILELIST For Input As #iFileNumber
'    For i = 0 To lngFileListCnt
'       If i = lngFileListCnt Then
'          Exit For
'       End If
'       '�t�@�C�������擾����B
'       Input #iFileNumber, strWork
'       '�t�@�C������`�Ȃ�
'       If strWork = "" Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'          bRet = False
'          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'          Exit For
'       '�t�H�[�}�b�g�ُ�
'       ElseIf " " <> Mid(strWork, 2, 1) And Left$(strWork, 1) <> "/" Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       '�t�H�[�}�b�g�ُ�
'       ElseIf (InStr(strWork, ".") - 1) = -1 And Left$(strWork, 1) <> "/" Then
'           MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       '�u/*--�v���̃R�����g���͏���
'       ElseIf Left$(strWork, 1) = "/" Then
'               '�������Ȃ��B
'       Else
'          '�t�@�C�����݂̂𒊏o
'          sGetFileListName = Mid(strWork, 3, 16)
'          '�擾�t�@�C�����̃T�C�Y���擾
'          myLen = LenB(StrConv(Trim(sGetFileListName), vbFromUnicode))                                              '���p���Z�̃o�C�g�����擾
'          If FILE_NAME_MAX_SIZE < myLen Then
'            '13�o�C�g�ȏ�̏ꍇ
'            MsgBox "�t�@�C�������ُ�ł��B" _
'                   & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                   vbOKOnly + vbExclamation, _
'                   sJverName & "�������D�@ �o�[�W�����Ǘ�"
'            bRet = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'            Exit For
'           End If
'        End If
'     Next
'    '�t�@�C�����X�g���N���[�Y�B
'    Close #iFileNumber
'V2.6.0.1 DEL END
'V2.6.0.1 ADD START
    For i = 0 To UBound(FileList) - 1
       '�擾�t�@�C�����̃T�C�Y���擾
       myLen = LenB(StrConv(Trim(FileList(i)), vbFromUnicode))                                              '���p���Z�̃o�C�g�����擾
       If FILE_NAME_MAX_SIZE < myLen Then
          '13�o�C�g�ȏ�̏ꍇ
          MsgBox "�t�@�C�������ُ�ł��B" _
                 & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
                  vbOKOnly + vbExclamation, _
                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
           bRet = False
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
           Exit For
       End If
    Next
'V2.6.0.1 ADD END

    pfInstallSeitouseiChck = bRet
Exit Function

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pfInstallSeitouseiChck = False
End Function
'V1.6.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyWRKtoNOW
'//  �@�\����  : �ŐV�o�[�W�����R�s�[
'//  �@�\�T�v  : ���[�N�t�H���_���̃t�@�C�����A���s�t�H���_�ɃR�s�[
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sCopyWRKtoNOW() As Boolean
    
    Dim i As Integer                '�J�E���^
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim bError As Boolean           '�t���O
    Dim bRet As Boolean             '�߂�l
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^
  
    '�߂�l������
    sCopyWRKtoNOW = True
    
    '****************************
    '* �t�@�C�����X�g���R�s�[���� *
    '****************************
    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
                                    '���[�N�t�H���_���t�@�C�������쐬����
    sDstFileName = FolderName(1, FolderSyubetu) & "\" & MN_FILELIST
                                    '���s�t�H���_���t�@�C�������쐬����
    'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������   'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then     '�t�@�C���̌���������   'V1.20.0.1 ADD
        '�t�@�C�����X�g���u���[�N�v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
        FileCopy sSrcFileName, sDstFileName
    Else                                '�t�@�C�������݂��Ȃ�
     '�u���[�N�t�H���_�t�@�C�����X�g�Ȃ��v�|�b�v�A�b�v��ʕ\��
     MsgBox "�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
             & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
             vbOKOnly + vbExclamation, _
             TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
     sCopyWRKtoNOW = False
     Set objFso = Nothing    'V1.20.0.1 ADD
     Exit Function                   '�������I������
    End If

    bError = False                  '�G���[�t���O���u�U�v�ɂ���
    For i = 0 To UBound(FileList) - 1
                                    '�t�@�C�����X�g�ꗗ�����J��Ԃ�
        sSrcFileName = FolderName(0, FolderSyubetu) & "\" & FileList(i)
                                    '���[�N�t�H���_���t�@�C�������쐬����
        sDstFileName = FolderName(1, FolderSyubetu) & "\" & FileList(i)
                                    '���s�t�H���_���t�@�C�������쐬����

        '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
        'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������   'V1.20.0.1 DEL
        If objFso.FileExists(sSrcFileName) = True Then   '�t�@�C���̌���������   'V1.20.0.1 ADD
            '�t�@�C�����u���[�N�v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
            FileCopy sSrcFileName, sDstFileName
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    Exit Function                           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    Select Case Err.Number
        Case 53 '�u���[�N�����s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
            MsgBox "�ُ�I�����܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
            
            sCopyWRKtoNOW = False
            Set objFso = Nothing    'V1.20.0.1 ADD
            Exit Function
        Case Else
                ' ���̃G���[�����������ɋL�q���܂��B
    End Select
    sCopyWRKtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyNOWtoOLD
'//  �@�\����  : ���s�o�[�W�����ۑ�����
'//  �@�\�T�v  : ���s�t�H���_���̃t�@�C�����A���t�H���_�ɃR�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sCopyNOWtoOLD() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim sSrcFileName As String          '�R�s�[���t�@�C���̃t���p�X��
    Dim sDstFileName As String          '�R�s�[��t�@�C���̃t���p�X��
    Dim iResponse As Integer            'MsgBox�{�^���R�[�h
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    On Error GoTo ErrorHandler              '�G���[�n���h���ݒ�
  
    '�߂�l������
    sCopyNOWtoOLD = True
   
    '���s�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'
'                '���s�t�H���_���t�@�C�������쐬����
'                sSrcFileName = gstrMyPath & MyName
'
'                '���t�H���_���t�@�C�������쐬����
'                sDstFileName = FolderName(2, FolderSyubetu) & "\" & MyName
'
'                '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
'                FileCopy sSrcFileName, sDstFileName
'
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                '���s�t�H���_���t�@�C�������쐬����
                sSrcFileName = gstrMyPath & MyName

                '���t�H���_���t�@�C�������쐬����
                sDstFileName = FolderName(2, FolderSyubetu) & "\" & MyName

                '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
                FileCopy sSrcFileName, sDstFileName

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
           ' �u���[�N�����s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
            MsgBox "�ُ�I�����܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
    
    sCopyNOWtoOLD = False
    
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyOLDtoNOW
'//  �@�\����  : ���o�[�W�����ɖ߂�����
'//  �@�\�T�v  : ���t�H���_���̃t�@�C�����A���s�t�H���_�ɃR�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sCopyOLDtoNOW() As Boolean
    Dim i As Integer                '�J�E���^
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim bError As Boolean           '�G���[�t���O
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler
    
    '�����l�ݒ�
    sCopyOLDtoNOW = True

    '****************************
    '* �t�@�C�����X�g���R�s�[���� *
    '****************************
    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '�u���v�t�H���_���t�@�C�������쐬����
    sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '�u���s�v�t�H���_���t�@�C�������쐬����
    'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������  'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then '�t�@�C���̌���������   'V1.20.0.1 ADD
        '�t�@�C�����X�g���u���v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
        FileCopy sSrcFileName, sDstFileName
    Else
       '�u���t�H���_�t�@�C�����X�g�Ȃ��v�|�b�v�A�b�v��ʕ\��
        MsgBox "�u���v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
                   & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '�������I������
    End If

    bError = False                  '�G���[�t���O���u�U�v�ɂ���
    For i = 0 To UBound(FileList) - 1
                                    '�t�@�C�����X�g�����J��Ԃ�
        '���t�H���_���t�@�C�������쐬����
        sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)

        '���s�t�H���_���t�@�C�������쐬����
        sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)

        '���t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
        'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������  'V1.20.0.1 DEL
        If objFso.FileExists(sSrcFileName) = True Then '�t�@�C���̌���������   'V1.20.0.1 ADD
            '�t�@�C�����u���v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
            FileCopy sSrcFileName, sDstFileName
        Else                                '�t�@�C�������݂��Ȃ�
            bError = True                   '�G���[�t���O���u�^�v�ɂ���
        End If
    Next
    If bError = True Then
        '�u���t�H���_�t�@�C�����X�g�o�^�Ȃ��v�|�b�v�A�b�v��ʕ\��
        MsgBox "�u���v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
                   & Chr(vbKeyReturn) & "�t�@�C�����X�g�ɓo�^����Ă��āA���݂��Ȃ��t�@�C��������܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function
    End If

    Set objFso = Nothing    'V1.20.0.1 ADD
    Exit Function       '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    '�u�������s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  �������s �R�s�["
        
    sCopyOLDtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sOldFolderRemove
'//  �@�\����  : ���t�H���_���t�@�C���폜����
'//  �@�\�T�v  : ���t�H���_���̃t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sOldFolderRemove() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^
   
   '�߂�l������
    sOldFolderRemove = True
 
    '�u���s�v�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    gstrMyPath = FolderName(2, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'                '�t�@�C�����폜����
'                Kill gstrMyPath & MyName
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
                '�t�@�C�����폜����
                Kill gstrMyPath & MyName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    Exit Function           '�������I������

ErrorHandler:   ' �G���[�������[�`���B
    '�u���[�N�����s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
    '�u�����ް�ޮ݁F���t�H���_̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLDFILE_DELETE_ERROR, lngErrCode)

    sOldFolderRemove = False
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sNowFolderRemove
'//  �@�\����  : ���s�t�H���_���̃t�@�C���폜����
'//  �@�\�T�v  : ���s�t�H���_���̃t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sNowFolderRemove() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END

    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    '�����l�ݒ�
    sNowFolderRemove = True
    
    '�u���s�v�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'
'                Kill gstrMyPath & MyName        '�t�@�C�����폜����
'
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                Kill gstrMyPath & MyName        '�t�@�C�����폜����

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    '�u�������s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  �������s �R�s�["

    '�u�����ް�ޮ݁F���s�t�H���_̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOWFILE_DELETE_ERROR, lngErrCode)

    sNowFolderRemove = False
    
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sWrkFolderRemove
'//  �@�\����  : ���[�N�t�H���_���t�@�C���폜����
'//  �@�\�T�v  : ���[�N�t�H���_���̃t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    '�����l�ݒ�
    sWrkFolderRemove = True
   
    '���[�N�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    gstrMyPath = FolderName(0, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'                '�t�@�C�����폜����
'                Kill gstrMyPath & MyName
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
                '�t�@�C�����폜����
                Kill gstrMyPath & MyName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END

    '�u���[�N�N���A����I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "���A" _
               & Chr(vbKeyReturn) & "�S�č폜���܂����B", _
               vbOKOnly + vbExclamation, _
               TitleBox(FolderSyubetu) & "  ���[�N �N���A"

    Exit Function '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    '�u���[�N�N���A�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbCritical, _
           "���[�N �N���A"
           
   '�u�����ް�ޮ݁Fܰ�̫���̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeOutPutFile
'//  �@�\����  : �}�̏o�͏������s���B
'//  �@�\�T�v  : �}�̏o�̓t�@�C���쐬�Əo�͂��s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeOutPutFile()
'   Dim sOutFileName As String '�}�̏o�̓t�@�C����[��ʕ�]
'   Dim iFileNumber As Integer '�t�@�C���ԍ�
'   Dim i As Integer           '�J�E���^�[
'   Dim bFlag As Boolean       '�t���O
'   Dim iResponse As Integer   'MsgBox�߂�l
'   Dim lngErrCode As Long     '�G���[�R�[�h
'   Dim fso         As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
'   Dim strWriteDir As String               '�o�͐�t�H���_
'
'   On Error Resume Next 'V1.21.0.1 ADD
'
'  '�t�H���_�I�𕔂Ɏw��L���`�F�b�N
'  bFlag = False                                 '�t���O���u�U�v�ɂ���
'  For i = 0 To 2                                '�t�H���_�����J��Ԃ�
'     If chkFolder(i).Value = CHECKBOX_ON Then   '�u�H�H�v�t�H���_���w�肳��Ă���
'        bFlag = True                            '�t���O���u�^�v�ɂ���
'        Exit For                                '���[�v�𔲂���
'     End If
'  Next
'
'  If bFlag = False Then                       '�t�H���_�w�薳��
'     If gStrCurrentForm = sFormName_EJVer Then
'       '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
'         MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
'                 vbOKOnly + vbExclamation, _
'                 "EG-R�������D�@ �o�[�W�����Ǘ�"
'     Else
'       '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
'         MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
'                vbOKOnly + vbExclamation, _
'                "NEG�������D�@ �o�[�W�����Ǘ�"
'     End If
'         '�����𔲂���
'     Exit Function
'   End If
'
'  '�t�H���_�I���|�b�v�A�b�v��ʕ\��
''  strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")                         'V1.12.0.1 DEL
'  strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.12.0.1 ADD
'
'  '�w��t�H���_�Ȃ�
'  If Len(strWriteDir) = 0 Then
'       Exit Function
'  End If
'
'  '�R�s�[��t�H���_�̗L���m�F
'  If fso.FolderExists(strWriteDir) = False Then
'     '�R�s�[��t�H���_�쐬
'     fso.CreateFolder (strWriteDir)
'  End If
'
'   '�������t�H�[���ɂ��A�}�̏o�͂���t�@�C�����쐬
'   If gStrCurrentForm = sFormName_EJVer Then
'       '���\�[�X�I�𕔕���
'       Select Case FolderSyubetu
'        Case 0      '����CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_EJHANTEIPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJHANTEIPRO      'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJHANTEIPRO 'V1.8.0.1 ADD
'        Case 1      '���C��CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINPRO   'V1.8.0.1 ADD
'        Case 2      '�T�uCPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_EJSUBPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJSUBPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJSUBPRO   'V1.8.0.1 ADD
'        Case 3      '���C��CPU-OS
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINOS        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINOS   'V1.8.0.1 ADD
'        Case 4      '�\��1
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI1         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
'        Case 5      '�\��2
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI2
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI2         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI2    'V1.8.0.1 ADD
'        End Select
'  Else
'       '���\�[�X�I�𕔕���
'       Select Case FolderSyubetu
'        Case 0      '����CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJHANTEIPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJHANTEIPRO      'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJHANTEIPRO 'V1.8.0.1 ADD
'        Case 1      '���C��CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJMAINPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJMAINPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJMAINPRO   'V1.8.0.1 ADD
'        Case 2      '�T�uCPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJSUBPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJSUBPRO         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJSUBPRO    'V1.8.0.1 ADD
'        Case 3      '���C��CPU-OS
'          sOutFileName = PATH_WORK & VER_TXT_NJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_NJMAINOS         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJMAINOS    'V1.8.0.1 ADD
'        Case 4      '�\��1
'          sOutFileName = PATH_WORK & VER_TXT_NJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_NJYOBI1          'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJYOBI1     'V1.8.0.1 ADD
'        Case 5      '�\��2
'          sOutFileName = PATH_WORK & VER_TXT_NJYOBI2
'          'strWriteDir = strWriteDir & VER_TXT_NJYOBI2          'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJYOBI2     'V1.8.0.1 ADD
'        End Select
'  End If
'
'  iFileNumber = FreeFile              '���g�p�̃t�@�C���ԍ����擾����
'
'  '�Ώۃt�@�C�����I�[�v������B
'  Open sOutFileName For Output Access Write As #iFileNumber
'
'  For i = 0 To lstKan.ListCount - 1
'  '���X�g�{�b�N�X�ɕ\������Ă��镪�����A�������ށB
'       Print #iFileNumber, lstKan.List(i) & Chr(vbKeyReturn)
'  Next
'
'  '�Ώۃt�@�C�����N���[�Y����B
'  Close #iFileNumber
'
'  '�t�@�C���̗L���m�F
'  If fso.FileExists(sOutFileName) = False Then
'     '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
'     MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
'     Exit Function
'  End If
'
'  On Error GoTo COPY_ERROR
'  '�t�@�C���R�s�[
'  fso.CopyFile sOutFileName, strWriteDir
'  '�u�}�̏o�͐���I���v�|�b�v�A�b�v��ʕ\��
'  'V1.8.0.1 DEL START
'  'iResponse = MsgBox("����I�����܂����B", vbOKOnly, _
'  '                   "�o�͌���")
'  'V1.8.0.1 DEL END
'  MsgBox "����I�����܂����B", vbInformation, "�o�͌���"   'V1.8.0.1 ADD
'
'  '�u�����ް�ޮ݁F�}�̏o�͏�������v���O�o��
'  Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_OK, 0)
'
'  Set fso = Nothing
'
'  Exit Function
'
''*******************************
''VB�G���[����
'COPY_ERROR:
'        '�����ُ�̏ꍇ�A�o�͌��ʃ|�b�v�A�b�v(�ُ�)�\��
'        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
'        '�u�����ް�ޮ݁F�}�̏o�͏����ُ�v���O�o��
'        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_ERROR, lngErrCode)
'        Set fso = Nothing
''*******************************
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v������
'//  �@�\�T�v  : ���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
   On Error Resume Next
    
    '�ėp���[����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmJVer.Caption, False
    End If
End Sub
'V1.4.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGetGoukiNo
'//  �@�\����  : �_�����@�ԍ����擾����B
'//  �@�\�T�v  : GATE.INI���_�����@�ԍ����擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer            [OUT]�擪���@�ԍ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-04-18   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Function pfGetGoukiNo() As Integer                        'V1.6.0.1 DEL
Private Function pfGetGoukiNo(iGoukiCunter As Integer) As Integer  'V1.6.0.1 ADD

    Dim lngRet As Long          '�֐��̕Ԃ�l
    Dim iGate As Integer        '����INDEX
    Dim j As Integer            '���[�NINDEX
    Dim sGoukiNo As String      'GLT�t�@�C�����R�[�h�f�[�^(���@�ԍ��\������)
    Dim cWork As Byte           '���[�N�G���A
    Dim lngErrCode As Long      '�G���[�R�[�h
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim intGLTFileNo As Integer     '̧�ٔԍ�
'   Dim iGoukiCunter As Integer�@�@ 'V1.6.0.1 DEL
    

    On Error Resume Next

 '   For iGoukiCunter = 1 To MAX_GATE_NO   'V1.6.0.1 DEL
         '�������D�@���擾
         sKeyName = "gate" & Format(iGoukiCunter, "00")
         iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                        sKeyName, _
                                        DEFAILT, sGateData, Len(sGateData), _
                                        PATH_GATE_FILE)
         If iRet = 0 Then
            '�uEG-R�������D�@�o�[�W�����Ǘ���ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
            pfGetGoukiNo = 0
            Exit Function
         End If
             
         If Len(sGateData) <> 0 Then
            '�f�[�^�̎擾
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
'        If Trim(sFData(4)) = EGR Then                          'V1.0.6.1 DEL
        If Trim(sFData(4)) = EGR Or Trim(sFData(4)) = NEG Then  'V1.0.6.1 ADD
           pfGetGoukiNo = iGoukiCunter
           Exit Function
        End If
 'Next 'V1.6.0.1 DEL
End Function
'V1.4.0.1 ADD END

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetChkFile
'//  �@�\����  : ���[�N�����s�R�s�[�Ŏg�p���鐳�����`�F�b�NINI�Ǎ���
'//  �@�\�T�v  : INI�t�@�C���ɂ̓��e���G���A�ɓW�J����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String    �Z�N�V������
'//              String    �L�[��
'//              String    �t�@�C����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String    �������`�F�b�NINI�̓��e�i�ُ펞�̓u�����N�j
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sSetChkFile(sSec As String, sKey As String, sFilePath As String) As String

    Dim iRet As Integer             '�֐��̖߂�l
    Dim sIni_Data As String * 128   'INI�t�@�C�����1�s���擾
    
    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
    Dim sErrExe As String              '�G���[���O�pINI�g���q
    Dim lngErrCode As Long             '�G���[�R�[�h

    
    '�G���[���[�`����錾
    On Error Resume Next

    'ini�t�@�C���擾
    sIni_Data = ""
    iRet = GetPrivateProfileString(sSec, sKey, DEFAILT, sIni_Data, Len(sIni_Data), sFilePath)
    
    '�ُ폈��
    If iRet = 0 Then
        
        '���O�o�́uINI�t�@�C���Ǎ��ُ�v
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_INI_READ, lngErrCode)
        '���O�o�́@���t�@�C����
        Call psFileNameGet(sFilePath, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)
        '���O�o�́@���L�[��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��Key:" & sKey, lngErrCode)
        
    End If
    
    sSetChkFile = Left$(sIni_Data, iRet)
    
End Function
'V1.20.0.1 ADD END

