VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKVer 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�o�[�W�����Ǘ��i�Ď��Ձj"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -60
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer tmrLogTimer 
      Left            =   11520
      Top             =   1560
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   11520
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   �� �� ���s     �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   19
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyWork_Jikko 
      Caption         =   " ���[�N �� ���s �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   18
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "���[�N�N���A"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   17
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " �}�� �� ���[�N �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   16
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton CmdRemove 
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
      Height          =   615
      Left            =   9360
      TabIndex        =   14
      Top             =   6960
      Width           =   2415
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
      Height          =   6060
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   8655
   End
   Begin VB.CommandButton cmdOutPut 
      Caption         =   " �o�[�W������� �}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   5
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdRefresh 
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
      Height          =   615
      Left            =   9360
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame fraVersion 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   9360
      TabIndex        =   7
      Top             =   480
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1380
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
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1380
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
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " �o�[�W�����Ǘ�   ��ʂ֖߂�"
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
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   9000
      Top             =   7200
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�����Ď��Ճo�[�W�����Ǘ�"
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
      TabIndex        =   15
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblKansibanVersion 
      Caption         =   "�S�̃o�[�W����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�t�@�C����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   2535
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
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����(�޲�)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�X�V���t"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�o�[�W�����ԍ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   8
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmKVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmKVer.frm
'//  �p�b�P�[�W���F�o�[�W�����Ǘ�(�Ď���)���
'//
'//  �T�v�F�o�[�W�����Ǘ�(�Ď���)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �EEG10�ێ���A�o�[�W�����Ǘ�(�Ď���)���(frmKVer.frm)���p�B
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.100�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.100�z�y����TR-No.184�z
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.273�C���Ή��z
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.22�C���Ή��z
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                 �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y2,3�����Ή�
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//     REVISIONS :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS �F(EG20 V30.3.0.1)2014-10-23  CODED BY  [TCC] T.Nakajima
'//                  �k���V�����t�F�[�Y�Q�Ή��i�}�̎�O���G���[�Ή��j
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
Dim uVersion() As MN_VERSION_LIST       '�o�[�W�������i�[�G���A

'�t�H���_��ʕ�
Public mlngChkFolderType        As Long

Private Const VERSION_STA = 28
Private Const VERSION_SIZE = 12
Private Const VERMOJI_STA = 1
Private Const FOLDER_STS = 27
Private Const HIDUKE_STA = 40
Private Const VERSION_END = 30

' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
Private Const HEADERTITLE_WRK = "�����Ď��Ճo�[�W�����i���[�N�j�F"
Private Const HEADERTITLE_NOW = "�@�@�@�@�@�@�@�@�@�@�i���s�j�@�F"
Private Const HEADERTITLE_OLD = "�@�@�@�@�@�@�@�@�@�@�i���j�@�@�F"
Private Const HEADERVERSION_NON = "--.--.--.--"
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��

' EG20 V3.3.0.1�y����TR-No.184�z �ǉ��J�n
Private Const APL_INTERVAL = 390000         ' �A�v���N���^�C�}�f�t�H���g�l
Private Const LOG_INTERVAL = 30000          ' ���O�N���^�C�}�f�t�H���g�l(30�b)
Dim lngAplMAX_Time As Long                  ' INI�擾�ݒ�l�i�`�o�k�j
Dim lngLogMAX_Time As Long                  ' INI�擾�ݒ�l�i���O�j
Dim lngtime        As Long                  ' ���݃^�C�}�l
Dim lngChangeKind  As Long                  ' �o�[�W�����ؑ֎��
' EG20 V3.3.0.1�y����TR-No.184�z �ǉ��I��

' EG20 V7.3.0.1�yEG20_KANSI03_01�zADD START
Private Const DESHU_ID = 242                              '�f�W1�R�[�iID
Private Const WAIT_TIME_OUT = 180000                      '�^�C���A�E�g�l�i�R���j
Private Const DESHU_CONNECT = 1                           '�f�W�ڑ��ݒ�
Private Const GATE_CONNECT = 2                            '���D�@�ڑ��ݒ�
Private Const ERROR_TUSHIN_DISP = 1                       '�ʐM�ؒf�ُ탁�b�Z�[�W�\��
Private Const ERROR_MISOU_DISP = 2                        '�����f�[�^�o�͎��s���b�Z�[�W�\��
Private Const ERROR_END_DISP = 3                          '�ُ�I�����b�Z�[�W�\��
Private udtMail          As MAIL_CONECT_CMD               '�ʐM�ݒ�v��CMD
Public miCornerNo        As Integer                       '�R�[�i�[�ԍ�
Public mbMisouResult     As Boolean                       '�����f�[�^�쐬���ʁ@TRUE�F����@FALSE�F�ُ�
Public miErrorSts        As Integer                       '�ُ펞�ʐM���
Public miErrorDisp       As Integer                       '�ُ펞�\������
Private byDeshuCnctSet(CONECT_CORNER_MAXINDEX)  As Byte   '�f�W�ؗ��ݒ�
Private byGateCnctSet(CONECT_JIKAI_CHK_MAX)   As Byte     '�����ؗ��ݒ�
' EG20 V7.3.0.1�yEG20_KANSI03_01�zADD END

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdClear_Click
'//  �@�\����    �F���[�N�N���A
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 �y�^���\�����P�Ή��z
'//  REVISIONS   �F (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F (EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()
   
    Dim iResponse As Integer         ' MsgBox�{�^���R�[�h
    Dim bResult As Boolean           ' ��������
    
    On Error Resume Next

    '�u�o�[�W�����Ǘ���ʁF���[�N�N���A�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_CREA_BUTTOM, 0)
    
    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�u���[�N�v�t�H���_���̃t�@�C�����A" _
           & Chr(vbKeyReturn) & "�S�č폜���܂��B    ��낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "���[�N �N���A")
    
    If iResponse <> vbCancel Then
        sCmdBtnEnabled False                        ' ��ʑ���s��
        '[�͂�] �{�^����I�������ꍇ
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '���[�N�t�H���_���̃t�@�C�����폜����
        bResult = sWrkFolderRemove
        sCmdBtnEnabled True                         ' ��ʑ����
        If bResult = True Then
            ' �Ď��Ղ̃o�[�W��������\������
            Call psVersionDisp
            
' EG20 V5.8.0.1�폜�J�n
'            ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'            Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
            ' �^����ԍX�V
'            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_KANSI, BOOTINFO_UNKAI_NASHI)      ' EG20 V5.11.0.1�폜
            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_KANSI, BOOTINFO_UNKAI_CLEAR)       ' EG20 V5.11.0.1�ǉ�
            Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMEKANSI, BOOTINFO_UNKAI_NASHI)
' EG20 V5.8.0.1�ǉ��I��

' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD START
            ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
            Call funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_CLEAR)
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD END

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        End If
    End If

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1�ǉ��I��

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdCopyBaitai_Work_Click
'//  �@�\����    �F�}�́����[�N�R�s�[
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()

    On Error Resume Next
    '�u�}�́����[�N�R�s�[�v�{�^���̏ꍇ�B
    '�u�o�[�W�����Ǘ���ʁF�}�́����[�N�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_USB_COPY_WRK_BUTTOM, 0)
    
    sCmdBtnEnabled False                        ' ��ʑ���s��
    '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����
    Call sFDInstall
    sCmdBtnEnabled True                         ' ��ʑ����
    Call psVersionDisp

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdCopyOld_Jikko_Click
'//  �@�\����    �F�������s�R�s�[
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.184�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.22�C���Ή��z
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.372�C���Ή��z
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    
    Dim udtSendData As ML_KANEND_REQ_CMD  ' ���ʃG���A
    Dim lngSendSize As Long               ' ���M���郁�[���T�C�Y
    Dim lngErrCode  As Long               ' �G���[�R�[�h
    Dim bRet        As Boolean            ' ���[�����M�����߂�l
    Dim iResponse   As Integer            ' MsgBox�{�^���R�[�h
    Dim iAplChk     As Integer            ' �A�v���N���`�F�b�N�߂�l    'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
    
    On Error Resume Next
    
    '�u�o�[�W�����Ǘ���ʁF�������s�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OLD_COPY_NOW_BUTTOM, 0)

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1�ǉ��I��

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�u���v�t�H���_�̓��e���A" _
            & Chr(vbKeyReturn) & "�u���s�v�t�H���_�ɖ߂����Ƃɂ��A" _
            & Chr(vbKeyReturn) & "�����Ď��Ղ̈ꐢ��O�o�[�W��������s�o�[�W�����Ƃ��܂��" _
            & Chr(vbKeyReturn) & "��낵���ł����H", _
           (vbOKCancel + vbExclamation), _
           "�������s�@�R�s�[")
    If iResponse = vbCancel Then
        Exit Sub
    End If
        
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��J�n
    ' ���o�[�W�����t�H���_�ɑ�\�o�[�W�����t�@�C�������݂��Ȃ��ꍇ�ُ͈�Ƃ���B
    ' ���o�[�W�����EKANSI�E�����Ď��ՁE
    bRet = dllCheckAplVersion(4, PATH_KANSI, 2)
    If bRet = False Then
        MsgBox "�ُ�I�����܂����B", vbCritical, "�������s�@�R�s�["
        Exit Sub
    End If
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��I��
        
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD START
    ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONDOWN)
    If bRet = False Then
        MsgBox "�ُ�I�����܂����B", vbCritical, "�������s�@�R�s�["
        Exit Sub
    End If

    ' �I���m�F
    iResponse = MsgBox("���s�R�s�[��K�p���邽�߂ɓ����Ď��Ղ�" & Chr(vbKeyReturn) _
                        & "�ċN�����܂����H", _
                        vbOKCancel + vbExclamation, _
                        "�������s�@�R�s�[")
    If iResponse = vbCancel Then
        Exit Sub
    End If
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD END
        
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL START
'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��J�n
'    ' �����Ď��Ղ��N�����̏ꍇ�Ƀ��b�Z�[�W�{�b�N�X��\������B
'    iAplChk = CheckAppStart(PROC_KANRI)
'    If iAplChk <> 0 Then
''EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��I��
'        '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
'        iResponse = MsgBox("�����Ď��դ�h�c�t��k�c�t�A�v���P�[�V������" _
'                & Chr(vbKeyReturn) & "�I�����܂��B��낵���ł����H", _
'               (vbOKCancel + vbExclamation), _
'               "�I���m�F")
'
'        If iResponse = vbCancel Then
'            Exit Sub
'        End If
'    End If  'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL END

' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
' AplVersionChangeProc�Ƀ��W���[����
'    ' ���[���̑��M���e��ҏW����
'    udtSendData.udtlHeader.dwId = ML_ID_KANEND_REQ      ' ���[���h�c�@=�h"�Ď����u�I���v��"
'    udtSendData.udtlHeader.dwSize = MlSize.KANEND_REQ   ' ���[���T�C�Y=�h"�Ď����u�I���v��"
'    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' ���M���v���Z�X�h�c=�h�ێ�h
'    udtSendData.udtlHeader.dwSubArea = 0                ' �⏕���@=�@0
'
'    udtSendData.dwStartProc = ML_DT_VERSIONDOWN         ' �N���v���Z�X��� = �o�[�W�����_�E��
'    ' ���M�T�C�Y��ݒ肷��B
'    lngSendSize = udtSendData.udtlHeader.dwSize
'
'    ' �ă}�ɑ΂��āA�ݒ���v�����[���𑗐M����B
'    bRet = DssSendMail(MAIL_SLOT_KANRI, lngSendSize, udtSendData.udtlHeader)
'    ' ���[���𐳏�ɑ��M�������̃��O
'    If bRet = False Then
'        '�u�ݒ���v�����[�����M�ُ�v���O�o��
'        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, lngErrCode)
'    Else
'        '�u�ݒ���v�����[�����M����v���O�o��
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, 0)
'    End If
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
' EG20 V3.3.0.1�y����TR-No.184�z�폜�J�n
'' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
'    ' �A�v���P�[�V�����o�[�W�����ؑ֎��s����
'    If (AplVersionChangeProc(ML_DT_VERSIONDOWN) = False) Then
'        ' // �ێ���I������B
'        Call psEndHoshuProc
'        '�ێ�v���Z�X�I��
'        End
'    End If
'' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
' EG20 V3.3.0.1�y����TR-No.184�z�폜�I��
' EG20 V3.3.0.1�y����TR-No.184�z�ǉ��J�n

    sCmdBtnEnabled False                            ' ��ʑ���s��
    ' �����Ď��ՂփA�v���I���v���̑��M
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
               vbOKOnly + vbExclamation, _
               "�Ď��Ճo�[�W�����Ǘ�"
        sCmdBtnEnabled True                         ' ��ʑ����
    Else

        lngtime = MN_MAIL_INTERVAL                  ' ���݃^�C�}�l������
        tmrAplTimer.Enabled = True                  ' ���݃^�C�}�N��
    
        lngChangeKind = ML_DT_VERSIONDOWN           ' �ؑ֎�ʂ�ݒ�
    End If
' EG20 V3.3.0.1�y����TR-No.184�z�ǉ��I��


End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdCopyWork_Jikko_Click
'//  �@�\����    �F���[�N�����s�R�s�[
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.184�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.22�C���Ή��z
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.372�C���Ή��z
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//  REVISIONS    :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
    
    Dim udtSendData As ML_KANEND_REQ_CMD  ' ���ʃG���A
    Dim lngSendSize As Long               ' ���M���郁�[���T�C�Y
    Dim lngErrCode  As Long               ' �G���[�R�[�h
    Dim bRet        As Boolean            ' ���[�����M�����߂�l
    Dim iResponse   As Integer            ' MsgBox�{�^���R�[�h
    Dim iAplChk     As Integer            ' �A�v���N���`�F�b�N�߂�l    'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
    
    On Error Resume Next
    
    '�u�o�[�W�����Ǘ���ʁF���[�N�����s�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_COPY_NOW_BUTTOM, 0)

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1�ǉ��I��

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�u���[�N�v�t�H���_�̓��e���A" _
            & Chr(vbKeyReturn) & "�u���s�v�t�H���_�ɓo�^���邱�Ƃɂ��A" _
            & Chr(vbKeyReturn) & " �����Ď��Ղ̍ŐV�o�[�W�������A���s�o�[�W�����Ƃ��܂��B" _
            & Chr(vbKeyReturn) & "��낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "���[�N�����s �R�s�[")
    If iResponse = vbCancel Then
        Exit Sub
    End If
        
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��J�n
    ' ���[�N�o�[�W�����t�H���_�ɑ�\�o�[�W�����t�@�C�������݂��Ȃ��ꍇ�ُ͈�Ƃ���B
    ' ���[�N�o�[�W�����EKANSI�E�����Ď���
    bRet = dllCheckAplVersion(1, PATH_KANSI, 2)
    If bRet = False Then
        MsgBox "�ُ�I�����܂����B", vbCritical, "���[�N�����s �R�s�["
        Exit Sub
    End If
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��I��

' EG20 V7.3.0.1�yEG20_KANSI03_01�zADD START
        '���[�N�����s�R�s�[�O����
    bRet = fWorktoNow_Before1
    If bRet = False Then
        Exit Sub
    End If
' EG20 V7.3.0.1�yEG20_KANSI03_01�zADD END

' EG20 V7.3.0.1�yEG20_KANSI03_01�zDEL START
'' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD START
'    ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
'    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONUP)
'    If bRet = False Then
'        MsgBox "�ُ�I�����܂����B", vbCritical, "���[�N�����s �R�s�["
'        Exit Sub
'    End If
'
'    ' �I���m�F
'    iResponse = MsgBox("���s�R�s�[��K�p���邽�߂ɓ����Ď��Ղ�" & Chr(vbKeyReturn) _
'                        & "�ċN�����܂����H", _
'                        vbOKCancel + vbExclamation, _
'                        "���[�N�����s �R�s�[")
'    If iResponse = vbCancel Then
'        Exit Sub
'    End If
'' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD END
'
'' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL START
'''EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��J�n
''    ' �����Ď��Ղ��N�����̏ꍇ�Ƀ��b�Z�[�W�{�b�N�X��\������B
''    iAplChk = CheckAppStart(PROC_KANRI)
''    If iAplChk <> 0 Then
'''EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��I��
''        '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
''        iResponse = MsgBox("�����Ď��դ�h�c�t��k�c�t�A�v���P�[�V������" _
''                & Chr(vbKeyReturn) & "�I�����܂��B��낵���ł����H", _
''               vbOKCancel + vbExclamation, _
''               "�I���m�F")
''
''        If iResponse = vbCancel Then
''            Exit Sub
''        End If
''    End If  'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
'' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL END
'
'' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
'' AplVersionChangeProc�Ƀ��W���[����
''    ' ���[���̑��M���e��ҏW����
''    udtSendData.udtlHeader.dwId = ML_ID_KANEND_REQ      ' ���[���h�c�@=�h"�Ď����u�I���v��"
''    udtSendData.udtlHeader.dwSize = MlSize.KANEND_REQ   ' ���[���T�C�Y=�h"�Ď����u�I���v��"
''    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' ���M���v���Z�X�h�c=�h�ێ�h
''    udtSendData.udtlHeader.dwSubArea = 0                ' �⏕���@=�@0
''
''    udtSendData.dwStartProc = ML_DT_VERSIONUP           ' �N���v���Z�X��� = �o�[�W�����A�b�v
''    ' ���M�T�C�Y��ݒ肷��B
''    lngSendSize = udtSendData.udtlHeader.dwSize
''
''    ' �ă}�ɑ΂��āA�ݒ���v�����[���𑗐M����B
''    bRet = DssSendMail(MAIL_SLOT_KANRI, lngSendSize, udtSendData.udtlHeader)
''    ' ���[���𐳏�ɑ��M�������̃��O
''    If bRet = False Then
''        '�u�ݒ���v�����[�����M�ُ�v���O�o��
''        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
''        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, lngErrCode)
''    Else
''        '�u�ݒ���v�����[�����M����v���O�o��
''        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, 0)
''    End If
'' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
'' EG20 V3.3.0.1�y����TR-No.184�z�폜�J�n
''' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
''    ' �A�v���P�[�V�����o�[�W�����ؑ֎��s����
''    If (AplVersionChangeProc(ML_DT_VERSIONUP) = False) Then
''        ' // �ێ���I������B
''        Call psEndHoshuProc
''        '�ێ�v���Z�X�I��
''        End
''    End If
''' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
'' EG20 V3.3.0.1�y����TR-No.184�z�폜�I��
'' EG20 V3.3.0.1�y����TR-No.184�z�ǉ��J�n
'
'    sCmdBtnEnabled False                            ' ��ʑ���s��
'    ' �����Ď��ՂփA�v���I���v���̑��M
'    bRet = pubFuncAplEndRequest()
'    If bRet = False Then
'        MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
'               vbOKOnly + vbExclamation, _
'               "�Ď��Ճo�[�W�����Ǘ�"
'        sCmdBtnEnabled True                         ' ��ʑ����
'    Else
'
'        lngtime = MN_MAIL_INTERVAL                  ' ���݃^�C�}�l������
'        tmrAplTimer.Enabled = True                  ' ���݃^�C�}�N��
'
'        lngChangeKind = ML_DT_VERSIONUP             ' �ؑ֎�ʂ�ݒ�
'    End If
'' EG20 V3.3.0.1�y����TR-No.184�z�ǉ��I��
' EG20 V7.3.0.1�yEG20_KANSI03_01�zDEL END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����Ǘ�(�Ď���)���(�A�N�e�B�u��)
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
 
    '�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����Ǘ�(�Ď���)���(�f�B�A�N�e�B�u��)
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
    
    '�^�C�}���~����
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����Ǘ�(�Ď���)���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.184�z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
   
   '�u�Ď��Ճo�[�W�����Ǘ���ʁF�\���v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_KANRI_GAMEN_START, 0)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    '������
    lstKan.Clear
    mlngChkFolderType = 0

    '�t�H���_�I�𕔁F�I��L��
    chkFolder(0).Value = 1
    chkFolder(1).Value = 1
    chkFolder(2).Value = 1
    
    mlngChkFolderType = 7
    
' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
''    �Ď��Ղ̃o�[�W�����ԍ���\������
'    sKansibanVersionSet
'
''   �o�[�W�������̃��X�g�{�b�N�X���쐬����
'    fMakeListbox
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    ' �����Ď��Ղ̃o�[�W��������\������
    Call psVersionDisp
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   
'   ���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

' EG20 V3.3.0.1�y����TR-No.184�z �ǉ��J�n
    ' INI�t�@�C�����A�v���N���^�C�}�l���擾
    lngAplMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                       APL_INTERVAL, HOSHU_FILE)
    ' �擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
    If lngAplMAX_Time = 0 Then
       lngAplMAX_Time = APL_INTERVAL
    End If

    ' �^�C�}�l�ݒ�
    tmrAplTimer.Interval = MN_MAIL_INTERVAL
    tmrAplTimer.Enabled = False

    ' INI�t�@�C����胍�O�N���^�C�}�l���擾
    lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
    ' �擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
    If lngLogMAX_Time = 0 Then
       lngLogMAX_Time = LOG_INTERVAL
    End If

    ' �^�C�}�l�ݒ�
    tmrLogTimer.Interval = MN_MAIL_INTERVAL
    tmrLogTimer.Enabled = False

    ' �ؑ֎�ʂ�������
    lngChangeKind = 0
' EG20 V3.3.0.1�y����TR-No.184�z �ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : chkFolder_Click
'//  �@�\����  : �u�t�H���_�`�F�b�N�v�`�F�b�N��������
'//  �@�\�T�v  : �t�H���_�I�𕔃`�F�b�N���s���B
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
Private Sub chkFolder_Click(Index As Integer)
  
  Dim ValueCnt                As Integer

    '��ނɂ���đ����l��ύX����
    ValueCnt = 0
    '���[�N
    If Index = 0 Then
        ValueCnt = 1
    '���s
    ElseIf Index = 1 Then
        ValueCnt = 2
    '��
    ElseIf Index = 2 Then
        ValueCnt = 4
    End If

    '�`�F�b�N���͂����ꂽ��
    If chkFolder(Index).Value = 0 Then
        mlngChkFolderType = mlngChkFolderType - ValueCnt
    '�`�F�b�N���ꂽ��
    ElseIf chkFolder(Index).Value = 1 Then
        mlngChkFolderType = mlngChkFolderType + ValueCnt
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdRefresh_Click
'//  �@�\����  : �u�\���X�V�v�t��������
'//  �@�\�T�v  : �ŐV�̏�Ԃ�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdRefresh_Click()
    Dim i As Integer        '�J�E���^�[
    Dim bFlag As Boolean    '�\���t�H���_�I���`�F�b�N(TRUE�F�`�F�b�N�L�BFALSE�F�`�F�b�N��)
   
    On Error Resume Next
    
    '�u�Ď��Ճo�[�W�����Ǘ���ʁF�\���X�V�t�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
 
    '�\���t�H���_�I���`�F�b�N���`�F�b�N�����ŏ���������B
    bFlag = False
    '�\���t�H���_�I���`�F�b�N�L�����`�F�b�N����B
    For i = 0 To 2
      If chkFolder(i).Value = CHECKBOX_ON Then
        '�P�ł��`�F�b�N�L��̏ꍇ�A�\���t�H���_�I���`�F�b�N���A�`�F�b�N�L�ɂ���B
         bFlag = True
          Exit For
       End If
    Next
   
    '�\���t�H���_�I���̃`�F�b�N���Ȃ��ꍇ�́A�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
    If bFlag = False Then
       MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
               vbOKOnly + vbExclamation, _
                "�Ď��Ճo�[�W�����Ǘ�"
        Exit Sub
    End If
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   
   '�o�[�W�������̃��X�g�{�b�N�X���쐬����
'    fMakeListbox           ' EG20 V2.1.0.1[Mainte_03_01]�폜
    Call psVersionDisp      ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdOutPut_Click
'//  �@�\����  : �u�o�[�W�������}�̏o�́v�t��������
'//  �@�\�T�v  : �\�����ꂽ�o�[�W��������}�̂ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.100�z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()
'*******************************
'VB�G���[����
On Error GoTo Error_cmdOutPut_Click
'*******************************
    Dim iRet        As Integer                '�߂�l
    Dim strCopySaki As String                 '�o�͐�t�@�C���p�X
    Dim strWriteDir As String                 '�o�͐�t�H���_
    Dim fso         As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim iFileNumber As Integer                '�t�@�C���ԍ�
    Dim iMaxLine As Integer                   '���X�g�{�b�N�X�̍s��
    Dim iLine As Integer                      '���X�g�{�b�N�X�̍s�J�E���^
    Dim sCopymoto As String                   '�o�͌��t�@�C���p�X
    Dim lngErrCode  As Long              '�G���[�R�[�h
    
    Dim strStationName       As String          ' �w����                ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim szCornerName         As String          ' �R�[�i����            ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim nNullIndex           As Integer         ' ���������[�N          ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim strWork              As String          ' ���[�N                ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim strFileName         As String           ' �t�@�C����            ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim bRet                As Boolean          ' �߂�l                ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�

   '�u�Ď��Ճo�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͖t�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OUTPUT, 0)

' EG20 V3.3.0.1 �y����TR-No.100�z�ǉ��J�n
    ' ���X�g�ɂP�����f�[�^���Ȃ��ꍇ�ُ͈�I��
    If lstKan.ListCount = 0 Then
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Exit Sub
    End If
' EG20 V3.3.0.1 �y����TR-No.100�z�ǉ��I��

' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    strStationName = gsGetStationEkiName
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)

    '�w��t�H���_�Ȃ�
    If Len(strWriteDir) = 0 Then
        Exit Sub
    End If

    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ �ێ��p�֐�:�����o�[�W�����t�@�C���i��ʕ\���p�j�쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllKansiCreateVerFile(mlngChkFolderType, MN_VERSI_FILE, VERLISTKIND_REPORT)
    ' �o�[�W�����t�@�C������
    If bRet Then
        '�u�o�[�W�������t�@�C���쐬����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    ' �o�[�W�����t�@�C�����s
    Else
        '�u�o�[�W�������t�@�C���쐬�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
       Exit Sub
    End If

    '�t�@�C���̗L���m�F
    If fso.FileExists(MN_VERSI_FILE) = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Exit Sub
    End If
    strFileName = Dir(MN_VERSI_FILE)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�J�n�i�����ړ��j
'    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
'    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
'
'    '�w��t�H���_�Ȃ�
'    If Len(strWriteDir) = 0 Then
'        Exit Sub
'    End If
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�I��

    '�R�s�[��t�H���_�̗L���m�F
    If fso.FolderExists(strWriteDir) = False Then
        '�R�s�[��t�H���_�쐬
        fso.CreateFolder (strWriteDir)
    End If

    '�R�s�[��t�@�C�����쐬
    strCopySaki = strWriteDir & "\" & strStationName & "_" & strFileName

    '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
    fso.CopyFile MN_VERSI_FILE, strCopySaki, True
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
''V1.7.0.1 DEL START
''    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
''    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")
''
''    '�w��t�H���_�Ȃ�
''    If Len(strWriteDir) = 0 Then
''        Exit Sub
''    End If
''V1.7.0.1 DEL END
'    iFileNumber = FreeFile              '���g�p�̃t�@�C���ԍ����擾����
'
'    sCopymoto = PATH_WORK + VER_TXT_NAME
'
'    '�o�[�W�����e�L�X�g�t�@�C�����I�[�v������B�t�@�C�����Ȃ���ΐV�K�ɍ쐬����B
'    Open sCopymoto For Output Access Write As #iFileNumber
'
'    iMaxLine = lstKan.ListCount
'    For iLine = 0 To lstKan.ListCount - 1
'        '���X�g�{�b�N�X�P�s�������o�[�W�����e�L�X�g�t�@�C���ɏ������ށB
'        Print #iFileNumber, lstKan.List(iLine) & Chr(vbKeyReturn)
'    Next
'    '�o�[�W�����e�L�X�g�t�@�C�����N���[�Y����B
'    Close #iFileNumber
'
''V1.7.0.1 DEL START
''    '�R�s�[��t�H���_�̗L���m�F
''    If fso.FolderExists(strWriteDir) = False Then
''        '�R�s�[��t�H���_�쐬
''        fso.CreateFolder (strWriteDir)
''    End If
''V1.7.0.1 DEL END
'
'   '�t�@�C���̗L���m�F
'    If fso.FileExists(sCopymoto) = False Then
'        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
'        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
'        Exit Sub
'    End If
'
''V1.7.0.1 ADD  START
'    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
''    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")                       'V1.12.0.1 DEL
'    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.12.0.1 ADD
'
'    '�w��t�H���_�Ȃ�
'    If Len(strWriteDir) = 0 Then
'        Exit Sub
'    End If
'
'    '�R�s�[��t�H���_�̗L���m�F
'    If fso.FolderExists(strWriteDir) = False Then
'        '�R�s�[��t�H���_�쐬
'        fso.CreateFolder (strWriteDir)
'    End If
''V1.7.0.1 ADD END
'
'    '�R�s�[��t�@�C�����쐬
'    strCopySaki = strWriteDir & "\" & VER_TXT_NAME
'
'    '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
'    fso.CopyFile sCopymoto, strCopySaki, True
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
  
    '�o�͌��ʃ|�b�v�A�b�v(����)�\��
    MsgBox "����I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
    '�u�Ď��Ճo�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͏�������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_INFO_OUTPUT_OK, 0)
    
    Set fso = Nothing
    
    Exit Sub
'*******************************
'VB�G���[����
Error_cmdOutPut_Click:
' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
'        'V1.21.0.1 ADD  START
'        If iFileNumber > 0 Then
'           Close #iFileNumber
'        End If
'        'V1.21.0.1 ADD  END
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '�����ُ�̏ꍇ�A�o�͌��ʃ|�b�v�A�b�v(�ُ�)�\��
        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
        '�u�Ď��Ճo�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͏����ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OUTPUT_ERROR, lngErrCode)
        Set fso = Nothing
'*******************************
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdRemove_Click
'//  �@�\����  : �u�}�̎�O�v�t��������
'//  �@�\�T�v  : �}�̂̎��O�����s���B
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
Private Sub cmdRemove_Click()
   
   On Error Resume Next
       
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
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
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '�u�Ď��Ճo�[�W�����Ǘ���ʁF�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_KANRI_GAMEN_END, 0)
 
 ' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_KANSI_APLNEW)
    pubSubCreateFolder (PATH_KANSI_APLOLD)
' EG20 V5.6.0.1�ǉ��I��

    '�o�[�W�����Ǘ��i�Ď��Ձj��ʂ����
    Unload Me
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FpsVersionDisp
'//  �@�\����    �F�o�[�W�������쐬����
'//  �@�\�T�v    �F�o�[�W�������t�@�C���쐬/��ʕ\�����s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub psVersionDisp()
    Dim bRet            As Boolean  '�߂�l
    Dim intFileNo       As Integer  '�t�@�C���ԍ�
    Dim strWork         As String   '��ƃG���A
    Dim strVerData      As String   '�S�̃o�[�W����
    Dim lngErrCode      As Long     '�G���[�R�[�h
    Dim strList         As String
    Dim strVer          As String
    Dim strWork1        As String
    Dim strWork2        As String
    Dim strWork3        As String
    Dim sFileName       As String


'*******************************
'VB�G���[����
On Error GoTo Error_psVersionDisp
'*******************************

    '�}�̏o�͖t�����s��
    cmdOutPut.Enabled = False

    '���X�g������
    lstKan.Clear
    
    '��ƃG���A������
    strWork = ""

    '�S�̃o�[�W����������
    strVerData = ""

    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ �ێ��p�֐�:�����o�[�W�����t�@�C���i��ʕ\���p�j�쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllKansiCreateVerFile(mlngChkFolderType, KANSI_VERSION_CSVFILE, VERLISTKIND_DISP)

    ' �o�[�W�����t�@�C������
    If bRet Then
       '�u�o�[�W�������t�@�C���쐬����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    ' �o�[�W�����t�@�C�����s
    Else
       '�u�o�[�W�������t�@�C���쐬�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
       Exit Sub
    End If
    
    ' �o�[�W�����t�@�C���̗L���m�F
    If Len(Trim(Dir(KANSI_VERSION_CSVFILE))) = 0 Then
        Exit Sub
    End If

    ' �o�[�W�����t�@�C���̃t�@�C���ԍ����擾����B
    intFileNo = FreeFile

    ' �o�[�W�����t�@�C���I�[�v��
    Open KANSI_VERSION_CSVFILE For Input As #intFileNo
    
        '���[�N
        Line Input #intFileNo, strWork
        
        If (Trim(strWork) = "") Then
            strVerData = HEADERTITLE_WRK & HEADERVERSION_NON & vbCrLf
        Else
            '�S�̃o�[�W����������쐬
            strVerData = strWork & vbCrLf
        End If

        '���s
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
            strVerData = strVerData & HEADERTITLE_NOW & HEADERVERSION_NON & vbCrLf
        Else
            strVerData = strVerData & strWork & vbCrLf
        End If

        '��
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
            strVerData = strVerData & HEADERTITLE_OLD & HEADERVERSION_NON & vbCrLf
        Else
            strVerData = strVerData & strWork & vbCrLf
        End If

        '�S�̃o�[�W�����o��
        lblKansibanVersion.Caption = strVerData

        strWork = ""

        '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1�폜
        Do While Not EOF(intFileNo)                         ' EG20 V3.3.0.1�ǉ�

            Line Input #intFileNo, strWork

            '���s�R�[�h�݂͓̂ǂ݂Ƃ΂�
            If Trim(strWork) <> "" Then

                strWork1 = Right(strWork, 42)
                strWork2 = Mid(strWork1, 1, 12)   '�T�C�Y�̂ݒ��o
                strWork3 = Mid(strWork1, 13, 30)
                strVer = Format(strWork2, "#,##0")
                strVer = Format(strVer, "@@@@@@@@@@@@")
                sFileName = StrConv(MidB(StrConv(Mid(strWork, 1, 27) & Space(20), vbFromUnicode), 1, 27), vbUnicode)
                strList = sFileName & strVer & strWork3
                '���X�g�ɏo��
                lstKan.AddItem (strList)

            End If

        Loop

    '�t�@�C���N���[�Y
    Close #intFileNo

    '�}�̏o�͖t������
    cmdOutPut.Enabled = True

    Exit Sub

'*******************************
'VB�G���[����
Error_psVersionDisp:
    '�o�[�W�������t�@�C���쐬�ُ탍�O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
    '�t�@�C���N���[�Y
    Close #intFileNo
'*******************************

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sKansibanVersionSet
'//  �@�\����  : �Ď��Ղ̃o�[�W�����擾�\������B
'//  �@�\�T�v  : KansiVersion.ini���A�o�[�W�������擾�E�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sKansibanVersionSet()
    Dim lSts As Long                            '�֐��߂�l
    Dim strKansiVersion As String * 128         '�Ď��ՑS�̃o�[�W����
    Dim strKansiVersionNow As String            ' �Ď��ՑS�̃o�[�W�����i���s�j  EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim strKansiVersionOld As String            ' �Ď��ՑS�̃o�[�W�����i���j    EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim strKansiVersionWrk As String            ' �Ď��ՑS�̃o�[�W�����i���[�N�jEG20 V2.1.0.1[Mainte_03_01]�ǉ�
    
    On Error Resume Next
        
' EG20 V2.1.0.1[Mainte_03_01] �R�����g�ǉ��J�n
    ' /////////////////////////////////////////////////////
    ' // ���s�o�[�W����
' EG20 V2.1.0.1[Mainte_03_01] �R�����g�ǉ��I��
    
    strKansiVersion = ""
    
    ' KansiVersion.ini����Ď��Ղ̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSION_FILE)
     If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
'        lblKansibanVersion.Caption = "�S�̃o�[�W�����F" & Left$(strKansiVersion, lSts)     ' EG20 V2.1.0.1[Mainte_03_01] �폜
        strKansiVersionNow = HEADERTITLE_NOW & Left$(strKansiVersion, lSts)                 ' EG20 V2.1.0.1[Mainte_03_01] �ǉ�
    Else
        '�o�[�W�����ԍ��̎擾�ُ�̏ꍇ�A�u--,--,--,--�v��\��
'        lblKansibanVersion.Caption = "�S�̃o�[�W�����F--.--.--.-- "                        ' EG20 V2.1.0.1[Mainte_03_01] �폜
        strKansiVersionNow = HEADERTITLE_NOW & HEADERVERSION_NON                            ' EG20 V2.1.0.1[Mainte_03_01] �ǉ�
    End If

' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
    ' /////////////////////////////////////////////////////
    ' // ���o�[�W����
    strKansiVersion = ""
    
    ' KansiVersion.ini����Ď��Ղ̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSIONOLD_FILE)
     If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
        strKansiVersionOld = HEADERTITLE_OLD & Left$(strKansiVersion, lSts)
    Else
        '�o�[�W�����ԍ��̎擾�ُ�̏ꍇ�A�u--,--,--,--�v��\��
        strKansiVersionOld = HEADERTITLE_OLD & HEADERVERSION_NON
    End If
    
    ' /////////////////////////////////////////////////////
    ' // ���[�N�o�[�W����
    strKansiVersion = ""
    
    ' KansiVersion.ini����Ď��Ղ̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSIONWRK_FILE)
     If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
        strKansiVersionWrk = HEADERTITLE_WRK & Left$(strKansiVersion, lSts)
    Else
        '�o�[�W�����ԍ��̎擾�ُ�̏ꍇ�A�u--,--,--,--�v��\��
        strKansiVersionWrk = HEADERTITLE_WRK & HEADERVERSION_NON
    End If

    
    ' /////////////////////////////////////////////////////
    ' // �\�����e�̍���
    lblKansibanVersion.Caption = strKansiVersionWrk & vbCrLf & _
                                strKansiVersionNow & vbCrLf & _
                                strKansiVersionOld
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��


End Sub

' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : fMakeListbox
''//  �@�\����  : �o�[�W�����\���Ώۂ��o�[�W�������擾�\������B
''//  �@�\�T�v  : ���A���s�A���[�N�AINI���ɂ���A
''//              *.exe�A*.dll�A*.OCX�A*.INI�̃o�[�W�������擾����B
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
''//                 �t�F�[�Y�R�@���������@�s��C��
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Function fMakeListbox() As Boolean
'    Dim strFilePath     As String   '�o�[�W�����t�@�C���p�X
'    Dim bRet            As Boolean  '�߂�l
'    Dim intFileNo       As Integer  '�t�@�C���ԍ�
'    Dim strWork         As String   '��ƃG���A
'    Dim strVerData      As String   '�S�̃o�[�W����
'    Dim intCnt          As Integer  '�J�E���^�[
'    Dim lngErrCode      As Long     '�G���[�R�[�h
'    Dim strVerformat As String
'    Dim strList As String
'    Dim strVer As String
''V1.8.0.1 ADD START
'    Dim strWork1 As String
'    Dim strWork2 As String
'    Dim strWork3 As String
'    Dim strWork4 As String
'    Dim sFileName As String
''V1.8.0.1 ADD END
''*******************************
''VB�G���[����
'On Error GoTo Error_psVersionDisp
''*******************************
'
'    fMakeListbox = False
'
''    �}�̏o�͖t�����s��
'    cmdOutPut.Enabled = False
'
''    ���X�g������
'    lstKan.Clear
'
''    ��ƃG���A������
'    strWork = ""
'
''    �Ď��Չ�ʕ\���p�o�[�W�����t�@�C���p�X�쐬
'    strFilePath = KANSI_VERSION_CSVFILE
'
'    bRet = True
''    ///////////////////////////////////////////////////////////////////////////////////////////
''    / ����DA:LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���쐬
''    ///////////////////////////////////////////////////////////////////////////////////////////
'    bRet = dllKansiCreateVerFile(mlngChkFolderType, strFilePath)
'
''    �Ď��Չ�ʕ\���p�o�[�W�����t�@�C���쐬����
'    If bRet Then
''       �u�Ď��Ճo�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬����v���O�o��
'       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
''    �Ď��Չ�ʕ\���p�o�[�W�����t�@�C���쐬���s
'    Else
''       �u�Ď��Ճo�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
'       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'       Exit Function
'    End If
'
''    �Ď��Չ�ʕ\���p�o�[�W�����t�@�C���̗L���m�F
'    If Len(Trim(Dir(strFilePath))) = 0 Then
'        Exit Function
'    End If
'
''    �Ď��Չ�ʕ\���p�o�[�W�����t�@�C���̃t�@�C���ԍ����擾����
'    intFileNo = FreeFile
'
''    �Ď��Չ�ʕ\���p�o�[�W�����t�@�C���I�[�v��
'    Open strFilePath For Input As #intFileNo
'
'    strWork = ""
'
''    ���X�g�\�����ǂݍ��� (�t�@�C���I�[�܂Ń��[�v���J��Ԃ�)
'    Do While Not EOF(1)
'
'        Line Input #intFileNo, strWork
'
''        ���s�R�[�h�݂͓̂ǂ݂Ƃ΂�
'        If Trim(strWork) <> "" Then
'            '�o�[�W�����t�@�C�����̃o�[�W�����l���uzzz,zzz,zzz�v�t�H�[�}�b�g�ɕϊ����鏈��
'            'V1.8.0.1 DEL START
''            strVer = Mid(strWork, VERSION_STA, VERSION_SIZE)
''            strVerformat = Format(strVer, "#,##0")
''            strVerformat = Format(strVerformat, "@@@@@@@@@@@@")
''            strList = Mid(strWork, VERMOJI_STA, FOLDER_STS)
''            strList = strList & strVerformat
''            strList = strList & Mid(strWork, HIDUKE_STA, VERSION_END)
'            'V1.8.0.1 DEL END
'            'V1.8.0.1 ADD START
'            strWork1 = Right(strWork, 42)
'            strWork2 = Mid(strWork1, 1, 12)   '�T�C�Y�̂ݒ��o
'            strWork3 = Mid(strWork1, 13, 30)
'            strVer = Format(strWork2, "#,##0")
'            strVer = Format(strVer, "@@@@@@@@@@@@")
'            sFileName = StrConv(MidB(StrConv(Mid(strWork, 1, 27) & Space(20), vbFromUnicode), 1, 27), vbUnicode)
'            strList = sFileName & strVer & strWork3
'            'V1.8.0.1 ADD END
''           ���X�g�ɏo��
'            lstKan.AddItem (strList)
'        End If
'    Loop
'
''    �t�@�C���N���[�Y
'    Close #intFileNo
'
'    fMakeListbox = True
'
''    �}�̏o�͖t������
'    cmdOutPut.Enabled = True
'
'    Exit Function
'
''*******************************
''VB�G���[����
'Error_psVersionDisp:
''   �u�Ď��Ճo�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
''    �t�@�C���N���[�Y
'   Close #intFileNo
''*******************************
'End Function
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��i�Ď��Ճo�[�W�����A�b�v�Ή��j
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    '�ėp���[����M�������s��
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then          ' EG20 V3.0.0.2�폜
'    If pfVersionDispMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then   ' EG20 V3.0.0.2�ǉ� ' EG20 V7.3.0.1�yEG20_KANSI03_01�zDEL
    If pfMailRecieve_KansiVerDisp = ML_ID_HOSHU_ACTIVE_REQ Then  ' EG20 V7.3.0.1�yEG20_KANSI03_01�zADD
        AppActivate frmKVer.Caption, False
        pfFormActive (frmKVer.hwnd)                              ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sWrkFolderRemove
'//  �@�\����    �F ���[�N�t�H���_���t�@�C���폜����
'//  �@�\�T�v    �F ���[�N�t�H���_���̃t�@�C�����폜����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F �Ȃ�
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F���D�@�o�[�W�����Ǘ���ʂ�sWrkFolderRemove���p
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim stringWorkFolder As String      ' �t�H���_��
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    '�����l�ݒ�
    sWrkFolderRemove = True
   
    '//////////////////////////////////////////////////////////////////////////
    '// ���[�N�t�H���_���̑����t�H���_������
    ' ���[�N�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    stringWorkFolder = PATH_KANSI_APLNEW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    Set objFso = Nothing

'    '�u����I���v�|�b�v�A�b�v��ʕ\��
'    MsgBox "����I�����܂����B", _
'           vbOKOnly + vbInformation, _
'           "���s����"

    Exit Function '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�u���[�N�N���A�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbCritical, _
           "���s����"
           
   '�u�����ް�ޮ݁Fܰ�̫���̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    Set objFso = Nothing
End Function

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
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.273�C���Ή��z
'//  REVISIONS   �F(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y2,3�����Ή�
'//  REVISIONS   �F (EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//  REVISIONS   �F (EG20 V30.3.0.1)2014-10-23  CODED BY  [TCC] T.Nakajima
'//                  �k���V�����t�F�[�Y�Q�Ή��i�}�̎�O���G���[�Ή��j
'//  ���l        �F���D�@�o�[�W�����Ǘ���ʂ�sFDInstall���p
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall()
    Dim MyName As String            '�t�@�C���t���p�X��
    Dim iResponse As Integer        'MsgBox�{�^���R�[�h
    Dim sInputPass As String        '�C���X�g�[�����f�B���N�g����(STD)or�t�@�C����(LZH)
    Dim lngErrCode As Long          '�G���[�R�[�h
    
    Dim lngProcId As Long                ' �v���Z�XID
    Dim hProc As Variant                 ' �v���Z�X�n���h��
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                    '�t�@�C���I�u�W�F�N�g
    Dim FileName As String               ' ���o�t�@�C����                ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim FileKaku As String               ' �g���q                        ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim ExecCommand As String            ' ���s������                    ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim CurrentDirectory As String       ' �J�����g�f�B���N�g��          ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim ExecDirectory As String          ' ���s�t�@�C���f�B���N�g��      ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    
    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^

    '���k�t�@�C���w��̎�:
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
    CommonDialog1.Filter = "���s�t�@�C���i*.EXE�j|*.exe|" & _
                            "�o�b�`�t�@�C���i*.BAT�j|*.bat|" & _
                            "�X�N���v�g�t�@�C���i*.VBS�j|*.vbs|"
    '�t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    '�I�������t�@�C�������擾
    sInputPass = CommonDialog1.FileName
    If sInputPass = "" Then '�t�@�C�����I��
        Set objFso = Nothing
        Set objFi = Nothing
        Exit Sub    '�t�@�C�����I������Ȃ���Ώ������f
    End If
        
    '�u���[�N�R�s�[�m�F�v�|�b�v�A�b�v��ʕ\��
    iResponse = MsgBox("�I�����ꂽ�C���X�g�[�����ނ̓��e�𓝍��Ď��ՃA�v���P�[�V������" _
                       & Chr(vbKeyReturn) _
                       & "�ؑ֗̈�ɓW�J���܂��B��낵���ł����H", _
                       (vbOKCancel + vbExclamation), _
                       "�}�́����[�N�@�R�s�[")
        
    If iResponse = vbCancel Then
    '[������] �{�^����I��:�������Ȃ��B
        'V1.20.0.1 ADD START
        Set objFso = Nothing
        Set objFi = Nothing
        'V1.20.0.1 ADD END
        Exit Sub
    End If
    
    Set objFso = Nothing
    Set objFi = Nothing
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    
'    lngProcId = Shell(sInputPass, vbNormalFocus)       ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�폜
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��J�n
    ' �J�����g�f�B���N�g���擾
    CurrentDirectory = CurDir$()
    Call psFolderPathGet(sInputPass, ExecDirectory)
    Call ChDir(ExecDirectory)
    ' �t�@�C�����O�擾
    psFileNameGet sInputPass, FileName, FileKaku
    If UCase(FileKaku) = "VBS" Then
        ExecCommand = "wscript.exe " & sInputPass
    Else
        ExecCommand = sInputPass
    End If
    lngProcId = Shell(ExecCommand, vbNormalFocus)
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��I��
    
    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, lngProcId)   ' �v���Z�X�n���h�����擾���܂��B
    If hProc > 0 Then                                           ' �v���Z�X�n���h�����擾�ł����ꍇ
        Call dllWaitForSingleObject(hProc)                      ' �v���Z�X���V�O�i����ԂɂȂ�܂ő҂��܂��B
        CloseHandle hProc                                       ' �v���Z�X�n���h����������܂��B
    End If

    'EG20 V30.0.3.1 ADD START
    'ChDir�ł�CommonDialog�̏ꍇ�AH�h���C�u���I�����ꂽ�܂ܕύX���ꂸ�A�}�̎�O�����ł��Ȃ��Ȃ邽�߁AChDrive�ɕύX
    ChDrive "C"
    'EG20 V30.0.3.1 ADD END
    Call ChDir(CurrentDirectory)                        ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    
' EG20 V5.8.0.1�폜�J�n
'    ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
    ' �^����ԍX�V
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_KANSI, BOOTINFO_UNKAI_ARI)
    Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMEKANSI, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1�ǉ��I��
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD START
    ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
    Call funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_CLEAR)
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD END

' EG20 V5.8.0.1 ADD START
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (PATH_KANSI_APLNEW)
' EG20 V5.8.0.1 ADD END
    
    Exit Sub    '�������I������

ErrorHandler:   ' �G���[�����B
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing

' EG20 V5.8.0.1 ADD START
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (PATH_KANSI_APLNEW)
' EG20 V5.8.0.1 ADD END

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    MsgBox "�C���X�g�[���}�̂���̃R�s�[�G���[���������܂����B" _
            & Chr(vbKeyReturn) & "�G���[�R�[�h��" _
            & str$(Err.Number), _
            vbOKOnly + vbExclamation, _
            "�}�́����[�N�@�R�s�["
    
    '�u�����ް�ޮ݁F�}�́�ܰ���߰�����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_USB_COPY_WRK_ERROR, lngErrCode)
End Sub


'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'/
'/  �֐�����     : sCmdBtnEnabled
'/  �@�\����     : �R�}���h�{�^�������E�s����
'/  �@�\�T�v     : �R�}���h�{�^���������Ɋ�ĉ����E�s�������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                EG20�t�F�[�Y�Q�Ή�
'//                EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)
    Dim iLoopCnt    As Integer

    '�t�H���_�I�𕔁F�I��L��
    chkFolder(0).Enabled = blnFlg
    chkFolder(1).Enabled = blnFlg
    chkFolder(2).Enabled = blnFlg

    cmdRefresh.Enabled = blnFlg                     ' �\���X�V
    cmdClear.Enabled = blnFlg                       ' ���[�N�N���A
    cmdCopyBaitai_Work.Enabled = blnFlg             ' �}�́����[�N�R�s�[
    cmdCopyWork_Jikko.Enabled = blnFlg              ' ���[�N�����s�R�s�[
    cmdCopyOld_Jikko.Enabled = blnFlg               ' �������s�R�s�[
    cmdOutPut.Enabled = blnFlg                      ' �o�[�W�������}�̏o��
    cmdRemove.Enabled = blnFlg                      ' �}�̎�O
    cmdReturn.Enabled = blnFlg                      ' �o�[�W�����Ǘ���ʂ֖߂�

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : tmrAplTimer_Timer
'//  �@�\����  : �A�v���N���`�F�b�N�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v  : �^�C���A�b�v���ɃA�v���N����Ԃ��`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  : (EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή��y����TR-No.184�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplTimer_Timer()

    Dim bIDURet As Boolean
    Dim bLDURet As Boolean

    On Error Resume Next

    If CheckAppStart(PROC_KANRI) <> 0 Then
        If lngtime >= lngAplMAX_Time Then
            tmrAplTimer.Enabled = False
            '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
            '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)

            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
                    vbOKOnly + vbExclamation, _
                    "�Ď��Ճo�[�W�����Ǘ�"
            sCmdBtnEnabled True                         ' ��ʑ����
        Else
            '�^�C�}���蒼��
            tmrAplTimer.Interval = MN_MAIL_INTERVAL
            lngtime = lngtime + MN_MAIL_INTERVAL
        End If
    Else
        tmrAplTimer.Enabled = False
        If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
            bIDURet = EndIDULog 'IDU���O�N������IDU���O�ɑ΂��ă��O�I���v��CMD���M
        Else
            bIDURet = True
        End If

        If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
            bLDURet = EndLDULog  'LDU���O�N������LDU���O�ɑ΂��ă��O�I���v��CMD���M
        Else
            bLDURet = True
        End If

        If bIDURet = True And bLDURet = True Then
            lngtime = MN_MAIL_INTERVAL
            tmrLogTimer.Enabled = True
        Else
            '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
            '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
                    vbOKOnly + vbExclamation, _
                    "�Ď��Ճo�[�W�����Ǘ�"
            sCmdBtnEnabled True                         ' ��ʑ����
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : tmrLogTimer_Timer
'//  �@�\����  : ���O�N���`�F�b�N�^�C�}����
'//  �@�\�T�v  : ���O�N���`�F�b�N�^�C�}�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  : (EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή��y����TR-No.184�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()

    On Error Resume Next

    If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
        Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then

        If lngtime >= lngLogMAX_Time Then
            '���O�N���`�F�b�N�^�C�}���~����B
            tmrLogTimer.Enabled = False
            '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
                    vbOKOnly + vbExclamation, _
                    "�Ď��Ճo�[�W�����Ǘ�"
            sCmdBtnEnabled True                         ' ��ʑ����
        Else
            '�^�C�}���蒼��
            tmrLogTimer.Interval = MN_MAIL_INTERVAL
            lngtime = lngtime + MN_MAIL_INTERVAL
        End If
    Else
        tmrLogTimer.Enabled = False
        '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)

        '�ؑփc�[���N��
        Call AplVersionChangeProc(lngChangeKind)

        '�I������
        psEndHoshuProc
        '�ێ�v���Z�X�I��
        End
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : fWorktoNow_Before1
'//  �@�\����  : ���[�N�����s�R�s�[�O����1
'//  �@�\�T�v  : ���[�N�����s�R�s�[�O�ɉ��L�������s��
'//�@�@�@�@�@�@�@�E�Ď��ՃA�v�����N���`�F�b�N
'//�@�@�@�@�@�@�@�E���ؖ����f�[�^���݃`�F�b�N
'//�@�@�@�@�@�@�@�E�f�W�ʐM�ؒf
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@ TRUE      ����I��
'//�@�@�@�@�@�@�@�@�@�@�@�@FALSE     �ُ�I��
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fWorktoNow_Before1() As Boolean

    Dim iCnt As Integer

    fWorktoNow_Before1 = False

    '-------------------------------------------------------------------------------------------
    '�Ď��ՃA�v�����N�����́A�{�V�[�P���X�����{���Ȃ�
    '-------------------------------------------------------------------------------------------
    If CheckAppStart(PROC_KANRI) = 0 Then
        MsgBox "�ێ�P�ƋN���̂��߁A���[�N�����s�R�s�[���s���܂���B", _
                vbOKOnly + vbCritical, _
                "���[�N�����s �R�s�["
        Exit Function
    End If

    '-------------------------------------------------------------------------------------------
    '���ؖ����f�[�^�����݂���ꍇ�A�{�V�[�P���X�����{���Ȃ�
    '-------------------------------------------------------------------------------------------
    If fChkSimekiriMisouUmu = False Then
        MsgBox "���ؖ����f�[�^�����邽�߁A���[�N�����s�R�s�[���s���܂���B", _
                vbOKOnly + vbCritical, _
                "���[�N�����s �R�s�["
        Exit Function
    End If

    sCmdBtnEnabled False                            ' ��ʑ���s��

    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
    
    Erase byDeshuCnctSet  '�f�W�ؗ��ݒ菉����
    Erase byGateCnctSet   '�����ؗ��ݒ菉����
    miErrorSts = 0        '�ُ펞�ʐM��ʏ�����
    miErrorDisp = 0       '�ُ펞�ُ펞�\������������
    
    '�f�W�̐ڑ��^�ؒf�ݒ���擾
    For iCnt = CNT_MIN To CONECT_CORNER_MAXINDEX
        If gblnCornerSet(iCnt) = True Then
            If (0 = pfGetJyouiKikiConectSet(DESHU_ID + iCnt)) Then
                byDeshuCnctSet(iCnt) = 1
            End If
        End If
    Next
    
    '�����̐ڑ��^�ؒf�ݒ���擾
    For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
        If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
            If (0 = pfGetGateConectSet(iCnt + 1)) Then
                byGateCnctSet(iCnt) = 1
            End If
        End If
    Next

    '-------------------------------------------------------------------------------------------
    '�f�W�ʐM�ؒf
    '-------------------------------------------------------------------------------------------
    If False = pfKill_TusinConect(ML_DT_DESHU) Then
        '�ʐM�ؒf�ُ폈��
        Call ConnctErrorProc(DESHU_CONNECT, ERROR_TUSHIN_DISP)
        Exit Function
        
    End If

    fWorktoNow_Before1 = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : fWorktoNow_Before2
'//  �@�\����  : ���[�N�����s�R�s�[�O����2
'//  �@�\�T�v  : �ʐM�ݒ�v��RES�i�f�W�A�ؒf�j��M��A���L�������s��
'//�@�@�@�@�@�@�@�E�����ʐM�ؒf
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@ TRUE      ����I��
'//�@�@�@�@�@�@�@�@�@�@�@�@FALSE     �ُ�I��
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fWorktoNow_Before2() As Boolean

    fWorktoNow_Before2 = False

    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TRANS_KANRI)
    
    '-------------------------------------------------------------------------------------------
    '�����ʐM�ؒf
    '-------------------------------------------------------------------------------------------
    If False = pfKill_TusinConect(ML_DT_JIKAI) Then
        '�ʐM�ؒf�ُ폈��
        Call ConnctErrorProc(GATE_CONNECT, ERROR_TUSHIN_DISP)
        Exit Function
    End If
    
    fWorktoNow_Before2 = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : fWorktoNow_Start
'//  �@�\����  : ���[�N�����s�R�s�[����
'//  �@�\�T�v  : �������؃f�[�^�쐬��A���[�N�����s�R�s�[�������s��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@ TRUE      ����I��
'//�@�@�@�@�@�@�@�@�@�@�@�@FALSE     �ُ�I��
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY ]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fWorktoNow_Start() As Boolean

    Dim bRet        As Boolean            ' ���[�����M�����߂�l
    Dim iResponse   As Integer            ' MsgBox�{�^���R�[�h
    
    On Error Resume Next
    
    fWorktoNow_Start = True
    
    ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONUP)
    If bRet = False Then
        '�f�W�������ʐM�ڑ�
        Call ConnctErrorProc(GATE_CONNECT, ERROR_END_DISP)
        Exit Function
    End If

    '���[�N�����s�R�s�[�ؒf�ݒ�p�����[�^�X�V����
    bRet = funcUpdateConnectSetParam(byDeshuCnctSet, byGateCnctSet)
    If bRet = False Then
        '�f�W�������ʐM�ڑ�
        Call ConnctErrorProc(GATE_CONNECT, ERROR_END_DISP)
        Exit Function
    End If

    ' �I���m�F
    iResponse = MsgBox("���s�R�s�[��K�p���邽�߂ɓ����Ď��Ղ�" & Chr(vbKeyReturn) _
                        & "�ċN�����܂����H", _
                        vbOKCancel + vbExclamation, _
                        "���[�N�����s �R�s�[")
    If iResponse = vbCancel Then
        fWorktoNow_Start = False
        Exit Function
    End If
    
    sCmdBtnEnabled False                            ' ��ʑ���s��
    
    ' �����Ď��ՂփA�v���I���v���̑��M
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
               vbOKOnly + vbExclamation, _
               "�Ď��Ճo�[�W�����Ǘ�"
        sCmdBtnEnabled True                         ' ��ʑ����
    Else

        lngtime = MN_MAIL_INTERVAL                  ' ���݃^�C�}�l������
        tmrAplTimer.Enabled = True                  ' ���݃^�C�}�N��
    
        lngChangeKind = ML_DT_VERSIONUP             ' �ؑ֎�ʂ�ݒ�
        
    End If
        
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : fChkSimekiriMisouUmu
'//  �@�\����  : ���ؖ����f�[�^���݃`�F�b�N
'//  �@�\�T�v  : ���ؖ����f�[�^�����݂���ꍇ�A�{�V�[�P���X�����{���Ȃ�
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@ TRUE      ���ؖ����f�[�^�Ȃ�
'//�@�@�@�@�@�@�@�@�@�@�@�@FALSE     ���ؖ����f�[�^����
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fChkSimekiriMisouUmu() As Boolean

    Dim objFso As New FileSystemObject                  ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim nLoop As Integer                                ' ���[�v
    Dim bEnable As Boolean                              ' �{�^�����
    Dim szFileName As String

    On Error GoTo ErrorHandler                          ' �G���[�n���h���̓o�^
    
    fChkSimekiriMisouUmu = True

    For nLoop = 0 To UBound(gblnCornerSet)

        bEnable = False
        If gblnCornerSet(nLoop) = True Then
            ' /////////////////////////////////////////////////////////////////////////
            ' // ���؏o�̓f�[�^�͑��݂��邩�H�iD:\KANSI\SHUKEI\OUT_DATA\CORNER##\SIME##.DAT�j
            szFileName = Replace(PATH_SHUKEI_SHIMEDAT, "##", Format(nLoop + 1, "0#"))
            If objFso.FileExists(szFileName) = True Then
                fChkSimekiriMisouUmu = False
                Exit Function
            End If
        End If
        
    Next nLoop
    
    Set objFso = Nothing
    
    Exit Function

' /////////////////////////////////////////////////////////
' // �G���[����
ErrorHandler:
    Set objFso = Nothing
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfKill_TusinConect
'//  �@�\����  : �ʐM����ؒf����
'//  �@�\�T�v  : �w�肵���O���@��̒ʐM�����ؒf����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Long      dwKiki    �O���@��v�����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@ TRUE      ���b�Z�[�W���M����
'//�@�@�@�@�@�@�@�@�@�@�@�@FALSE     ���b�Z�[�W���M�ُ�
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfKill_TusinConect(dwKiki As Long) As Boolean

    Dim bRet As Boolean                 '���[�����M�߂�l
    Dim iCnt As Integer                 '�J�E���^�[
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    pfKill_TusinConect = False

    '-------------------------------------------------------------------------------------------
    '�ʐM�ݒ�v��CMD���b�Z�[�W�쐬
    '-------------------------------------------------------------------------------------------
    '�w�b�_�����ʍ쐬����
    Call SendMailHeader(ML_ID_CONECT_CMD, MlSize.CONECT_CMD)
    
    '�f�[�^���ݒ�
    udtMail.dwRequestKIKI = dwKiki
    udtMail.dwRequestConectType = ML_REQUEST_SETUDAN
    For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
        udtMail.dwGouki(iCnt) = ML_TARGET_OFF
    Next
    
    '�O���@��v����ʂ������H
    If dwKiki = ML_DT_JIKAI Then
        '�O���@��v����ʂ������̏ꍇ
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
            '���D�@����������Ă��邩�H
            If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
                udtMail.dwGouki(iCnt) = ML_TARGET_ON
            End If
        Next
    Else
        '�O���@��v����ʂ������ȊO�̏ꍇ
        For iCnt = 0 To UBound(gblnCornerSet)
            '�R�[�i�ڑ�����Ă��邩�H
            If gblnCornerSet(iCnt) = True Then
                udtMail.dwGouki(iCnt) = ML_TARGET_ON
            End If
        Next
    End If
    
    '-------------------------------------------------------------------------------------------
    '�ʐM�ݒ�v��CMD(�Ώ�ID)���ă}�v���Z�X�ɑ��M����
    '-------------------------------------------------------------------------------------------
    bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
    If False = bRet Then
        '�u�ʐM�ڑ��E�ؒf��ʁF�ʐM�ݒ�v��CMD���M�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
        Exit Function
    Else
        '�u�ʐM�ڑ��E�ؒf��ʁF�ʐM�ݒ�v��CMD���M����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
    End If

    pfKill_TusinConect = True

End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfConnect_TusinConect
'//  �@�\����  : �ʐM����ڑ�����
'//  �@�\�T�v  : �w�肵���O���@��̒ʐM�����ڑ�����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Long      dwKiki    �O���@��v�����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@ TRUE      ���b�Z�[�W���M����
'//�@�@�@�@�@�@�@�@�@�@�@�@FALSE     ���b�Z�[�W���M�ُ�
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfConnect_TusinConect(dwKiki As Long) As Boolean

    Dim bRet As Boolean                 '���[�����M�߂�l
    Dim iCnt As Integer                 '�J�E���^�[
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    pfConnect_TusinConect = False
    
    '-------------------------------------------------------------------------------------------
    '�ʐM�ݒ�v��CMD���b�Z�[�W�쐬
    '-------------------------------------------------------------------------------------------
    '�w�b�_�����ʍ쐬����
    Call SendMailHeader(ML_ID_CONECT_CMD, MlSize.CONECT_CMD)
    
    '�f�[�^���ݒ�
    udtMail.dwRequestKIKI = dwKiki
    udtMail.dwRequestConectType = ML_REQUEST_CONECT
    For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
        udtMail.dwGouki(iCnt) = ML_TARGET_OFF
    Next
    
    '�O���@��v����ʂ������H
    If dwKiki = ML_DT_JIKAI Then
        '�O���@��v����ʂ������̏ꍇ
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
            '���D�@����������Ă��邩�H
            If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
                udtMail.dwGouki(iCnt) = byGateCnctSet(iCnt)
            End If
        Next
    Else
        '�O���@��v����ʂ������ȊO�̏ꍇ
        For iCnt = 0 To UBound(gblnCornerSet)
            '�R�[�i�ڑ�����Ă��邩�H
            If gblnCornerSet(iCnt) = True Then
                udtMail.dwGouki(iCnt) = byDeshuCnctSet(iCnt)
            End If
        Next
    End If
    
    '-------------------------------------------------------------------------------------------
    '�ʐM�ݒ�v��CMD(�Ώ�ID)���ă}�v���Z�X�ɑ��M����
    '-------------------------------------------------------------------------------------------
    bRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.CONECT_CMD, udtMail.mlHeader)
    If False = bRet Then
        '�u�ʐM�ڑ��E�ؒf��ʁF�ʐM�ݒ�v��CMD���M�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, lngErrCode)
        Exit Function
    Else
        '�u�ʐM�ڑ��E�ؒf��ʁF�ʐM�ݒ�v��CMD���M����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, CONECT_CONECTSETTEI_CMD_SEND, 0)
    End If

    pfConnect_TusinConect = True

End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : SendMailHeader
'//  �@�\����  : ���M���[���쐬����
'//  �@�\�T�v  : ���M���[��(�w�b�_��)�쐬���s��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SendMailHeader(dwId As Long, dwSize As Long)

    Dim bytWork()   As Byte
    Dim i           As Integer
    
    Erase bytWork
    
      udtMail.mlHeader.dwId = dwId
      udtMail.mlHeader.dwSize = dwSize
      udtMail.mlHeader.dwProid = RHOSHU_ID
      udtMail.mlHeader.dwSubArea = 0
      
      bytWork = StrConv(MAIL_SLOT_HOSHU, vbFromUnicode)
      '���I�z��̓��e�����O�p�����[�^�\���̂̐ÓI�z��Ɋi�[����B
      For i = 0 To UBound(bytWork)
        'Null�l�ɂȂ����珈���𔲂���B
         If bytWork(i) = vbVEmpty Then Exit For
               
            udtMail.byMailName(i) = bytWork(i)
                
            '���I�z��̍ő�v�f�ɂȂ����珈���𔲂���
             If i = UBound(bytWork) Then Exit For
      Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfMailRecieve_KansiVerDisp
'//  �@�\����  : �ėp���[����M�����i�Ď��Ճo�[�W�����Ǘ��j
'//  �@�\�T�v  : �ێ烁�[���X���b�g����A���[������M�B
'//              ���v���Z�X�I���w�����͋����I��
'//              ���v���Z�X�I���w������M�����ꍇ�ɉ�����ʒm����B
'//
'//              �^        ����      �Ӗ�
'//   ����     :�Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :Long�@�@�@�@�@�@�@�@[OUT]�߂�l
'//
'//  ORIGINAL  :(EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//              2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function pfMailRecieve_KansiVerDisp() As Long

    Dim lLen As Long                    '���[���T�C�Y
    Dim uMail As ML_KYOTU_INF           '�ėp���[���t�H�[�}�b�g
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim bRet As Boolean                 '�߂�l
    Dim iCnt As Integer                 '�J�E���^
   
    On Error Resume Next
    
    pfMailRecieve_KansiVerDisp = 0      '�߂�l�𐳏�Ƃ���

    '�ێ烁�[����X���b�g���烁�[�������o��
    lLen = DssMailRead(plMSlot_MN, uMail)
    If lLen > 0 Then                            '��M?
    
        '------------------------------------------------------------------------------
        '�v���Z�X�I���w���̏ꍇ
        '------------------------------------------------------------------------------
        If uMail.udtlHeader.dwId = ML_ID_PROEND_ORD Then
           
           '�u�v���Z�X�I���w����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            
            '�v���Z�X�I���ʒm�𑗐M����
            uMail.udtlHeader.dwId = ML_ID_PROEND_INF
            uMail.udtlHeader.dwSize = MlSize.PROEND_INF
            uMail.udtlHeader.dwProid = RHOSHU_ID
            uMail.udtlHeader.dwSubArea = 0
            bRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.PROEND_INF, uMail.udtlHeader)
            If bRet = True Then
               '�u�v���Z�X�I���ʒm���M�F����v���O�o��
               Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, PROCESS_END_REQ_SEND, 0)
            Else
               '�u�v���Z�X�I���ʒm���M�F�ُ�v���O�o��
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, PROCESS_END_REQ_SEND, lngErrCode)
            End If
            
            '�����I���������s��
            pfAbortProc
            Exit Function       '�������I������
            
        '------------------------------------------------------------------------------
        '�ێ��ʃA�N�e�B�u�\���̏ꍇ
        '------------------------------------------------------------------------------
        ElseIf uMail.udtlHeader.dwId = ML_ID_HOSHU_ACTIVE_REQ Then
           
           '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
            
            pfMailRecieve_KansiVerDisp = ML_ID_HOSHU_ACTIVE_REQ
        
        '------------------------------------------------------------------------------
        '�ʐM�ݒ�v��RES�̏ꍇ
        '------------------------------------------------------------------------------
        ElseIf uMail.udtlHeader.dwId = ML_ID_CONECT_RES Then
           '�u�ʐM�ݒ�v��RES��M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, CONECT_CONECTSETTEI_CMD_RECV, 0)
            
            '�G���[���ʐM��ʂ�0�i����j�łȂ��ꍇ
            If (miErrorSts <> 0) Then
                '�O���@��v����ʂ������H
                If uMail.lngData(0) = ML_DT_JIKAI Then
                    Call ErrorProc
                 
                '�O���@��v����ʂ������łȂ��ꍇ
                Else
                    '�G���[��ʃf�W�̏ꍇ
                    If (miErrorSts = DESHU_CONNECT) Then
                        Call ErrorProc

                    '�G���[��ʃf�W�łȂ��ꍇ
                    Else
                        '�����ƒʐM�ڑ�
                        pfConnect_TusinConect (ML_DT_JIKAI)
                    End If
                End If
                Exit Function       '�������I������
            End If
                    
            '�O���@��v����ʂ������H
            If uMail.lngData(0) = ML_DT_JIKAI Then
                If uMail.lngData(1) = ML_CONECT_ERROR Then
                    Call ConnctErrorProc(GATE_CONNECT, ERROR_TUSHIN_DISP) '�ʐM�ؒf�ُ폈��
                    Exit Function                                   '�������I������
                End If

                sCmdBtnEnabled False                                ' ��ʑ���s��

                '�����ʐM�ؒf�ҋ@����
                bRet = pfCheakGateConectSts
                If bRet = False Then
                    Call ConnctErrorProc(GATE_CONNECT, ERROR_TUSHIN_DISP) '�ʐM�ؒf�ُ폈��
                    Exit Function                                   '�������I������
                End If

                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                
                '���ؖ����f�[�^�쐬�i�ڑ��R�[�i�[�����J��Ԃ��j
                For iCnt = 0 To UBound(gblnCornerSet)
                    '�R�[�i�ڑ�����Ă��邩�H
                    If gblnCornerSet(iCnt) = True Then
                        miCornerNo = iCnt
                        '���؃f�[�^�o�͒���ʕ\��
                        frmShimekiriOutPut2.Show vbModal
                        If (mbMisouResult = False) Then
                            tmrMail.Enabled = True
                            Call ConnctErrorProc(GATE_CONNECT, ERROR_MISOU_DISP) '�ʐM�ؒf�ُ폈��
                            Exit Function
                        End If
                    End If
                Next
                tmrMail.Enabled = True
               
                '���[�N�����s�R�s�[����
                bRet = fWorktoNow_Start
                If bRet = False Then
                    sCmdBtnEnabled True                             ' ��ʑ����
                End If
            
            Else
                If uMail.lngData(1) = ML_CONECT_ERROR Then
                    Call ConnctErrorProc(DESHU_CONNECT, ERROR_TUSHIN_DISP) ' �ʐM�ؒf�ُ폈��
                    Exit Function                                   ' �������I������
                End If
                
                '��ʋ@��ʐM�ؒf�ҋ@����
                bRet = pfCheakJyouiKikiConectSts
                If (bRet = False) Then
                    Call ConnctErrorProc(DESHU_CONNECT, ERROR_TUSHIN_DISP) ' �ʐM�ؒf�ُ폈��
                    Exit Function                                   ' �������I������
                End If
                
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
                '�O���@��v����ʂ��f�W�̏ꍇ�A�����ʐM�ؒf���������{
                bRet = fWorktoNow_Before2
                If bRet = False Then
                    sCmdBtnEnabled True                     ' ��ʑ����
                End If
            End If
            
        '------------------------------------------------------------------------------
        '��L�ȊO
        '------------------------------------------------------------------------------
        Else
        
           '�u���[��ID�s���v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        
        End If
        
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfCheakJyouiKikiConectSts
'//  �@�\����  : ��ʋ@��ʐM�ؒf�ҋ@����
'//  �@�\�T�v  : ��ʋ@��ʐM��Ԃ��S�R�[�i�[�ʐM�ُ�ɂȂ�܂ŌJ��Ԃ�
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean   True      �S�R�[�i�[�ʐM�ؒf
'//�@�@�@�@�@�@�@�@�@�@�@�@False�@�@ �ʐM�ؒf�҂��^�C���A�E�g����
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfCheakJyouiKikiConectSts() As Boolean

    Dim iAreId(0 To 5) As Integer       '��ʋ@��ʐM��ԃG���AID
    Dim bCheakAns As Boolean            '�`�F�b�N����
    Dim iConectSts As Integer           '�ʐM���
    Dim iCnt As Integer                 '�J�E���^
    Dim LngSleepTotal As Long           '�X���[�v�J�E���g���v�l
    
    On Error Resume Next
    
    pfCheakJyouiKikiConectSts = False
    bCheakAns = True
    LngSleepTotal = 0
    
    iAreId(0) = 1                       '��ʋ@��ʐM��ԃG���AID�F�f�[�^�W�v�@(�R�[�i�P)
    iAreId(1) = 9                       '��ʋ@��ʐM��ԃG���AID�F�f�[�^�W�v�@(�R�[�i�Q)
    iAreId(2) = 10                      '��ʋ@��ʐM��ԃG���AID�F�f�[�^�W�v�@(�R�[�i�R)
    iAreId(3) = 11                      '��ʋ@��ʐM��ԃG���AID�F�f�[�^�W�v�@(�R�[�i�S)
    iAreId(4) = 12                      '��ʋ@��ʐM��ԃG���AID�F�f�[�^�W�v�@(�R�[�i�T)
    iAreId(5) = 13                      '��ʋ@��ʐM��ԃG���AID�F�f�[�^�W�v�@(�R�[�i�U)
    
    '------------------------------------------------------------------------------------
    '��ʋ@��ʐM��ԃ`�F�b�N
    '------------------------------------------------------------------------------------
    '�S�R�[�i�[�ʐM�ؒf�����܂Ő����J��Ԃ�
    '���S�R�[�i�[�̒ʐM���ؒf�����܂ŁA�R���ȏ�o�߂����ꍇ�ُ�I���Ƃ���B
    Do While LngSleepTotal < WAIT_TIME_OUT
    
        '�ڑ��R�[�i�[�����J��Ԃ�
        For iCnt = 0 To UBound(gblnCornerSet)
            
            bCheakAns = True
            
            '�R�[�i�ڑ�����Ă��邩�H
            If gblnCornerSet(iCnt) = True Then
                
                '��ʋ@��ʐM��Ԏ擾
                iConectSts = pfGetJyouiKikiConectSts(iAreId(iCnt))
                '�ʐM��Ԃ��u0:���㒆�vor�u1:�ʐM����v���H
                If iConectSts = 1 Then
                    bCheakAns = False
                    Exit For
                End If
                
            End If
        
        Next
        
        '�S�R�[�i�[�ʐM�ؒf���ꂽ���H
        If bCheakAns = True Then
            Exit Do
        End If
        
        Sleep (100)
        LngSleepTotal = LngSleepTotal + 100
    
    Loop
    
    '�R���ȏ�o�߂������H
    If LngSleepTotal >= WAIT_TIME_OUT Then
        Exit Function
    End If
    
    pfCheakJyouiKikiConectSts = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfCheakGateConectSts
'//  �@�\����  : �����ʐM�ؒf�ҋ@����
'//  �@�\�T�v  : �����ʐM��Ԃ��S�R�[�i�[�ʐM�ُ�ɂȂ�܂ŌJ��Ԃ�
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean   True      �S���@�ʐM�ؒf
'//�@�@�@�@�@�@�@�@�@�@�@�@False�@�@ �ʐM�ؒf�҂��^�C���A�E�g����
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfCheakGateConectSts() As Boolean

    Dim bCheakAns As Boolean            '�`�F�b�N����
    Dim iConectSts As Integer           '�ʐM���
    Dim iCnt As Integer                 '�J�E���^
    Dim LngSleepTotal As Long           '�X���[�v�J�E���g���v�l
    
    On Error Resume Next
    
    pfCheakGateConectSts = False
    bCheakAns = True
    LngSleepTotal = 0
    
    '------------------------------------------------------------------------------------
    '�����ʐM��ԃ`�F�b�N
    '------------------------------------------------------------------------------------
    '�S���@�ʐM�ؒf�����܂Ő����J��Ԃ�
    '�������̒ʐM���S���@�ؒf�����܂ŁA�R���ȏ�o�߂����ꍇ�ُ�I���Ƃ���B
    Do While LngSleepTotal < WAIT_TIME_OUT
    
        '���@�����J��Ԃ�
        For iCnt = CNT_MIN To CONECT_JIKAI_CHK_MAX
            
            bCheakAns = True
            
            '���D�@����������Ă��邩�H
            If gudtDisp(iCnt).intJiso = JissouUmu.jissou Then
                
                '�����@��ʐM��Ԏ擾
                iConectSts = pfGetGateConectSts(iCnt + 1)
                '�ʐM��Ԃ��u0:���㒆�vor�u1:�ʐM����v���H
                If iConectSts = 0 Or iConectSts = 1 Then
                    bCheakAns = False
                    Exit For
                End If
                
            End If
        
        Next
        
        '�S���@�ʐM�ؒf���ꂽ���H
        If bCheakAns = True Then
            Exit Do
        End If
        
        Sleep (100)
        LngSleepTotal = LngSleepTotal + 100
    
    Loop
    
    '�R���ȏ�o�߂������H
    If LngSleepTotal >= WAIT_TIME_OUT Then
        Exit Function
    End If
    
    pfCheakGateConectSts = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfGetJyouiKikiConectSts
'//  �@�\����  : ��ʋ@��ʐM��Ԏ擾����
'//  �@�\�T�v  : ��ʋ@��̒ʐM��Ԃ��擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iAreId  �@[IN]��ʋ@��ʐM��ԃG���AID
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@�@�@�@�@�@��ʋ@��ʐM���
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetJyouiKikiConectSts(iAreId As Integer) As Integer
    
    On Error Resume Next
    
    pfGetJyouiKikiConectSts = -1
    
    '�h�c�ʏ�񑀍�N���X�̐���
    Set Idinf_Jyoui = New IdInfProc
   '�Q��(��ʋ@��ʐM���)�G���A����ݒ�
    Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui
    Idinf_Jyoui.IdOpen
    If Idinf_Jyoui.Errsts <> 0 Then
      Set Idinf_Jyoui = Nothing
      Exit Function
    End If
    
    '�Q��(��ʋ@��ʐM���)�G���A���k�n�b�j����B
    Idinf_Jyoui.IdLock
    If Idinf_Jyoui.Errsts <> 0 Then
      Idinf_Jyoui.IdFree
      Set Idinf_Jyoui = Nothing
      Exit Function
    End If
    
     '�G���A�̓��e��ǂݍ��ށB
    Idinf_Jyoui.id = iAreId
    Idinf_Jyoui.GetInf (CONECT)
    If Idinf_Jyoui.Errsts <> 0 Then
       Idinf_Jyoui.IdFree
       Set Idinf_Jyoui = Nothing
       Exit Function
    End If
    
    '��ʋ@��ʐM��Ԃ��擾
    pfGetJyouiKikiConectSts = CInt(Idinf_Jyoui.DataArea(0))
    
    Idinf_Jyoui.IdFree
    Set Idinf_Jyoui = Nothing
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfGetGateConectSts
'//  �@�\����  : �����ʐM��Ԏ擾����
'//  �@�\�T�v  : �����̒ʐM��Ԃ��擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iGouki  �@[IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@�@�@�@�@�@�����ʐM���
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetGateConectSts(iGouki As Integer) As Integer
    
    On Error Resume Next
    
    pfGetGateConectSts = -1
    
    Set Idinf_JikaiTuushin = New IdInfProc
    '�Q��(�����ʐM���)�G���A����ݒ�
    Idinf_JikaiTuushin.ProcMode = DATA_ID.Data_Id_JikaiTuushinJyotai
    Idinf_JikaiTuushin.IdOpen
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       Set Idinf_JikaiTuushin = Nothing
       Exit Function
    End If
     
    '�Q��(�����ʐM���)�G���A���k�n�b�j����B
    Idinf_JikaiTuushin.IdLock
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       Idinf_JikaiTuushin.IdFree
       Set Idinf_JikaiTuushin = Nothing
       Exit Function
    End If
    
    '�G���A�̓��e��ǂݍ��ށB
    Idinf_JikaiTuushin.id = IdGateComSts.GATE_COM
    Idinf_JikaiTuushin.GetJikai_Tuusin iGouki - 1
    If Idinf_JikaiTuushin.Errsts <> 0 Then
       Idinf_JikaiTuushin.IdFree
       Set Idinf_JikaiTuushin = Nothing
       Exit Function
    End If
        
    pfGetGateConectSts = CInt(Idinf_JikaiTuushin.DataArea(iGouki - 1))
    
    Idinf_JikaiTuushin.IdFree
    Set Idinf_JikaiTuushin = Nothing
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfGetJyouiKikiConectSet
'//  �@�\����  : ��ʋ@��ʐM�ڑ��^�ؒf�ݒ�擾����
'//  �@�\�T�v  : ��ʋ@��̒ʐM�ڑ��^�ؒf�ݒ���擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iKansiId  �@[IN]�G���AID
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@1�@�@�@�@�@�ؗ��ݒ�
'//�@�@�@�@�@�@�@�@�@�@�@�@ 0          �ڑ��ݒ�
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetJyouiKikiConectSet(iKansiId As Integer) As Integer
    
    On Error Resume Next
    
    pfGetJyouiKikiConectSet = -1
    
    '�h�c�ʏ�񑀍�N���X�̐���
    Set Idinf_KansiSettei = New IdInfProc
    '���L�G���A�I�[�v��
    Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
    Idinf_KansiSettei.IdOpen
    If Idinf_KansiSettei.Errsts <> 0 Then
        Set Idinf_KansiSettei = Nothing
        Exit Function
    End If
       
    '�Ď��ݒ�G���A���k�n�b�j����B
    Idinf_KansiSettei.IdLock
    If Idinf_KansiSettei.Errsts <> 0 Then
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing
        Exit Function
    End If
    
    '�Ď��ݒ�G���AID��ݒ�
    Idinf_KansiSettei.id = iKansiId
    Idinf_KansiSettei.IdGet
    If Idinf_KansiSettei.Errsts <> 0 Then
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing
        Exit Function
    End If
    
    pfGetJyouiKikiConectSet = Idinf_KansiSettei.DataArea(0)   '�ݒ���e
    
    Idinf_KansiSettei.IdFree
    Set Idinf_KansiSettei = Nothing

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : pfGetGateConectSet
'//  �@�\����  : �����ʐM�ڑ��^�ؒf�ݒ�擾����
'//  �@�\�T�v  : �����̒ʐM�ڑ��^�ؒf�ݒ���擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iGouki  �@�@[IN]���@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@1�@�@�@�@�@�ؗ��ݒ�
'//�@�@�@�@�@�@�@�@�@�@�@�@ 0          �ڑ��ݒ�
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetGateConectSet(iGouki As Integer) As Integer
    
    On Error Resume Next
    
    pfGetGateConectSet = -1
    
    Set Idinf_JikaiSettei = New IdInfProc
    '�����ݒ�G���A���I�[�v������B
    Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
    Idinf_JikaiSettei.IdOpen
    If Idinf_JikaiSettei.Errsts <> 0 Then
        Set Idinf_JikaiSettei = Nothing
        Exit Function
    End If
    
    '�����ݒ�G���A���k�n�b�j����B
    Idinf_JikaiSettei.IdLock
    If Idinf_JikaiSettei.Errsts <> 0 Then
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing
        Exit Function
    End If
    
    '�G���A�̓��e��ǂݍ��ށB
    Idinf_JikaiSettei.id = IdGate.JIKAI_CONECT_SETTEI
    Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
    If Idinf_JikaiSettei.Errsts <> 0 Then
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing
        Exit Function
    End If
    
    '�ݒ���e���擾
    pfGetGateConectSet = Idinf_JikaiSettei.DataArea(iGouki - 1)
    
    '��ԁF����
    Idinf_JikaiSettei.IdFree
    Set Idinf_JikaiSettei = Nothing
    
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : ConnctErrorProc
'//  �@�\����  : �ʐM�ؒf�ُ폈��
'//  �@�\�T�v  : �f�W�܂��͎����̒ʐM�ؒf�ňُ픭�����̏������s��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iGouki  �@�@[IN]���@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Long�@�@  1�@�@�@�@�@�ؗ��ݒ�
'//�@�@�@�@�@�@�@Long�@�@�@0          �ڑ��ݒ�
'//�@�@�@�@�@�@�@Long�@�@�@0          �ڑ��ݒ�
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub ConnctErrorProc(iTusinSts As Integer, iErrorDisp As Integer)

    On Error Resume Next

    '�ُ펞�ʐM��ʐݒ�
    miErrorSts = iTusinSts

    '�ُ펞�\�����ݒ�
    miErrorDisp = iErrorDisp
    
    '�f�W�ڑ�
    pfConnect_TusinConect (ML_DT_DESHU)

End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  �֐�����  : ErrorProc
'//  �@�\����  : �ُ폈��
'//  �@�\�T�v  :
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iGouki  �@�@[IN]���@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Long�@�@  1�@�@�@�@�@�ؗ��ݒ�
'//�@�@�@�@�@�@�@Long�@�@�@0          �ڑ��ݒ�
'//�@�@�@�@�@�@�@Long�@�@�@0          �ڑ��ݒ�
'//
'//  ORIGINAL  : (EG20 V7.3.0.1) 2013-07-08  CODED BY  [TCC] S.Kuroda
'//                 2013�N�x�{�� ���u�Ή��yEG20_KANSI03_01�z
'//  REVISIONS : (EG20 VX.X.0.X) ----------  REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub ErrorProc()

    On Error Resume Next

                                        
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                                                                                
    '�ُ핶���\��
    If (miErrorDisp = ERROR_TUSHIN_DISP) Then
        '�ʐM�ؒf���s���b�Z�[�W�\��
        MsgBox "�ʐM�ؒf�������Ɉُ킪�������܂����B���[�N�����s�R�s�[���s���܂���B", _
                        vbOKOnly + vbCritical, _
                        "���[�N�����s �R�s�["
    ElseIf (miErrorDisp = ERROR_MISOU_DISP) Then
        '�����f�[�^�쐬���s���b�Z�[�W�\��
        MsgBox "�����f�[�^�쐬�������Ɉُ킪�������܂����B���[�N�����s�R�s�[���s���܂���B", _
                        vbOKOnly + vbCritical, _
                        "���[�N�����s �R�s�["
    Else
        '�ُ�I�����b�Z�[�W�\��
        MsgBox "�ُ�I�����܂����B", vbCritical, "���[�N�����s �R�s�["
    End If
    
    ' ��ʑ����
    sCmdBtnEnabled True

End Sub

