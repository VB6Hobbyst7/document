VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKansenGateVerUpdate 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   8160
      Top             =   3600
   End
   Begin VB.ListBox LstStatus 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   360
      TabIndex        =   42
      Top             =   5880
      Width           =   7935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGateComConf 
      Caption         =   " �����؂藣��"
      Height          =   525
      Left            =   9240
      Style           =   1  '���̨���
      TabIndex        =   23
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   �� �� ���s   �R�s�["
      Height          =   525
      Left            =   9240
      Style           =   1  '���̨���
      TabIndex        =   22
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyWork_Jikko 
      Caption         =   " ���[�N �� ���s �R�s�["
      Height          =   525
      Left            =   9240
      Style           =   1  '���̨���
      TabIndex        =   21
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " ���k�t�@�C�� �� ���[�N�R�s�["
      Height          =   525
      Left            =   9240
      Style           =   1  '���̨���
      TabIndex        =   20
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "���[�N�N���A"
      Height          =   525
      Left            =   9240
      Style           =   1  '���̨���
      TabIndex        =   19
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdUSBRemove 
      Caption         =   "�}�̎�O"
      Height          =   525
      Left            =   9240
      TabIndex        =   18
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work2 
      Caption         =   " �}�� �� ���[�N�@�R�s�["
      Height          =   525
      Left            =   9240
      Style           =   1  '���̨���
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Frame fraDataSelect 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   13
      Top             =   4440
      Width           =   7935
      Begin VB.CheckBox optData 
         Caption         =   "�\���R"
         Height          =   240
         Index           =   8
         Left            =   5520
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "�\���Q"
         Height          =   240
         Index           =   7
         Left            =   5520
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "�\���P"
         Height          =   240
         Index           =   6
         Left            =   5520
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "�\���Q"
         Height          =   240
         Index           =   5
         Left            =   2760
         TabIndex        =   26
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "�\���P"
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   25
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "�n�r"
         Height          =   240
         Index           =   3
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "�T�u�b�o�t"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "���C���b�o�t"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox optData 
         Caption         =   "����b�o�t"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdModoru 
      Caption         =   " �o�[�W�����Ǘ� ��ʂ֖߂�"
      Height          =   855
      Left            =   9240
      Style           =   1  '���̨���
      TabIndex        =   0
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdSelectNone 
      Caption         =   "�S�R�[�i��I��"
      Height          =   525
      Left            =   3000
      Style           =   1  '���̨���
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "�S�R�[�i�I��"
      Height          =   525
      Left            =   360
      Style           =   1  '���̨���
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Frame fraICMDLL 
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   11775
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "���I��"
         Height          =   855
         Index           =   5
         Left            =   9840
         Style           =   1  '���̨���
         TabIndex        =   35
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "���I��"
         Height          =   855
         Index           =   4
         Left            =   7920
         Style           =   1  '���̨���
         TabIndex        =   34
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "���I��"
         Height          =   855
         Index           =   3
         Left            =   6000
         Style           =   1  '���̨���
         TabIndex        =   33
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "���I��"
         Height          =   855
         Index           =   2
         Left            =   4080
         Style           =   1  '���̨���
         TabIndex        =   32
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "���I��"
         Height          =   855
         Index           =   1
         Left            =   2160
         Style           =   1  '���̨���
         TabIndex        =   31
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdate 
         BackColor       =   &H0080FF80&
         Caption         =   "�I��"
         Height          =   855
         Index           =   0
         Left            =   240
         Style           =   1  '���̨���
         TabIndex        =   30
         Top             =   1080
         Value           =   1  '����
         Width           =   1515
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '��������
         Caption         =   "�R�[�i�U"
         Height          =   255
         Index           =   5
         Left            =   9720
         TabIndex        =   41
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '��������
         Caption         =   "�R�[�i�T"
         Height          =   255
         Index           =   4
         Left            =   7800
         TabIndex        =   40
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '��������
         Caption         =   "�R�[�i�S"
         Height          =   255
         Index           =   3
         Left            =   5895
         TabIndex        =   39
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '��������
         Caption         =   "�R�[�i�R"
         Height          =   255
         Index           =   2
         Left            =   3945
         TabIndex        =   38
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '��������
         Caption         =   "�R�[�i�Q"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   37
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblCornerNo 
         Alignment       =   2  '��������
         Caption         =   "�R�[�i�P"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         Height          =   855
         Index           =   1
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         Height          =   855
         Index           =   2
         Left            =   3945
         TabIndex        =   10
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         Height          =   855
         Index           =   3
         Left            =   5895
         TabIndex        =   9
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         Height          =   855
         Index           =   4
         Left            =   7800
         TabIndex        =   8
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         Height          =   855
         Index           =   5
         Left            =   9720
         TabIndex        =   7
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.Frame fraMakerName 
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   10575
      Begin VB.Label lblMakerName 
         BackStyle       =   0  '����
         Caption         =   "�R�[�i�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   10335
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�V�����������D�@�o�[�W�����ꊇ�X�V"
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
      TabIndex        =   1
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmKansenGateVerUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmGateVerUpdate.frm
'//  �p�b�P�[�W���F�����o�[�W�����ꊇ�X�V���
'//  �T�v        �F�����o�[�W�����ꊇ�X�V���
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-17  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.79�z
'//  REVISIONS   �F (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 �y�^���\�����P�Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z�yTOMAS�p�̈�R�s�[�Ή��z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////

Option Explicit

'�R�[�i�I��t �g�p�萔
Private Const SELECTSW_ON_MESSAGE = "�I��"    ' �t���b�Z�[�W�F�I��
Private Const SELECTSW_OFF_MESSAGE = "���I��" ' �t���b�Z�[�W�F���I��
Private Const SELECTSW_ON_COLOR = &H80FF80    ' �t�F�F�I��
Private Const SELECTSW_OFF_COLOR = &H80FFFF   ' �t�F�F���I��
Private Const SELECTSW_ON_VALUE = 1           ' �t��ԁF�I��
Private Const SELECTSW_OFF_VALUE = 0          ' �t��ԁF���I��

Private Const MN_FOLD_WRK = 0                   '�u���[�N�v�t�H���_
Private Const MN_FOLD_NOW = 1                   '�u���s�v�t�H���_
Private Const MN_FOLD_OLD = 2                   '�u���v�t�H���_

Private Const MN_MAIL_INTERVAL = 1000           '���[���^�C�}�̃C���^�[�o���l

Private Const FILE_NAME_MAX_SIZE = 12

Private Const EG20_JIKAI_KISHU = "EG6000"       'EG20 �����@�햼
Private Const EG30_JIKAI_KISHU = "EG7000"       'EG30 �����@�햼
Private Const HANKUKA_KUK = "HAN_KUKA.KUK"
Private Const INI_MAX = 5

Private Const DATA_KIND_MAX = 6                 '�f�[�^��ʐ�       'EG20 V30.1.0.1 ADD

'�yNG�ʒu�z
Private Const ERROR_HEDER = "�w�b�_"  '�w�b�_
Private Const ERROR_FOTTER = "�t�b�^" '�t�b�^
'�yNG���ځz
Private Const KISHU_NAME_ERROR = "�@�햼"       '�@�햼
Private Const FILE_NAME_ERRORE = "�t�@�C����"   '�t�@�C����
Private Const CREATE_DATA_ERROR = "�쐬���t"    '�쐬���t
Private Const VERSION_ERROR = "�o�[�W����"      '�o�[�W����

Dim FolderSyubetu As Integer                    ' �I�����\�[�X���

Dim FolderName(0 To 2, 0 To 8) As String        ' �t�H���_��
Dim TitleBox(0 To 8) As String                  ' �^�C�g����
Dim LogBox(0 To 8) As String                    ' ���O�o�͗p�^�C�g����
Dim FileList() As String                        '�t�@�C�������X�g�ꗗ�i�[�G���A
Dim FileListType() As String                 '�t�@�C�����X�g�ꗗ�i�[�G���A�i�����㎩���^�C�v���܂ށj
Dim gintUnkaiKind(0 To 8) As Integer            ' �^�����    ' EG20 V5.11.0.1�ǉ�
Dim gintProgramJudgeKind(0 To 8) As Integer     ' �v���O����������    ' EG20 V6.9.0.1�y�ʎY�Ή��zADD

Private sNGSts As String        'NG�ʒu
Private sNGKoumoku As String    'NG����
Dim HAN_KUKA_DATA As HANTEI_DATA
Private Type HANTEI_DATA
    sHederKisyu(0 To 4) As String
    sHederFile(0 To 4) As String
    sFotterKisyu(0 To 4) As String
    sFotterFile(0 To 4) As String
End Type

' ����p�t�@�C�����i�[�G���A
'EG20 V30.1.0.1 DEL START
'Private EG20_HANTEI_CPU_CHK_FILE As String
'Private EG20_MAIN_CPU_CHK_FILE As String
'Private EG20_SUB_CPU1_CHK_FILE As String
'Private EG20_SUB_CPU2_CHK_FILE As String
'Private EG20_SUB_CPU3_CHK_FILE As String
'Private EG20_MAIN_OS_CHK_FILE As String
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
'�V��������
Private EG30_HANTEI_CPU_CHK_FILE As String
Private EG30_MAIN_CPU_CHK_FILE As String
Private EG30_SUB_CPU_CHK_FILE As String
Private EG30_MAIN_OS_CHK_FILE As String
'EG20 V30.1.0.1 ADD END

Private Const NGATE_00 = -1         'TOMAS�̈�t�H���_  EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD


'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    : Form_Load
'//  �@�\����    : Form_Load������
'//  �@�\�T�v    : Form_Load���������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����        :
'//  �߂�l      :
'//
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                EG20�t�F�[�Y�Q�Ή�
'//                EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   : (EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   :(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    Dim intLoop         As Integer          ' ���[�v�J�E���^
    
    On Error Resume Next
    
    '�u�����o�[�W�����ꊇ�X�V��ʁF�\���v���O�o��
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_GAMEN_START, 0)      'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_GAMEN_START, 0)      'EG20 V30.1.0.1 ADD
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���[����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
        
  
    ' /////////////////////////////////////////////////////////////////////////
    ' // �R�[�i�ݒ�
    ' /////////////////////////////////////////////////////////////////////////
    ' �R�[�i���̐ݒ菈��
    Call gsGetCornerName
    
    For intLoop = 0 To CONECT_CORNER_MAXINDEX
    
        '�ݒ肠��̃R�[�i�������ɂ���
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
            ' /////////////////////////////////////////////////
            ' // ���x���i�R�[�i�[���̕\���j
            lblCornerNo(intLoop).Visible = True
            lblGokiBetsuNumber(intLoop).Caption = gstrCornerName(intLoop)
            lblGokiBetsuNumber(intLoop).Visible = True
            
            ' /////////////////////////////////////////////////
            ' // �t
            chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
            chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
            chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
            
            chkUpdate(intLoop).Visible = True
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
            'EG20 V30.1.0.1 ADD START
            '�����R�[�i�[�̂݉����\�Ƃ���B
'            If gintCornerType(intLoop) = CORNER_TYPE_KANSEN Then
'                chkUpdate(intLoop).Enabled = True
'            Else
'                chkUpdate(intLoop).Enabled = False
'            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�z DEL END

        Else
            lblCornerNo(intLoop).Visible = False
            lblGokiBetsuNumber(intLoop).Caption = ""
            lblGokiBetsuNumber(intLoop).Visible = False
        
            chkUpdate(intLoop).Visible = False
        End If
    
    Next intLoop
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // ���̑��R���g���[���ݒ�
    ' /////////////////////////////////////////////////////////////////////////
    LstStatus.Clear
   
    'For intLoop = 0 To 8       'EG20 V30.1.0.1 DEL
    For intLoop = 0 To DATA_KIND_MAX - 1    'EG20 V30.1.0.1 ADD
        optData(intLoop).Value = SELECTSW_ON_VALUE      ' ���`�F�b�N
    Next intLoop
    
    ' �����c�k�k�t�H���_�ݒ�
    sSetFolderName
    
    ' �ϐ��̏�����
    FolderSyubetu = 0
    
    ' �R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(True)
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FForm_Activate
'//  �@�\����    �F�����o�[�W�����ꊇ�X�V���(�A�N�e�B�u��)
'//  �@�\�T�v    �F��ʍĕ\���������s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    ' ���݃t�H�[�����X�V
    'gStrCurrentForm = sFormName_GateVerUpdate      'EG20 V30.1.0.1 DEL
    gStrCurrentForm = sFormName_KGateVerUpdate      'EG20 V30.1.0.1 ADD
    
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FForm_Deactivate
'//  �@�\����    �F�����o�[�W�����ꊇ�X�V���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v    �F���[����M�p�̃^�C�}��~
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    On Error Resume Next
    
    If blnCabfrmOpenFlg = True Then
        Call fnTsbCabCallDiverge
        Exit Sub
    End If
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdModoru_Click
'//  �@�\����    �F�u�o�[�W�����Ǘ���ʂ֖߂�v�{�^����������
'//  �@�\�T�v    �F�u�o�[�W�����Ǘ���ʂ֖߂�v�{�^�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Click()
    On Error Resume Next
    
    '�u�����o�[�W�����ꊇ�X�V��ʁF�����v���O�o��
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_GAMEN_END, 0)        'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_GAMEN_END, 0)        'EG20 V30.1.0.1 ADD

    '��ʂ�Unload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����     : cmdSelectAll_Click
'//  �@�\����     : �u�S�R�[�i�I���v�{�^����������
'//  �@�\�T�v     : �u�S�R�[�i�I���v�{�^�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����         :
'//  �߂�l       :
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   : (EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdSelectAll_Click()

    Dim intLoop     As Integer          ' ���[�v�J�E���^

    On Error Resume Next
    
    '�u�����o�[�W�����ꊇ�X�V��ʁF�S�R�[�i�I��t�����v���O�o��
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_SELECTALL_BUTTON, 0)     'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_SELECTALL_BUTTON, 0)     'EG20 V30.1.0.1 ADD

    For intLoop = 0 To CONECT_CORNER_MAXINDEX

        If chkUpdate(intLoop).Visible = True Then
            'EG20 V30.1.0.1 DEL START
'            chkUpdate(intLoop).Caption = SELECTSW_ON_MESSAGE
'            chkUpdate(intLoop).BackColor = SELECTSW_ON_COLOR
'            chkUpdate(intLoop).Value = SELECTSW_ON_VALUE
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
            'EG20 V30.1.0.1 ADD START
            '�V�����R�[�i�[�ɑ΂��Ă̂ݐ؂�ւ����s���B
'            If gintCornerType(intLoop) = CORNER_TYPE_KANSEN Then
'                chkUpdate(intLoop).Caption = SELECTSW_ON_MESSAGE
'                chkUpdate(intLoop).BackColor = SELECTSW_ON_COLOR
'                chkUpdate(intLoop).Value = SELECTSW_ON_VALUE
'            End If
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�zDEL END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�zADD START
            chkUpdate(intLoop).Caption = SELECTSW_ON_MESSAGE
            chkUpdate(intLoop).BackColor = SELECTSW_ON_COLOR
            chkUpdate(intLoop).Value = SELECTSW_ON_VALUE
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�zADD END
        End If
    Next

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����     : cmdSelectAll_Click
'//  �@�\����     : �u�S�R�[�i��I���v�{�^����������
'//  �@�\�T�v     : �u�S�R�[�i��I���v�{�^�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����         :
'//  �߂�l       :
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdSelectNone_Click()

    Dim intLoop     As Integer          ' ���[�v�J�E���^

    On Error Resume Next
    
    '�u�����o�[�W�����ꊇ�X�V��ʁF�S�R�[�i��I��t�����v���O�o��
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_SELECTALLOFF_BUTTON, 0)      'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_SELECTALLOFF_BUTTON, 0)      'EG20 V30.1.0.1 ADD

    For intLoop = 0 To CONECT_CORNER_MAXINDEX

        If chkUpdate(intLoop).Visible = True Then
            'EG20 V30.1.0.1 DEL START
'            chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
'            chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
'            chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_01�z DEL START
            'EG20 V30.1.0.1 ADD START
            '�����R�[�i�[�ɑ΂��Ă̂ݐؑւ��s���B
'            If gintCornerType(intLoop) = CORNER_TYPE_KANSEN Then
'                chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
'                chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
'                chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
'            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
            chkUpdate(intLoop).Caption = SELECTSW_OFF_MESSAGE
            chkUpdate(intLoop).BackColor = SELECTSW_OFF_COLOR
            chkUpdate(intLoop).Value = SELECTSW_OFF_VALUE
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
        End If
    Next

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����     : chkUpdate_Click
'//  �@�\����     : �u�R�[�i�ʑI���v�{�^����������
'//  �@�\�T�v     : �u�R�[�i�ʑI���v�{�^�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����         :
'//  �߂�l       :
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub chkUpdate_Click(Index As Integer)

    '�u�����o�[�W�����ꊇ�X�V��ʁF�R�[�i�I��t�����v���O�o��
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKAI_VERASION_IKKATSU_KANRI_SELECT_BUTTON, 0)        'EG20 V30.1.0.1 DEL
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_IKKATSU_KANRI_SELECT_BUTTON, 0)        'EG20 V30.1.0.1 ADD

    If chkUpdate(Index).Value = SELECTSW_ON_VALUE Then
        chkUpdate(Index).Caption = SELECTSW_ON_MESSAGE
        chkUpdate(Index).BackColor = SELECTSW_ON_COLOR
'        chkUpdate(Index).Value = SELECTSW_ON_VALUE
    Else
        chkUpdate(Index).Caption = SELECTSW_OFF_MESSAGE
        chkUpdate(Index).BackColor = SELECTSW_OFF_COLOR
'        chkUpdate(Index).Value = SELECTSW_OFF_VALUE
    End If

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F cmdClear_Click
'//  �@�\����    �F�u���[�N�N���A�v�{�^����������
'//  �@�\�T�v    �F�u���[�N�N���A�v�{�^�������������s��
'//
'//                 �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F (EG20 V30.3.0.1) 2014-11-13 REVISED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()
    
    Dim iResponse As Integer        ' MsgBox�{�^���R�[�h
    Dim iCornerLoop As Integer      ' ���[�v
    Dim iSelctLoop As Integer       ' ���[�v
    Dim bStatus As Boolean          ' ��������
    Dim iTomasFlg   As Integer      ' TOMAS�����σt���O�i�R�[�i�ꊇ�������Ɉ�̃R�[�i����TOMAS�̈�̃R�s�[���s���j    'EG20 V30.3.0.1 ADD
    
    
    On Error Resume Next
    
    '�u�����o�[�W�����ꊇ�X�V��ʁF���[�N�N���A�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_CREA_BUTTOM, 0)
    
    ' �R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(False)

    ' �m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�I�����ꂽ�R�[�i�E��ʂ́u���[�N�v�t�H���_���̃t�@�C����S�č폜���܂��B" _
                         & Chr(vbKeyReturn) & "��낵���ł����H", _
                        vbYesNo + vbExclamation, _
                        "���[�N �N���A")
    If iResponse = vbYes Then

        ' �R�[�i�I���E��ʑI���`�F�b�N
        If sSelectChk = False Then
            '�R�}���h�{�^�������E�s����
             Call sCmdBtnEnabled(True)
            Exit Sub
        End If

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

        LstStatus.Clear

        ' /////////////////////////////////////////////////
        ' // �R�[�i�P�ʂł̏���
        iTomasFlg = 0       ' EG20 V30.3.0.1 ADD
        For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
            If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
'                For iSelctLoop = 0 To 8        'EG20 V30.1.0.1 DEL
                For iSelctLoop = 0 To DATA_KIND_MAX - 1
                    If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                        FolderSyubetu = iSelctLoop
                        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
                        If iTomasFlg = 0 Then
                            bStatus = sWrkFolderRemove(NGATE_00)
                            'Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
                        End If
                        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
                        bStatus = sWrkFolderRemove(iCornerLoop)
                        Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
                    End If
                Next
                'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
                '�ŏ��̃R�[�i1��̂�TOMAS�̈�̃t�H���_�ɃR�s�[�����OK�Ȃ̂ŁA�ȍ~���Ȃ��悤�Ƀt���O��ON
                iTomasFlg = 1
                'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
            End If
        Next
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    End If

    '�R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(True)
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F cmdCopyBaitai_Work_Click
'//  �@�\����    �F�u���k�t�@�C�������[�N�R�s�[�v�t����������
'//  �@�\�T�v    �F�u���k�t�@�C�������[�N�R�s�[�v�t�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()
    
    On Error Resume Next
    
    '�R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(False)
    
    '�u�����ް�ޮ݁F���ķ�ف�ܰ���߰�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_CAB_COPY_WRK_BUTTOM, 0)
 
    ' �R�[�i�I���E��ʑI���`�F�b�N
    If sSelectChk = False Then
        '�R�}���h�{�^�������E�s����
        Call sCmdBtnEnabled(True)
        Exit Sub
    End If
    
    LstStatus.Clear
    
    '���k�t�@�C������C���X�g�[������B
    sFDInstall "LZH"

   '�R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F cmdCopyBaitai_Work2_Click
'//  �@�\����    �F�u�}�� �����[�N�R�s�[�v�t����������
'//  �@�\�T�v    �F�u�}�� �����[�N�R�s�[�v�t�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work2_Click()

    On Error Resume Next
    
    '�R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(False)
    
    '�u�����ް�ޮ݁F���ķ�ف�ܰ���߰�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_LZH_COPY_WRK_BUTTOM, 0)
 
    ' �R�[�i�I���E��ʑI���`�F�b�N
    If sSelectChk = False Then
        '�R�}���h�{�^�������E�s����
        Call sCmdBtnEnabled(True)
        Exit Sub
    End If
    
    LstStatus.Clear
    
    '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����B
    sFDInstall "STD"

   '�R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����     : cmdCopyWork_Jikko_Click
'//  �@�\����     :�u���[�N �� ���s�R�s�[�v�t����������
'//  �@�\�T�v     :�u���[�N �� ���s�R�s�[�v�t�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����         :
'//  �߂�l       :
'//
'//  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS    :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS    :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS    :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS    :(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 �y�^���\�����P�Ή��z
'//  REVISIONS   �F (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(EG20 V30.3.0.1) 2014-11-11  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
    Dim iResponse               As Integer      'MsgBox�{�^���R�[�h
    Dim iCornerLoop As Integer      ' ���[�v
    Dim iSelctLoop As Integer       ' ���[�v
    Dim bStatus As Boolean          ' ��������
    Dim iTomasFlg   As Integer      ' TOMAS�����σt���O�i�R�[�i�ꊇ�������Ɉ�̃R�[�i����TOMAS�̈�̃R�s�[���s���j'EG20 V30.3.0.1 ADD

    
    On Error Resume Next
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_COPY_NOW_BUTTOM, 0)

    Call sCmdBtnEnabled(False)

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�I�����ꂽ�R�[�i�E��ʂ̎��s�t�H���_���N���A����" _
                        & Chr(vbKeyReturn) & "���[�N�t�H���_�̃t�@�C�����R�s�[���܂���" _
                        & Chr(vbKeyReturn) & "��낵���ł����H", _
                        vbYesNo + vbExclamation, _
                        "���[�N�����s�R�s�[")
    If iResponse = vbYes Then

        ' �R�[�i�I���E��ʑI���`�F�b�N
        If sSelectChk = False Then
            '�R�}���h�{�^�������E�s����
             Call sCmdBtnEnabled(True)
            Exit Sub
        End If

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        LstStatus.Clear
        ' /////////////////////////////////////////////////
        ' // �R�[�i�P�ʂł̏���
        iTomasFlg = 0   'EG20 V30.3.0.1 ADD
        For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
            If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
                'For iSelctLoop = 0 To 8        'EG20 V30.1.0.1 DEL
                For iSelctLoop = 0 To DATA_KIND_MAX - 1     'EG20 V30.1.0.1 ADD
                    If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                        FolderSyubetu = iSelctLoop
                        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
                        If iTomasFlg = 0 Then
                            bStatus = fNewVersion(NGATE_00)
                            'Call AddMessageLstStatus(NGATE_00, FolderSyubetu, bStatus)
                        End If
                        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
                        bStatus = fNewVersion(iCornerLoop)
                        Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
' EG20 V3.0.0.2�ǉ��J�n
                        If bStatus = True Then
                            '���D�@���ʃG���A�X�V�����i����j
                            Call pubfuncCommonAreaUpdate
' EG20 V5.8.0.1�폜�J�n
'                            ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'                            Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
                            ' �^����ԍX�V
                            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1�ǉ��I��
'                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1)   ' EG20 V5.6.0.1�ǉ�           ' EG20 V5.11.0.1�폜
                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1�ǉ�
                        End If
' EG20 V3.0.0.2�ǉ��I��
                    End If
                Next
                'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
                '�ŏ��̃R�[�i1��̂�TOMAS�̈�̃t�H���_�ɃR�s�[�����OK�Ȃ̂ŁA�ȍ~���Ȃ��悤�Ƀt���O��ON
                iTomasFlg = 1
                'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
            End If
        Next

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    End If

   '�R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����     : cmdCopyOld_Jikko_Click
'//  �@�\����     :�u�� �� ���s�R�s�[�v�t����������
'//  �@�\�T�v     :�u�� �� ���s�R�s�[�v�t�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����         :
'//  �߂�l       :
'//
'//  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS    :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS    :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS    : (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 �y�^���\�����P�Ή��z
'//  REVISIONS    : (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(EG20 V30.3.0.1) 2014-11-11  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    Dim iResponse               As Integer      'MsgBox�{�^���R�[�h
    Dim iCornerLoop As Integer      ' ���[�v
    Dim iSelctLoop As Integer       ' ���[�v
    Dim bStatus As Boolean          ' ��������
    
    Dim iTomasFlg   As Integer      ' TOMAS�����σt���O�i�R�[�i�ꊇ�������Ɉ�̃R�[�i����TOMAS�̈�̃R�s�[���s���j'EG20 V30.3.0.1 ADD


    On Error Resume Next
    '�u�����ް�ޮ݁F�������s�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OLD_COPY_NOW_BUTTOM, 0)

    Call sCmdBtnEnabled(False)
    
    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�I�����ꂽ�R�[�i�E��ʂ̎��s�t�H���_���N���A����" _
                        & Chr(vbKeyReturn) & "���t�H���_�̃t�@�C�����R�s�[���܂���" _
                        & Chr(vbKeyReturn) & "��낵���ł����H", _
                        vbYesNo + vbExclamation, _
                        "�������s�R�s�[")
    If iResponse = vbYes Then

        ' �R�[�i�I���E��ʑI���`�F�b�N
        If sSelectChk = False Then
            '�R�}���h�{�^�������E�s����
             Call sCmdBtnEnabled(True)
            Exit Sub
        End If

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

        LstStatus.Clear

        ' /////////////////////////////////////////////////
        ' // �R�[�i�P�ʂł̏���
        iTomasFlg = 0       ' EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD
        For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
            If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
                'For iSelctLoop = 0 To 8                    'EG20 V30.1.0.1 DEL
                For iSelctLoop = 0 To DATA_KIND_MAX - 1     'EG20 V30.1.0.1 ADD
                    If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                        FolderSyubetu = iSelctLoop
                        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
                        If iTomasFlg = 0 Then
                            bStatus = fOldVersion(NGATE_00)
                            'Call AddMessageLstStatus(NGATE_00, FolderSyubetu, bStatus)
                        End If
                        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
                        bStatus = fOldVersion(iCornerLoop)
                        Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
' EG20 V3.0.0.2�ǉ��J�n
                        If bStatus = True Then
                            '���D�@���ʃG���A�X�V�����i����j
                            Call pubfuncCommonAreaUpdate
' EG20 V5.8.0.1�ǉ��J�n
                            ' �^����ԍX�V
                            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1�ǉ��I��
'                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1)   ' EG20 V5.6.0.1�ǉ�          ' EG20 V5.11.0.1�폜
                            Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iCornerLoop + 1, gintUnkaiKind(FolderSyubetu))   ' EG20 V5.11.0.1�ǉ�
                        Else
                            Call pubfuncErrorOccur(MN_FOLD_NOW)
                        End If
' EG20 V3.0.0.2�ǉ��I��
                    End If
                Next
                'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
                '�ŏ��̃R�[�i1��̂�TOMAS�̈�̃t�H���_�ɃR�s�[�����OK�Ȃ̂ŁA�ȍ~���Ȃ��悤�Ƀt���O��ON
                iTomasFlg = 1       ' EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD
                'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
            End If
        Next

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    End If

   '�R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F cmdGateComConf_Click
'//  �@�\����    �F�u�����؂藣���v�t����������
'//  �@�\�T�v    �F�u�����؂藣���v�t�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdGateComConf_Click()
    '�u�����ް�ޮ݁F�����؂藣���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_KIRIHANASI_BUTTOM, 0)

    '�ʐM�ڑ��E�ؒf��ʂ�\������B
    Load frmConectSts
    frmConectSts.Show 1
End Sub


'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����     : cmdUSBRemove_Click
'//  �@�\����     :�u�}�̎��O���v�t����������
'//  �@�\�T�v     :�u�}�̎��O���v�t�������������s��
'//
'//                   �^          ����            �Ӗ�
'//  ����         :
'//  �߂�l       :
'//
'//  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdUSBRemove_Click()
    On Error Resume Next
   
    '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
    ' �R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(False)
 
    ' �}�̎�O����
    Call pfRemove(Me)

    ' �R�}���h�{�^�������E�s����
    Call sCmdBtnEnabled(True)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'/
'/  �֐�����     : sSelectChk
'/  �@�\����     : �R�[�i�I���E��ʑI���`�F�b�N
'/  �@�\�T�v     : �R�[�i�I���E��ʑI���`�F�b�N�������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.79�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Function sSelectChk() As Boolean
    Dim iCnt            As Integer
    Dim bRet            As Boolean

    '�����l�ݒ�
    sSelectChk = True

    '�R�[�i�I���`�F�b�N
    bRet = False
    For iCnt = 0 To CONECT_CORNER_MAXINDEX
        
        If chkUpdate(iCnt).Value = SELECTSW_ON_VALUE Then
            bRet = True
        End If
    Next
    
    If bRet = False Then
        sSelectChk = False
        MsgBox "�R�[�i���I������Ă��܂���B", _
                vbOKOnly + vbExclamation, _
                "�R�[�i�I��"
        Exit Function
    End If

    '��ʑI���`�F�b�N
    bRet = False
    'For iCnt = 0 To 8                      'EG20 V30.1.0.1 DEL
    For iCnt = 0 To DATA_KIND_MAX - 1       'EG20 V30.1.0.1 ADD
        
        If optData(iCnt).Value = SELECTSW_ON_VALUE Then
            bRet = True
        End If
    Next
    
    If bRet = False Then
        sSelectChk = False
' EG20 V3.6.0.1�y03����TR-No.79�z�폜�J�n
'        MsgBox "��ʂ�����Ă��܂���B", _
'                vbOKOnly + vbExclamation, _
'                "��ʑI��"
' EG20 V3.6.0.1�y03����TR-No.79�z�폜�I��
' EG20 V3.6.0.1�y03����TR-No.79�z�ǉ��J�n
        MsgBox "��ʂ��I������Ă��܂���B", _
                vbOKOnly + vbExclamation, _
                "��ʑI��"
' EG20 V3.6.0.1�y03����TR-No.79�z�ǉ��I��
        Exit Function
    End If

End Function

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
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/                  �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)
    Dim iLoopCnt    As Integer

    '�R�[�i�I��t
    For iLoopCnt = 0 To CONECT_CORNER_MAXINDEX
        'chkUpdate(iLoopCnt).Enabled = blnFlg       'EG20 V30.1.0.1 DEL
        chkUpdate(iLoopCnt).Enabled = blnFlg       'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�zADD
        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
        'EG20 V30.1.0.1 ADD START
        '�����R�[�i�[�ɑ΂��Ă̂ݐݒ�\�Ƃ���B
'        If gintCornerType(iLoopCnt) = CORNER_TYPE_KANSEN Then
'            chkUpdate(iLoopCnt).Enabled = blnFlg
'        End If
        'EG20 V30.1.0.1 ADD END
        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
    Next
    
    '�S�R�[�i�I���E��I��t
    cmdSelectAll.Enabled = blnFlg
    cmdSelectNone.Enabled = blnFlg
    
    '��ʑI��
    'For iLoopCnt = 0 To 8                  'EG20 V30.1.0.1 DEL
    For iLoopCnt = 0 To DATA_KIND_MAX - 1   'EG20 V30.1.0.1 ADD
        optData(iLoopCnt).Enabled = blnFlg
    Next
    
    '�e�R�}���h�t
    cmdClear.Enabled = blnFlg             ' �u���[�N�N���A�v
    cmdCopyBaitai_Work.Enabled = blnFlg   ' �u���k�t�@�C�������[�N�R�s�[�v
    cmdCopyBaitai_Work2.Enabled = blnFlg  ' �u�}�́����[�N�R�s�[�v
    cmdCopyWork_Jikko.Enabled = blnFlg    ' �u���[�N�����s�R�s�[�v
    cmdCopyOld_Jikko.Enabled = blnFlg     ' �u�������s�R�s�[�v
    cmdGateComConf.Enabled = blnFlg       ' �u�����؂藣���v
    cmdUSBRemove.Enabled = blnFlg         ' �u�}�̎�O�v
    cmdModoru.Enabled = blnFlg            ' �u�߂�v

End Sub
'EG20 V30.1.0.1 DEL START
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
''//
''//  �֐�����    �F sSetFolderName
''//  �@�\����    �F �f�[�^�W�J
''//  �@�\�T�v    �F �t�H���_���Ȃǂ̃f�[�^���O���[�o���G���A�ɓW�J����B
''//
''//                 �^        ����      �Ӗ�
''//  ����        �F �Ȃ�
''//
''//                 �^        �l        �Ӗ�
''//  �߂�l      �F �Ȃ�
''//
''//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
''//                 EG20�t�F�[�Y�Q�Ή�
''//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
''//  REVISIONS   �F (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
''//                 �y�^���\�����P�Ή��z
''//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
''//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z
''//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
''//  ���l        �F
''///////////////////////////////////////////////////////////////////
'Private Sub sSetFolderName()
'
'    TitleBox(0) = "����f�[�^  "
'    TitleBox(1) = "�v���O����  "
'    TitleBox(2) = "�T�uCPU-Pro1"
'    TitleBox(3) = "�T�uCPU-Pro2"
'    TitleBox(4) = "�T�uCPU-Pro3"
'    TitleBox(5) = "�����i�n�r�j"
'    TitleBox(6) = "�\��1       "
'    TitleBox(7) = "�\��2       "
'    TitleBox(8) = "�\��3       "
'
'    LogBox(0) = "����"
'    LogBox(1) = "�v��"
'    LogBox(2) = "�T�u1"
'    LogBox(3) = "�T�u2"
'    LogBox(4) = "�T�u3"
'    LogBox(5) = "OS"
'    LogBox(6) = "�\��1"
'    LogBox(7) = "�\��2"
'    LogBox(8) = "�\��3"
'
'    '�t�H���_���ɐݒ���s��
'    FolderName(MN_FOLD_WRK, 0) = EG20_NHAN1WRK       ' ����f�[�^CPU-PRO�i���[�N�j
'    FolderName(MN_FOLD_NOW, 0) = EG20_NHAN1NOW       ' ����f�[�^CPU-PRO�i���s�j
'    FolderName(MN_FOLD_OLD, 0) = EG20_NHAN1OLD       ' ����f�[�^CPU-PRO�i���j
'
'    FolderName(MN_FOLD_WRK, 1) = EG20_NPRO1WRK       ' ���D�@�v���O�������i���[�N�j
'    FolderName(MN_FOLD_NOW, 1) = EG20_NPRO1NOW       ' ���D�@�v���O�������i���s�j
'    FolderName(MN_FOLD_OLD, 1) = EG20_NPRO1OLD       ' ���D�@�v���O�������i���j
'
'    FolderName(MN_FOLD_WRK, 2) = EG20_NSCP1WRK       ' �T�uCPU1-PRO�i���[�N�j
'    FolderName(MN_FOLD_NOW, 2) = EG20_NSCP1NOW       ' �T�uCPU1-PRO�i���s�j
'    FolderName(MN_FOLD_OLD, 2) = EG20_NSCP1OLD       ' �T�uCPU1-PRO�i���j
'
'    FolderName(MN_FOLD_WRK, 3) = EG20_NSCP2WRK       ' �T�uCPU2-PRO�i���[�N�j
'    FolderName(MN_FOLD_NOW, 3) = EG20_NSCP2NOW       ' �T�uCPU2-PRO�i���s�j
'    FolderName(MN_FOLD_OLD, 3) = EG20_NSCP2OLD       ' �T�uCPU2-PRO�i���j
'
'    FolderName(MN_FOLD_WRK, 4) = EG20_NSCP3WRK       ' �T�uCPU3-PRO�i���[�N�j
'    FolderName(MN_FOLD_NOW, 4) = EG20_NSCP3NOW       ' �T�uCPU3-PRO�i���s�j
'    FolderName(MN_FOLD_OLD, 4) = EG20_NSCP3OLD       ' �T�uCPU3-PRO�i���j
'
'    FolderName(MN_FOLD_WRK, 5) = EG20_NOSWRK         ' ���D�@�iOS�j���i���[�N�j
'    FolderName(MN_FOLD_NOW, 5) = EG20_NOSNOW         ' ���D�@�iOS�j���i���s�j
'    FolderName(MN_FOLD_OLD, 5) = EG20_NOSOLD         ' ���D�@�iOS�j���i���j
'
'    FolderName(MN_FOLD_WRK, 6) = EG20_NYOBI1WRK      ' �\��1�i���[�N�j
'    FolderName(MN_FOLD_NOW, 6) = EG20_NYOBI1NOW      ' �\��1�i���s�j
'    FolderName(MN_FOLD_OLD, 6) = EG20_NYOBI1OLD      ' �\��1�i���j
'
'    FolderName(MN_FOLD_WRK, 7) = EG20_NYOBI2WRK      ' �\��2�i���[�N�j
'    FolderName(MN_FOLD_NOW, 7) = EG20_NYOBI2NOW      ' �\��2�i���s�j
'    FolderName(MN_FOLD_OLD, 7) = EG20_NYOBI2OLD      ' �\��2�i���j
'
'    FolderName(MN_FOLD_WRK, 8) = EG20_NYOBI3WRK      ' �\��3�i���[�N�j
'    FolderName(MN_FOLD_NOW, 8) = EG20_NYOBI3NOW      ' �\��3�i���s�j
'    FolderName(MN_FOLD_OLD, 8) = EG20_NYOBI3OLD      ' �\��3�i���j
'
'    ' /////////////////////////////////////////////////////
'    ' // EG20����
'    ' �L�[��:����CPU-PRO��\
'    EG20_HANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
'
'    ' �L�[��:���C��CPU-PRO��\
'    EG20_MAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_MAIN_PRO, PATH_GATEVER_FILE)
'
'    ' �L�[���F�T�uCPU-PRO��\
'    EG20_SUB_CPU1_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO1, PATH_GATEVER_FILE)
'    EG20_SUB_CPU2_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO2, PATH_GATEVER_FILE)
'    EG20_SUB_CPU3_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_SUB_PRO3, PATH_GATEVER_FILE)
'
'    ' �L�[��:���C��CPU-OS��\
'    EG20_MAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_EG20, GATE_MAIN_OS, PATH_GATEVER_FILE)
'
'' EG20 V5.11.0.1�y�^���\�����P�Ή��z�ǉ��J�n
'    gintUnkaiKind(0) = BootInfoGateType.TYPE_NHAN
'    gintUnkaiKind(1) = BootInfoGateType.TYPE_NPRO
'    gintUnkaiKind(2) = BootInfoGateType.TYPE_NSCP1
'    gintUnkaiKind(3) = BootInfoGateType.TYPE_NSCP2
'    gintUnkaiKind(4) = BootInfoGateType.TYPE_NSCP3
'    gintUnkaiKind(5) = BootInfoGateType.TYPE_NOS
'    gintUnkaiKind(6) = BootInfoGateType.TYPE_NYOBI1
'    gintUnkaiKind(7) = BootInfoGateType.TYPE_NYOBI2
'    gintUnkaiKind(8) = BootInfoGateType.TYPE_NYOBI3
'' EG20 V5.11.0.1�y�^���\�����P�Ή��z�ǉ��I��
'
'' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD START
'    gintProgramJudgeKind(0) = ProgramJudgeKind.JUDGE_NHAN       ' ����f�[�^
'    gintProgramJudgeKind(1) = ProgramJudgeKind.JUDGE_NPRO       ' �v���O����
'    gintProgramJudgeKind(2) = ProgramJudgeKind.JUDGE_NSCP1      ' �T�uCPU-Pro1
'    gintProgramJudgeKind(3) = ProgramJudgeKind.JUDGE_NSCP2      ' �T�uCPU-Pro2
'    gintProgramJudgeKind(4) = ProgramJudgeKind.JUDGE_NSCP3      ' �T�uCPU-Pro3
'    gintProgramJudgeKind(5) = ProgramJudgeKind.JUDGE_NOS        ' �����iOS�j
'    gintProgramJudgeKind(6) = ProgramJudgeKind.JUDGE_NOCHECK    ' �\��1
'    gintProgramJudgeKind(7) = ProgramJudgeKind.JUDGE_NOCHECK    ' �\��2
'    gintProgramJudgeKind(8) = ProgramJudgeKind.JUDGE_NOCHECK    ' �\��3
'' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD END
'
'End Sub
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-18  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSetFolderName()

        TitleBox(0) = "����b�o�t"
        TitleBox(1) = "���C���b�o�t"
        TitleBox(2) = "�T�u�b�o�t"
        TitleBox(3) = "�n�r"
        TitleBox(4) = "�\���P"
        TitleBox(5) = "�\���Q"
    
        LogBox(0) = "����"
        LogBox(1) = "�v���O���C��"
        LogBox(2) = "�T�u"
        LogBox(3) = "�n�r"
        LogBox(4) = "�\���P"
        LogBox(5) = "�\���Q"
        
        '�t�H���_���ɐݒ���s��
        FolderName(0, 0) = EG30_JHANWRK
        FolderName(1, 0) = EG30_JHANNOW
        FolderName(2, 0) = EG30_JHANOLD
        FolderName(0, 1) = EG30_JPROWRK
        FolderName(1, 1) = EG30_JPRONOW
        FolderName(2, 1) = EG30_JPROOLD
        FolderName(0, 2) = EG30_JSCPUWRK
        FolderName(1, 2) = EG30_JSCPUNOW
        FolderName(2, 2) = EG30_JSCPUOLD
        FolderName(0, 3) = EG30_JOSWRK
        FolderName(1, 3) = EG30_JOSNOW
        FolderName(2, 3) = EG30_JOSOLD
        FolderName(0, 4) = EG30_JYOBIWK1
        FolderName(1, 4) = EG30_JYOBINW1
        FolderName(2, 4) = EG30_JYOBIOD1
        FolderName(0, 5) = EG30_JYOBIWRK
        FolderName(1, 5) = EG30_JYOBINOW
        FolderName(2, 5) = EG30_JYOBIOLD

'-------�V��������-------
    ' �L�[��:����CPU-PRO��\
    EG30_HANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    ' �L�[��:���C��CPU-PRO��\
    EG30_MAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_MAIN_PRO, PATH_GATEVER_FILE)
        
    ' �L�[���F�T�uCPU-PRO��\
    EG30_SUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_SUB_PRO1, PATH_GATEVER_FILE)
    
    ' �L�[��:���C��CPU-OS��\
    EG30_MAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_MAIN_OS, PATH_GATEVER_FILE)

    gintUnkaiKind(0) = BootInfoGateType.TYPE_JHAN
    gintUnkaiKind(1) = BootInfoGateType.TYPE_JPRO
    gintUnkaiKind(2) = BootInfoGateType.TYPE_JSCPU
    gintUnkaiKind(3) = BootInfoGateType.TYPE_JOS

    gintProgramJudgeKind(0) = ProgramJudgeKind.JUDGE_JHAN       'a:����CPU�p�v���O�����i�����j
    gintProgramJudgeKind(1) = ProgramJudgeKind.JUDGE_JPRO       'b:���C��CPU�p�v���O�����i�����j
    gintProgramJudgeKind(2) = ProgramJudgeKind.JUDGE_JSCPU     'c:�T�uCPU�v���O�����i�����j
    gintProgramJudgeKind(3) = ProgramJudgeKind.JUDGE_JOS        ' d:OS�v���O�����i�����j
    gintProgramJudgeKind(4) = ProgramJudgeKind.JUDGE_YOBI1      'e:�\���P�i�����j �`�F�b�N����
    gintProgramJudgeKind(5) = ProgramJudgeKind.JUDGE_YOBI       'f:�\���i�����j �`�F�b�N����
    gintProgramJudgeKind(6) = ProgramJudgeKind.JUDGE_NOCHECK
    gintProgramJudgeKind(7) = ProgramJudgeKind.JUDGE_NOCHECK
    gintProgramJudgeKind(8) = ProgramJudgeKind.JUDGE_NOCHECK

End Sub
'EG20 V30.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sWrkFolderRemove
'//  �@�\����    �F ���[�N�t�H���_���t�@�C���폜����
'//  �@�\�T�v    �F ���[�N�t�H���_���̃t�@�C�����폜����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-17  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 �y�^���\�����P�Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove(nCorner As Integer) As Boolean
    Dim MyName As String                    ' �t�@�C����
    Dim lngErrCode As Long                  ' �G���[�R�[�h
    Dim lngPgmHanteiStsWork As Long         ' �v���O���������ԁi���[�N�j   ' EG20 V3.6.0.1�ǉ�
    
    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                       ' �t�@�C���I�u�W�F�N�g
    
    On Error GoTo ErrorHandler              ' �G���[�n���h���̓o�^

    '�����l�ݒ�
    sWrkFolderRemove = True
   
    '���[�N�t�H���_���̃f�B���N�g���̖��O��\�����܂��B�i�R�[�i�P�ʁj
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
    
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

' EG20 V3.6.0.1�ǉ��J�n
    '�Ď��ݒ�G���A�u�v���O��������ُ��ԁi���[�N�j�v�̏�Ԃ��擾����
    lngPgmHanteiStsWork = pfGetKansiSet(IdKansiSet.PG_HANTEI_ERR_STS_WORK)

    '�u�v���O��������ُ��ԁi���[�N�j�v�i����j
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_ERR_STS_WORK, ErrCode.Normal)
    
    '�ω����������ꍇ�A�u��ԕω��ʒm�v�𑗐M����
    If lngPgmHanteiStsWork <> ErrCode.Normal Then
        Call sSendMailStsChgInf(MailSts.stsNormal, lngPgmHanteiStsWork)
    End If
' EG20 V3.6.0.1�ǉ��I��

' EG20 V5.11.0.1�폜�J�n
'' EG20 V5.8.0.1�폜�J�n
''    ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
''    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1�ǉ�
'' EG20 V5.8.0.1�폜�I��
'' EG20 V5.8.0.1�ǉ��J�n
'    ' �^����ԍX�V
'    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_NASHI)
'' EG20 V5.8.0.1�ǉ��I��
'    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, nCorner + 1)   ' EG20 V5.6.0.1�ǉ�
' EG20 V5.11.0.1�폜�I��
' EG20 V5.11.0.1�ǉ��J�n
    ' �^����ԍX�V
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_CLEAR)
    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, nCorner + 1, gintUnkaiKind(FolderSyubetu))
' EG20 V5.11.0.1�ǉ��I��

    Exit Function '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    Call pubfuncErrorOccur(MN_FOLD_WRK)             ' EG20 V3.6.0.1�ǉ�
           
   '�u�����ް�ޮ݁Fܰ�̫���̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    Set objFso = Nothing
    Set objFi = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sFDInstall
'//  �@�\����    �F �}�̃C���X�g�[������
'//  �@�\�T�v    �F �C���X�g�[���}�̃t�@�C�����A���[�N�t�H���_�ɃR�s�[����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F String    sFlag     �������
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 �y�^���\�����P�Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall(sFlag As String)
    Dim iResponse As Integer        'MsgBox�{�^���R�[�h
    Dim sInputPass As String        '�C���X�g�[�����f�B���N�g����(STD)or�t�@�C����(LZH)
    Dim sInputFolder As String      '�C���X�g�[�����t�H���_���BLZH�̎��A�𓀐�t�H���_�B
    Dim lngErrCode As Long          '�G���[�R�[�h
    Dim bRet As Boolean             '�������`�F�b�N�߂�l
    Dim bStatus As Boolean          ' ��������
    Dim iCornerLoop As Integer      ' ���[�v
    Dim iSelctLoop As Integer       ' ���[�v
   
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                    '�t�@�C���I�u�W�F�N�g

    Dim lngPgmHanteiStsWork As Long     '�v���O���������ԁi���[�N�j   ' EG20 V3.0.0.2�ǉ�

    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^

    If sFlag = "STD" Then
    '�W���i�񈳏k�j�t�@�C���w��̎�:
    '�f�B���N�g���I����ʂ�\�������A���̓t�@�C���i�[�f�B���N�g�����𓾂�B
        sInputPass = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
        If sInputPass = "" Then
        '�f�B���N�g�����w��Ȃ����͏����I��
            Set objFso = Nothing
            Set objFi = Nothing
            Exit Sub
        End If
        sInputFolder = sInputPass
    Else
    '���k�t�@�C���w��̎�:
    '���k�t�@�C���I����ʂ�\�������ALZH�t�@�C���t���p�X���𓾂�i�f�t�H���g�͂e�c��\���B�j�B
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
        
        Call ChDrive("D")
        
       '�𓀗p�ꎞ�t�H���_���쐬����B
       psMakeFolder MELTED_FOLDER_FULLPASS
       '���k�t�@�C�����A�𓀗p�ꎞ�t�H���_�ɉ𓀁E�i�[������B
        Call psCabReqest(CABREQEST.CAB_THAW, sInputPass, MELTED_FOLDER_FULLPASS)
        If glngCabErrCd <> 0 Then
            psDeleteFolder MELTED_FOLDER_FULLPASS
            Set objFso = Nothing
            Set objFi = Nothing
            Exit Sub
        End If
        sInputFolder = MELTED_FOLDER_FULLPASS
    End If
    
    Set objFso = Nothing
    Set objFi = Nothing
    
    '�u���[�N�R�s�[�m�F�v�|�b�v�A�b�v��ʕ\��
    iResponse = MsgBox(sInputPass & " �̑S�Ẵt�@�C�����A" _
                       & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                       & "�u���[�N�v�t�H���_�ɃR�s�[���܂��B " _
                       & "��낵���ł����H", _
                       vbYesNo + vbExclamation, _
                       "�}�́����[�N �R�s�[")
    If iResponse = vbNo Then
    '[������] �{�^����I��:�������Ȃ��B
    '�A���A���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B
        If sFlag = "LZH" Then
            psDeleteFolder MELTED_FOLDER_FULLPASS
        End If
        Exit Sub
    End If
    
' EG20 V6.9.0.1 �y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zDEL START
'    '�O�����̓v�����������`�F�b�N
'    If sFlag = "STD" Then
'       '�}�́����[�N �R�s�[��
'       bRet = pfInstallSeitouseiChck(sInputPass)
'    Else
'       '���k�t�@�C�������[�N �R�s�[��
'       bRet = pfInstallSeitouseiChck(MELTED_FOLDER_FULLPASS & "\")
'    End If
'    If bRet = False Then
'       Exit Sub
'    End If
' EG20 V6.9.0.1 �y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zDEL END
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    ' /////////////////////////////////////////////////
    ' // �R�[�i�P�ʂł̏���
    For iCornerLoop = 0 To CONECT_CORNER_MAXINDEX
        If chkUpdate(iCornerLoop).Value = SELECTSW_ON_VALUE Then
            'For iSelctLoop = 0 To 8                    'EG20 V30.1.0.1 DEL
            For iSelctLoop = 0 To DATA_KIND_MAX - 1     'EG20 V30.1.0.1 ADD
                If optData(iSelctLoop).Value = SELECTSW_ON_VALUE Then
                    FolderSyubetu = iSelctLoop
                    bStatus = sFDInstall2(iCornerLoop, sFlag, sInputFolder)
                    Call AddMessageLstStatus(iCornerLoop, FolderSyubetu, bStatus)
' EG20 V3.0.0.2�ǉ��J�n
                    If bStatus = True Then
                        '�Ď��ݒ�G���A�u�v���O��������ُ��ԁi���[�N�j�v�̏�Ԃ��擾����
                        lngPgmHanteiStsWork = pfGetKansiSet(IdKansiSet.PG_HANTEI_ERR_STS_WORK)

                        '�u�v���O��������ُ��ԁi���[�N�j�v�i����j
                        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_ERR_STS_WORK, ErrCode.Normal)
                        
                        '�ω����������ꍇ�A�u��ԕω��ʒm�v�𑗐M����
                        If lngPgmHanteiStsWork <> ErrCode.Normal Then
                            Call sSendMailStsChgInf(MailSts.stsNormal, lngPgmHanteiStsWork)
                        End If
                    
' EG20 V5.8.0.1�폜�J�n
'                        ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'                        Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
                        ' �^����ԍX�V
                        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1�ǉ��I��
'                        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iCornerLoop + 1)   ' EG20 V5.6.0.1�ǉ�           ' EG20 V5.11.0.1�폜
                        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iCornerLoop + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1�ǉ�
                    Else
                        Call pubfuncErrorOccur(MN_FOLD_WRK)
                    End If
' EG20 V3.0.0.2�ǉ��I��
                End If
            Next
        End If
    Next
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B(�g�p�ς݂̂���)
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
    
    Exit Sub    '�������I������

ErrorHandler:   ' �G���[�����B
    Set objFso = Nothing
    Set objFi = Nothing
    Call pubfuncErrorOccur(MN_FOLD_WRK)             ' EG20 V3.0.0.2�ǉ�

' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
    '���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
    '�u�����ް�ޮ݁F�}�́�ܰ���߰�����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sFDInstall2
'//  �@�\����    �F �}�̃C���X�g�[������
'//  �@�\�T�v    �F �C���X�g�[���}�̃t�@�C�����A���[�N�t�H���_�ɃR�s�[����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F String    sFlag     �������
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z�yTOMAS�p�̈�R�s�[�Ή��z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function sFDInstall2(nCorner As Integer, sFlag As String, sInputFolder As String) As Boolean

    Dim MyName As String            '�t�@�C���t���p�X��
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim lngErrCode As Long          '�G���[�R�[�h
    Dim sChkName As String          '�`�F�b�N�t�@�C��
    Dim szTargetFolder As String    ' �����ύX��t�H���_��            ' EG20 V5.8.0.1�ǉ�
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                    '�t�@�C���I�u�W�F�N�g
    Dim bRet As Boolean             '�������`�F�b�N�߂�l             ' EG20 V6.9.0.1ADD
    
    Dim sTomasPath As String        ' TOMAS�p�̈�t�@�C���p�X
    
    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^

    
    sFDInstall2 = True

' EG20 V6.9.0.1 �y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD START
    bRet = pfInstallSeitouseiChck(sInputFolder & "\")
    If bRet = False Then
        Set objFso = Nothing
        Set objFi = Nothing
        sFDInstall2 = False
        Exit Function           '�������I������
    End If
' EG20 V6.9.0.1 �y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD END

' EG20 V5.8.0.1�ǉ��J�n
    szTargetFolder = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu)
' EG20 V5.8.0.1�ǉ��I��
    '�o�[�W�����`�F�b�N�t�@�C���L���`�F�b�N���s���B
    sChkName = fSelectFile
    
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
   
    If objFso.FileExists(gstrMyPath & sChkName) = True Then
        '�w��t�@�C�������݂���
        sChkName = objFso.GetFileName(gstrMyPath & sChkName)
        Kill gstrMyPath & sChkName
    Else
        sChkName = ""
    End If
    
    '�w��t�H���_���̃t�@�C�����A�S�āu���[�N�v�t�H���_�ɃR�s�[����B
    For Each objFi In objFso.GetFolder(sInputFolder).files   '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then  '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            '�}�̓��t�@�C�������쐬
            sSrcFileName = sInputFolder & "\" & MyName
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
                '���[�N�t�H���_���t�@�C�������쐬����
                sDstFileName = gstrMyPath & MyName
                '�}�̓��̃t�@�C�������[�N�t�H���_�ɃR�s�[����
                FileCopy sSrcFileName, sDstFileName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    
'    '���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B(�g�p�ς݂̂���)
'    If sFlag = "LZH" Then
'        psDeleteFolder MELTED_FOLDER_FULLPASS
'    End If

' EG20 V5.8.0.1�ǉ��J�n
    ' �����ύX����
    dllChangeAttributeContents (szTargetFolder)
' EG20 V5.8.0.1�ǉ��I��
    
' EG20 V6.9.0.1 �y�ʎY�Ή��FTOMAS�p�̈�R�s�[�Ή��zADD START
    ' �������ׂ��Ώۂ��R�[�i1�̏ꍇ
    ' TOMAS�̈�iN_GATE00�j��N_GATE01�̓��e�ŃR�s�[
    'If nCorner = 0 Then                            'EG20 V30.1.0.1 DEL
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
    '���[�N�R�s�[���悤�Ƃ��邽�тɂ��̃R�[�i����00�փR�s�[���邽�߁A�擪�R�[�i�̔�����폜
    'If nCorner = gintKansenFirstCornerIdx Then      'EG20 V30.1.0.1 ADD
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
        ' �폜��̃t�H���_�iTOMAS�̈�j���w��
        sTomasPath = PATH_GATE_EG20 & "00" & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
        
        ' TOMAS�̈���폜
        If funcRemoveFile(sTomasPath) = False Then
           
            sFDInstall2 = False
            '�u�����ް�ޮ݁FTOMAS̫���̧�ٍ폜�ُ�v���O�o��
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_DELETE_ERROR, lngErrCode)
        
            Exit Function '�������I������
        End If
        
        ' TOMAS�̈�փR�s�[
        If funcCopyFile(gstrMyPath, sTomasPath, lngErrCode) = False Then
            
            sFDInstall2 = False
            '�u�����ް�ޮ݁FTOMAS�̈��߰�����ُ�v���O�o��
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_COPY_ERROR, lngErrCode)
        
            Exit Function '�������I������
        End If
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
    'End If
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
' EG20 V6.9.0.1 �y�ʎY�Ή��FTOMAS�p�̈�R�s�[�Ή��zADD END
    
    '�u�����ް�ޮ݁F�}�́�ܰ���߰��������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_OK, 0)
    
    Exit Function '�������I������

ErrorHandler:   ' �G���[�����B
    Set objFso = Nothing
    Set objFi = Nothing
    
    sFDInstall2 = False

' EG20 V5.8.0.1�ǉ��J�n
    ' �����ύX����
    dllChangeAttributeContents (szTargetFolder)
' EG20 V5.8.0.1�ǉ��I��

    '�u�����ް�ޮ݁F�}�́�ܰ���߰�����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F AddLstStatus
'//  �@�\����    �F �������ʃ��X�g�ւ̃��b�Z�[�W�o�͏���
'//  �@�\�T�v    �F �������ʃ��X�g�ւ̃��b�Z�[�W�o�͏���
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nProc     �����ԍ��i�����t�j
'//                 Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//                 Integer   nDataKind ������ʁi0�`�j
'//                 Boolean   bResult   �������ʁiTRUE:����AFALSE:�ُ�j
'//
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F Boolean   bResult   �������ʁiTRUE:����AFALSE:�ُ�j
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Sub AddMessageLstStatus(nCorner As Integer, _
                                      nDataKind As Integer, _
                                      bResult As Boolean)
    Dim szOutMessae As String       ' �o�̓��b�Z�[�W
    Dim szWorkMsg As String         ' �o�̓��b�Z�[�W�i���[�N�j
    
    On Error Resume Next

    szOutMessae = "�R�[�i" & Format(nCorner + 1, "00") & "�F" & _
                    TitleBox(nDataKind) & "�F"
    
    If bResult = True Then
        szWorkMsg = "����I��"
    Else
        szWorkMsg = "�ُ�I��"
    End If
    
    szOutMessae = szOutMessae & szWorkMsg

    LstStatus.AddItem (szOutMessae)
    LstStatus.Selected(LstStatus.ListCount - 1) = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F pfInstallSeitouseiChck
'//  �@�\����    �F �O�����̓v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v    �F �O�����̓v���O��������f�[�^�������`�F�b�N�������s���B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F �Ȃ�
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F Boolean   bResult   �������ʁiTRUE:����AFALSE:�ُ�j
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z
'//     REVISIONS :(EG20 V6.11.0.1) 2013-03-27 REVISED BY  [TCC] H.Kondoh
'//                 �}�̓����@�\�ύX�Ή�
'//                   ��ʂO�̏ꍇ���ُ�Ƃ���悤�ɕύX
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function pfInstallSeitouseiChck(sInputPass As String) As Boolean
    Dim myLen As Long                        '������̒���
    Dim lngSumRet As Long
    Dim i As Integer
    Dim bRet As Boolean
    Dim lngCnt As Long
    Dim sSrcFileName As String               '�t�@�C�����X�g��
    Dim lngErrCode   As Long
    Dim intCheckKind As Integer              ' �`�F�b�N���     ' EG20 V6.9.0.1ADD
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error Resume Next
    
    pfInstallSeitouseiChck = True
    
    '********************************
    '*�v�����������`�F�b�N
    '********************************
    '�O���}�̃t�H���_���t�@�C�������쐬
    sSrcFileName = sInputPass & MN_FILELIST
    '�O���}�̂̌���������
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    Else
     '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      pfInstallSeitouseiChck = False
      Set objFso = Nothing
      Exit Function
    End If

   '����[�N��t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(sInputPass & MN_FILELIST)

    '�T���l�`�F�b�N
    For lngCnt = 0 To UBound(FileList) - 1
        If pfFileSumChk(sInputPass & FileList(lngCnt), lngSumRet) <> True Then
            pfInstallSeitouseiChck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    '�t�@�C�����ő�`�F�b�N
    If UBound(FileList) > FILECNT_MAX Then
      pfInstallSeitouseiChck = False

      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)

      Exit Function
    End If
    For i = 0 To UBound(FileList) - 1
       '�擾�t�@�C�����̃T�C�Y���擾
       myLen = LenB(StrConv(Trim(FileList(i)), vbFromUnicode))                                              '���p���Z�̃o�C�g�����擾
       If FILE_NAME_MAX_SIZE < myLen Then
          '13�o�C�g�ȏ�̏ꍇ
           bRet = False
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
           Exit For
       End If
    Next

' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD START
    If bRet = False Then
        pfInstallSeitouseiChck = bRet
        Exit Function
    End If

    For i = 0 To UBound(FileList) - 1
        ' �t�@�C�����X�g���̎�ʂ𒊏o
        'intCheckKind = CInt(Left$(FileListType(i), 1))         'EG20 V30.1.0.1 DEL
        intCheckKind = Asc(Left$(FileListType(i), 1))   'EG20 V30.1.0.1 ADD
'EG20 V6.11.0.1 DEL Start
'        If ((gintProgramJudgeKind(FolderSyubetu) = intCheckKind) Or _
'            (intCheckKind = ProgramJudgeKind.JUDGE_NOCHECK)) Then
'            ' �f�[�^��ʑI�𕔂̑I����e�ƃt�@�C�����X�g���̎�ʂ̔�r���ʂ��u��v�v�A��������
'            ' �t�@�C�����X�g���̎�ʂ��u�`�F�b�N�Ȃ��v
'            ' ���`�F�b�N���ʐ���
'EG20 V6.11.0.1 DEL End
'EG20 V6.11.0.1 ADD Start
        If (gintProgramJudgeKind(FolderSyubetu) = intCheckKind) Then
            ' �f�[�^��ʑI�𕔂̑I����e�ƃt�@�C�����X�g���̎�ʂ̔�r���ʂ��u��v�v
            ' ���`�F�b�N���ʐ���
'EG20 V6.11.0.1 ADD End
            bRet = True
        Else
            ' ��L�ȊO
            ' ���`�F�b�N���ʈُ�
            bRet = False
            ' �G���[���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_PRGKIND_ERROR, 0)
            Exit For
        End If
    Next
' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD END

    pfInstallSeitouseiChck = bRet
Exit Function

FileGetError:
    pfInstallSeitouseiChck = False
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F fNewVersion
'//  �@�\����    �F �ŐV�o�[�W��������
'//  �@�\�T�v    �F �ŐV(���[�N)�o�[�W�������A���s(���s)�o�[�W�����ɓo�^
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V5.0.2.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��yTR-51�C���Ή��z
'//  REVISIONS   �F(EG20 V30.3.0.1) 2014-11-11  CODED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function fNewVersion(nCorner As Integer) As Boolean
    Dim bRet As Boolean                      '�߂�l
    Dim sSrcFileName            As String    '���[�N�t�H���_���t�@�C�����X�g
    Dim lngErrCode As Long                   '�G���[�R�[�h
    Dim iKansiAplChk As Integer              '�A�v���N���`�F�b�N�߂�l�@'V1.6.0.1 ADD
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error Resume Next
    
    '����[�N��t�H���_�̃t�@�C�����X�g����������
    '���[�N�t�H���_���t�@�C�������쐬
    '���[�N�t�H���_���̃f�B���N�g���̖��O��\�����܂��B�i�R�[�i�P�ʁj
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\" & MN_FILELIST
    
    '�t�@�C���̌���������
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    Else
        '�t�@�C�������݂��Ȃ�
        '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

        fNewVersion = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function
    End If
    
    bRet = pfSeitouseiChck(nCorner)    'V1.4.0.1�@ADD
    '����[�N��t�H���_����t�@�C�����X�g���A�o�^�t�@�C�������J�E���g����
    If bRet = True Then
       bRet = fReadFileList(sSrcFileName)
    End If

    If bRet = True Then
        '�����t�H���_���̃t�@�C����S�č폜����
        If sOldFolderRemove(nCorner) <> True Then
'            Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�    EG20 V3.6.0.1�폜
            Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1�ǉ�
            fNewVersion = False
            Exit Function
        End If

        '�����t�H���_���̃t�@�C������s��t�H���_�̓��e�ɒu������
        If sCopyNOWtoOLD(nCorner) <> True Then
'            Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�    EG20 V3.6.0.1�폜
            Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1�ǉ�
            fNewVersion = False
            Exit Function
        End If

        '����s��t�H���_���̃t�@�C���𢃏�[�N��t�H���_�̓��e�ɒu������
        If sCopyWRKtoNOW(nCorner) <> True Then
            Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�
            fNewVersion = False
            Exit Function
        End If
    
        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
'        '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
'        '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
'        iKansiAplChk = CheckAppStart(PROC_KANRI)
'        If iKansiAplChk <> 0 Then
'            '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
''            psVersionUpdateReqest (ML_REQUEST_EG20GATE)
'            frmVerUpdateIkkatsu.Show vbModal
'        Else
'            '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
'            gintGateVerInfUpdRes = MailSts.stsNormal
'        End If
'
'        '���D�@�o�[�W�����X�V��������
'        If gintGateVerInfUpdRes = MailSts.stsNormal Then
'            '����
'            fNewVersion = True
'        Else
'            '�ُ�
'            fNewVersion = False
'        End If
        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
        
        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
        If nCorner <> NGATE_00 Then
            '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
            '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
            iKansiAplChk = CheckAppStart(PROC_KANRI)
            If iKansiAplChk <> 0 Then
                '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
    '            psVersionUpdateReqest (ML_REQUEST_EG20GATE)
                frmVerUpdateIkkatsu.Show vbModal
            Else
                '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
                gintGateVerInfUpdRes = MailSts.stsNormal
            End If
        
            '���D�@�o�[�W�����X�V��������
            If gintGateVerInfUpdRes = MailSts.stsNormal Then
                '����
                fNewVersion = True
            Else
                '�ُ�
                fNewVersion = False
            End If
        Else
            fNewVersion = True
        End If
        'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END
  
'        fNewVersion = True             ' EG20 V5.0.2.1�폜
    Else
        fNewVersion = False
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F fOldVersion
'//  �@�\����    �F ���o�[�W��������
'//  �@�\�T�v    �F �ꐢ��O�̃o�[�W���������s(���s)�o�[�W�����ɕԂ��B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V5.0.2.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��yTR-51�C���Ή��z
'//  REVISIONS   �F(EG20 V30.3.0.1) 2014-11-11  CODED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function fOldVersion(nCorner As Integer) As Boolean
    Dim bRet As Boolean                     '�߂�l
    Dim sSrcFileName            As String   '���t�H���_���t�@�C�����X�g
    Dim lngErrCode              As Long     '�G���[�R�[�h
    Dim iKansiAplChk As Integer              '�A�v���N���`�F�b�N�߂�l�@'V1.6.0.1 ADD

    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error Resume Next
 
   '���t�H���_���̃t�@�C�����X�g����������B
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    Else                                '�t�@�C�������݂��Ȃ�
        '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)
 
        fOldVersion = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '�������I������
    End If
    
    '�����t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(sSrcFileName)

' EG20 V3.6.0.1 �y����TR-No.260�z�ǉ��J�n
    bRet = fDataFileCheck(sSrcFileName)
    If bRet = False Then
       Call pubfuncErrorOccur(MN_FOLD_OLD)
       fOldVersion = False
       Exit Function
    End If
' EG20 V3.6.0.1 �y����TR-No.260�z�ǉ��I��

  ' EG20 V3.0.0.2�ǉ��J�n
    ' ���D�@���ʔ��菈��
    bRet = pubfuncCommonGateCheck(nCorner, MN_FOLD_OLD)
    If bRet = False Then
        fOldVersion = False
       Exit Function
    End If
  ' EG20 V3.0.0.2�ǉ��I��

    '����s��t�H���_���̃t�@�C����S�č폜����
    If sNowFolderRemove(nCorner) <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.6.0.1�ǉ�
        fOldVersion = False
        Exit Function
    End If
    
    '�����t�H���_���̃t�@�C������s��t�H���_�̓��e�ɒu������
    If sCopyOLDtoNOW(nCorner) <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.6.0.1�ǉ�
        fOldVersion = False
        Exit Function
    End If
    
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
'    '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
'    '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
'     iKansiAplChk = CheckAppStart(PROC_KANRI)
'     If iKansiAplChk <> 0 Then
'        '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
''         psVersionUpdateReqest (ML_REQUEST_EG20GATE)
'        frmVerUpdateIkkatsu.Show vbModal
'    Else
'        '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
'        gintGateVerInfUpdRes = MailSts.stsNormal
'    End If
'
'     '���D�@�o�[�W�����X�V�����ُ�
'    If gintGateVerInfUpdRes = MailSts.stsNormal Then
'        '����
'        fOldVersion = True
'    Else
'        '�ُ�
'        fOldVersion = False
'    End If
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
    
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD START
    If nCorner <> NGATE_00 Then
        '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
        '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
         iKansiAplChk = CheckAppStart(PROC_KANRI)
         If iKansiAplChk <> 0 Then
            '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
    '         psVersionUpdateReqest (ML_REQUEST_EG20GATE)
            frmVerUpdateIkkatsu.Show vbModal
        Else
            '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
            gintGateVerInfUpdRes = MailSts.stsNormal
        End If
         
         '���D�@�o�[�W�����X�V�����ُ�
        If gintGateVerInfUpdRes = MailSts.stsNormal Then
            '����
            fOldVersion = True
        Else
            '�ُ�
            fOldVersion = False
        End If
    Else
        fOldVersion = True
    End If
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z ADD END

'    fOldVersion = True                 ' EG20 V5.0.2.1�폜
End Function


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sOldFolderRemove
'//  �@�\����    �F ���t�H���_���t�@�C���폜����
'//  �@�\�T�v    �F ���t�H���_���̃t�@�C�����폜����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function sOldFolderRemove(nCorner As Integer) As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^
   
   '�߂�l������
    sOldFolderRemove = True
 
    '�u���s�v�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\"
    
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
    
    Exit Function           '�������I������

ErrorHandler:   ' �G���[�������[�`���B
    '�u�����ް�ޮ݁F���t�H���_̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLDFILE_DELETE_ERROR, lngErrCode)

    sOldFolderRemove = False
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sNowFolderRemove
'//  �@�\����    �F ���s�t�H���_���̃t�@�C���폜����
'//  �@�\�T�v    �F ���s�t�H���_���̃t�@�C�����폜����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function sNowFolderRemove(nCorner As Integer) As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g

    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    '�����l�ݒ�
    sNowFolderRemove = True
    
    '�u���s�v�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\"
    
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

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    '�u�����ް�ޮ݁F���s�t�H���_̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOWFILE_DELETE_ERROR, lngErrCode)

    sNowFolderRemove = False
    
    Set objFso = Nothing
    Set objFi = Nothing
    
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sCopyNOWtoOLD
'//  �@�\����    �F ���s�o�[�W�����ۑ�����
'//  �@�\�T�v    �F ���s�t�H���_���̃t�@�C�����A���t�H���_�ɃR�s�[����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function sCopyNOWtoOLD(nCorner As Integer) As Boolean
    Dim MyName As String                '�t�@�C����
    Dim sSrcFileName As String          '�R�s�[���t�@�C���̃t���p�X��
    Dim sDstFileName As String          '�R�s�[��t�@�C���̃t���p�X��
    
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    
    On Error GoTo ErrorHandler              '�G���[�n���h���ݒ�
  
    '�߂�l������
    sCopyNOWtoOLD = True
   
    '���s�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    gstrMyPath = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\"
    
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                '���s�t�H���_���t�@�C�������쐬����
                sSrcFileName = gstrMyPath & MyName

                '���t�H���_���t�@�C�������쐬����
                sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MyName

                '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
                FileCopy sSrcFileName, sDstFileName

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    
    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    
    sCopyNOWtoOLD = False
    
    Set objFso = Nothing
    Set objFi = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sCopyWRKtoNOW
'//  �@�\����    �F �ŐV�o�[�W�����R�s�[
'//  �@�\�T�v    �F ���[�N�t�H���_���̃t�@�C�����A���s�t�H���_�ɃR�s�[�B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��iPASSINF�R�s�[�Ή��j
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function sCopyWRKtoNOW(nCorner As Integer) As Boolean
    
    Dim i As Integer                '�J�E���^
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim bError As Boolean           '�t���O
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^
  
    '�߂�l������
    sCopyWRKtoNOW = True
    
    '****************************
    '* �t�@�C�����X�g���R�s�[���� *
    '****************************
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\" & MN_FILELIST
                                    '���[�N�t�H���_���t�@�C�������쐬����
    sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '���s�t�H���_���t�@�C�������쐬����
    If objFso.FileExists(sSrcFileName) = True Then     '�t�@�C���̌���������   'V1.20.0.1 ADD
        '�t�@�C�����X�g���u���[�N�v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
        FileCopy sSrcFileName, sDstFileName
    Else                                '�t�@�C�������݂��Ȃ�
        sCopyWRKtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '�������I������
    End If

    bError = False                  '�G���[�t���O���u�U�v�ɂ���
    For i = 0 To UBound(FileList) - 1
                                    '�t�@�C�����X�g�ꗗ�����J��Ԃ�
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\" & FileList(i)
                                    '���[�N�t�H���_���t�@�C�������쐬����
        sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)
                                    '���s�t�H���_���t�@�C�������쐬����

        '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
        If objFso.FileExists(sSrcFileName) = True Then   '�t�@�C���̌���������   'V1.20.0.1 ADD
            '�t�@�C�����u���[�N�v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
            FileCopy sSrcFileName, sDstFileName
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2�ǉ��J�n
    If pfuncCopyPASSINF(nCorner, MN_FOLD_WRK) = False Then
        sCopyWRKtoNOW = False
    End If
' EG20 V3.0.0.2�ǉ��I��
    
    Exit Function                           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    sCopyWRKtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F sCopyOLDtoNOW
'//  �@�\����    �F ���o�[�W�����ɖ߂�����
'//  �@�\�T�v    �F ���t�H���_���̃t�@�C�����A���s�t�H���_�ɃR�s�[����B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��iPASSINF�R�s�[�Ή��j
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function sCopyOLDtoNOW(nCorner As Integer) As Boolean
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
    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '���[�N�t�H���_���t�@�C�������쐬����
    sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '���s�t�H���_���t�@�C�������쐬����
    
    If objFso.FileExists(sSrcFileName) = True Then '�t�@�C���̌���������   'V1.20.0.1 ADD
        '�t�@�C�����X�g���u���v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
        FileCopy sSrcFileName, sDstFileName
    Else
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '�������I������
    End If

    bError = False                  '�G���[�t���O���u�U�v�ɂ���
    For i = 0 To UBound(FileList) - 1
                                    '�t�@�C�����X�g�����J��Ԃ�
        '���t�H���_���t�@�C�������쐬����
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)

        '���s�t�H���_���t�@�C�������쐬����
        sDstFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)

        '���t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
        If objFso.FileExists(sSrcFileName) = True Then '�t�@�C���̌���������   'V1.20.0.1 ADD
            '�t�@�C�����u���v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
            FileCopy sSrcFileName, sDstFileName
        Else                                '�t�@�C�������݂��Ȃ�
            bError = True                   '�G���[�t���O���u�^�v�ɂ���
        End If
    Next
    If bError = True Then
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function
    End If

    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2�ǉ��J�n
    If pfuncCopyPASSINF(nCorner, MN_FOLD_OLD) = False Then
        sCopyOLDtoNOW = False
    End If
' EG20 V3.0.0.2�ǉ��I��
    
    Exit Function       '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    sCopyOLDtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F pfSeitouseiChck
'//  �@�\����    �F �v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v    �F �v���O��������f�[�^�������`�F�b�N�������s���B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function pfSeitouseiChck(nCorner As Integer) As Boolean
    Dim bRet As Boolean
    
    Dim szTargetFolder As String            ' �Ώۃt�H���_
    
    On Error Resume Next
    
    pfSeitouseiChck = True
    
    szTargetFolder = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_WRK, FolderSyubetu) & "\"
    '********************************
    '*�v�����������`�F�b�N
    '********************************
    '�����v���O��������f�[�^�������`�F�b�N���s��(�Ώۃt�@�C���FHAN_KUKA.KUK)
    bRet = fDataFileCheck(szTargetFolder & MN_FILELIST)
    If bRet = False Then
'       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�     EG20 V3.6.0.1�폜
       Call pubfuncErrorOccur(MN_FOLD_WRK)          ' EG20 V3.6.0.1�ǉ�
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2�ǉ��J�n
    ' ���D�@���ʔ��菈��
    bRet = pubfuncCommonGateCheck(nCorner, MN_FOLD_WRK)
    If bRet = False Then
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2�ǉ��I��

    '�@�퐳�����`�F�b�N(�Ώۃt�@�C���FXX_GATEY.VEF�@XX:���[�U�[���@Y�F�f�[�^���)
    bRet = fKishuCheck(szTargetFolder)
    If bRet = False Then
       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�
       pfSeitouseiChck = False
       Exit Function
    End If

    pfSeitouseiChck = bRet
Exit Function

FileGetError:
    pfSeitouseiChck = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����    �F fDataFileCheck
'//  �@�\����    �F �����v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v    �F �ΏۂƂȂ�HAN_KUKA.KUK�L���`�F�b�N���s���B
'//
'//                 �^        ����      �Ӗ�
'//  ����        �F String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//                 �^        �l        �Ӗ�
'//  �߂�l      �F �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
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
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
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
    
    '�t�@�C���@�퐳�����`�F�b�N���s���B
    For i = 0 To iCnt - 1
         '�`�F�b�N�Ώۃt�@�C���p�X�쐬
        sChkFileName = sFileList & FileList(i)
    
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
            
            '�������o
            'sChkData = Left(sKisyu, Len(EG20_JIKAI_KISHU))      'EG20 V30.1.0.1 DEL
            sChkData = Left(sKisyu, Len(EG30_JIKAI_KISHU))       'EG20 V30.1.0.1 DEL
            'If EG20_JIKAI_KISHU = sChkData Then        'EG20 V30.1.0.1 DEL
            If EG30_JIKAI_KISHU = sChkData Then         'EG20 V30.1.0.1 ADD
                bRet = True  '�@�퐳�����F����
            Else
                bRet = False '�@�퐳�����F�ُ�
                fKishuCheck = bRet
                Set objFso = Nothing    'V1.20.0.1 ADD
                Exit Function
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.sSetChkFile���p
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.fSelectFile���p
'///////////////////////////////////////////////////////////////////
Private Function fSelectFile() As String
 
    fSelectFile = ""
    '�o�[�W�����`�F�b�N�t�@�C������ݒ肷��B
    Select Case FolderSyubetu
        'EG20 V30.1.0.1 DEL START
'       Case 0 '����CPU-Pro
'            fSelectFile = EG20_HANTEI_CPU_CHK_FILE
'
'       Case 1 '���C��CPU-Pro
'            fSelectFile = EG20_MAIN_CPU_CHK_FILE
'
'       Case 2 '�T�uCPU1-Pro
'            fSelectFile = EG20_SUB_CPU1_CHK_FILE
'
'       Case 3 '�T�uCPU2-Pro
'            fSelectFile = EG20_SUB_CPU2_CHK_FILE
'
'       Case 4 '�T�uCPU3-Pro
'            fSelectFile = EG20_SUB_CPU3_CHK_FILE
'
'       Case 5 '���C��CPU-OS
'            fSelectFile = EG20_MAIN_OS_CHK_FILE
       'EG20 V30.1.0.1 DEL END
       'EG20 V30.1.0.1 ADD START
       Case 0 '����CPU-Pro
            fSelectFile = EG30_HANTEI_CPU_CHK_FILE
       
       Case 1 '���C��CPU-Pro
            fSelectFile = EG30_MAIN_CPU_CHK_FILE
       
       Case 2 '�T�uCPU-Pro
            fSelectFile = EG30_SUB_CPU_CHK_FILE
       
       Case 3 '���C��CPU-OS
            fSelectFile = EG30_MAIN_OS_CHK_FILE
       'EG20 V30.1.0.1 ADD END
     
     End Select


End Function

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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.fReadFileList���p
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
            'EG20 V30.1.0.1 DEL START
'            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
'            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.1.0.1 ADD START
            '�t�@�C����ʂ͑啶���ɕϊ������A�t�@�C����������啶���ɕϊ�����悤�ɂ���B�i���܂ł͎�ʂ�����������������Ȃ������j
            FileListType(iListCnt - 1) = Trim$(Left$(sFileName, 18))
            FileList(iListCnt - 1) = UCase(Mid$(FileListType(iListCnt - 1), 3, 16))
            'EG20 V30.1.0.1 ADD�@END
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'//  REVISIONS   �F(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.fHankukaChck���p
'///////////////////////////////////////////////////////////////////
Private Function fHankukaChck(sFilePath As String) As Boolean
    Dim iFileNumber As Integer           '�t�@�C���ԍ�
    Dim i As Integer
    Dim lSts As Long
    Dim sKeyName As String
    Dim lPos As Long                     '�o�[�W�������i�[�ʒu
    Dim lLen As Long                     '�t�@�C���T�C�Y
    'Dim uFooter As MN_FOOT          '�t�b�^���i�[�G���A      'EG20 V30.1.0.1 DEL
    Dim uFooter As MN_KAN_FOOT          '�t�b�^���i�[�G���A   'EG20 V30.1.0.1 ADD
    Dim sDateTime As String
    Dim j As Integer
    Dim lngErrCode As Long          '�G���[�R�[�h
    Dim uHeder As HAN_KUKA_KUK_HEADER       '�w�b�_���i�[�G���A
    Dim sGetInfo As String * MAX_PATH_SIZE  'INI�t�@�C���擾�p
    Dim sChkFileData As String
    Dim iMojisu As Integer
    
    Dim bChkSts As Boolean              '�`�F�b�N���ʃt���O
    Dim sChkData As String              '��r�������o
    
   '�������F����(�u�����N�j
    sNGSts = ""
    sNGKoumoku = ""
    'V1.4.0.1 ADD END
    Dim oFs As New FileSystemObject 'V2.5.0.1 ADD
    
    fHankukaChck = False
    
 '�t�@�C���L���`�F�b�N���s���B
 If oFs.FileExists(sFilePath) = False Then
    '�t�@�C����������ΐ������`�F�b�N���s��Ȃ��B
    fHankukaChck = True
    Set oFs = Nothing
    Exit Function
 End If

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
      'EG20 V30.1.0.1 DEL START
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 DEL END
      'EG20 V30.1.0.1 ADD START
      lSts = GetPrivateProfileString(EG30_HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 ADD END
      If lSts = False Then
       
      Else
        HAN_KUKA_DATA.sHederKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      '�w�b�_�F���Ғl�t�@�C�����擾
      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
      'EG20 V30.1.0.1 DEL START
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 DEL END
      'EG20 V30.1.0.1 ADD START
      lSts = GetPrivateProfileString(EG30_HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 ADD END
      If lSts = False Then
        
      Else
         HAN_KUKA_DATA.sHederFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      'EG20 V30.1.0.1 DEL START�i�V�����̓t�b�^�����j
      '�t�b�^�F���Ғl�@�햼�擾
'      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'
'      Else
'        HAN_KUKA_DATA.sFotterKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
'      End If
'      '�t�b�^�F���Ғl�t�@�C�����擾
'      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'
'      Else
'        HAN_KUKA_DATA.sFotterFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
'      End If
       'EG20 V30.1.0.1 DEL END
    Next i
    'V1.4.0.1 ADD END

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�
    
    'HAN_KUKA.KUK�t�@�C���T�C�Y�擾
    lLen = FileLen(sFilePath)
    
    '���g�p�̃t�@�C���ԍ����擾����
    iFileNumber = FreeFile
    
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
   
   '�w�b�_���F�@�햼�`�F�b�N
   iMojisu = InStr(uHeder.sKisyuName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sKisyuName, 1)
   Else
     sChkFileData = Mid(uHeder.sKisyuName, 1, iMojisu)
   End If
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

   '�w�b�_���F�t�@�C�����`�F�b�N
   iMojisu = InStr(uHeder.sProgrumName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sProgrumName, 1)
   Else
     sChkFileData = Mid(uHeder.sProgrumName, 1, iMojisu)
   End If

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
    
   'EG20 V30.1.0.1 DEL START
'   '�t�b�^���F�@�햼�`�F�b�N
'   iMojisu = InStr(uFooter.sKisyu, " ") - 1
'   If iMojisu < 0 Then
'     sChkFileData = Mid(uFooter.sKisyu, 1)
'   Else
'     sChkFileData = Mid(uFooter.sKisyu, 1, iMojisu)
'   End If
'
'   bChkSts = False
'    For i = 0 To INI_MAX - 1
'       If HAN_KUKA_DATA.sFotterKisyu(i) <> "" Then
'          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterKisyu(i)))
'          If sChkData = HAN_KUKA_DATA.sFotterKisyu(i) Then
'             bChkSts = True
'           Exit For
'          End If
'      End If
'    Next
'    '�`�F�b�N���ʃt���O����
'    If bChkSts = False Then
'       '�@�햼���Ғl�S�s��v�F
'        sNGSts = ERROR_FOTTER
'        sNGKoumoku = KISHU_NAME_ERROR
'         GoTo ErrorHandler
'    End If
'
'   '�t�b�^���F�t�@�C�����`�F�b�N
'   iMojisu = InStr(uFooter.sFileName, " ") - 1
'   If iMojisu < 0 Then
'     sChkFileData = Mid(uFooter.sFileName, 1)
'   Else
'     sChkFileData = Mid(uFooter.sFileName, 1, iMojisu)
'   End If
'
'   bChkSts = False
'    For i = 0 To INI_MAX - 1
'       If HAN_KUKA_DATA.sFotterFile(i) <> "" Then
'          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterFile(i)))
'          If sChkData = HAN_KUKA_DATA.sFotterFile(i) Then
'             bChkSts = True
'           Exit For
'          End If
'      End If
'    Next
'    '�`�F�b�N���ʃt���O����
'    If bChkSts = False Then
'       '�@�햼���Ғl�S�s��v�F
'        sNGSts = ERROR_FOTTER
'        sNGKoumoku = FILE_NAME_ERRORE
'         GoTo ErrorHandler
'    End If
    'EG20 V30.1.0.1 DEL END
      
    '�t�b�^���F�쐬���t�����l���ǂ���
     sDateTime = ""
     For j = 0 To 3
         sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
     Next
    'sDateTime = sDateTime & " " 'V1.4.0.1 DEL
     For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_FOTTER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    'EG20 V30.1.0.1 DEL START �V�����̃t�b�^���ɂ̓o�[�W�����͑��݂��Ȃ�
'    '�o�[�W�����l�`�F�b�N
'    '�t�b�^���F�o�[�W�����l�����l���ǂ���
'    If IsNumeric(uFooter.sVersion) = False Then
'       sNGSts = ERROR_FOTTER
'       sNGKoumoku = VERSION_ERROR
'       GoTo ErrorHandler
'       Exit Function
'    End If
'    'V1.4.0.1 ADD END
    'EG20 V30.1.0.1 DEL END
    
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
        AppActivate frmKansenGateVerUpdate.Caption, False
        pfFormActive (frmKansenGateVerUpdate.hwnd)
    End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : pubfuncCommonGateCheck
'//  �@�\����  : ���D�@���ʔ��菈��
'//  �@�\�T�v  : �T���l�`�F�b�N�A�t�@�C�����ő�`�F�b�N�̎��s
'//
'//              �^        ����             �Ӗ�
'//  ����      �FInteger   nCorner   �R�[�i�ԍ��i0�`5�j
'//  ����      : Integer    nKind           MN_FOLD_WRK(0):���[�N
'//                                         MN_FOLD_NOW(1):���s
'//                                         MN_FOLD_OLD(2):��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : BOOL      TRUE      ����
'//                        FALSE     �ُ�
'//
'//  ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                  EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pubfuncCommonGateCheck(nCorner As Integer, nKind As Integer) As Boolean

    Dim lngSumRet As Long
    Dim lngCnt As Long
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim i As Integer
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim bRet As Boolean
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    Dim myLen As Long                        '������̒���
    Dim sSrcFileName            As String    '���[�N�t�H���_���t�@�C�����X�g
    Dim lTotalCount As Long                  ' ���ʌ���

    Dim lngPgmHanteiRcvErrSts   As Long     '�v���O���������M�ُ���
    Dim lngPgmHanteiSndErrSts   As Long     '�v���O��������z�M�ُ���
    Dim lngPgmHanteiErrSts      As Long     '�v���O��������ُ��ԁi���s�j
    Dim lngPgmHanteiErrStsOld   As Long     '�v���O��������ُ��ԁi���j
    Dim lngPgmHanteiElseErrSts  As Long     '�v���O�������肻�̑��ُ���

    
    On Error Resume Next

    ' /////////////////////////////////////////////////////
    ' // �T���l�`�F�b�N
    For lngCnt = 0 To UBound(FileList) - 1
        
        '
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(nKind, FolderSyubetu) & "\" & FileList(lngCnt)
        If pfFileSumChk(sSrcFileName, lngSumRet) <> True Then
            
            '�u�v���O���������M�ُ��ԁv�擾
            lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
        
            '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
            Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_SumChk)
                    
            '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
            If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_SumChk Then
                Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_SumChk)
            End If
            
' ���b�Z�[�W�{�b�N�X�͕\�����Ȃ�
'            '�T���l�ُ�
'            If lngSumRet = SUM_CHK.SumErr Then
'               MsgBox "�T���l���ُ�ł��B" _
'                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                      vbOKOnly + vbExclamation, _
'                      "�������D�@ �o�[�W�����Ǘ�"
'
'            '�T���l�ُ�ȊO�ُ�
'            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
'               MsgBox "�ُ�I�����܂����B", _
'                     vbOKOnly + vbExclamation, _
'                      "�������D�@ �o�[�W�����Ǘ�"
'            End If
            pubfuncCommonGateCheck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    ' /////////////////////////////////////////////////////
    ' // �t�@�C�����ő�`�F�b�N
    If UBound(FileList) > FILECNT_MAX Then

        '�u�v���O���������M�ُ��ԁv
        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)

        '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
                
        '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
        End If

' ���b�Z�[�W�{�b�N�X�͕\�����Ȃ�
'        MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
'                & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                vbOKOnly + vbExclamation, _
'                "�������D�@ �o�[�W�����Ǘ�"
        pubfuncCommonGateCheck = False

        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
        Exit Function
    End If

    'EG20 V30.1.0.1 DEL START �k���V�����ł͑S��ʂ̏���l�������Ă��Ȃ��̂Ń`�F�b�N�͕s�v�Ƃ���
'    ' /////////////////////////////////////////////////////
'    ' // �S�t�@�C�����ő�`�F�b�N�i���s�{�ǉ����j
'    bRet = True
'    lTotalCount = pfuncTotalListCount(nCorner)
'    lTotalCount = lTotalCount + UBound(FileList)
'    If lTotalCount > TOTALFILECNT_MAX Then
'        bRet = False
'    End If
'
'    If bRet = False Then
'        '�u�v���O���������M�ُ��ԁv
'        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
'
'        '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
'        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
'
'        '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
'        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
'            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
'        End If
'
'' ���b�Z�[�W�{�b�N�X�͕\�����Ȃ�
''        MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
''                & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
''                vbOKOnly + vbExclamation, _
''                "�������D�@ �o�[�W�����Ǘ�"
'        pubfuncCommonGateCheck = False
'
'        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
'        Exit Function
'    End If
    'EG20 V30.1.0.1 DEL END
    
    pubfuncCommonGateCheck = True
    Exit Function

' �����{
'    ' /////////////////////////////////////////////////////
'    ' // �t�@�C�����T�C�Y�`�F�b�N
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
'
'    sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(nKind, FolderSyubetu) & "\" & MN_FILELIST
'    '�t�@�C�����X�g���I�[�v���B
'    Open sSrcFileName For Input As #iFileNumber
'
'    bRet = True
'    For i = 0 To lngFileListCnt
'        If i = lngFileListCnt Then
'            Exit For
'        End If
'
'        '�t�@�C�������擾����B
'        Input #iFileNumber, strWork
'        If strWork <> "" And Left$(strWork, 1) <> "/" Then  '�t�@�C���������݂���
'            '�t�@�C������`�Ȃ�
'            If strWork = "" Then
'                '���[�v����
'' ���b�Z�[�W�{�b�N�X�͕\�����Ȃ�
''                MsgBox "�t�@�C�������ُ�ł��B" _
''                        & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
''                        vbOKOnly + vbExclamation, _
''                        "�������D�@ �o�[�W�����Ǘ�"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            '�t�H�[�}�b�g�ُ�
'            ElseIf " " <> Mid(strWork, 2, 1) Then
'              '���[�v����
'' ���b�Z�[�W�{�b�N�X�͕\�����Ȃ�
''                MsgBox "�t�@�C�������ُ�ł��B" _
''                        & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
''                        vbOKOnly + vbExclamation, _
''                        "�������D�@ �o�[�W�����Ǘ�"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            ElseIf (InStr(strWork, ".") - 1) = -1 Then
'' ���b�Z�[�W�{�b�N�X�͕\�����Ȃ�
''                MsgBox "�t�@�C�������ُ�ł��B" _
''                        & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
''                        vbOKOnly + vbExclamation, _
''                        "�������D�@ �o�[�W�����Ǘ�"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            Else
'                '�t�@�C�����݂̂𒊏o
'                sGetFileListName = Mid(strWork, 3, 16)
'                '�擾�t�@�C�����̃T�C�Y���擾
'                myLen = LenB(StrConv(sGetFileListName, vbFromUnicode))      '���p���Z�̃o�C�g�����擾
'                If FILE_NAME_MAX_SIZE < myLen Then
'                    '13�o�C�g�ȏ�̏ꍇ
'' ���b�Z�[�W�{�b�N�X�͕\�����Ȃ�
''                    MsgBox "�t�@�C�������ُ�ł��B" _
''                            & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
''                            vbOKOnly + vbExclamation, _
''                            "�������D�@ �o�[�W�����Ǘ�"
'                    bRet = False
'                    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                    Exit For
'                End If
'            End If
'        End If
'    Next
'
'    If bRet = False Then
'        '�u�v���O���������M�ُ��ԁv
'        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
'
'        '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
'        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
'
'        '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
'        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
'            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
'        End If
'    End If
'    '�t�@�C�����X�g���N���[�Y�B
'    Close #iFileNumber
'    pubfuncCommonGateCheck = bRet

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pubfuncCommonGateCheck = False
    
    '�u�v���O���������M�ُ��ԁv
    lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)

    '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
            
    '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
    If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
        Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfuncTotalListCount
'//  �@�\����  : �����X�g���̎擾
'//  �@�\�T�v  : �w���ʈȊO�̑��t�@�C�������Z�o����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//
'//              �^        �l               �Ӗ�
'//  �߂�l    : LONG      lResultCount     ����
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.fReadFileList���p
'///////////////////////////////////////////////////////////////////
Private Function pfuncTotalListCount(nCorner As Integer) As Long
    Dim lResultCount As Long                ' ���ʌ���
    Dim iLoop As Integer                    ' ���[�v
    
    Dim iFileNumber As Integer              '�t�@�C���ԍ�
    Dim sFileName As String                 '�t�@�C����
    Dim sSrcFileName As String              '�t�@�C����
    Dim iListCnt As Integer                 '�t�@�C���i�[��
    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�
    
    lResultCount = 0
    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����
    For iLoop = 0 To 8
        
        iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����
        sSrcFileName = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(0, iLoop) & "\" & MN_FILELIST
   
        If objFso.FileExists(sSrcFileName) = True Then
            Open sSrcFileName For Input Access Read As #iFileNumber     '�t�@�C�����X�g�̃I�[�v��
            iListCnt = 0
            Do While Not EOF(iFileNumber)                               '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
                Line Input #iFileNumber, sFileName                      '�f�[�^��ǂݍ��݂܂��B
                If sFileName <> "" And Left$(sFileName, 1) <> "/" Then  '�t�@�C���������݂���
                    iListCnt = iListCnt + 1                             '�t�@�C�����̃J�E���^���A�b�v����
                End If
            Loop
            Close #iFileNumber      '�t�@�C������܂��B
            iFileNumber = 0
            If iLoop <> FolderSyubetu Then
                lResultCount = lResultCount + iListCnt
            End If
        End If
    Next

    pfuncTotalListCount = lResultCount    '�߂�l��ݒ肷��
    Set objFso = Nothing

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    pfuncTotalListCount = 0    '�߂�l��ݒ肷��
    Set objFso = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : pfuncCopyPASSINF
'//  �@�\����  : ���s�t�H���_�ւ�PASSINF�R�s�[
'//  �@�\�T�v  : �w���ʈȊO�̑��t�@�C�������Z�o����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//  ����      : Integer    nKind           MN_FOLD_WRK(0):���[�N
'//                                         MN_FOLD_NOW(1):���s
'//                                         MN_FOLD_OLD(2):��
'//
'//              �^        �l               �Ӗ�
'//  �߂�l    : BOOL      TRUE             ����
'//            : BOOL      FALSE            �ُ�
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.fReadFileList���p
'///////////////////////////////////////////////////////////////////
Private Function pfuncCopyPASSINF(nCorner As Integer, nKind As Integer) As Boolean
    
    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim szSrcFile As String                 ' �R�s�[���t�@�C��
    Dim szDstFile As String                 ' �R�s�[��t�@�C��

    On Error GoTo ErrorHandler              ' �G���[�n���h���̓o�^

    ' �Ώۂ�����f�[�^�̏ꍇ�̂ݏ������s��
    ' ��L�ɊY�����Ȃ��ꍇ�͐���I��
    If FolderSyubetu <> 0 Then
        pfuncCopyPASSINF = True
        Set objFso = Nothing
        Exit Function
    End If

    ' �R�s�[���t�@�C��
    szSrcFile = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(nKind, 0) & "\" & "PASSINF"
    szDstFile = PATH_GATE_EG20 & Format(nCorner + 1, "00") & FolderName(MN_FOLD_NOW, 0) & "\" & "PASSINF"

    If objFso.FileExists(szSrcFile) = True Then
        '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
        objFso.CopyFile szSrcFile, szDstFile, True
        pfuncCopyPASSINF = True
    Else
        pfuncCopyPASSINF = False
    End If

    Set objFso = Nothing
    Exit Function

ErrorHandler:
    pfuncCopyPASSINF = False
    Set objFso = Nothing
End Function


