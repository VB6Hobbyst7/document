VERSION 5.00
Begin VB.Form frmFirmWareVer 
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
   Begin VB.CommandButton cmdCancel 
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   2520
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
      TabIndex        =   8
      Top             =   1680
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
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
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
      Height          =   7500
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8295
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
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "�}�̓���"
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
      TabIndex        =   1
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Timer tmrMail 
      Left            =   8760
      Top             =   8040
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
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1815
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
      Left            =   7320
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�q�x�s�o�[�W�����Ǘ�"
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
      TabIndex        =   11
      Top             =   0
      Width           =   12015
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
      Left            =   5040
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "��۸��і�"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   1575
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
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmFirmWareVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmFirmWareVer.frm
'//  �p�b�P�[�W���F�q�x�s�o�[�W�����Ǘ����
'//
'//  �T�v�F�q�x�s�o�[�W�������
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@��ʃ��C�A�E�g�ύX�ɂ��C��
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 �Ď��t�@�[��������f�B���N�g���ʒu�ύX
'//                 �C���X�g�[���}�̃f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �@�t�H���_�I����ʂ�OS�d�l�ɕύX
'//                 �A�o�[�W�����\���̍X�V�����ǉ�
'//  ���l�F�t�F�[�Y�P�A�Q���́u�Ď��t�@�[���o�[�W�����Ǘ��v
'//        �t�F�[�Y�R�ɂāu�q�x�s�o�[�W�����Ǘ��v�ɉ�ʖ��̕ύX�̂���
'//        �e���̃R�����g�ɂ��Ắu�Ď��t�@�[���v�̂܂܂Ƃ���B
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir�֐���FileSystemObject�ɕύX
'///////////////////////////////////////////////////////////////////
Option Explicit
'V1.6.0.1 DEL START
'Private Const KANSI_FIRM = 0            '�Ď��t�@�[��CPU
'Private Const RAS_MICO = 1              'RAS�}�C�R��
'Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
'Private Chk_OptButtom As Integer        '�I�����W�I�t�l
'V1.6.0.1 DEL END
'V1.6.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
'V1.6.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �q�x�s�o�[�W�����Ǘ����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    On Error Resume Next
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �q�x�s�o�[�W�����Ǘ����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    On Error Resume Next
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �q�x�s�o�[�W�����Ǘ����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@��ʃ��C�A�E�g�ύX�ɂ��C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
'    '�u�Ď�̧���ް�ޮ݊Ǘ���ʁF�\���v���O�o��'V1.6.0.1 DEL
    '�uRYT�ް�ޮ݊Ǘ���ʁF�\���v���O�o��      'V1.6.0.1 ADD
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_START, 0)

   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    'V1.6.0.1 DEL START
    'optSyubetu(0).Value = True
    
    'Chk_OptButtom = KANSI_FIRM
    'V1.6.0.1 DEL END
    
    '�o�[�W�������\������
    Call psVersionDisp
    
 End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdVer_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �u�\���X�V�v�u�e�L�X�g�\���v�u�}�̏o�́v�u�}�̓��́v
'//              �t�����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   Index    [IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-29   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�e�L�X�g�\�����Ƀt�@�C���L���`�F�b�N���s���B
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdVer_Click(Index As Integer)
    Dim lRetVal As Long             '�߂�l
    Dim sCommand As String          '�R�}���h������
    Dim lngErrCode As Long
    Dim bRet As Boolean
    Dim sFile As String             '�t�@�C����

    On Error Resume Next
 
 Select Case Index
    Case 0
         '�u�\���X�V�t�F�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
         bRet = UpData_Info
         If bRet = True Then
            '�u�\���X�V����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_UPDATA_OK, 0)
         Else
            '�u�\���X�V�ُ�v���O�o��
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_UPDATA_ERROR, lngErrCode)
         End If
      
    Case 1
         '�u�e�L�X�g�\���t�F�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_TEXT_BUTTOM, 0)
         'V1.6.0.1 ADD START
         sFile = Dir(MN_VERSI_FILE, vbNormal)
         If sFile = "" Then
            '�t�@�C���������O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_TEXT_ERROR, 0)
            Exit Sub
         End If
         'V1.6.0.1 ADD END
         
         '���������s�R�}���h���쐬
         sCommand = MN_EXE_MEMO & MN_VERSI_FILE
         '���������N������
         lRetVal = Shell(sCommand, vbMaximizedFocus)
         '���������A�N�e�B�u�i�O�ʕ\���j�ɂ���
         AppActivate lRetVal, True
         SendKeys "{LEFT}", True
    
    Case 2
         '�u�}�̏o�͖t�F�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_OUTPUT_BUTTOM, 0)
         bRet = Text_OutPut
         If bRet = True Then
            '�u�}�̏o�͐���v���O�o��
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_UPDATA_OK, 0) 'V1.6.0.1 DEL
            Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_OUTPUT_OK, 0) 'V1.6.0.1 ADD
         Else
            '�u�}�̏o�ُ͈�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_OUTPUT_ERROR, 0)
         End If
    
    Case 3
        '�u�}�̓��͖t�F�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_INPUT_BUTTOM, 0)
        bRet = File_InPut
        If bRet = True Then
           '�u�}�̏o�͐���v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KANSI_FIRMWARE_VER_INPUT_OK, 0)
        Else
           '�u�}�̏o�ُ͈�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_FIRMWARE_VER_INPUT_ERROR, 0)
        End If
 End Select
End Sub
'V1.6.0.1 DEL START
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : optSyubetu_Click
''//  �@�\����  : ���W�I�t�I������
''//  �@�\�T�v  : �Ď��t�@�[���ARAS�}�C�R���I�������X�V�ێ�����B
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Sub optSyubetu_Click(Index As Integer)
'   On Error Resume Next
'   Chk_OptButtom = Index
'End Sub
'V1.6.0.1 DEL END

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
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
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
'//  �֐�����  : cmdCancel_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@��ʃ��C�A�E�g�ύX�ɂ��C��
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �o�[�W�����\���X�V�����ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    On Error Resume Next
    
'    '�u�Ď��t�@�[���o�[�W�����Ǘ���ʁF�����v���O�o��  'V1.6.0.1 DEL
     '�uRYT�ް�ޮ݊Ǘ���ʁF�����v���O�o��              'V1.6.0.1 ADD
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_FIRMWARE_VER_GAMEN_END, 0)
    
    'V1.20.0.1 ADD START
    '�o�[�W�����Ǘ���ʂ̃o�[�W�����\���X�V�������s���B
    frmVersion.psGetVersion
    'V1.20.0.1 ADD END
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psVersionDisp
'//  �@�\����  : �o�[�W�������\������
'//  �@�\�T�v  : �o�[�W�������\�����̕\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@��ʃ��C�A�E�g�ύX�ɂ��C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function psVersionDisp() As Boolean

    Dim strFilePath     As String   '�o�[�W�����t�@�C���p�X
    Dim bRet            As Boolean  '�߂�l
    Dim intFileNo       As Integer  '�t�@�C���ԍ�
    Dim strWork         As String   '��ƃG���A
    Dim strVerData      As String   '�S�̃o�[�W����
    Dim intCnt          As Integer  '�J�E���^�[
    Dim lngErrCode      As Long     '�G���[�R�[�h
    Dim strFolderPath   As String   '�t�H���_�p�X
   
'*******************************
'VB�G���[����
On Error GoTo Error_psVersionDisp
'*******************************

    '���X�g������
    lstKan.Clear

    '��ƃG���A������
    strWork = ""

    '�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�o�[�W�����t�@�C���p�X�쐬
    strFilePath = MN_VERSI_FILE
    
    'V1.6.0.1�@DEL�@START
    ''RAS�@or�@�Ď��t�@�[��
    'If Chk_OptButtom = RAS_MICO Then
    '   '�\����RAS�}�C�R��
    '   strFolderPath = PATH_KANSI_FIRMWARE_RAS & "*.*"
    'Else
    '   '�\�����Ď��t�@�[��
    '   strFolderPath = PATH_KANSI_FIRMWARE & "*.*"
    'End If
    'V1.6.0.1�@DEL�@END
    strFolderPath = PATH_KANSI_FIRMWARE & "*.*" 'V1.6.0.1�@ADD
    
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ ����DA:�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�o�[�W�����t�@�C���쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllKANSIFRMVER(strFolderPath, lngErrCode, strFilePath)

    '�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�o�[�W�����t�@�C������
    If lngErrCode = 1 Then
       '�u�Ď��t�@�[���o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    '�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�o�[�W�����t�@�C�����s
    Else
       '�u�Ď��t�@�[���o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
       psVersionDisp = False
       Exit Function
    End If

    '�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�o�[�W�����t�@�C���̗L���m�F
    If Len(Trim(Dir(strFilePath))) = 0 Then
       psVersionDisp = False
       Exit Function
    End If

    '�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�o�[�W�����t�@�C���̃t�@�C���ԍ����擾����B
    intFileNo = FreeFile

    '�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�o�[�W�����t�@�C���I�[�v��
    Open strFilePath For Input As #intFileNo

    strWork = ""

    '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
    Do While Not EOF(1)
       
       Line Input #intFileNo, strWork

       '���s�R�[�h�݂͓̂ǂ݂Ƃ΂�
       If Trim(strWork) <> "" Then
          '���X�g�ɏo��
          lstKan.AddItem (strWork)
       End If
     Loop

    '�t�@�C���N���[�Y
    Close #intFileNo
    
    psVersionDisp = True

    Exit Function

'*******************************
'VB�G���[����
Error_psVersionDisp:
   '�u�Ď��t�@�[���o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
   '�t�@�C���N���[�Y
   Close #intFileNo
   psVersionDisp = False
'*******************************
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : UpData_Info
'//  �@�\����  : �u�\���X�V�v�t��������
'//  �@�\�T�v  : �o�[�W�������\�����̍ĕ`����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function UpData_Info() As Boolean
     
     On Error Resume Next

     Dim bUpData As Boolean
     
     '�o�[�W�����\���������s���B
     bUpData = psVersionDisp
     
     UpData_Info = bUpData
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Text_OutPut
'//  �@�\����  : �u�}�̏o�́v�t��������
'//  �@�\�T�v  : �o�[�W�����e�L�X�g�t�@�C����}�̂ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-23   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 �Ď��t�@�[��������f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function Text_OutPut() As Boolean

'*******************************
'VB�G���[����
On Error GoTo Error_cmdOutPut_Click
'*******************************

    Dim iRet        As Integer      '�߂�l
    Dim strVerFile  As String       'ID���p���j�b�g�t�@�C���p�X
    Dim strCopySaki As String       '�o�̓t�@�C���p�X
    Dim strWriteDir As String       '�o�͐�t�H���_
    Dim strEkimei   As String       '�ݒu�w��
    Dim strWork     As String * 256 '��ƃG���A
'   Dim fso         As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.6.0.1�@DEL
    Dim lngErrCode  As Long              '�G���[�R�[�h
    Dim sLzhDirName As String        '�w��t�H���_�@'V1.6.0.1�@ADD
    
    '�Ď��t�@�[���o�[�W�����Ǘ���ʕ\���p�t�@�C��
    strVerFile = MN_VERSI_FILE

'V1.6.0.1 DEL START
'    '�t�@�C���̗L���m�F
'    If fso.FileExists(strVerFile) = False Then
'        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
'        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
'        Text_OutPut = False
'        Set fso = Nothing
'        Exit Function
'    End If
'V1.6.0.1 DEL END
    
    'V1.6.0.1 ADD START
    '�t�H���_�I����ʂ�\�������A�t�@�C���i�[�f�B���N�g�����𓾂�B
'    sLzhDirName = pfDirSelection("a:", "�Ď��t�@�[�������ݐ�f�B���N�g���I��")     'V1.12.0.1 DEL
    'sLzhDirName = pfDirSelection("H:", "�Ď��t�@�[�������ݐ�f�B���N�g���I��")      'V1.12.0.1 ADD  'V1.20.0.1 DEL
    sLzhDirName = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
    If sLzhDirName = "" Then
       '�}�̃t�H���_�w��Ȃ���
       Text_OutPut = True
       Exit Function
    End If
    'V1.6.0.1 ADD END

    'V1.6.0.1 DEL START
'    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
'    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")
'
'    '�w��t�H���_�Ȃ�
'    If Len(strWriteDir) = 0 Then
'       Text_OutPut = False
'       Set fso = Nothing
'       Exit Function
'    End If
'
'    '�R�s�[��t�H���_�̗L���m�F
'    If fso.FolderExists(strWriteDir) = False Then
'        '�R�s�[��t�H���_�쐬
'        fso.CreateFolder (strWriteDir)
'    End If
'
'    '�R�s�[��t�@�C�����쐬
'    strCopySaki = strWriteDir & "\" & VER_TXT_NAME
'
'   '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
'    fso.CopyFile strVerFile, strCopySaki, True
   'V1.6.0.1 DEL END
   'V1.6.0.1 ADD START
   strCopySaki = sLzhDirName & "\" & VER_TXT_NAME
   
   FileCopy strVerFile, strCopySaki
   'V1.6.0.1 ADD END
 
    MsgBox "�}�̏o�͂͐���I�����܂����B", vbInformation + vbOKOnly, "�}�̏o�͌���"
    
    Text_OutPut = True
'   Set fso = Nothing       'V1.6.0.1 DEL

    Exit Function
'*******************************
'VB�G���[����
Error_cmdOutPut_Click:
     MsgBox "�}�̏o�ُ͈͂�I�����܂����B", vbCritical, "�}�̏o�͌���"
'    Set fso = Nothing      'V1.6.0.1 DEL

     Text_OutPut = False
'*******************************
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : File_InPut
'//  �@�\����  : �u�}�̓��́v�t��������
'//  �@�\�T�v  : �t�@�C����}�̓��͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@��ʃ��C�A�E�g�ύX�ɂ��C��
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 �C���X�g�[���}�̃f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                 �o�[�W�����\���X�V�����ǉ�
'//                 FileSystemObject�̎g�p���~�߁AFileCopy�ɕύX
'//                 �ǂݎ���p������ύX���鏈����ǉ�
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir�֐���FileSystemObject�ɕύX
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function File_InPut() As Boolean
  Dim lngErrCode As Long
  Dim bRet As Boolean
  Dim sLzhDirName As String
  Dim strFileName As String
'  Dim fso         As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g        'V1.20.0.1 DEL
  Dim FolderName As String
'V2.6.0.1 ADD START
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                    '�t�@�C���I�u�W�F�N�g
    Dim MyName As String
    Dim sSrcFileName As String
    Dim sDstFileName As String
'V2.6.0.1 ADD END
  
  On Error Resume Next
  
  '�t�H���_�I����ʂ�\�������A�t�@�C���i�[�f�B���N�g�����𓾂�B
'  sLzhDirName = pfDirSelection("a:", "�C���X�g�[���}�̂̃f�B���N�g���I��")     'V1.12.0.1 DEL
  'sLzhDirName = pfDirSelection("H:", "�C���X�g�[���}�̂̃f�B���N�g���I��")      'V1.12.0.1 ADD     'V1.20.0.1 DEL
  sLzhDirName = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)      'V1.20.0.1 ADD
  If sLzhDirName = "" Then
     '�}�̃t�H���_�w��Ȃ���
     File_InPut = True
'     Set fso = Nothing     'V1.20.0.1 DEL
     Exit Function
  End If

  '�t�H���_�I��ʓW�J�������s���B
  'If Chk_OptButtom = KANSI_FIRM Then 'V1.6.0.1 DEL
     
     '�ꎞ�t�H���_���쐬���A�ꎞ�t�H���_�ɃR�s�[(���J�o���΍�)
     FolderName = Mid(PATH_KANSI_FIRMWARE_WORK, 1, Len(PATH_KANSI_FIRMWARE_WORK) - 2)
     MkDir FolderName
     On Error GoTo Recovary_Error
     strFileName = Dir(PATH_KANSI_FIRMWARE & "*.*", vbNormal)
     Do While strFileName <> ""
'        fso.CopyFile PATH_KANSI_FIRMWARE & strFileName, PATH_KANSI_FIRMWARE_WORK & strFileName        'V1.20.0.1 DEL
        FileCopy PATH_KANSI_FIRMWARE & strFileName, PATH_KANSI_FIRMWARE_WORK & strFileName        'V1.20.0.1 ADD
        strFileName = Dir
     Loop
     
     'V1.6.0.1 ADD START
     strFileName = Dir(PATH_KANSI_FIRMWARE & "*.*", vbNormal)
     If strFileName <> "" Then
     'V1.6.0.1 ADD END
          Kill PATH_KANSI_FIRMWARE & "*.*"
     End If 'V1.6.0.1 ADD
     
     '�}�̂��A�Ď��t�@�[��CPU�t�H���_�ɃR�s�[
     On Error GoTo In_Put_Error
'V2.6.0.1 DEL START
'     strFileName = Dir(sLzhDirName & "*.*", vbNormal)
'     Do While strFileName <> ""
''        fso.CopyFile sLzhDirName & strFileName, PATH_KANSI_FIRMWARE & strFileName        'V1.20.0.1 DEL
'        FileCopy sLzhDirName & strFileName, PATH_KANSI_FIRMWARE & strFileName        'V1.20.0.1 ADD
'        strFileName = Dir
'     Loop
'V2.6.0.1 DEL END
    'V2.6.0.1 ADD START
    For Each objFi In objFso.GetFolder(sLzhDirName).files   '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then  '�t�@�C�����̎擾�`�F�b�N
           '�f�B���N�g�������擾
           MyName = objFi.Name
           '�}�̓��t�@�C�������쐬
           sSrcFileName = sLzhDirName & MyName
           ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
           If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
               '���[�N�t�H���_���t�@�C�������쐬����
               sDstFileName = PATH_KANSI_FIRMWARE & MyName
               '�}�̓��̃t�@�C�������[�N�t�H���_�ɃR�s�[����
               FileCopy sSrcFileName, sDstFileName
           End If
        End If
    Next
    Set objFso = Nothing
    Set objFi = Nothing
    'V2.6.0.1 ADD END
'V1.6.0.1 DEL START
'  Else
'
'     '�ꎞ�t�H���_���쐬���A�ꎞ�t�H���_�ɃR�s�[(���J�o���΍�)
'     FolderName = Mid(PATH_KANSI_FIRMWARE_RAS_WORK, 1, Len(PATH_KANSI_FIRMWARE_RAS_WORK) - 2)
'     MkDir FolderName
'     strFileName = Dir(PATH_KANSI_FIRMWARE_RAS & "*.*", vbNormal)
'     Do While strFileName <> ""
'        On Error GoTo Recovary_Error
'        fso.CopyFile PATH_KANSI_FIRMWARE_RAS & strFileName, PATH_KANSI_FIRMWARE_RAS_WORK & strFileName
'        strFileName = Dir
'     Loop
'
'     Kill PATH_KANSI_FIRMWARE_RAS & "*.*"
'
'     '�}�̂��A�Ď��t�@�[��CPU�t�H���_�ɃR�s�[
'     strFileName = Dir(sLzhDirName & "*.*", vbNormal)
'     Do While strFileName <> ""
'        On Error GoTo In_Put_Error
'        fso.CopyFile sLzhDirName & strFileName, PATH_KANSI_FIRMWARE_RAS & strFileName
'        strFileName = Dir
'     Loop
'  End If
'V1.6.0.1 DEL END

  '�ꎞ�t�H���_���폜
    'V1.20.0.1 DEL START
'  fso.DeleteFolder FolderName, False
'  Set fso = Nothing
   'V1.20.0.1 DEL END
   
   'V1.20.0.1 ADD START
   psDeleteFolder FolderName
   
   '�ǂݎ���p�̏ꍇ�ɑ����ύX���s��
   Folder_SetAttr (PATH_KANSI_FIRMWARE)
  
  '�o�[�W�������\������
  Call psVersionDisp
  'V1.20.0.1 ADD END
    
  '�}�̓��͐���|�b�v�A�b�v��ʕ\��
  MsgBox "�}�̓��͂͐���I�����܂����B", vbInformation + vbOKOnly, "�}�̓��͌���"
  File_InPut = True
  Exit Function

 
Recovary_Error:
  '�}�̓��ُ͈�|�b�v�A�b�v��ʕ\��
  MsgBox "�}�̓��ُ͈͂�I�����܂����B", vbCritical, "�}�̓��͌���"
  File_InPut = False
  
  '�ꎞ�t�H���_���폜
    'V1.20.0.1 DEL START
'  fso.DeleteFolder FolderName, False
'  Set fso = Nothing
   'V1.20.0.1 DEL END
   psDeleteFolder FolderName        'V1.20.0.1 ADD
  Exit Function

In_Put_Error:
  
  '���J�o���������s���B
  'If Chk_OptButtom = KANSI_FIRM Then 'V1.6.0.1 DEL
  'V2.6.0.1 ADD START
   Set objFso = Nothing
   Set objFi = Nothing
  'V2.6.0.1 ADD END

     Kill PATH_KANSI_FIRMWARE & "*.*"
 
     '�ꎞ�t�H���_���A�Ď��t�@�[��CPU�փR�s�[
     strFileName = Dir(PATH_KANSI_FIRMWARE_WORK & "*.*", vbNormal)
     Do While strFileName <> ""
        FileCopy PATH_KANSI_FIRMWARE_WORK & strFileName, PATH_KANSI_FIRMWARE & strFileName
        strFileName = Dir
     Loop
'V1.6.0.1 DEL START
'  Else
'     Kill PATH_KANSI_FIRMWARE_RAS & "*.*"
'
'     '�ꎞ�t�H���_���ARAS�}�C�R���փR�s�[
'     strFileName = Dir(PATH_KANSI_FIRMWARE_RAS_WORK & "*.*", vbNormal)
'     Do While strFileName <> ""
'        FileCopy PATH_KANSI_FIRMWARE_RAS_WORK & strFileName, PATH_KANSI_FIRMWARE_RAS & strFileName
'        strFileName = Dir
'     Loop
'  End If
'V1.6.0.1 DEL END
  
'�}�̓��ُ͈�|�b�v�A�b�v��ʕ\��
  MsgBox "�}�̓��ُ͈͂�I�����܂����B", vbCritical, "�}�̓��͌���"
  File_InPut = False

 '�ꎞ�t�H���_���폜
    'V1.20.0.1 DEL START
'  fso.DeleteFolder FolderName, False
'  Set fso = Nothing
   'V1.20.0.1 DEL END
   psDeleteFolder FolderName        'V1.20.0.1 ADD

End Function

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
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    On Error Resume Next
 
    '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmFirmWareVer.Caption, False
        pfFormActive (frmFirmWareVer.hwnd)
    End If
End Sub

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Folder_SetAttr
'//  �@�\����  : �t�@�C�������ύX
'//  �@�\�T�v  : �t�H���_���̓ǂݎ��t�@�C��������ʏ�ɐݒ肷��
'//
'//              �^      ����         �Ӗ�
'//  ����      : String  sFolderName  �t�H���_�p�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-11  CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Folder_SetAttr(sFolderName As String)
    On Error Resume Next
    
    Dim iAttr As Integer
    Dim lIndex As Long
    Dim foName As Folder
    Dim fName As File
    Dim fsoObj As FileSystemObject
    
    Set fsoObj = New FileSystemObject
    Set foName = fsoObj.GetFolder(sFolderName)
    lIndex = 0
    
    For Each fName In foName.files
        '�������擾
        iAttr = GetAttr(fName.Path)
        '�ʏ�t�@�C���A�܂��̓A�[�J�C�u�t�@�C���ɓǂݎ�葮�����t���Ă�����
        If iAttr = vbReadOnly Or iAttr = vbArchive + vbReadOnly Then
            '�ǂݎ�葮������菜���ăZ�b�g
            Call SetAttr(fName.Path, iAttr - vbReadOnly)
        End If
    lIndex = lIndex + 1
    Next fName
    
    Set fsoObj = Nothing
    Set fName = Nothing
    Set foName = Nothing
    
End Sub
'V1.20.0.1 ADD END
