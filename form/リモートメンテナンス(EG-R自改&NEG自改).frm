VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRMente 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "���O�g���[�X�iEG-R�������D�@�j"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
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
      Left            =   8520
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   735
      Left            =   9500
      TabIndex        =   11
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "�t�@�C���폜"
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
      Index           =   5
      Left            =   9500
      TabIndex        =   8
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "���k�}�̏o��"
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
      Left            =   9500
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.ListBox lstTraceFile 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   240
      MultiSelect     =   2  '�g��
      TabIndex        =   5
      Top             =   1080
      Width           =   9135
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "���k���ʊm�F"
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
      Index           =   4
      Left            =   9500
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "   �t�@�C��     �}�̏o��"
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
      Left            =   9500
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
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
      Index           =   1
      Left            =   9500
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "�f�[�^���W "
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
      Left            =   9500
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�����[�g�����e�i���X��ʂ֖߂�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9500
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   9120
      Top             =   8040
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�������D�@�����[�g�����e�i���X"
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
      TabIndex        =   12
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblListItem 
      BorderStyle     =   1  '����
      Caption         =   "    �g���[�X�t�@�C����"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   7455
   End
   Begin VB.Label lblListItem 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�o�C�g�T�C�Y"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Caption         =   "�������D�@  �g���[�X�t�@�C��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   450
      Width           =   4335
   End
End
Attribute VB_Name = "frmRMente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmRMente.frm
'//  �p�b�P�[�W���F�������D�@�����[�g�����e�i���X���
'//
'//  �T�v�F�������D�@�����[�g�����e�i���X���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-07-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 �g���[�X�t�@�C�������ݐ�f�B���N�g���ʒu�ύX
'//                 �g���[�X�t�@�C�����k�����ݐ�f�B���N�g���ʒu�ύX
'//                 ���k�t�@�C���I���f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.272�C���Ή��z
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 �y���k�t�H���_�w��Ή��z
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

'���X�g�{�b�N�X�Ɋւ���l
Private Const LIST_FILE_SIZE_LENGTH = 11   '�޲Ļ��ޗ��̕�����
Private Const LIST_FILE_ELIMITTER = " -- " '�޲Ļ��ނ��ڰ�̧�يԂ̋�ؕ�����
Private Const LIST_HEDDER_LENGTH = LIST_FILE_SIZE_LENGTH + 4 '��L�A�Q�̕��������v
Private sTOOLPass As String
'Private sHyoujiGoukiNo(0 To 18) As String         '�\�����@�ԍ��i�[�G���A          ' EG20 V3.6.0.1�y����TR-No.272�C���Ή��z�폜
Private sHyoujiGoukiNo(0 To 31) As String         '�\�����@�ԍ��i�[�G���A           ' EG20 V3.6.0.1�y����TR-No.272�C���Ή��z�ǉ�
Private Const TITLENAME_CORNER = "�R�[�i#"        ' �R�[�i��                        ' EG20 V6.6.0.1�ǉ�
Private sRonriCornerNo(0 To 31) As String         '�_���R�[�i�ԍ��i�[�G���A         ' EG20 V6.6.0.1�ǉ�

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : CmdRemove_Click
'//  �@�\����  : �u�}�̎�O�v�t����������
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
'//  �֐�����  : Form_Activate
'//  �@�\����  : �������D�@�����[�g�����e�i���X(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}�A�N��
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
'//  �@�\����  : �������D�@�����[�g�����e�i���X(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}�A��~
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

    '�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �������D�@�����[�g�����e�i���X(���[�h��)
'//  �@�\�T�v  : �����������s���B
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
Private Sub Form_Load()
    Dim iRet As Integer
    
On Error Resume Next
    '�u�������D�@�Ӱ�����ݽ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_START, 0)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

   'GLT�t�@�C�����쐬���A���e���X�V����B
    iRet = fMakeGLTFile
    
    If iRet = 0 Then
        '���X�g�{�b�N�X�Ƀg���[�X�t�@�C������\������B
        fListDisplay
    End If
    
    '���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdTraceFile_Click
'//  �@�\����  : �e�t����������
'//  �@�\�T�v  : �e�t���̂̏������s���B
'//              �u�f�[�^���W�v�u�\���X�V�v�u�t�@�C���}�̏o�́v
'//              �u���k�}�̏o�́v�u���k���ʊm�F�v�u�t�@�C���폜�v
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@ [IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-07-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdTraceFile_Click(Index As Integer)
    Dim lRetVal As Double      'Shell�֐��߂�l
    Dim iResponse As Integer   'MsgBox�߂�l
    Dim sWriteDir As String    '�g���[�X�t�@�C�������ݐ�̃f�B���N�g��
    Dim lngErrCode As Long   '�G���[�R�[�h
   
   On Error Resume Next

    Select Case Index
    Case 0
       '�u�������D�@�Ӱ�����ݽ��ʁF�f�[�^���W�t�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_DATA_SHUSHU_BUTTOM, 0)
        '���w.GLT�t�@�C���֎������������ށB
        fMakeGLTFile
        '����SW�ێ�f�[�^�쐬�������s���B
        'If sSWFileCopy > 0 Then   'V1.6.0.1 DEL
        sSWFileCopy  'V1.6.0.1 ADD
          '�����[�g�����e�c�[�����N������B
          psGATERMenteTool
          '�������D�@�c�[���N��
          lRetVal = Shell(sTOOLPass, vbNormalFocus)
          If 0 = lRetVal Then
             GoTo ERROR_MSG_RMENTE
          End If
          '�����[�g�����e�c�[�����A�N�e�B�u�i�O�ʕ\���j�ɂ���
        '  AppActivate lRetVal, True 'V1.7.0.1 DEL
        'V1.6.0.1 DEL START
        'Else
        '  '�u�Ӱ�����ݽ��ʁF�����ێ�SW�f�[�^�t�@�C���R�s�[�ُ�v���O�o��
        '  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        '  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
        'End If
        'V1.6.0.1 DEL END
    Case 1
      '�u�������D�@�Ӱ�����ݽ��ʁF�\���X�V�t�����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
       '���X�g�{�b�N�X�Ƀg���[�X�t�@�C������\������B
       fListDisplay
    Case 2
      '�u�������D�@�Ӱ�����ݽ��ʁF�t�@�C���}�̏o�͖t�����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_FILE_OUTPUT_BUTTOM, 0)
      sCopyTraceFile
    Case 3
      '�u�������D�@�Ӱ�����ݽ��ʁF���k�}�̏o�͖t�����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_LZH_OUTPUT_BUTTOM, 0)
      sLzhFileWrite
    Case 4
      '�u�������D�@�Ӱ�����ݽ��ʁF���k���ʊm�F�t�����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_LZH_KAKUNIN_BUTTOM, 0)
      '���k�t�@�C���̓��e��\������B
      sLzhFileDisplay
    Case 5
      '�u�������D�@�Ӱ�����ݽ��ʁF�t�@�C���폜�t�����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_FILE_DELETE_BUTTOM, 0)
       '�I�𒆃t�@�C�����폜����B
        If fSelectedFilesDelete = True Then
            '�폜�t�@�C�����������Ȃ�A���X�g�{�b�N�X��\���X�V����B
            fListDisplay
        End If
    Case Else
 End Select

 Exit Sub

ERROR_MSG_RMENTE:
'===�g���[�X�f�[�^���W�G���[�̏ꍇ�A
    '�u�������D�@�Ӱ�����ݽ��ʁF�����[�g�����e�c�[���N���ُ�v���O�o��
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, RMENTE_GAMEN_KIDOU_ERROR, 0)
    '�u�����[�g�����e�c�[���N���ُ�v�|�b�v�A�b�v��\������B
    iResponse = MsgBox("�����[�g�����e�c�[���iR_Mente.exe�j���N���ł��܂���B", _
                vbYes, _
               "�����[�g�����e�c�[�����s�G���[")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
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
Private Sub cmdReturn_Click()
On Error Resume Next
    '�u�������D�@�Ӱ�����ݽ��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_END, 0)
    '����ʂ������B
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeGLTFile
'//  �@�\����  : ���w.GLT�t�@�C���ւ̎��������������ݏ���
'//  �@�\�T�v  : GATE.INI���Q�Ƃ��A���w.GLT�t�@�C���ցA
'//              ���@�ԍ��A�\�������AIP�A�h���X���������ށB
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.6.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.7.0.1)  2012-06-28  CODED BY  [TCC] H.Sugimoto
'//                 �y���ڃ`�F�b�N�̑Ώۂ����D�@���݂̂Ƃ���C���z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeGLTFile() As Integer
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
    Dim szCorner As String      ' �R�[�i�ԍ�
    Dim szTitleName As String                       ' �^�C�g����                    ' EG20 V6.7.0.1�ǉ�
    Dim fso As New FileSystemObject                 '�t�@�C���V�X�e���I�u�W�F�N�g   ' EG20 V6.7.0.1�ǉ�

    On Error Resume Next
    MkDir PATH_RMENTE_GATE_DEN   '�����p�d�S�t�H���_���쐬����B�iGLT�t�@�C���p�j
    
' EG20 V6.7.0.1�ǉ��J�n
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH_RMENTE_GATE_DEN_JIEKI) = False Then
        '�R�s�[��t�H���_�쐬
        fso.CreateFolder (PATH_RMENTE_GATE_DEN_JIEKI)
    End If
    Set fso = Nothing
' EG20 V6.7.0.1�ǉ��I��
    
    
    'GLT�t�@�C�����J���B�t�@�C�������݂��Ȃ���ΐV�K�ɍ쐬�����B
    On Error GoTo ErrorHandlerGLTFile
    intGLTFileNo = FreeFile        ' ���g�p�̃t�@�C���ԍ����擾����B
    Open GATE_GLT_FILE_FULLPASS For Output As #intGLTFileNo     ' GLT�t�@�C�����J���B

    For iGate = CNT_MIN To MAX_GATE_NO - 1
      '�������D�@���擾
      sKeyName = "gate" & Format(iGate + 1, "00")
      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                     sKeyName, _
                                     DEFAILT, sGateData, Len(sGateData), _
                                     PATH_GATE_FILE)
      If iRet = 0 Then
         '�u�Ӱ�����ݽ��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
         Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
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
      
      If Len(Trim(sFData(1))) = 1 Then
         '���@�ԍ����P���Ȃ�΁A�擪�ɂO��t������B
'         sGoukiNo = "0" & Trim(sFData(1)) & "���@"                                 ' EG20 V6.7.0.1�폜
         sGoukiNo = "0" & Trim(sFData(1))                                           ' EG20 V6.7.0.1�ǉ�
      Else
'         sGoukiNo = Trim(sFData(1)) & "���@"                                       ' EG20 V6.7.0.1�폜
         sGoukiNo = Trim(sFData(1))                                                 ' EG20 V6.7.0.1�ǉ�
      End If
        
' EG20 V6.6.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
'        szCorner = Replace(TITLENAME_CORNER, "#", Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))) ' EG20 V6.7.0.1�폜
        szCorner = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))                                  ' EG20 V6.7.0.1�ǉ�
        sRonriCornerNo(iGate) = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.6.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��I��
' EG20 V6.7.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
        ' �R�[�i�ԍ��ϊ�
        szTitleName = Replace(RMENTE_GOKITITLENAME, "$", szCorner)
        ' ���@�ԍ��ϊ�
        szTitleName = Replace(szTitleName, "##", sGoukiNo)
' EG20 V6.7.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
      
      If Trim(sFData(4)) <> "��" Then
         'Gate.ini�t�@�C���̍��@�ԍ��\�������AIP�A�h���X��GLT�t�@�C���ɏ������ށB
'         Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(5))                     ' EG20 V6.6.0.1�폜
'         Print #intGLTFileNo, szCorner & "_" & sGoukiNo & "," & Trim(sFData(5))    ' EG20 V6.6.0.1�ǉ� ' EG20 V6.7.0.1�폜
         Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))                   ' EG20 V6.7.0.1�ǉ�
      End If
      
      '�\�����@�ԍ�
      If Len(Trim(sFData(1))) = 1 Then
         '���@�ԍ����P���Ȃ�΁A�擪�ɂO��t������B
         sHyoujiGoukiNo(iGate) = "0" & Trim(sFData(1))
      Else
         sHyoujiGoukiNo(iGate) = Trim(sFData(1))
      End If
    
    Next
    
    'GLT�t�@�C�������B
    Close #intGLTFileNo
    
    fMakeGLTFile = 0    '����I��
    Exit Function

ErrorHandlerGateIni:
   '�u�������D�@�Ӱ�����ݽ��ʁF�t�@�C���A�N�Z�X�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 1
   'GLT�t�@�C�������B
   Close #intGLTFileNo
   Exit Function
ErrorHandlerGLTFile:
   '�u�������D�@�Ӱ�����ݽ��ʁF�t�@�C���A�N�Z�X�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 2
   'GLT�t�@�C�������B
   Close #intGLTFileNo

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSWFileCopy
'//  �@�\����  : �����ێ�SW�ݒ�f�[�^�t�@�C���쐬����
'//  �@�\�T�v  : �����ێ�SW�ݒ�f�[�^���A�����ێ�SW�f�[�^�t�@�C����
'//              �R�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.6.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sSWFileCopy() As Integer

     Dim iCnt As Integer                     '�J�E���^�[
     Dim sSWDataPath As String               '�����ێ�SW�f�[�^�t�@�C��
     Dim sMyPath As String                   '�����ێ�SW�ݒ�f�[�^
     
     On Error Resume Next
   
     sSWFileCopy = 0                         '�t�@�C�����ݐ�
    
    '�����ő吔�����[�v����B
    For iCnt = 1 To MAX_GATE_NO
     '�uGATE_SW##.dat�v�́u##�v��01�`16�ɕϊ�����B
     sMyPath = Replace(GATE_SW_FILE, "##", Format(iCnt, "0#"))
     '�����ێ�SW�ݒ�f�[�^�̌������s���B
     If Dir(sMyPath) <> "" Then
        '�����ێ�SW�f�[�^�t�@�C���̃p�X���쐬����B
        sSWDataPath = PATH_RMENTE_GATE_DEN_JIEKI_GOUKI
' EG20 V6.6.0.1�ǉ��J�n
        '�u�R�[�i$�v�́u$�v��1�`6�ɕϊ�����B
        sSWDataPath = Replace(sSWDataPath, "$", sRonriCornerNo(iCnt - 1))
' EG20 V6.6.0.1�ǉ��I��
        '�u##���@�v�́u##�v��01�`16�ɕϊ�����B
        sSWDataPath = Replace(sSWDataPath, "##", Format(sHyoujiGoukiNo(iCnt - 1), "0#"))
        '�t�H���_�쐬
        MkDir sSWDataPath
        sSWDataPath = sSWDataPath & TOOL_SW_File
        
        '�����ێ�SW�f�[�^�������ێ�SW�f�[�^�t�@�C���ɃR�s�[����B
        FileCopy sMyPath, sSWDataPath
        sSWFileCopy = sSWFileCopy + 1
     End If
   Next
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fListDisplay
'//  �@�\����  : ���X�g�{�b�N�X�̓��e��\���X�V����B
'//  �@�\�T�v  : ���X�g�{�b�N�X�̕\�����e��������A
'//              �ŐV�̃g���[�X�t�@�C������\������B
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
Private Function fListDisplay()
    Dim sInFolder(1) As String  '�g���[�X�f�[�^�t�H���_��

    On Error Resume Next

    '���X�g�{�b�N�X����ɂ���B
    lstTraceFile.Clear
    '�g���[�X�f�[�^�t�H���_�ȉ��̃t�@�C�������X�g�{�b�N�X�ɕ\������B
    sInFolder(0) = PATH_RMENTE_GATE_DEN_JIEKI  '�{�d�S�t�H���_����J�n����B
    sFileDisplay 1, sInFolder()
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFileDisplay
'//  �@�\����  : ���X�g�{�b�N�X�\������
'//  �@�\�T�v  : �w��t�H���_�����̃t�@�C���������X�g�{�b�N�X�ɕ\������B
'//              �ŐV�̃g���[�X�t�@�C������\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFolder
'//        �@�@: Integer �@iFolderNo
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sFileDisplay(iFolderNo As Integer, sFolder() As String)
    Dim iInFileNo As Integer   '�����Ώۃt�H���_�����̃t�@�C���̌�
    Dim sInFile() As String    '  ����  �t�@�C�����i�t���p�X�j
    Dim iInFolderNo As Integer '�����Ώۃt�H���_�����̃t�@�C���̌�
    Dim sInFolder() As String  '  ����  �t�H���_���i�t���p�X�F�ŏI�����́��j
    Dim i     As Integer       '���[�N�J�E���^
    Dim j     As Integer       '���[�N�J�E���^
    Dim sFileSize As String * LIST_FILE_SIZE_LENGTH  '�\���t�@�C���̃o�C�g�T�C�Y
    Dim sDisplay As String     '���X�g�{�b�N�X�֕\������P�s���̕�����

    On Error Resume Next

    '�w�肳�ꂽ�t�H���_�̑S�Ăɂ��Ď��{����B
    For i = CNT_MIN To iFolderNo - 1
        '�����Ώۃt�H���_�����̃t�@�C���E�t�H���_���擾����B
        psFolderSearch sFolder(i), iInFileNo, sInFile(), iInFolderNo, sInFolder()
        '�����Ώۃt�H���_�����̃t�@�C�������X�g�{�b�N�X�֕\������B
        For j = 0 To iInFileNo - 1
            '̧�ٻ��ނ͉E�l�߁A�R���̃J���}��؂�ŕ\������B
            RSet sFileSize = Format$(FileLen(sInFile(j)), "#,###")
            '�t�@�C�����́A�E�E\���d�S\���w\�܂ł̃t�H���_:RMENTE_DIR_TRACE�͕\�����Ȃ��B
            '            �i�擪�ɋ�؂蕶��:LIST_FILE_ELIMITTER��t����B�j
            sDisplay = sFileSize & LIST_FILE_ELIMITTER & _
                       Right(sInFile(j), Len(sInFile(j)) - Len(PATH_RMENTE_GATE_DEN_JIEKI))
            lstTraceFile.AddItem sDisplay
        Next
        '�����Ώۃt�H���_�����̃t�H���_�ȉ��̃t�@�C�������X�g�{�b�N�X�ɕ\������B
        sFileDisplay iInFolderNo, sInFolder()
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyTraceFile
'//  �@�\����  : �u�t�@�C���}�̏o�́v�t����������
'//  �@�\�T�v  : �t�@�C�����w��f�B���N�g���ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 �g���[�X�t�@�C�������ݐ�f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sCopyTraceFile()
    Dim iLine As Integer         '�ڰ�̧��ؽ��ޯ���̍s���ޯ��
    Dim iMaxLine As Integer      '�ڰ�̧��ؽ��ޯ���̍s��
    Dim iFlag As Integer         '�I�𒆃t�@�C���L���i�P�^�O�j
    Dim iResponse As Integer     'MsgBox�{�^���R�[�h
    Dim sFullPass As String      '�R�s�[���t�@�C���t���p�X��
    Dim sFileName As String      '�R�s�[���t�@�C����
    Dim sCopyDir As String       '�R�s�[��f�B���N�g��
    Dim sCopyFileName As String  '�R�s�[��t�@�C����
    Dim lSts As Long             '���[�N�i�߂�l�j
    Dim sWork As String          '���[�N
    Dim i As Integer             '���[�N
    Dim j As Integer             '���[�N
    Dim lngErrCode As Long       '�G���[�R�[�h
    Dim iFileCounter As Integer  '�Ώ�̧�ِ��J�E���^    ' EG20 V5.9.0.1�y���O�I������Ή��zADD

On Error GoTo COPY_ERROR
    iFlag = 0   '�I�𒆃t�@�C�����Ƃ��Ă����B
    '���X�g�{�b�N�X�\�����̑S�s�ɂ��Ĉȉ������{����B
    iMaxLine = lstTraceFile.ListCount  '�ڰ�̧��ؽ��ޯ���̍s���𓾂�B
    
' EG20 V5.9.0.1�y���O�I������Ή��zADD START
    iFileCounter = 0
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
            iFileCounter = iFileCounter + 1
        End If
    Next

    If iFileCounter > LOG_FILECNT_MAX Then
        ' �x�������\��
        MsgBox "�I�����ꂽ�t�@�C����������𒴂��܂����B" _
               & Chr(vbKeyReturn) & "�I���ł���t�@�C������[" & LOG_FILECNT_MAX & "]���܂łł��B", _
               vbOKOnly + vbCritical, _
               "�t�@�C���w��ُ�"
        Exit Sub
    End If
' EG20 V5.9.0.1�y���O�I������Ή��zADD END
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        '�I�����ꂽ�s�Ȃ�΁A
            If iFlag = 0 Then
                ' ��o����f�B���N�g����I������
'                sCopyDir = pfDirSelection("a:", "�g���[�X�t�@�C�������ݐ�̃f�B���N�g���I��")  'V1.12.0.1 DEL
                'sCopyDir = pfDirSelection("H:", "�g���[�X�t�@�C�������ݐ�̃f�B���N�g���I��")   'V1.12.0.1 ADD�@'V1.20.0.1 DEL
                sCopyDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER) 'V1.20.0.1 ADD
                If sCopyDir = "" Then
                '�f�B���N�g���w�肪�Ȃ���΁A �������I����B
                    Exit Sub
                End If
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
                '�v���O���X�o�[��\������
                Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
            End If
            iFlag = 1  '�I�𒆃t�@�C���L��Ƃ���B
            '�R�s�[���t�@�C�����\�����e���Z�b�g����B�isWork���޲Ļ���--01���@\CTRC2000xxx.xxx�j
            sWork = lstTraceFile.List(iLine)
            '�擪�����޲Ļ��ޕ����i"�޲Ļ���--" ����=LIST_HEDDER_LENGTH�j�����O����B
            '                                     �isFileName��01���@\CTRC2000xxx.xxx�j
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            '�R�s�[���t�@�C�����t���p�X���Z�b�g����B�isFullPass��C:\tool\R_Mente\DATA\�{�d�S\���w\01���@\CTRC2000xxx.xxx�j
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            '�����ݐ�f�B���N�g���{�t�@�C���i�R�s�[ ���Ɠ����j�����Z�b�g����B
            '                                 �isCopyFileName��a:\01���@\CTRC2000xxx.xxx�j
            sCopyFileName = sCopyDir & sFileName
            '�R�s�[��f�B���N�g���Ƀt�H���_���쐬����B
            On Error Resume Next
            i = 1
            sWork = sCopyDir
            Do
                j = InStr(i, sFileName, "\")
                If j = 0 Then Exit Do
                j = j + 1
                sWork = sWork & Mid$(sFileName, i, j - i)
                MkDir sWork
                i = j
            Loop
            '���O�g���[�X�t�@�C�����w��f�B���N�g���ɏ����o���B
            On Error GoTo COPY_ERROR
            FileCopy sFullPass, sCopyFileName
        End If
    Next
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    If iFlag = 0 Then
    '�t�@�C�����I������Ă��Ȃ���΁A�G���[���b�Z�[�W��\�����A�������I������B
        MsgBox "��o���t�@�C�����I������Ă��܂���B" _
               & Chr(vbKeyReturn) & "�I�����Ă��������B", _
               vbOKOnly + vbExclamation, _
                "�����[�g�����e�i���X�i�������D�@�j"
        Exit Sub
    End If
    
    '�u�Ӱ�����ݽ��ʁF�t�@�C���}�̏o�͏�������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RMENTE_GAMEN_FILE_OUTPUT_OK, 0)
    Exit Sub

COPY_ERROR:
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    Select Case Err.Number
        Case 61 ' �R�s�[��󂫗e�ʕs��
            iResponse = MsgBox("�󂯑��̃h���C�u�̃f�B�X�N�������ς��ł��B" _
               & Chr(vbKeyReturn) & "�V�����f�B�X�N��}�����Ă��������B", _
               vbOKOnly, _
               "���O�t�@�C����o��")
        Case 70 ' ���C�g�v���e�N�g
            lSts = CopyFile(sFullPass, sCopyFileName, 0)
            If (lSts = 0) Then
                iResponse = MsgBox("�t�@�C�����쐬�܂��͒u���ł��܂���B���̃f�B�X�N�̓��C�g�v���e�N�g����Ă܂��B" _
                   & Chr(vbKeyReturn) & "���C�g�v���e�N�g���������邩�@�ʂ̃f�B�X�N���g���Ă��������B", _
                   vbOKOnly, _
                   "���O�t�@�C����o��")
            End If
        Case 71 ' �f�B�X�N��}�����Ă�������
            iResponse = MsgBox("�h���C�u�Ƀf�B�X�N�������Ă܂���B" _
               & Chr(vbKeyReturn) & "�f�B�X�N��}�����Ă����蒼���Ă��������B", _
               vbOKOnly, _
               "���O�t�@�C����o��")
         Case Else
            iResponse = MsgBox("�\�����ʃG���[���������܂����B" _
               & Chr(vbKeyReturn) & "�������蒼���Ă��������B", _
               vbOKOnly, _
               "���O�t�@�C����o��")
    End Select
    
    '�u�Ӱ�����ݽ��ʁF�t�@�C���}�̏o�͏����ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RMENTE_GAMEN_FILE_OUTPUT_ERROR, lngErrCode)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sLzhFileWrite
'//  �@�\����  : �u���k�}�̏o�́v�t����������
'//  �@�\�T�v  : ���X�g�{�b�N�X�Ŏw�肳�ꂽ�t�@�C�������k���A
'//              �w��f�B���N�g���ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 �g���[�X�t�@�C�����k�����ݐ�f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sLzhFileWrite()
    Dim iLine As Integer         '�ڰ�̧��ؽ��ޯ���̍s���ޯ��
    Dim iMaxLine As Integer      '�ڰ�̧��ؽ��ޯ���̍s��
    Dim iFlag As Integer         '�I�𒆃t�@�C���L���i�P�^�O�j
    Dim iResponse As Integer     'MsgBox�{�^���R�[�h
    Dim sFullPass As String      '���k���t�@�C���t���p�X��
    Dim sFileName As String      '���k���t�@�C����
    Dim sLzhDirName As String    '.LZḨ�يi�[�f�B���N�g����
    Dim sLzhFileName As String   '.LZḨ�ٖ�
    Dim iSts As Integer          '�֐��߂�l
    Dim sWork As String          '���[�N
    Dim i As Integer             '���[�N
    Dim j As Integer             '���[�N
    Dim lngErrCode As Long       '�G���[�R�[�h
    Dim nIndex As Integer        ' ������                    ' EG20 V5.6.0.1�ǉ�
    Dim iFileCounter As Integer  '�Ώ�̧�ِ��J�E���^    ' EG20 V5.9.0.1�y���O�I������Ή��zADD
    
On Error GoTo WRITE_ERROR
    iFlag = 0   '�I�𒆃t�@�C�����Ƃ��Ă����B
    '���X�g�{�b�N�X�\�����̑S�s�ɂ��Ĉȉ������{����B
    iMaxLine = lstTraceFile.ListCount  '�ڰ�̧��ؽ��ޯ���̍s���𓾂�B
    
' EG20 V5.9.0.1�y���O�I������Ή��zADD START
    iFileCounter = 0
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
            iFileCounter = iFileCounter + 1
        End If
    Next

    If iFileCounter > LOG_FILECNT_MAX Then
        ' �x�������\��
        MsgBox "�I�����ꂽ�t�@�C����������𒴂��܂����B" _
               & Chr(vbKeyReturn) & "�I���ł���t�@�C������[" & LOG_FILECNT_MAX & "]���܂łł��B", _
               vbOKOnly + vbCritical, _
               "�t�@�C���w��ُ�"
        Exit Sub
    End If
' EG20 V5.9.0.1�y���O�I������Ή��zADD END
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        '�I�����ꂽ�s�Ȃ�΁A
            If iFlag = 0 Then
                ' ��o����f�B���N�g����I������
'                sLzhDirName = pfDirSelection("a:", "�g���[�X�t�@�C�����k�����ݐ�̃f�B���N�g���I��")   'V1.12.0.1 DEL
                'sLzhDirName = pfDirSelection("H:", "�g���[�X�t�@�C�����k�����ݐ�̃f�B���N�g���I��")    'V1.12.0.1 ADD 'V1.20.0.1 DEL
                sLzhDirName = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
                If sLzhDirName = "" Then
                '�f�B���N�g���w�肪�Ȃ���΁A �������I����B
                    Exit Sub
                End If
' EG20 V5.6.0.1�y���k�t�H���_�w��Ή��z�ǉ��J�n
                ' �o�̓t�H���_�ɔ��p�X�y�[�X���܂܂�Ă���ꍇ�A���k�ňُ킪�������Ă��܂�����
                ' ���k�O�Ƀ`�F�b�N���Ĉُ��\������B
                nIndex = InStr(sLzhDirName, " ")
                If nIndex <> 0 Then
                    ' �x���|�b�v�A�b�v�E�B���h�E��\������B
                    Call MsgBox(CABFOLDERSELECT_ERRORMESSAGE, vbCritical, CABFOLDERSELECT_ERRORTITLE)
                    Exit Sub  '�f�B���N�g�����w�肳��Ȃ���΁A�����I��
                End If
' EG20 V5.6.0.1�y���k�t�H���_�w��Ή��z�ǉ��I��

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
                '�v���O���X�o�[��\������
                Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
            
            End If
            iFlag = 1  '�I�𒆃t�@�C���L��Ƃ���B
            '���k���t�@�C�����\�����e���Z�b�g����B�isWork���޲Ļ���--01���@\CTRC2000xxx.xxx�j
            sWork = lstTraceFile.List(iLine)
            '�擪�����޲Ļ��ޕ����i"�޲Ļ���--" ����=LIST_HEDDER_LENGTH�j�����O����B
            '                                  �isFileName��01���@\CTRC2000xxx.xxx�j
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            '���k���t�@�C�����t���p�X���Z�b�g����B�isFullPass��C:\tool\R_Mente\DATA\�{�d�S\���w\01���@\CTRC2000xxx.xxx�j
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            '�����ݐ�f�B���N�g���{�t�@�C���i���k���Ɠ����j�����Z�b�g���A�g���q�ɁA.CAB��t������B
            '                                 �isLzhFileName��a:\01���@\CTRC2000xxx.xxx.CAB�j
            sLzhFileName = sLzhDirName & sFileName & ".CAB"
            '���k��f�B���N�g���Ƀt�H���_���쐬����B
            On Error Resume Next
            i = 1
            sWork = sLzhDirName
            Do
                j = InStr(i, sFileName, "\")
                If j = 0 Then Exit Do
                j = j + 1
                sWork = sWork & Mid$(sFileName, i, j - i)
                MkDir sWork
                i = j
            Loop
            On Error GoTo WRITE_ERROR
            '�Ώۃt�@�C�����A���k��.CAB�t�@�C���Ɋi�[����B
            Call psCabReqest(CABREQEST.CAB_COMPRESSION, sLzhFileName, sFullPass)
        End If
    Next
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    If iFlag = 0 Then
    '�t�@�C�����I������Ă��Ȃ���΁A�G���[���b�Z�[�W��\�����A�������I������B
        MsgBox "��o���t�@�C�����I������Ă��܂���B" _
               & Chr(vbKeyReturn) & "�I�����Ă��������B", _
               vbOKOnly + vbExclamation, _
               "�����[�g�����e�i���X�i�������D�@�j"
        Exit Sub
    End If
    
    '�u�Ӱ�����ݽ��ʁF���k�}�̏o�͏�������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RMENTE_GAMEN_LZH_OUTPUT_OK, 0)
  
    Exit Sub

WRITE_ERROR:
    '�u�Ӱ�����ݽ��ʁF���k�}�̏o�͏����ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RMENTE_GAMEN_LZH_OUTPUT_ERROR, lngErrCode)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sLzhFileDisplay
'//  �@�\����  : �u���k���ʊm�F�v�t����������
'//  �@�\�T�v  : �w�肳�ꂽ���k�t�@�C���̓��e���擾���A�������\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ���k�t�@�C���I���f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sLzhFileDisplay()
    Dim sLzhFileName As String   '.LZḨ�ٖ�
    Dim sLzhDataFile As String   '.LZḨ�ٓ��e�����݃t�@�C�����i���߽�j
    Dim sCommand As String
    Dim lRetVal As Long
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g  'V1.20.0.1 ADD
    
    On Error Resume Next

    '���k�t�@�C���I����ʂ�\�����A���k�t�@�C����I��������B
'    sLzhFileName = pfCabFileSelection("a:")        'V1.12.0.1 DEL
    'sLzhFileName = pfCabFileSelection("H:")         'V1.12.0.1 ADD 'V1.20.0.1 DEL
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
    Set objFso = Nothing
    '�g���q��ݒ�
    CommonDialog1.Filter = "���k�t�@�C���i*.cab�j|*.cab|"
    '�t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    '�I�������t�@�C�������擾
    sLzhFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
    If sLzhFileName = "" Then Exit Sub   '�t�@�C�����I������Ȃ���΁A�߂�B
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�I�����ꂽ���k�t�@�C���̓��e���擾����B
    Call psCabReqest(CABREQEST.CAB_DRAFT, sLzhFileName, vbNullString)
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�t�@�C�����e�擾�l���
    sLzhDataFile = gstrCabErrCd
    If sLzhDataFile = "" Then Exit Sub   '�t�@�C�����e�擾�G���[�ł���΁A�߂�B
    '�������̎��s�R�}���h���쐬����
    sCommand = MN_EXE_MEMO & sLzhDataFile
    lRetVal = Shell(sCommand, vbMaximizedFocus)
    '���������A�N�e�B�u�i�O�ʕ\���j�ɂ���
    AppActivate lRetVal, True
    SendKeys "{LEFT}", True
    
    Call ChDrive("D")  'V2.5.0.1 ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fSelectedFilesDelete
'//  �@�\����  : �u�t�@�C���폜�v�t����������
'//  �@�\�T�v  : �I�𒆂̃t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//                                   True:�t�@�C���폜�@FALSE�F�t�@�C�����폜
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     ORIGINAL  :(1.1.0.2) 2009-02-XX   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSelectedFilesDelete() As Boolean
    Dim iLine As Integer         '�ڰ�̧��ؽ��ޯ���̍s���ޯ��
    Dim iMaxLine As Integer      '�ڰ�̧��ؽ��ޯ���̍s��
    Dim iDelLine As Integer      '�ڰ�̧��ؽ��ޯ���̑I���s��
    Dim iResponse As Integer     'MsgBox�{�^���R�[�h
    Dim sFullPass As String      '�폜�Ώۃt�@�C���t���p�X��
    Dim sFileName As String      '�폜�Ώۃt�@�C����
    Dim sWork As String          '���[�N

On Error GoTo ErrorDeleteFile
    
    '�t�@�C���폜�Ȃ��Ƃ��Ă����B
    fSelectedFilesDelete = False
    iDelLine = 0
    '���X�g�{�b�N�X�\�����̑S�s�ɂ��Ĉȉ������{����B
    iMaxLine = lstTraceFile.ListCount  '�ڰ�̧��ؽ��ޯ���̍s���𓾂�B
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        '�I�����ꂽ�s�Ȃ�΁A
            If iDelLine = 0 Then
                '�폜�m�F���b�Z�[�W��\������B
                iResponse = MsgBox("�I�𒆂̃t�@�C�����폜���܂��B" _
                                    & Chr(vbKeyReturn) & " ��낵���ł����H", _
                                    vbYesNo + vbExclamation, _
                                    "�g���[�X�t�@�C���̍폜")
                If iResponse = vbNo Then
                ' [������] �{�^����I�������ꍇ�A�폜�����I������B
                    Exit Function
                End If
            End If
            '�Y���s�t�@�C�����\�����e���Z�b�g����B�isWork���޲Ļ���--01���@\CTRC2000xxx.xxx�j
            sWork = lstTraceFile.List(iLine)
            '�擪�����޲Ļ��ޕ����i"�޲Ļ���--" ����=LIST_HEDDER_LENGTH�j�����O����B
            '                                   �isFileName��01���@\CTRC2000xxx.xxx�j
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            '�R�s�[���t�@�C�����t���p�X���Z�b�g����B�isFullPass��:\tool\R_Mente\DATA\�{�d�S\���w\01���@\CTRC2000xxx.xxx�j
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            '�Y���s�̃t�@�C�����폜����B
            Kill sFullPass
            iDelLine = iDelLine + 1
            '�t�@�C�����폜�����B
            fSelectedFilesDelete = True
            '�u�������D�@�Ӱ�����ݽ��ʁF�t�@�C���폜�v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, FILE_DELETE, 0)
        End If
    Next
Exit Function

ErrorDeleteFile:

    MsgBox "�t�@�C���̍폜�ŃG���[���������܂����B", _
           vbOKOnly + vbExclamation, _
           "�g���[�X�t�@�C���̍폜"

    fSelectedFilesDelete = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v������
'//  �@�\�T�v  : ���[����M�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
On Error Resume Next
    '�ėp���[����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmRMente.Caption, False
        pfFormActive (frmRMente.hwnd)           ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psGATERMenteTool
'//  �@�\����  : �������D�@�̃����[�g�����e�i���X�c�[���p�X���擾����
'//  �@�\�T�v  : �������D�@�����[�g�����e�i���X�c�[���p�X���擾���s���B
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
Public Sub psGATERMenteTool()
 
    Dim sPath As String * MAX_PATH_SIZE
    Dim iRet As Integer
    
    On Error Resume Next

    ' HOSHU.INI��莩�����D�@�c�[���p�X���擾����B
    iRet = GetPrivateProfileString(KANSI_HOSHU_GATE_RMENTE_SEC, _
                                    KANSI_HOSHU_GATE_RMENTE_KEY, _
                                    DEFAILT, sPath, Len(sPath), _
                                    HOSHU_FILE)

      If iRet = 0 Then
        sTOOLPass = ""
      Else
        sTOOLPass = sPath
      End If
      
End Sub


