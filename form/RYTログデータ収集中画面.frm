VERSION 5.00
Begin VB.Form frmRYTSyusyuCyu 
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdOK 
      Caption         =   "�n �j"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmRYTSyusyuCyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmRYTSyusyuCyu.frm
'//  �p�b�P�[�W���F�q�x�s���O�f�[�^���W�����
'//
'//  �T�v�F�q�x�s���O�f�[�^���W�����
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�q�x�s���O�f�[�^���W����ʒǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �q�x�s���O�f�[�^���W�����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
      
    Dim uMail As MAIL_RYT_LOG_CMD           '���[��
    Dim bRtn As Boolean                 '���[���̖߂�l
    Dim lExitCode As Long
      
    On Error Resume Next
   
    '�q�r�v���Z�X�Ɂu�q�x�s���O�f�[�^���W�v���b�l�c�v�𑗐M����B
    uMail.mlHeader.dwId = ML_ID_RYT_LOG_CMD
    uMail.mlHeader.dwSize = MlSize.RYT_LOG_CMD
    uMail.mlHeader.dwProid = RHOSHU_ID
    uMail.mlHeader.dwSubArea = 0
    uMail.dwRequestType = MailRYTType.ML_DT_LOGDATA_ID
    bRtn = DssSendMail(MAIL_SLOT_RYT, MlSize.RYT_LOG_CMD, uMail.mlHeader)
    If bRtn <> 0 Then
       '�u�q�x�s���O�f�[�^���W�v���b�l�c���M����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, RYT_LOG_CMD_OK, 0)
         
       '���W���̃K�C�h��\������
       lblMessage(0) = "�q�x�s���O�f�[�^�����W���ł��B"
       lblMessage(1) = "���΂炭���҂��������B"
       tmrMail.Enabled = True
    Else
       '�u�q�x�s���O�f�[�^���W�v���b�l�c���M�ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, RYT_LOG_CMD_ERROR, 0)
        '�q�x�s���O�f�[�^���W��������(�ُ�I��)��ʂ�\��
        sSyusyuEnd (1)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �q�x�s���O�f�[�^���W�����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
  
  On Error Resume Next
  
  cmdOK.Visible = False
  cmdOK.Enabled = False
  
  '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
  tmrMail.Interval = MN_MAIL_INTERVAL
  tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �q�x�s���O�f�[�^���W�����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    
    On Error Resume Next
    
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdOK_Click
'//  �@�\����  : �uOK�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    
    On Error Resume Next
    
    '����ʂ������B
    Unload Me
End Sub

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
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    Dim udtReadMail As ML_KYOTU_INF  '���[����M�G���A
    Dim lngLength As Long            '��M���[���o�C�g�T�C�Y
    Dim intStatus As Integer         '��M���[���`�F�b�N����

    On Error Resume Next
    
    '���[������M����B
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
    '��M���[��������΁A���[���h�c���̏���������B
        Select Case udtReadMail.udtlHeader.dwId        '���[���h�c
            Case ML_ID_PROEND_ORD
                 '�u�v���Z�X�I���w���v����M�����ꍇ�A
                 '�u�v���Z�X�I���w����M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                 '�v���Z�X�̏I���������s��
                 pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                 '�u�ێ��ʃA�N�e�B�u�\���v����M�����ꍇ
                 '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                 '�\������ʁi�ێ�f�[�^���W��ʁj���A�N�e�B�u�\������B
                 AppActivate frmRYTSyusyuCyu.Caption, False
            Case ML_ID_RYT_LOG_RES
                 '�u�q�x�s���O�f�[�^���W�v��RES�v����M�����ꍇ
                 '�u�q�x�s���O�f�[�^���W�v���q�d�r��M�v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, RYT_LOG_RES_RECV, 0)
                '���W���ʂ�\������B
                If udtReadMail.lngData(0) = 0 Then
                   '���W���ʁF����
                       sSyusyuEnd (0)
                Else
                   '���W���ʁF�ُ�
                       sSyusyuEnd (1)
                End If
            Case Else
                 '���̑��̃��[������M�����ꍇ
                 '�u���[��ID�s���v���O�o��
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSyusyuEnd
'//  �@�\����  : ���W���ʕ\������
'//  �@�\�T�v  : RYT���O�f�[�^���W���ʂ̌��ʕ�����\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iEnd�@�@�@[IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSyusyuEnd(iEnd As Integer)
    Dim i As Integer       '�J�E���^
    Dim lngErrCode As Long '�G���[�R�[�h

    On Error Resume Next

    If iEnd = 0 Then
       '����I�����̕�����\������B
       lblMessage(0) = "����I�����܂����B"
       lblMessage(1) = ""
       '�u�q�x�s���O�f�[�^���W��������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RYT_LOG_KANRI_GAMEN_SYUSYU_OK, 0)
    Else
       '���W���s���̕�����\������B
       lblMessage(0) = "�ُ�I�����܂����B"
       lblMessage(1) = ""
       '�u�q�x�s���O�f�[�^���W�����ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RYT_LOG_KANRI_GAMEN_SYUSYU_ERROR, lngErrCode)
    End If
    
    cmdOK.Visible = True
    cmdOK.Enabled = True
End Sub

