VERSION 5.00
Begin VB.Form frmJVerUpdate 
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Timer tmrMail 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��������
      Caption         =   "���΂炭���҂��������B"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblMessage 
      Caption         =   "���D�@�p�̃o�[�W���������X�V���ł��B"
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
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
End
Attribute VB_Name = "frmJVerUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmJVerUpData.frm
'//  �p�b�P�[�W���F���D�@�o�[�W�����X�V�����(EG-R����/NEG�����p)
'//
'//  �T�v�F���D�@�o�[�W�����X�V�����(EG-R����/NEG�����p)
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���D�@�o�[�W�����X�V�����(�A�N�e�B�u��)
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
    '���C����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ���D�@�o�[�W�����X�V�����(�f�B�A�N�e�B�u��)
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
Private Sub Form_Deactivate()
    '���C����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : ���D�@�o�[�W�����X�V�����(���[�h��)
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
Private Sub Form_Load()
    Dim udtMail As MAIL_GATE_VER_UPD_REQ  '�����o�[�W�������X�V�v�����[�����M�G���A
    Dim lngRet As Long                    '�֐��߂�l
  
    On Error Resume Next

    '�u���D�@�o�[�W�����X�V���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_UPDATA, 0)

    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
    udtMail.mlHeader.dwId = ML_ID_GATE_VER_UPD_REQ
    udtMail.mlHeader.dwSize = MlSize.GATE_VER_UPD_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequest = gintVerJikai  '�����o�[�W�������X�V�v���̎��
    lngRet = DssSendMail(MAIL_SLOT_KANRI, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       gintGateVerInfUpdRes = MailSts.stsErr
       Unload Me
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v����
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
    Dim udtReadMail As ML_KYOTU_INF  '���[����M�G���A
    Dim lngLength As Long            '��M���[���o�C�g�T�C�Y
    
    On Error Resume Next

    '���[������M����B
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
   '��M���[��������΁A���[���h�c���̏���������B
        Select Case udtReadMail.udtlHeader.dwId        '���[���h�c
            Case ML_ID_PROEND_ORD
                 '�u�v���Z�X�I���w����M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                 '�v���Z�X�̏I���������s��
                 pfAbortProc
                
            Case ML_ID_HOSHU_ACTIVE_REQ
                 '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                 '�\������ʁiEG-R����/NEG������ʁj���A�N�e�B�u�\������B
                 If gintVerJikai = ML_REQUEST_NGATE Then
                    gStrCurrentForm = sFormName_NJVer
                    AppActivate frmJVer.Caption, False
                    pfFormActive (frmJVerUpdate.hwnd)
                 Else
                    gStrCurrentForm = sFormName_EJVer
                    AppActivate frmJVer.Caption, False
                 End If
                
            Case ML_ID_GATE_VER_UPD_INF
                 '�u�����ް�ޮݏ��X�V�ʒm��M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GATE_VERSIONINFO_UPDATA_REQ_RECV, 0)
                 '�\������ʁiEG-R����/NEG������ʁj���A�N�e�B�u�\������B
                 If gintVerJikai = ML_REQUEST_NGATE Then
                    gStrCurrentForm = sFormName_NJVer
                    AppActivate frmJVer.Caption, False
                    pfFormActive (frmJVerUpdate.hwnd)
                 'EG20 V30.1.0.1 ADD START
                 ElseIf gintVerJikai = ML_REQUEST_EG20GATE Then
                    gStrCurrentForm = sFormName_EG20JVer
                    AppActivate frmGateVerKanri.Caption, False
                    pfFormActive (frmGateVerKanri.hwnd)
                 ElseIf gintVerJikai = ML_REQUEST_EG30GATE Then
                    gStrCurrentForm = sFormName_EG30JVer
                    AppActivate frmKansenGateVerKanri.Caption, False
                    pfFormActive (frmKansenGateVerKanri.hwnd)
                'EG20 V30.1.0.1 ADD END
                 Else
                    gStrCurrentForm = sFormName_EJVer
                    AppActivate frmJVer.Caption, False
                 End If
                 gintGateVerInfUpdRes = udtReadMail.lngData(1)
                 
                 '�{��ʂ��I������B
                 Unload Me
                
            Case Else
                '�u���[��ID�s���v���O�o��
                Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Sub
