VERSION 5.00
Begin VB.Form frmShimekiriOutPut 
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
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�n �j"
      Enabled         =   0   'False
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
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
      TabIndex        =   2
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
Attribute VB_Name = "frmShimekiriOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmShimekiriOutPut.frm
'//  �p�b�P�[�W���F���؃f�[�^�o�͒����
'//
'//  �T�v�F���؃f�[�^�o�͒����
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//                 �E���W�E�����e�f�[�^���W�����(frmSyusyuCyu.frm)�𗬗p
'//                 �E�t�F�[�Y�Q�Ή��yMainte_05_03�z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////

Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l
Private lngGateSts(1 To MAX_GATE_NO) As Long                '���@�����W���


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���؃f�[�^�o�͒����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    cmdOK.Enabled = False
    
    On Error Resume Next
    
    '���؃f�[�^�o�͎w�����W�v�֑��M����B
    If fSDATAMailSend = False Then
        lblMessage(0) = "�ُ�I�����܂����B"
        lblMessage(1) = ""
        cmdOK.Enabled = True
        Exit Sub
      
    End If
    
'    ���W���̃K�C�h��\������
    lblMessage(0) = "���؃f�[�^���o�͒��ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ���؃f�[�^�o�͒����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next
    
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : ���؃f�[�^�o�͒����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim i As Integer '�J�E���^
    Dim intCount As Integer
    Dim intCount2 As Integer
    
    On Error Resume Next
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_SHIMEKIRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()

    On Error Resume Next
    
    '����ʂ������B
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
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
                AppActivate frmShimekiriOutPut.Caption, False
                pfFormActive (frmShimekiriOutPut.hwnd)          ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
            Case ML_ID_SHIMEKIRI_OUT_RES
                '�u���؏o�͊����ʒm�v����M�����ꍇ
                '�u���؏o�͊����ʒm��M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, SHIMEKIRI_OUTPUT_REQ_RECV, 0)
                
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
                
                '�N���A�ʒm���e���`�F�b�N����B
                If fReadMailCheck(udtReadMail) = True Then
                    frmShimekiriData.gbShimekiriResult = True
                    lblMessage(0) = "����I�����܂����B"
                    lblMessage(1) = ""
                Else
                    frmShimekiriData.gbShimekiriResult = False
                    lblMessage(0) = "�ُ�I�����܂����B"
                    lblMessage(1) = ""
                End If
                cmdOK.Enabled = True
            Case Else
                 '���̑��̃��[������M�����ꍇ
                 '�u���[��ID�s���v���O�o��
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : fSDATAMailSend
'//  �@�\����  : ���؃f�[�^�o�͎w�����M����
'//  �@�\�T�v  : �����������F���[���𑗐M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 �y�@�\�������z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSDATAMailSend() As Boolean

    Dim udtMail As MAIL_HDATA_REQ  '�ێ�f�[�^���W�w�����[�����M�G���A
    Dim lngRet As Long              '�֐��߂�l
    Dim lngErrCode As Long          '�G���[�R�[�h
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCtlIndex As Integer
    Dim intDataIndex As Integer
    
    On Error Resume Next
 
    fSDATAMailSend = True
    
    '���؃f�[�^�o�͎w�����W�v�ɑ��M����B
    udtMail.mlHeader.dwId = ML_ID_SHIMEKIRI_OUT_CMD
    udtMail.mlHeader.dwSize = MlSize.SHIMEKIRI_OUTPUT_CMD
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequestType = ML_DT_W_SHIMEKIRI_H     '���؃f�[�^
    udtMail.dwStatus(0) = frmShimekiriData.SSTab1.Tab + 1  ' �R�[�i   ' EG20 V6.3.0.1
    
    lngRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       '�u���؉�ʁF���؃f�[�^�o�͎w�����M�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, SHIMEKIRI_OUTPUT_REQ_SEND, lngErrCode)
       
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
       '�v���O���X�o�[����������
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       fSDATAMailSend = False
       Exit Function
    Else
       '�u���؉�ʁF���؃f�[�^�o�͎w�����M����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, SHIMEKIRI_OUTPUT_REQ_SEND, 0)
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : fReadMailCheck
'//  �@�\����  : ���؃f�[�^�����ʒm���[���`�F�b�N����
'//  �@�\�T�v  : ���[����M���F���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim i    As Integer      '�J�E���^
    Dim iErr As Integer      '�����W���@�̗L���i1/0�j
    Dim intIndex As Integer
    On Error Resume Next
    
    iErr = 0
    If udtReadMail.lngData(0) <> ML_DT_W_SHIMEKIRI_H Then
        '�w�������f�[�^��ƈقȂ�ʒm�B
        fReadMailCheck = False
        Exit Function
    End If

    '�X�e�[�^�X�A�����t���O�`�F�b�N
    If udtReadMail.lngData(1) > 0 And iErr = 0 Then
        iErr = 1  '�X�e�[�^�X������ł͂Ȃ��B
    ElseIf udtReadMail.lngData(2) > 0 And iErr = 0 Then
        iErr = 1  '��������
    End If
    
    If iErr > 0 Then
       '�f�[�^�킪�ێ�f�[�^���W�w���̂��̂ƈقȂ�ꍇ
        '�w�������f�[�^��ƈقȂ�ʒm�B
        fReadMailCheck = False
        Exit Function
    End If
    
   fReadMailCheck = True
    
End Function
