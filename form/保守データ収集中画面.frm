VERSION 5.00
Begin VB.Form frmSyusyuCyu 
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
      TabIndex        =   1
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
Attribute VB_Name = "frmSyusyuCyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmSyusyuCyu.frm
'//  �p�b�P�[�W���F�ێ�f�[�^���W�����
'//
'//  �T�v�F�ێ�f�[�^���W�����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �E�������A�����ێ�f�[�^���W�����(frmSyusyuCyu.frm)�𗬗p
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'�ێ�f�[�^INDEX
Private Const SYUSYU_KADO = 1    '�ғ��f�[�^
Private Const SYUSYU_MENTE = 2   '�����e�f�[�^
Private Const SYUSYU_ERRLOG = 3  '�G���[���O�f�[�^
'Dim lngDataSyu(SYUSYU_KADO To SYUSYU_ERRLOG) As Long       '���W�f�[�^�̃f�[�^��B 'EG20 V2.1.0.1 DEL
Dim intSyusyuIni(SYUSYU_KADO To SYUSYU_ERRLOG)  As Integer '���W�v��(1/0)HosyuApl.INI��`�l�B
'Dim intSyusyuIndex  As Integer   '���W���̕ێ�f�[�^INDEX       'EG20 V2.1.0.1 DEL
Dim lngGateSts(1 To MAX_GATE_NO) As Long                '���@�����W���
Dim iErrSts As Integer                                  '0:INI��`�ُ�@2�F���[�����M�ُ�@'V1.7.0.1 ADD
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �ێ�f�[�^���W�����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
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
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intCount As Integer
    Dim blnSelected As Boolean
    'EG20 V2.1.0.1 ADD END
    
    cmdOK.Enabled = False
    
    On Error Resume Next
    
    'V1.7.0.1 ADD�@START
    '������
    iErrSts = 0
    'V1.7.0.1 ADD�@END
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    blnSelected = False
    For intCount = 0 To UBound(gintStatus)
        If gintStatus(intCount) = TAG_STATUS.STS_SENTAKU Then
            blnSelected = True
        End If
    Next
    
    '�w�荆�@�Ȃ��̏ꍇ�A���b�Z�[�W�{�b�N�X��\������
    If blnSelected = False Then
        lblMessage(0) = "�w�荆�@���I������Ă��܂���B"
        lblMessage(1) = "�I�����Ă��������B"
        cmdOK.Enabled = True
        Exit Sub
    End If
    
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_KBN_KADO_MAINTE)
    'EG20 V2.1.0.1 ADD END
    
    '���W���̕ێ�f�[�^INDEX������������B
'    intSyusyuIndex = 0         'EG20 V2.1.0.1 DEL

    '�ŏ��̕ێ�f�[�^�ɂ��āA�ێ�f�[�^���W�w�����ă}�֑��M����B
    If fHDATAMailSend = False Then
        'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'      If iErrSts = 0 Then 'V1.7.0.1 ADD INI��`�ُ�̂݉��L�������s���B
'    '�������M���Ȃ������i�S�f�[�^�Ƃ��Ɏ��W�s�v�j�Ȃ�΁A
'        '���b�Z�[�W�{�b�N�X��\�����A
'        MsgBox "HosyuApl.ini�ɁA���W���ׂ��f�[�^����`����Ă��܂���B"
'        '�u�ғ��E�����e�f�[�^���W��ʁF���W�f�[�^����`�v���O�o��
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KADO_MENTE_SYUSYU_GAMEN_DATA_SETTEI_ERROR, 0)
'        '����ʂ������B
'        Unload Me
'        Exit Sub
'      'V1.7.0.1 ADD START
'      Else
'         '����ʂ������B
'         Unload Me
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
        lblMessage(0) = "�ُ�I�����܂����B"
        lblMessage(1) = ""
        cmdOK.Enabled = True
        'EG20 V2.1.0.1 ADD END
         Exit Sub
'      End If           'EG20 V2.1.0.1 DEL
      'V1.7.0.1 ADD END
    End If
    
'    ���W���̃K�C�h��\������
    lblMessage(0) = "�ێ�f�[�^�����W���ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �ێ�f�[�^���W�����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
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
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �ێ�f�[�^���W�����(���[�h��)
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
    Dim strSyusyuKey(SYUSYU_KADO To SYUSYU_ERRLOG) As String 'HosyuApl.INI�L�[�l
    Dim i As Integer '�J�E���^
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intCount As Integer
    Dim intCount2 As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next
    
    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
    '���W�f�[�^�̃f�[�^����Z�b�g���Ă����B
'    lngDataSyu(SYUSYU_KADO) = ML_DT_W_KADO_H      '�ғ��f�[�^
'    lngDataSyu(SYUSYU_MENTE) = ML_DT_W_MENTE_H    '�����e�f�[�^
'    lngDataSyu(SYUSYU_ERRLOG) = ML_DT_W_ERRLOG_H  '�G���[���O�f�[�^
    
    ' HosyuApl.ini����u�ێ�f�[�^���W�v��`���e(���W�v��)����o���B
'    strSyusyuKey(SYUSYU_KADO) = PROFILE_KEY_NAME_HDATA_KADO     '�ғ��f�[�^ �L�[
'    strSyusyuKey(SYUSYU_MENTE) = PROFILE_KEY_NAME_HDATA_MENTE   '�����e�f�[�^ �L�[
'    strSyusyuKey(SYUSYU_ERRLOG) = PROFILE_KEY_NAME_HDATA_ERRLOG '�G���[���O�f�[�^ �L�[
    'EG20 V2.1.0.1 DEL END
    For i = SYUSYU_KADO To SYUSYU_ERRLOG
        intSyusyuIni(i) = GetPrivateProfileInt(PROFILE_KEY_HOSHU_DATA, _
                                               strSyusyuKey(i), DEFAILT_Int, HOSHUAPL_FILE)
    Next
    
    For i = 1 To MAX_GATE_NO
    '���@�����W��Ԃ��A����I���ŏ���������B
        lngGateSts(i) = ML_DT_SEIJO_SHUSHU
    Next
        
    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
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
                'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'EG20 V2.1.0.1 ADD END
                 '�v���Z�X�̏I���������s��
                 pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                 '�u�ێ��ʃA�N�e�B�u�\���v����M�����ꍇ
                 '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                 '�\������ʁi�ێ�f�[�^���W��ʁj���A�N�e�B�u�\������B
                 AppActivate frmSyusyuCyu.Caption, False
                 pfFormActive (frmSyusyuCyu.hwnd)           ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
            Case ML_ID_HDATA_ANS
                 '�u���؊J�n�v��RES�v����M�����ꍇ
                 '�u�ێ�f�[�^���W�ʒm��M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_DATA_SYUSYU_REQ_RECV, 0)
                'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'EG20 V2.1.0.1 ADD END
                '���W�ʒm���e���`�F�b�N����B
                If fReadMailCheck(udtReadMail) = True Then
                '���W�w���ɑ΂�����W�ʒm�ł���΁A
                    '���̃f�[�^��̎��W�w�����ă}�֑��M����B
'                    If fHDATAMailSend = False Then      'EG20 V2.1.0.1 DEL �yMainte_03_01�z
                       '�������M���Ȃ������i�S�f�[�^�Ƃ��Ɏ��W�ρj�Ȃ�΁A
                       '���W�I����Ԃ�\������B
                       sSyusyuEnd
'                    End If                              'EG20 V2.1.0.1 DEL �yMainte_03_01�z
                'EG20 V2.1.0.1 ADD START�yMainte_03_01�z
                Else
                    lblMessage(0) = "�ُ�I�����܂����B"
                    lblMessage(1) = ""
                    cmdOK.Enabled = True
                    '�v���O���X�o�[����������
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                'EG20 V2.1.0.1 ADD END
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
'//  �֐�����  : fHDATAMailSend
'//  �@�\����  : �ێ�f�[�^���W�w�����M����
'//  �@�\�T�v  : �����������F���[���𑗐M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fHDATAMailSend() As Boolean
    Dim udtMail As MAIL_HDATA_REQ  '�ێ�f�[�^���W�w�����[�����M�G���A
    Dim lngRet As Long             '�֐��߂�l
    Dim lngErrCode As Long         '�G���[�R�[�h
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCtlIndex As Integer
    Dim intDataIndex As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'    '������W�w������f�[�^����m�肷��B
'    For intSyusyuIndex = intSyusyuIndex + 1 To SYUSYU_ERRLOG
'        If intSyusyuIni(intSyusyuIndex) = 1 Then '���W�v�̃f�[�^���T���B
'            Exit For
'        End If
'    Next
'
'    '�S�Ă̎��W�f�[�^�Ɏw���ςł���΁A
'    If intSyusyuIndex > SYUSYU_ERRLOG Then
'       '�S�Ă̎��W�f�[�^�Ɏw���ςł���΁A
'       '������I����Ŗ߂�B
'        fHDATAMailSend = False
'        iErrSts = 0             'V1.7.0.1 ADD
'        Exit Function
'    End If
    'EG20 V2.1.0.1 DEL END
 
    '�ێ�f�[�^���W�w�����W�v�ɑ��M����B
    udtMail.mlHeader.dwId = ML_ID_HDATA_REQ
    udtMail.mlHeader.dwSize = MlSize.HOSHU_SYUSYU_CMD
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
'    udtMail.dwRequestType = lngDataSyu(intSyusyuIndex) '�Y���f�[�^�̃f�[�^��        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
    udtMail.dwRequestType = ML_DT_W_KADO_MAINTE_H       '���W�E�����e�f�[�^         'EG20 V2.1.0.1 ADD �yMainte_03_01�z
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    '���@�ʎ��W�X�e�[�^�X��ݒ肷��
    For intCount = 0 To 31
        If gintStatus(intCount) = TAG_STATUS.STS_SENTAKU Then
            udtMail.dwStatus(intCount) = TAG_STATUS.STS_SENTAKU
        Else
            udtMail.dwStatus(intCount) = TAG_STATUS.STS_MISENTAKU
        End If
    Next intCount
    'EG20 V2.1.0.1 ADD END
    
    lngRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
        '�u�ғ��E�����e�f�[�^���W��ʁF�ێ�f�[�^���W�w�����M�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_DATA_SYUSYU_CMD_SEND, lngErrCode)
        lblMessage(0) = "�ُ�I�����܂����B"
        lblMessage(1) = ""
        iErrSts = 2                   'V1.7.0.1 ADD
        'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        'EG20 V2.1.0.1 ADD END
        Exit Function
    Else
       '�u�ғ��E�����e�f�[�^���W��ʁF�ێ�f�[�^���W�w�����M����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_DATA_SYUSYU_CMD_SEND, 0)
    End If
        
    '����M����Ŗ߂�B
    fHDATAMailSend = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fReadMailCheck
'//  �@�\����  : �ێ�f�[�^���W�ʒm���[���`�F�b�N����
'//  �@�\�T�v  : ���[����M���F���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim iEnd As Integer      '
    Dim i    As Integer      '�J�E���^
    Dim iErr As Integer      '�����W���@�̗L���i1/0�j
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intIndex As Integer
    'EG20 V2.1.0.1 ADD END
    On Error Resume Next
    
    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'    iEnd = 0
'    If intSyusyuIndex > SYUSYU_ERRLOG Then
'        iEnd = 1  '�ێ�f�[�^���W�͊��ɏI�����Ă���B
'    ElseIf udtReadMail.lngData(0) <> lngDataSyu(intSyusyuIndex) Then
    'EG20 V2.1.0.1 DEL END
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    If udtReadMail.lngData(0) <> ML_DT_W_KADO_MAINTE_H Then
        iEnd = 2  '�w�������f�[�^��ƈقȂ�ʒm�B
    End If
    'EG20 V2.1.0.1 ADD END

'    If iEnd = 1 Then
'       '�f�[�^�킪�ێ�f�[�^���W�w���̂��̂ƈقȂ�ꍇ�A
'       '���W�ʒm�ُ�̃��O�o�͂��˗�����B
'       sLogRequest 2, udtReadMail
'       '��w���ɑ΂���ʒm�ł͂Ȃ���Ƃ��āA�߂�B
'       fReadMailCheck = False
'       Exit Function
'   End If
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    '�X�e�[�^�X�A�����t���O�`�F�b�N
    If udtReadMail.lngData(1) > 0 And iEnd = 0 Then
        iEnd = 1
    ElseIf udtReadMail.lngData(2) > 0 And iEnd = 0 Then
        iEnd = 1
    End If
    
    If iEnd = 2 Then
       '�f�[�^�킪�ێ�f�[�^���W�w���̂��̂ƈقȂ�ꍇ�A
       '���W�ʒm�ُ�̃��O�o�͂��˗�����B
       sLogRequest iErr, udtReadMail
       '��w���ɑ΂���ʒm�ł͂Ȃ���Ƃ��āA�߂�B
       fReadMailCheck = False
       Exit Function
    End If
    'EG20 V2.1.0.1 ADD END
    
  
   '����̎��W��Ԃ��A���@�����W��ԂɃ�������B
   iErr = 0       '�����W���@ �����A�Ƃ��Ă����B
   'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'   For i = 1 To MAX_GATE_NO
'       If udtReadMail.lngData(i) = ML_DT_MISHUSHU Then
'          '������W��ł���΁A��������B
'          lngGateSts(i) = ML_DT_MISHUSHU
'          iErr = 1
'       ElseIf udtReadMail.lngData(i) = ML_DT_GOUKI_NASI Then
'              '����@�Ȃ���ł���΁A��������B
'              lngGateSts(i) = ML_DT_GOUKI_NASI
'       End If
'    Next
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    For i = 3 To MAX_GATE_NO + 2
        intIndex = i - 3
        If gintStatus(intIndex) <> TAG_STATUS.STS_MISENTAKU Then
            Select Case udtReadMail.lngData(i)
            Case ML_DT_MISHUSHU, ML_DT_IJO_SHUSHU
                '������W��A�u�ُ�I���v�ł���΁A��������B
                lngGateSts(intIndex + 1) = udtReadMail.lngData(i)
                iErr = 1
                gintStatus(intIndex) = TAG_STATUS.STS_MISHUSHU
            Case ML_DT_GOUKI_NASI
                '����@�Ȃ���ł���΁A��������B
                lngGateSts(intIndex + 1) = ML_DT_GOUKI_NASI
                gintStatus(intIndex) = TAG_STATUS.STS_MISENTAKU
            Case ML_DT_SEIJO_SHUSHU
                '�u����I���v
                gintStatus(intIndex) = TAG_STATUS.STS_SHUSHU
            End Select
        End If
    Next
    'EG20 V2.1.0.1 ADD END
       
    '���W�`�F�b�N���s���B
    sLogRequest iErr, udtReadMail
    
    If iEnd = 0 Then
        fReadMailCheck = True
    Else
        fReadMailCheck = False
    End If
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sLogRequest
'//  �@�\����  : ���W���ʃ`�F�b�N����
'//  �@�\�T�v  : ���[����M���F���W���ʃ`�F�b�N���s���B
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
Private Sub sLogRequest(iErr As Integer, udtReadMail As ML_KYOTU_INF)
    Dim strDataSyu  As String    '�f�[�^��̕\��������
    Dim strErrorLog As String    '���O�g���[�X�˗����Ұ�
    
    On Error Resume Next

    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
    '���W���̃f�[�^��̕\����������Z�b�g����B
'    If intSyusyuIndex = SYUSYU_KADO Then
'        strDataSyu = "�ғ��f�[�^���W"
'    ElseIf intSyusyuIndex = SYUSYU_MENTE Then
'        strDataSyu = "�����e�f�[�^���W"
'    Else
'        strDataSyu = "�G���[���O���W"
'    End If
    'EG20 V2.1.0.1 DEL END
    
    If iErr = 0 Then
     '�Y���f�[�^��A����I���̏ꍇ�A
        strErrorLog = fLogStatusGet(udtReadMail)
    ElseIf iErr = 1 Then
    '�Y���f�[�^��A�����W����̏ꍇ�A
        strErrorLog = fLogStatusGet(udtReadMail)
    ElseIf iErr = 2 Then
    '�ʒm���[���ُ�̏ꍇ�A
        strErrorLog = Format$(udtReadMail.lngData(0), "000000")
    End If
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    '���O�o�͂��˗�����B
    Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_DATA_SYUSYU_REQ_RECV & ":" & strErrorLog, 0)
    'EG20 V2.1.0.1 ADD END
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fLogStatusGet
'//  �@�\����  : ���W���ʍ��@�ʏ�ԕ�����ҏW����
'//  �@�\�T�v  : ���[����M���F���@�ʕ\��������ҏW����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fLogStatusGet(udtReadMail As ML_KYOTU_INF)
    Dim i       As Integer     '�J�E���^
    Dim strWork As String      '�ҏW������

    On Error Resume Next
    
    strWork = ""
    '�S���@�ɂ��Ē��ׂ�B
    For i = 1 To MAX_GATE_NO
        If udtReadMail.lngData(i) <> ML_DT_GOUKI_NASI Then
           '�������̍��@�ɕt���ẮA�ҏW���Ȃ��B
           '���@�ԍ��������ށB
           strWork = strWork & "No" & Format$(i, "00")
           If udtReadMail.lngData(i) = ML_DT_SEIJO_SHUSHU Then
            '�����I����ł���΁AOK�������ށB
               strWork = strWork & "=OK,"
           Else
            '������W��ł���΁ANG�������ށB
               strWork = strWork & "=NG,"
           End If
        End If
    Next
    fLogStatusGet = strWork
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSyusyuEnd
'//  �@�\����  : ���W�I����ԕ\������
'//  �@�\�T�v  : ���[����M���F�ێ�f�[�^���W���ʕ�����\������B
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
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSyusyuEnd()
    Dim i As Integer       '�J�E���^
    Dim iEnd As Integer    '�I�����
    Dim lngErrCode As Long '�G���[�R�[�h

    On Error Resume Next

    iEnd = 0
    For i = 0 To MAX_GATE_NO - 1
        '�����W�̍��@���������Ȃ�΁A
        If gintStatus(i) = TAG_STATUS.STS_MISHUSHU Then
           '�ێ�f�[�^���W�́A���W���s�Ƃ���B
           iEnd = i
           Exit For
        End If
    Next
    If iEnd = 0 Then
       '����I�����̕�����\������B
       lblMessage(0) = "����I�����܂����B"
       lblMessage(1) = ""
       '�u�ێ�f�[�^���W��������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_SYUSYU_OK, 0)
    Else
       '���W���s���̕�����\������B
'       lblMessage(0) = "���W���s�B" & str(iEnd) & "���@�������W�ł��B"      'EG20 V2.1.0.1 DEL �yMainte_03_01�z
       lblMessage(0) = "���W���s�B�����W���@������܂��B"                    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
     '  lblMessage(1) = "-- �����۸��ڰ�(�Ď���)�Q�ƁB --" 'V1.8.0.1 DEL
'        lblMessage(1) = "-- ����͊Ď��Ճ��O�Ǘ��Q�ƁB --" 'V1.8.0.1 ADD   'EG20 V2.1.0.1 DEL �yMainte_03_01�z
        lblMessage(1) = "�w�荆�@�I��\���Ŋm�F���Ă��������B"              'EG20 V2.1.0.1 ADD �yMainte_03_01�z
       '�u�ێ�f�[�^���W�����ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_SYUSYU_ERROR, lngErrCode)
    End If
    cmdOK.Enabled = True
End Sub
