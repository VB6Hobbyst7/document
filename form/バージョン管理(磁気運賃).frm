VERSION 5.00
Begin VB.Form frmJikiUnkaiFD 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�� �C �^ �� �f �[ �^ �e �c �� ��"
   ClientHeight    =   9000
   ClientLeft      =   2700
   ClientTop       =   2220
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdKirikae 
      Caption         =   "�����ؑ�"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   3315
      Width           =   2175
   End
   Begin VB.CommandButton cmdFDInput 
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
      Height          =   1095
      Left            =   3120
      TabIndex        =   1
      Top             =   3315
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Left            =   600
      Top             =   7440
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "      ���j���[       ��ʂ֖߂�"
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
      Left            =   9360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "���C�^���o�[�W�����Ǘ�"
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
Attribute VB_Name = "frmJikiUnkaiFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmJikiUnkaiFD.frm
'//  �p�b�P�[�W���F�o�[�W�����Ǘ�(���C�^��)���
'//
'//  �T�v�F�o�[�W�����Ǘ�(���C�^��)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �E�������A�o�[�W�����Ǘ�(���C�^��)��ʗ��p�B
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.16.0.1) 2009-12-21  REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ̏����\���t�H���_���w��
'//                �u�����ؑցv�t�̕\���L����INI�t�@�C����
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir�֐���FileSystemObject�ɕύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

Dim mstrCopyFile()  As String           '�^���t�@�C�����ۑ��G���A

'���b�Z�[�W�{�b�N�X�\���p�h�c
Private Enum FDUNKAI
    FD_INSERT = 1                       '�@�P�F�e�c�}���˗��l�r�f
    REBOOT = 2                          '�@�Q�F�ċN���m�F�l�r�f
    FD_INSERT_ERR = 11                  '�P�P�F�}���t�@�C���ُ�l�r�f
    FD_INPUT_ERR = 12                   '�P�Q�F�e�c���͌��ʊm�F�l�r�f
    TODAY_CHANGE = 21                   '�Q�P�F���C�^�������ؑ֏����m�F�l�r�f
    CHANGE_OK = 22                      '�Q�Q�F���C�^�������ؑ֏������ʂl�r�f�i����j
    CHANGE_ERR = 31                     '�R�P�F���C�^�������ؑ֏������ʂl�r�f�i�ُ�j
End Enum

Private mIntFD      As Integer          '�e�c�}������
Private mIntFDTotal As Integer          '�e�c����

Private Const DEFAILT_HYOUJI_UMU = 1    '�u�����ؑցv�t�̃f�t�H���g�\��     'V1.20.0.1 ADD
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����Ǘ�(���C�^��)���(�A�N�e�B�u��)
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
    
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����Ǘ�(���C�^��)���(�f�B�A�N�e�B�u��)
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

    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����Ǘ�(���C�^��)���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                �u�����ؑցv�t�̕\���L����INI�t�@�C����
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim lSts As Long             '�֐��߂�l      'V1.20.0.1 ADD
    On Error Resume Next
    
    '�u���C�^���o�[�W�����Ǘ���ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

    'V1.20.0.1 ADD START
    '�ێ�.ini���A�u�����ؑցv�t�̕\���L�����擾����B
    lSts = GetPrivateProfileInt(KANS_JIKI, _
                                   KANSI_KIRIKAE_UMU, _
                                   DEFAILT_HYOUJI_UMU, _
                                   HOSHU_FILE)
    If lSts = 1 Then
        cmdKirikae.Visible = True
    Else
        cmdKirikae.Visible = False
    End If
    'V1.20.0.1 ADD END

    '���[����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFDInput_Click
'//  �@�\����  : �u�}�̓����v�t����������
'//  �@�\�T�v  : �}�̓������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 �s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFDInput_Click()

    Dim iResponse   As Integer          'MsgBox�{�^���R�[�h
    Dim bRet        As Boolean          '���[�����M����

    On Error Resume Next
    
    '�u���C�^���o�[�W�����Ǘ���ʁF�}�̓����t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDINPUT_BUTTOM, 0)

    '������
    bRet = False        '�����߂�l
    mIntFD = 0          '�}���t�@�C�������J�E���^
    mIntFDTotal = 0     '�t�@�C������

    '��ʂ̃{�^���������s�ɂ���
    Call sButtonEnabled(False)

    'V1.16.0.1 ADD START
    '�u�}�̓����v���t�@�C�����擾����
    m_sFILE_NAME1 = ""
    m_sFILE_NAME2 = ""
    m_sFILE_NAME1 = fGetFilName(VERJIKI_SEC1, VERJIKI_SEC1_KEY1)    '�擪�t�@�C����
    m_sFILE_NAME2 = fGetFilName(VERJIKI_SEC1, VERJIKI_SEC1_KEY2)    '�g���q��
    '�擾���ʃ`�F�b�N
    If m_sFILE_NAME1 = "" Or m_sFILE_NAME2 = "" Then
       '��ʂ̃{�^���������\�ɂ���
        Call sButtonEnabled(True)
       '���b�Z�[�W��\���u���C�^���f�[�^���͌��ʁ@�ُ�I���v
       fMessageBox (FDUNKAI.FD_INPUT_ERR)
       Exit Sub
    End If
    'V1.16.0.1 ADD END
    '���ɉ^���f�[�^������ꍇ�A�j������
    sFileDelete

    '���b�Z�[�W��\���u�e�c�}���@�˗��v
    iResponse = fMessageBox(FDUNKAI.FD_INSERT)

    If iResponse = vbOK Then        '�n�j����

        '�}�����ꂽ�e�c�̃t�@�C�������`�F�b�N�����[�N�t�H���_�ɃR�s�[����
        bRet = fFDFileNameCheck()

        '�t�@�C�����`�F�b�N�A�t�@�C���R�s�[������I���������A�^���t�@�C�����쐬����
        If bRet = True Then
            '�^���t�@�C�����쐬����
            bRet = fFileJoint

            '���[�����M���s��
            If bRet = True Then
                '���b�Z�[�W�{�b�N�X�\���u�ċN���v���@�m�F�v
                iResponse = fMessageBox(FDUNKAI.REBOOT)

                If iResponse = vbOK Then    'OK����
                    '���[�����M����
                    bRet = fSendMail(MAIL_SLOT_KANRI, ML_ID_KAN_PW_OFF_REQ)
                Else                        '�L�����Z���F��ōċN��
                    '��ʂ̃{�^���������\�ɂ���
                    Call sButtonEnabled(True)
                End If
            End If
        End If
    End If

    '�����L�����Z���܂��́A���[�����M���s
    If bRet = False Then
        '�^���t�@�C����S�č폜
        Call sFileDelete

        '��ʂ̃{�^���������\�ɂ���
        Call sButtonEnabled(True)
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdKirikae_Click
'//  �@�\����  : �u�����ؑցv�t����������
'//  �@�\�T�v  : �����ؑ֏������s���B
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
Private Sub cmdKirikae_Click()

    Dim iResponse   As Integer          'MsgBox�{�^���R�[�h
    Dim bSendMail   As Boolean          '���[�����M����
    On Error Resume Next
    
    '�u���C�^���o�[�W�����Ǘ���ʁF�����֖ؑt�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_KIRIKAE_BUTTOM, 0)

    '������
    bSendMail = False

    '��ʂ̃{�^���������s�ɂ���
    Call sButtonEnabled(False)

    '���b�Z�[�W�{�b�N�X�\���u���C�^�������ؑ֏����@�m�F�v
    iResponse = fMessageBox(FDUNKAI.TODAY_CHANGE)

    If iResponse = vbOK Then   'OK����
        bSendMail = fSendMail(MAIL_SLOT_KANMA, ML_ID_HOSHU_UNKAI_DAYCHG_REQ)
    End If

    '�����L�����Z���܂��́A���[�����M���s
    If bSendMail = False Then
        '��ʂ̃{�^���������\�ɂ���
        Call sButtonEnabled(True)
       '�u���C�^���o�[�W�����Ǘ���ʁF�����ؑ֏����ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JIKI_VERSION_KANRI_KIRIKAE_ERROR, 0)
    Else
       '�u���C�^���o�[�W�����Ǘ���ʁF�����ؑ֏�������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_KIRIKAE_OK, 0)
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂɖ߂�v�t����������
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
    '�u���C�^���o�[�W�����Ǘ���ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFileDelete
'//  �@�\����  : �^���t�@�C�����폜����B
'//  �@�\�T�v  : �u�}�̓����v�t�����������F���݂���^���t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 �s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sFileDelete()

    On Error Resume Next

    '�e�c�^���t�@�C���폜
'   Kill PATH_WORK & FILE_NAME1 & "*" & FILE_NAME2         'V1.16.0.1 DEL
    Kill PATH_WORK & m_sFILE_NAME1 & "*" & m_sFILE_NAME2   'V1.16.0.1 ADD
    '�u���C�^���o�[�W�����Ǘ���ʁFFD�^���t�@�C���폜�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDUNKAI_FILE_DELETE, 0)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fFDFileNameCheck
'//  �@�\����  : �^���t�@�C�����`�F�b�N
'//  �@�\�T�v  : �u�}�̓����v�t�����������F�����^���t�@�C�����`�F�b�N
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.16.0.1) 2009-12-21  REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ̏����\���t�H���_���w��
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 Dir�֐���FileSystemObject�ɕύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fFDFileNameCheck() As Boolean

    Dim strFDFile   As String           '�t�@�C����
    Dim iResponse   As Integer          'MsgBox�{�^���R�[�h
    Dim intFDNum    As Integer          '�e�c�ԍ�
    Dim bFileCHK    As Boolean          '�t�@�C���`�F�b�N
    Dim strWriteDir As String
   
    On Error Resume Next

    '������
    fFDFileNameCheck = False    '�֐��߂�l
    iResponse = 0               '���b�Z�[�W�{�b�N�X�{�^���R�[�h
    intFDNum = 0                '�擾����e�c�ԍ�

'V1.12.0.1 ADD  START
    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
    'strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")     'V1.20.0.1 DEL
    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)      'V1.20.0.1 ADD

    '�w��t�H���_�Ȃ�
    If Len(strWriteDir) = 0 Then
        iResponse = vbCancel
    End If
'V1.12.0.1 ADD  END
    '�����L�����Z���܂��́A�t�@�C�����擾�ł����Ƃ��ɁA���[�v�𔲂���
    Do Until mIntFD = 1 Or iResponse = vbCancel
        '�t�@�C�����擾�i�P��ڂ̃t�@�C���}���́A�K���P���ڂɂȂ�j
'        strFDFile = Dir(FDDRIVE & FILE_NAME1 & "*" & FILE_NAME2)   'V1.12.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & FILE_NAME1 & "*" & FILE_NAME2)    'V1.12.0.1 ADD  'V1.16.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & m_sFILE_NAME1 & "*" & m_sFILE_NAME2)    'V1.12.0.1 ADD  'V1.16.0.1 ADD 'V2.6.0.1 DEL
         strFDFile = sDoFileFind(strWriteDir & "\" & m_sFILE_NAME1 & "*" & m_sFILE_NAME2) 'V2.6.0.1 ADD
        '�t�@�C�����擾�`�F�b�N
        If strFDFile <> "" Then         '�擾����
            '�t�@�C���̑������擾����
'            mIntFDTotal = CInt(Mid(strFDFile, Len(FILE_NAME1) + 2, 1))  'V1.16.0.1 DEL
             mIntFDTotal = CInt(Mid(strFDFile, Len(m_sFILE_NAME1) + 2, 1))   ''V1.16.0.1 ADD
        End If

        '�擾�����t�@�C�������`�F�b�N
        If mIntFDTotal > 0 And mIntFDTotal < 10 Then
            mIntFD = 1          '�擾�����I�@���[�v�J�E���^���P�ɃZ�b�g
        Else
            '���b�Z�[�W��\���u�e�c�}���@�t�@�C���ُ�v
            iResponse = fMessageBox(FDUNKAI.FD_INSERT_ERR)

            If iResponse = vbCancel Then
                '���b�Z�[�W��\���u���C�^���f�[�^���͌��ʁ@�ُ�I���v
                fMessageBox (FDUNKAI.FD_INPUT_ERR)
'            End If 'V1.12.0.1 DEL
                        'V1.12.0.1 ADD  START
            Else
                '�t�H���_�I���|�b�v�A�b�v��ʕ\��
                'strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")   'V1.20.0.1 DEL
                strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
            
                '�w��t�H���_�Ȃ�
                If Len(strWriteDir) = 0 Then
                    iResponse = vbCancel
                End If
            End If
                        'V1.12.0.1 ADD  END

        End If
    Loop

    '�t�@�C�����ۑ��G���A�̍Ē�`
    ReDim mstrCopyFile(mIntFDTotal - 1)

    '�e�c�ԍ������������`�F�b�N����B�����L�����Z���܂��́A�t�@�C����S�Ď擾�ł�����A���[�v�𔲂���
    Do Until mIntFD > mIntFDTotal Or iResponse = vbCancel
        '�t�@�C�����擾
'        strFDFile = Dir(FDDRIVE & FILE_NAME1 & mIntFD & "*" & FILE_NAME2)      'V1.12.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & FILE_NAME1 & mIntFD & "*" & FILE_NAME2)   'V1.12.0.1 ADD 'V1.16.0.1 DEL
'        strFDFile = Dir(strWriteDir & "\" & m_sFILE_NAME1 & mIntFD & "*" & m_sFILE_NAME2)   'V1.12.0.1  'V1.16.0.1 ADD 'V2.6.0.1 DEL
         strFDFile = sDoFileFind(strWriteDir & "\" & m_sFILE_NAME1 & mIntFD & "*" & m_sFILE_NAME2) 'V2.6.0.1 ADD

        '�t�@�C�����擾�`�F�b�N
        'V1.16.0.1 DEL START
        'If Len(strFDFile) = Len(FILE_NAME1 & FILE_NAME2) + 2 Then      '�擾����
        '    intFDNum = CInt(Mid(strFDFile, Len(FILE_NAME1) + 1, 1))     '�e�c�ԍ�
        'End If
        'V1.16.0.1 DEL END
        'V1.16.0.1 ADD START
        If Len(strFDFile) = Len(m_sFILE_NAME1 & m_sFILE_NAME2) + 2 Then   '�擾����
            intFDNum = CInt(Mid(strFDFile, Len(m_sFILE_NAME1) + 1, 1))     '�e�c�ԍ�
        End If
        'V1.16.0.1 ADD END

        '�擾�����t�@�C���ԍ��iintFDNum�j�Ɗ��҂���t�@�C���ԍ��imIntFD�j�������Ă��邩�H
        If intFDNum = mIntFD Then       '�t�@�C���ԍ��@����
            '���[�N�t�H���_�ɁA�e�c�t�@�C�����R�s�[����
'            Call FileCopy(FDDRIVE & strFDFile, PATH_WORK & strFDFile)      'V1.12.0.1 DEL
'            Call FileCopy(strWriteDir & strFDFile, PATH_WORK & strFDFile)   'V1.12.0.1 ADD 'V1.16.0.1 DEL
             Call FileCopy(strWriteDir & "\" & strFDFile, PATH_WORK & strFDFile)   'V1.16.0.1 ADD

            '���[�N�t�@�C���̃t�@�C������ۑ�����
            mstrCopyFile(mIntFD - 1) = PATH_WORK & strFDFile

'V1.12.0.1 DEL START
'            '�e�c�ԍ���������菭�Ȃ��ꍇ�A���̂e�c�̑}���𑣂�
'            If mIntFD < mIntFDTotal Then
'                '���b�Z�[�W��\���u�e�c�}���@�˗��v
'                iResponse = fMessageBox(FDUNKAI.FD_INSERT)
'
'                If iResponse = vbCancel Then
'                    '���b�Z�[�W��\���u���C�^���f�[�^���͌��ʁ@�ُ�I���v
'                    fMessageBox (FDUNKAI.FD_INPUT_ERR)
'                End If
'            End If
'V1.12.0.1 DEL END

            '�e�c�������J�E���g�A�b�v
            mIntFD = mIntFD + 1
        Else                            '�t�@�C���ԍ��@�ُ�
            '���b�Z�[�W��\���u�e�c�}���@�t�@�C���ُ�v
            iResponse = fMessageBox(FDUNKAI.FD_INSERT_ERR)

            If iResponse = vbCancel Then
                '���b�Z�[�W��\���u���C�^���f�[�^���͌��ʁ@�ُ�I���v
                fMessageBox (FDUNKAI.FD_INPUT_ERR)
'            End If     'V1.12.0.1 DEL
'V1.13.0.1 ADD  START
            Else
                '�t�H���_�I���|�b�v�A�b�v��ʕ\��
                'strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")   'V1.20.0.1 DEL
                strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
            
                '�w��t�H���_�Ȃ�
                If Len(strWriteDir) = 0 Then
                    iResponse = vbCancel
                    '���b�Z�[�W��\���u���C�^���f�[�^���͌��ʁ@�ُ�I���v
                    fMessageBox (FDUNKAI.FD_INPUT_ERR)
                End If
            End If
'V1.13.0.1 ADD  END

        End If
    Loop

    '����I�������Ƃ��A�֐��̖߂�l��True�ɐݒ肷��B
    If iResponse <> vbCancel Then
        fFDFileNameCheck = True
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fFileJoint
'//  �@�\����  : �^���t�@�C�����쐬����B
'//  �@�\�T�v  : �u�}�̓����v�t�����������F
'//              �����t�@�C�����A�^���t�@�C�����쐬����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 �s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fFileJoint() As Boolean

    Dim intLoop         As Integer          '���[�v�J�E���^
    Dim intReadFileNo   As Integer          '�Ǎ��t�@�C���ԍ�
    Dim intWriteFileNo  As Integer          '�����t�@�C���ԍ�
    Dim bytReadFile()   As Byte             '�Ǎ��t�@�C���G���A
    Dim strFile         As String           '�Ǎ��t�@�C���擾

    On Error Resume Next

    '������
    fFileJoint = False      '�߂�l

    '�^���f�[�^�폜
    Kill FDUNKAI_FILE

    intWriteFileNo = FreeFile        '�����ݐ�p�t�@�C���̔ԍ����擾����B

    On Error GoTo Err_LOG

    '�����p�t�@�C��(FD_UNKAI.DAT)�������ݐ�p�o�C�i�����[�h�ŃI�[�v������B
    Open FDUNKAI_FILE For Binary As #intWriteFileNo
        For intLoop = 0 To mIntFDTotal - 1      '�e�c�����Ń��[�v����B

            '�t�@�C�������݂��邱�Ƃ��m�F����B
            strFile = Dir(mstrCopyFile(intLoop))
            If strFile = "" Then                '�t�@�C�����Ȃ�������A�G���[������
                GoTo Err_LOG
            End If

            intReadFileNo = FreeFile            '�Ǎ���p�̃t�@�C���ԍ����擾����B

            On Error GoTo Err_LOG

            '�Ǎ��t�@�C���G���A�̔z��̍Ē�`
            ReDim bytReadFile(FileLen(mstrCopyFile(intLoop)) - 1)


            '�Ǎ���p�t�@�C���̔ԍ����擾����B
            Open mstrCopyFile(intLoop) For Binary Access Read As intReadFileNo

            '�ǂݍ��݃t�@�C���f�[�^�擾
            Get intReadFileNo, , bytReadFile

            Close #intReadFileNo            '�Ǎ��t�@�C�������B

            '�^���f�[�^�t�@�C���ɏ�����
            Put #intWriteFileNo, , bytReadFile

        Next
    Close #intWriteFileNo                   '�����݃t�@�C�������B

    fFileJoint = True

    '�e�c�^���t�@�C���폜
    'Kill PATH_WORK & FILE_NAME1 & "*" & FILE_NAME2       'V1.16.0.1 DEL
    Kill PATH_WORK & m_sFILE_NAME1 & "*" & m_sFILE_NAME2  'V1.16.0.1 ADD
    '�u���C�^���o�[�W�����Ǘ���ʁFFD�^���t�@�C���폜�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDUNKAI_FILE_DELETE, 0)
    '�u���C�^���o�[�W�����Ǘ���ʁF�}�̓�����������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JIKI_VERSION_KANRI_FDINPUT_OK, 0)

    Exit Function

Err_LOG:

    '���b�Z�[�W��\���u���C�^���f�[�^���͌��ʁ@�ُ�I���v
    fMessageBox (FDUNKAI.FD_INPUT_ERR)

    '�Ǎ��t�@�C�������
    If intReadFileNo > 0 Then
        Close #intReadFileNo
    End If

    '�����t�@�C�������
    If intWriteFileNo > 0 Then
        Close #intWriteFileNo
    End If

    '�e�c�^���f�[�^�폜
    sFileDelete

    '�^���f�[�^�폜
    Kill FDUNKAI_FILE
    '�u���C�^���o�[�W�����Ǘ���ʁF�}�̓��������ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JIKI_VERSION_KANRI_FDINPUT_ERROR, 0)

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sButtonEnabled
'//  �@�\����  : �\����ʂ̖t�R���g���[�����s���B
'//  �@�\�T�v  : �u�}�̓����v�u�����ؑցv�t�����������F�t��������/�����s�ɂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Boolean�@bSet�@�@�@[IN]�t�̃R���g���[��(TRUE�F������,FALSE�F�����s��)
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sButtonEnabled(bSet As Boolean)

    On Error Resume Next

    cmdFDInput.Enabled = bSet       '�}�̓����{�^��
    cmdKirikae.Enabled = bSet       '�����ؑփ{�^��
    cmdReturn.Enabled = bSet        '���j���[��ʂ֖߂�{�^��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMessageBox
'//  �@�\����  : ��ʃ|�b�v�A�b�v�\������
'//  �@�\�T�v  : ���b�Z�[�WID�ɂ���ĕ\���|�b�v�A�b�v������/�\��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@ iMsgID   [IN]���b�Z�[�WID
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMessageBox(iMsgID As Integer) As Integer

    Dim strMessage  As String           'MSGBOX�̕���
    Dim strTitle    As String           'MSGBOX�̃^�C�g��
    Dim lngOption   As Long             'MSGBOX�̕\���{�^���ƃA�C�R��

    fMessageBox = 0
   
   On Error Resume Next

    Select Case iMsgID
        Case FDUNKAI.FD_INSERT      '�e�c�}���@�˗�
            strMessage = "���C�^����}�����Ă��������B"
            lngOption = vbOKCancel + vbInformation      '�u�n�j�v�u�L�����Z���v�{�^���A�u���v�A�C�R��
            strTitle = "�}�̑}��"

        Case FDUNKAI.REBOOT         '�ċN���v��
            strMessage = "�Ď��Ղ��ċN�����܂����A��낵���ł����H"
            lngOption = vbOKCancel + vbExclamation      '�u�n�j�v�u�L�����Z���v�{�^���A�u���Ӂv�A�C�R��
            strTitle = "�ċN���m�F"

        Case FDUNKAI.FD_INSERT_ERR  '�e�c�}���@�t�@�C���ُ�
            strMessage = "�ُ�ȃt�@�C�����}������܂����B" & Chr(vbKeyReturn) & _
                         "���������C�^����}�����Ă��������B"
            lngOption = vbOKCancel + vbCritical         '�u�n�j�v�u�L�����Z���v�{�^���A�u�x���v�A�C�R��
            strTitle = "�}�̑}��"

        Case FDUNKAI.FD_INPUT_ERR   '���C�^���f�[�^���͌��ʁ@�ُ�I��
            strMessage = "�^���f�[�^���ُ͈͂�I�����܂����B"
            lngOption = vbOKOnly + vbInformation        '�u�n�j�v�{�^���A�u���v�A�C�R��
            strTitle = "���C�^���f�[�^���͌���"

        Case FDUNKAI.TODAY_CHANGE   '���C�^�������ؑ֏����@�m�F
            strMessage = "���C�^�������ؑ֏������s���܂����A��낵���ł����H"
            lngOption = vbOKCancel + vbExclamation      '�u�n�j�v�u�L�����Z���v�{�^���A�u���Ӂv�A�C�R��
            strTitle = "���C�^�������ؑ֏����m�F"

        Case FDUNKAI.CHANGE_OK      '���C�^�������ؑ֏����@����I��
            strMessage = "���C�^�������ؑ֏����𐳏�I�����܂����B"
            lngOption = vbOKOnly + vbInformation        '�u�n�j�v�{�^���A�u���v�A�C�R��
            strTitle = "���C�^�������ؑ֌���"

        Case FDUNKAI.CHANGE_ERR     '���C�^�������ؑ֏����@�ُ�I��
            strMessage = "���C�^�������ؑ֏������ُ�I�����܂����B"
            lngOption = vbOKOnly + vbInformation        '�u�n�j�v�{�^���A�u���v�A�C�R��
            strTitle = "���C�^�������ؑ֌���"

        Case Else
    End Select

    If lngOption <> 0 Then
        '���b�Z�[�W�{�b�N�X��\�����A�߂�l��Function�̖߂�l�Ƃ���B
        fMessageBox = MsgBox(strMessage, lngOption, strTitle)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fSendMail
'//  �@�\����  : �w��v���Z�X�Ƀ��[���𑗐M������
'//  �@�\�T�v  : �w�胁�[���X���b�g�A���[��ID�ō쐬���M���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      :String�@�@MailSlot�@[IN]���[���X���b�g��
'//             Long      MailID    [IN]���M���[��ID
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSendMail(MailSlot As String, MailID As Long) As Boolean
    Dim lngMSlot    As Long             '���M���[���X���b�g�n���h��
    Dim lngRet      As Long             '�߂�l
    Dim udtMail     As MAIL_JIKI_UNKAI  '���C�^���p���[�����M�G���A

    On Error Resume Next

    fSendMail = False
    
    '���M���[���f�[�^�쐬
     udtMail.mlHeader.dwId = MailID              '���[���h�c�F����
     udtMail.mlHeader.dwSize = Len(udtMail)      '���[���T�C�Y
     udtMail.mlHeader.dwProid = RHOSHU_ID        '���M���v���Z�X�h�c�F�ێ�
     udtMail.mlHeader.dwSubArea = 0              '�⏕���F�O�i�Œ�j
     Select Case MailID                          '�f�[�^���F���[���h�c�œ��e�ݒ�
        Case ML_ID_HOSHU_UNKAI_DAYCHG_REQ        '�ێ�^�������ؑ֗v��
             udtMail.dwData = MlUnkaiJikiIC.ML_DT_UNKAI_JIKI     '�f�[�^��F���C
        Case ML_ID_KAN_PW_OFF_REQ                '�Ď��Փd���n�e�e�v��
             udtMail.dwData = Ml_SyoriType.ML_DT_REBOOT          '������ʁF���u�[�g
        Case Else
     End Select

     '���[�����M
      lngRet = DssSendMail(MailSlot, Len(udtMail), udtMail.mlHeader)
      If lngRet = False Then
         Select Case MailID                          '�f�[�^���F���[���h�c�œ��e�ݒ�
             Case ML_ID_HOSHU_UNKAI_DAYCHG_REQ        '�ێ�^�������ؑ֗v��
                  '�u���C�^���o�[�W�����Ǘ��F���[�����M�ُ�v���O�o��
                  Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_UNKAIKIRIKAE_CMD_SEND, 0)
             Case ML_ID_KAN_PW_OFF_REQ                '�Ď��Փd���n�e�e�v��
                  '�u���C�^���o�[�W�����Ǘ��F���[�����M�ُ�v���O�o��
                  Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
             Case Else
         End Select
      Else
         Select Case MailID                          '�f�[�^���F���[���h�c�œ��e�ݒ�
             Case ML_ID_HOSHU_UNKAI_DAYCHG_REQ        '�ێ�^�������ؑ֗v��
                  '�u���C�^���o�[�W�����Ǘ��F���[�����M����v���O�o��
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_UNKAIKIRIKAE_CMD_SEND, 0)
             Case ML_ID_KAN_PW_OFF_REQ                '�Ď��Փd���n�e�e�v��
                  '�u���C�^���o�[�W�����Ǘ��F���[�����M����v���O�o��
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
             Case Else
         End Select
         '���[�����M����I��
        fSendMail = True
      End If

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v  : ���[����M�������s���B
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
    Dim iResponse As Integer         'MsgBox�{�^���R�[�h

    On Error Resume Next

    '���[������M����B
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
        '���[����M�������s��
        Select Case udtReadMail.udtlHeader.dwId         '���[���h�c
            Case ML_ID_HOSHU_ACTIVE_REQ                 '�ێ�A�N�e�B�u�v��
               '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                AppActivate frmJikiUnkaiFD.Caption, False
            
            Case ML_ID_HOSHU_UNKAI_DAYCHG_INF               '�ێ�^�������ؑ֒ʒm
               '�u�ێ�^�������ؑ֒ʒm��M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_UNKAIKIRIKAE_REQ_RECV, 0)
                '�f�[�^��ʂ��g���C"�̎������A�������s��
                If udtReadMail.lngData(0) = MlUnkaiJikiIC.ML_DT_UNKAI_JIKI Then
                    If udtReadMail.lngData(1) = MlUnkaiKekka.ML_DT_UNKAI_NORMAL Then
                        iResponse = fMessageBox(FDUNKAI.CHANGE_OK)      '����I�����b�Z�[�W�\��
                    Else
                        iResponse = fMessageBox(FDUNKAI.CHANGE_ERR)     '�ُ�I�����b�Z�[�W�\��
                    End If
                    '��ʂ̃{�^���������\�ɂ���
                    Call sButtonEnabled(True)
                End If

            Case ML_ID_PROEND_ORD                       '�v���Z�X�I���w���̏ꍇ
               '�u�v���Z�X�I���w����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                '�����I���������s��
                pfAbortProc
            Case Else
        End Select
  End If
End Sub
'V1.16.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fGetFilName
'//  �@�\����  : �^���t�@�C�����擾����
'//  �@�\�T�v  : INI�t�@�C����蓊���t�@�C�������擾
'//
'//              �^        ����      �Ӗ�
'//  ����      : String   sSecName  [IN]�擾�Z�N�V������
'//  �@�@      : String   sKeyName  [IN]�擾�L�[��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :String�@�@�@�@�@�@�@[OUT]�擾�t�@�C����(����)
'//                                      �u�����N(�ُ�)
'//
'//     ORIGINAL :(1.16.0.1) 2009-12-20  REVISED BY [TCC] S.Terao
'//                 �s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fGetFilName(sSecName As String, sKeyName As String) As String

    Dim lSts As Long                                 '�֐��߂�l
    Dim strFileName As String * MAX_PATH_SIZE     '�擾�t�@�C����
    Dim lngErrCode As Long
    
    On Error Resume Next
  
    '���C�^��.ini���A�����^���f�[�^�t�@�C�������擾����B
    strFileName = ""
    lSts = GetPrivateProfileString(sSecName, _
                                   sKeyName, _
                                   DEFAILT, _
                                   strFileName, _
                                   Len(strFileName), _
                                   JIKIUNCHIN_FILE)
    If lSts > 0 Then
       fGetFilName = Left$(strFileName, lSts)
    Else
      '�u�o�[�W�����Ǘ����(���C�^��)�F�^���t�@�C�����擾�ُ�(INI�t�@�C���ǂݍ��ُ݈�)�v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_INI_READ, lngErrCode)
      fGetFilName = ""
    End If
End Function
'V1.16.0.1 ADD END
