VERSION 5.00
Begin VB.Form frmSyusyuOutPut 
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
   Begin VB.Timer tmrMail2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
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
Attribute VB_Name = "frmSyusyuOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmSyusyuOutPut.frm
'//  �p�b�P�[�W���F�ێ�f�[�^�o�͒����
'//
'//  �T�v�F�ێ�f�[�^�o�͒����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 �t�@�C���I�������Ή�
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.265�C���Ή��z
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-13  CODED BY  [TCC] H.Sugimoto
'//                 �y�R�[�i���X�y�[�X�����Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  REVISED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi06_005_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �ێ�f�[�^�o�͒����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intCount As Integer
    Dim blnSelected As Boolean
    
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
    'EG20 V2.1.0.1 ADD END
    
    cmdOK.Enabled = False
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_KBN_KADO_MAINTE)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
'    �o�͒��̃K�C�h��\������
    lblMessage(0) = "�ێ�f�[�^���o�͒��ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    tmrMail.Enabled = True
    tmrMail2.Enabled = True                  ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �ێ�f�[�^�o�͒����(���[�h��)
'//  �@�\�T�v  : �����������s���B
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
Private Sub Form_Load()
  On Error Resume Next
  '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
  tmrMail.Interval = MN_MAIL_INTERVAL
  tmrMail.Enabled = False
  
  tmrMail2.Interval = MN_MAIL_INTERVAL       ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
  tmrMail2.Enabled = False                   ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
    tmrMail2.Enabled = False                   ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
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
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
' V1.12.0.1 ADD START
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
' V1.12.0.1 ADD END
' V1.3.0.1 ADD START
    On Error Resume Next
' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL START
'    '�ėp���C����M�������s��
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
'        AppActivate frmSyusyuOutPut.Caption, False
'    End If
' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL END
' V1.3.0.1 ADD END
     '�o�̓t�@�C���쐬�������s���B
    sOutPutHoshuData

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
End Sub
' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail2_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v������
'//  �@�\�T�v  : ���[������M����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iEnd�@�@�@[IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()

    On Error Resume Next
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSyusyuOutPut.Caption, False
        pfFormActive (frmSyusyuOutPut.hwnd)
    End If

End Sub
' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSyusyuEnd
'//  �@�\����  : �o�͌��ʕ\������
'//  �@�\�T�v  : �ێ�f�[�^�o�͌��ʂ̌��ʕ�����\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iEnd�@�@�@[IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSyusyuEnd(iEnd As Integer)
    Dim i As Integer       '�J�E���^
    Dim lngErrCode As Long '�G���[�R�[�h

    On Error Resume Next
    
    Sleep (5000)
    
    If iEnd = 0 Then
       '����I�����̕�����\������B
       lblMessage(0) = "����I�����܂����B"
       lblMessage(1) = ""
       '�u�ێ�f�[�^�o�͏�������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_OUTPUT_OK, 0)
    Else
       '���W���s���̕�����\������B
       lblMessage(0) = "�ُ�I�����܂����B"
       lblMessage(1) = ""
       '�u�ێ�f�[�^�o�͗��ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KADO_MENTE_SYUSYU_GAMEN_OUTPUT_ERROR, lngErrCode)
    End If
    cmdOK.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sOutPutHoshuData
'//  �@�\����  : �ێ�f�[�^�o�͏���
'//  �@�\�T�v  : �ێ�f�[�^�o�͂��s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iEnd�@�@�@[IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 �t�@�C���I�������Ή�
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.265�C���Ή��z
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-23  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�o�b�N�A�b�v�t�@�C���Ή��z
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi06_005_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sOutPutHoshuData()
    Dim iIniKeka As Integer '�߂�l
    Dim iGate As Integer    '�쐬���@��
    Dim iCreate As Integer  '�쐬�t�@�C����
    Dim iCnt As Integer     '�J�E���^�[�P
    Dim i As Integer        '�J�E���^�[�Q
    Dim sMyFromPath As String '�쐬���t�@�C����
    Dim sMyToPath As String   '�쐬��t�@�C����
    Dim bRet As Boolean
    Dim sFromCreateFileName As String * MAX_PATH_SIZE    '�쐬���t�@�C����
    Dim sToCreateFileName As String * MAX_PATH_SIZE      '�쐬��t�@�C����
    Dim sFormatCreateFileName As String * MAX_PATH_SIZE  '�쐬�t�H�[�}�b�g�t�@�C����
    Dim sFormatCreateFileNameKan As String * MAX_PATH_SIZE  '�쐬�t�H�[�}�b�g�t�@�C�����i�����p�j   'EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�z ADD
    Dim sOut_Path As String * MAX_PATH_SIZE              '�쐬�t�H�[�}�b�g�t�@�C����
    Dim iRet As Integer
    Dim iNoFileCnt As Integer           ' V1.3.0.1 ADD
    Dim bErrChk As Boolean              ' V1.8.0.1 ADD�@�@'�o�̓t�@�C���쐬�G���[�`�F�b�N�t���O
    Dim nCorner As Integer              ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
    Dim nCornerGoki As Integer          ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
    Dim iStatus As Integer                              ' ��������          ' EG20 V5.4.0.1�y�o�b�N�A�b�v�t�@�C���Ή��z�ǉ�
    Dim sBackupFolderName As String * MAX_PATH_SIZE     '�쐬���t�H���_��   ' EG20 V5.4.0.1�y�o�b�N�A�b�v�t�@�C���Ή��z�ǉ�
      
    On Error Resume Next
    
    sOut_Path = ""
    iCreate = 0
    iGate = 0
    
    bErrChk = True                      'V1.8.0.1 ADD
    
    ' EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�z ADD START
    ' �e�R�[�i�̎�ʂ��擾����
    gsGetCornerType
    ' EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�z ADD END
    
    
    '�쐬�t�@�C�������擾����B
     iCreate = GetPrivateProfileInt(HOSHUPUT_FROM_SECTION_NAME, _
                                    HOSHUPUT_FROM_NUMBER_KEY_NAME, DEFAILT_Int, PATH_HOSHU_DATA_FILE)
    '�쐬���@�����擾����B
     iGate = GetPrivateProfileInt(HOSHUPUT_FROM_SECTION_NAME, _
                                    HOSHUPUT_FROM_GATE_NUMBER_KEY_NAME, DEFAILT_Int, PATH_HOSHU_DATA_FILE)
    
    '�R�s�[����擾����B
'V1.12.0.1 DEL START
'     iIniKeka = GetPrivateProfileString(KANSI_OUT_HOSHU_SEC, _
'                                        KANSI_OUT_HOSHU_KEY, DEFAILT, _
'                                        sOut_Path, Len(sOut_Path), _
'                                        HOSHU_FILE)
'V1.12.0.1 DEL END

'    MkDir sOut_Path                         'V1.12.0.1 DEL
     MkDir frmSyusyu.glbFilePath             'V1.12.0.1 ADD
    
     For iCnt = 1 To iCreate
         sFromCreateFileName = ""
         sToCreateFileName = ""
         sFormatCreateFileName = ""
         '�쐬���t�@�C�������擾����B
         iIniKeka = GetPrivateProfileString(HOSHUPUT_FROM_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sFromCreateFileName, Len(sFromCreateFileName), _
                                            PATH_HOSHU_DATA_FILE)
 'V1.7.0.1 DEL START
'        '�쐬��t�@�C�������擾����B
'         iIniKeka = GetPrivateProfileString(HOSHUPUT_TO_SECTION_NAME, _
'                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
'                                            sToCreateFileName, Len(sToCreateFileName), _
'                                            PATH_HOSHU_DATA_FILE)
 'V1.7.0.1 DEL END
        '�쐬�t�H�[�}�b�g�t�@�C�������擾����B
         iIniKeka = GetPrivateProfileString(HOSHUPUT_FORMAT_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sFormatCreateFileName, Len(sFormatCreateFileName), _
                                            PATH_HOSHU_DATA_FILE)
' EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�zADD START
         iIniKeka = GetPrivateProfileString(HOSHUPUT_FORMAT_KAN_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sFormatCreateFileNameKan, Len(sFormatCreateFileNameKan), _
                                            PATH_HOSHU_DATA_FILE)
' EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�zADD END

' EG20 V5.4.0.1�y�o�b�N�A�b�v�t�@�C���Ή��z�ǉ��J�n
        '�쐬�t�H�[�}�b�g�t�@�C�������擾����B
        iStatus = GetPrivateProfileString(HOSHUPUT_BACKUPFOLDER_SECTION_NAME, _
                                            HOSHUPUT_KEY_NAME & iCnt, DEFAILT, _
                                            sBackupFolderName, Len(sBackupFolderName), _
                                            PATH_HOSHU_DATA_FILE)

' EG20 V5.4.0.1�y�o�b�N�A�b�v�t�@�C���Ή��z�ǉ��I��
            
        iNoFileCnt = 0                     ' V1.3.0.1 ADD
        For i = 1 To iGate
            
            If gintStatus(i - 1) = TAG_STATUS.STS_SENTAKU Then      'EG20 V2.1.0.1 ADD �yMainte_03_01�z
                'V1.7.0.1 ADD START
                '������INI�t�@�C���擾����
'                sToCreateFileName = fGetGateInfoPath(i, iCnt, iIniKeka)        ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�폜
                sToCreateFileName = fGetGateInfoPath(i, iCnt, iIniKeka, _
                                                    nCorner, nCornerGoki)       ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
                If sToCreateFileName <> "" And iIniKeka <> 0 Then
                'V1.7.0.1 ADD END
                    sMyFromPath = ""
                    sMyToPath = ""
                    '�u##�v��01�`32�ɕϊ�����B
                    sMyFromPath = Replace(sFromCreateFileName, "##", Format(i, "0#"))
'                    sMyToPath = Replace(sToCreateFileName, "##", Format(i, "0#"))              ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�폜
                    ' �����Ď��՘_�����@�ԍ����R�[�i�ʘ_�����@�ԍ�
                    sMyToPath = Replace(sToCreateFileName, "##", Format(nCornerGoki, "0#"))     ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
                    'iRet = dllCreateCSVDataFile(sMyToPath, sFormatCreateFileName, sMyFromPath) ' EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�zDEL
                    
' EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�z ADD START
                    '���ݏ������Ă��鍆�@��������R�[�i��ʂɂ��DLL�֐��ɓn���t�H�[�}�b�g�t�@�C������؂�ւ���B
                    If gintCornerType(nCorner - 1) = CORNER_TYPE_KANSEN Then
                        '���̍��@�������R�[�i�̏ꍇ
                        iRet = dllCreateCSVDataFile(sMyToPath, sFormatCreateFileNameKan, sMyFromPath)
                    Else
                        '���̍��@���ݗ��R�[�i�̏ꍇ
                        iRet = dllCreateCSVDataFile(sMyToPath, sFormatCreateFileName, sMyFromPath)
                    End If
' EG20 V30.3.0.1 �yHKRK_Kansi06_005_01�z ADD END
                    If iRet = 0 Then
                        sSyusyuEnd (1)
                        Exit Sub
                    End If
    ' V1.3.0.1 DEL START
    '             If iRet = 2 And iGate = i And iCnt = iCreate Then
    '                sSyusyuEnd (1)
    '                Exit Sub
    '             End If
    ' V1.3.0.1 DEL START
    ' V1.3.0.1 ADD START
                    If iRet = 2 Then
                        '�쐬���t�@�C���Ȃ��ُ�
                        iNoFileCnt = iNoFileCnt + 1
                        'If iGate = iNoFileCnt And iCnt = iCreate Then  'V1.8.0.1 DEL
                        bErrChk = False                                 'V1.8.0.1  ADD
                        If i = iNoFileCnt And iCnt = iCreate Then      'V1.8.0.1 ADD
                            sSyusyuEnd (1)
                            Exit Sub
                        End If
                    End If
    ' V1.3.0.1 ADD END
                    If iRet = 1 Then
        '                bRet = HoshuCopy(sOut_Path, sMyToPath)             'V1.12.0.1 DEL
'                        bRet = HoshuCopy(frmSyusyu.glbFilePath, sMyToPath)  'V1.12.0.1 ADD
'                        bRet = HoshuCopy(frmSyusyu.glbFilePath, sMyToPath, nCorner)        ' V5.4.0.1�폜
                        bRet = HoshuCopy(frmSyusyu.glbFilePath, sMyToPath, nCorner, sBackupFolderName)
                        If False = bRet Then
                            sSyusyuEnd (1)
                            Exit Sub
                        End If
                    End If
                End If 'V1.7.0.1 ADD
            End If          'EG20 V2.1.0.1 ADD �yMainte_03_01�z
         Next
     Next
     
     'V1.8.0.1 ADD START
     If bErrChk = False Then
        sSyusyuEnd (1)
        Exit Sub
     End If
     'V1.8.0.1 ADD END
     
     sSyusyuEnd (0)
         
End Sub
'V1.7.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fGetGateInfoPath
'//  �@�\����  : �����ʃt�@�C���p�X�擾����
'//  �@�\�T�v  : �����ʃt�@�C���p�X�擾�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iGouki�@�@[IN]���@�ԍ�
'//              Integer�@iFilType�@[IN]�쐬�t�@�C�����
'//              Integer�@iIniKeka�@[OUT]�擾������
'//              Integer  nCorner      [OUT]�R�[�i�@�ԍ�
'//              Integer  nCornerGoki  [OUT]�R�[�i�_�����@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String�@�@�@�@�@�@[OUT]�t�@�C���p�X
'//
'//     ORIGINAL :(1.7.0.1) 2009-07-28   CODED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS:(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.265�C���Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fGetGateInfoPath(iGouki As Integer, iFilType As Integer, iIniKeka As Integer, _
                                  nCorner As Integer, nCornerGoki As Integer) As String
    Dim lngRet As Long          '�֐��̕Ԃ�l
    Dim iGate As Integer        '����INDEX
    Dim j As Integer            '���[�NINDEX
    Dim cWork As Byte           '���[�N�G���A
    Dim lngErrCode As Long      '�G���[�R�[�h
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim sToCreateFileName As String * MAX_PATH_SIZE      '�쐬��t�@�C����
 
    On Error Resume Next
    
    '�������D�@���擾
    sKeyName = "gate" & Format(iGouki, "00")
    iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                    sKeyName, _
                                    DEFAILT, sGateData, Len(sGateData), _
                                    PATH_GATE_FILE)
    If iRet = 0 Then
       '�u�Ӱ�����ݽ��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
       fGetGateInfoPath = ""
       iIniKeka = 0
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
     
    If Trim(sFData(4)) = MISETI Then
        '���ݒu�̏ꍇ
        fGetGateInfoPath = ""
        iIniKeka = 0
        nCorner = 0                                         ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
        nCornerGoki = 0                                     ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
        Exit Function
    'EG20 V2.1.0.1 DEL START
'     ElseIf Trim(sFData(4)) = EGR Then
'        'EG-R�����̏ꍇ
'        sKeyName = HOSHUPUT_KEY_E_NAME
'     ElseIf Trim(sFData(4)) = NEG Then
'        'NEG�����̏ꍇ
'        sKeyName = HOSHUPUT_KEY_N_NAME
'     End If
    'EG20 V2.1.0.1 DEL END
    'EG20 V2.1.0.1 ADD START
    Else
        sKeyName = HOSHUPUT_KEY_NAME
    End If
    'EG20 V2.1.0.1 ADD END
     sToCreateFileName = ""
     iIniKeka = GetPrivateProfileString(HOSHUPUT_TO_SECTION_NAME, _
                                        sKeyName & iFilType, DEFAILT, _
                                        sToCreateFileName, Len(sToCreateFileName), _
                                        PATH_HOSHU_DATA_FILE)
     If iIniKeka = 0 Then
       fGetGateInfoPath = ""
     Else
       fGetGateInfoPath = sToCreateFileName
     End If

    nCorner = CInt(sFData(GATE_IDX.IDX_RONRI_CORNER))     ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
    nCornerGoki = CInt(sFData(GATE_IDX.IDX_RONRI_GOKI))   ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�

End Function
'V1.7.0.1 ADD END
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : HoshuCopy
'//  �@�\����  : �ێ�f�[�^�R�s�[����
'//  �@�\�T�v  : �ێ�f�[�^�R�s�[���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@sOutPath�@  [IN]�o�͐�p�X
'//              String  sFromPath   [IN]�R�s�[���p�X
'//              Integer nCorner     [IN]�R�[�i�ԍ�
'//              String  sBackupPath [IN]�o�b�N�A�b�v�p�X   ' EG20 V5.4.0.1�ǉ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//                 �t�@�C���R�s�[�����^���s���f�ǉ�
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.265�C���Ή��z
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-23  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-13  CODED BY  [TCC] H.Sugimoto
'//                 �y�R�[�i���X�y�[�X�����Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Public Function HoshuCopy(sOutPath As String, sFromPath As String)                     ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�폜
' EG20 V5.4.0.1�폜�J�n
'Public Function HoshuCopy(sOutPath As String, sFromPath As String, nCorner As Integer)  ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
' EG20 V5.4.0.1�폜�I��
' EG20 V5.4.0.1�ǉ��J�n
Public Function HoshuCopy(sOutPath As String, sFromPath As String, nCorner As Integer, sBackupPath As String)
' EG20 V5.4.0.1�ǉ��I��

    Dim fso         As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim sCopyfile As String                 '�R�s�[��
    Dim FileName As String                  '���o�t�@�C����
    Dim FileKaku As String                  '�g���q
    Dim bRet As Boolean                     '�߂�l
    Dim strCorner As String                 ' �R�[�i��          ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�
    Dim szUnyouDate As String               ' �^�p���t
    Dim szBackupFolder As String            ' �o�b�N�A�b�v�t�H���_�̃p�X
    Dim nNullIndex As Integer               ' ���������[�N
    
    
'    On Error Resume Next       'V1.12.0.1 DEL
    On Error GoTo COPY_ERR        'V1.12.0.1 ADD
    
    HoshuCopy = False

    '�R�s�[��t�H���_�̗L���m�F
    If fso.FolderExists(sOutPath) = False Then
        '�R�s�[��t�H���_�쐬
        fso.CreateFolder (sOutPath)

    End If

' EG20 V6.1.0.1 �폜�J�n
'' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ��J�n
'    strCorner = gstrCornerName(nCorner - 1)
'' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ��I��
' EG20 V6.1.0.1 �폜�I��
' EG20 V6.1.0.1 �ǉ��J�n
    strCorner = Replace(gstrCornerName(nCorner - 1), " ", "")
' EG20 V6.1.0.1 �ǉ��I��

    '̧�ٖ��O�擾
    psFileNameGet sFromPath, FileName, FileKaku

    '�R�s�[��t�@�C�����쐬
'    sCopyfile = sOutPath & "\" & FileName & "." & FileKaku                 ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�폜
    sCopyfile = sOutPath & "\" & strCorner & FileName & "." & FileKaku      ' EG20 V3.4.0.1�y����TR-No.265�C���Ή��z�ǉ�

    '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
    fso.CopyFile sFromPath, sCopyfile, True
    
    HoshuCopy = True
    
    Set fso = Nothing

' EG20 V5.4.0.1�y�o�b�N�A�b�v�t�@�C���Ή��z�ǉ��J�n
    ' �o�b�N�A�b�v�t�@�C���̍쐬����
    ' �����́��^�p���t
    ' �����́��o�b�N�A�b�v�t�H���_�i�p�X�j
    ' �����́����̓t�@�C�����i�e�L�X�g�j
    If CheckAppStart(PROC_KANRI) <> 0 Then
        Set Idinf_KansiSettei = New IdInfProc             '�Ď����u�ݒ�G���A
        '�Q��(�����ʐM���)�G���A����ݒ�
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
            Exit Function
        End If
    
        '�G���AID�̐ݒ�l���擾
        Idinf_KansiSettei.IdLock
        Idinf_KansiSettei.id = IdKansiSet.SET_ID_KANSI_SET_UNYOU_DAY
        Idinf_KansiSettei.IdGet
        szUnyouDate = Idinf_KansiSettei.DataArea(0)
        Idinf_KansiSettei.IdFree

        nNullIndex = InStr(sBackupPath, Chr(0))
        If nNullIndex <> 0 Then
                szBackupFolder = Left(sBackupPath, nNullIndex - 1)
            Else
                szBackupFolder = sBackupPath
            End If

        If Len(szUnyouDate) > 4 Then
            szBackupFolder = szBackupFolder & Right(szUnyouDate, 4) & "\"
        Else
            szBackupFolder = szBackupFolder & szUnyouDate & "\"
        End If
        ' �o�b�N�A�b�v�t�@�C���쐬����
        Call dllSaveBackupFile(sFromPath, strCorner & FileName, szBackupFolder)
    End If
' EG20 V5.4.0.1�y�o�b�N�A�b�v�t�@�C���Ή��z�ǉ��I��

'V1.12.0.1 ADD START
    Exit Function
    
COPY_ERR:
    HoshuCopy = False
    Set fso = Nothing
'V1.12.0.1 ADD END

End Function


