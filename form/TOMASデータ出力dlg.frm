VERSION 5.00
Begin VB.Form frmTomasDataOut 
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Timer tmrOutput 
      Left            =   480
      Top             =   0
   End
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
      TabIndex        =   0
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
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmTomasDataOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmTomasDataOut.frm
'//  �p�b�P�[�W���FTOMAS�f�[�^�}�̏o�͉��
'//
'//  �T�v�F�o�[�W�����Ǘ����
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//                 �V�K�쐬�y�t�F�[�Y�R TOMAS�Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////

Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
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
'//  �֐�����  : Form_Activate
'//  �@�\����  : TOMAS�f�[�^�\�����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
    
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TOMAS_DATA_DISP)

    '�������̃K�C�h��\������
    lblMessage(0) = "�o�͒��ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    cmdOK.Enabled = False
    DoEvents
    
    tmrMail.Enabled = True
    tmrOutput.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : TOMAS�f�[�^�\�����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next
    
    '�^�C�}���~�߂�
    tmrMail.Enabled = False
    tmrOutput.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : TOMAS�f�[�^�\�����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '�o�͏����J�n�p�^�C�}�̒l��ݒ肷��B
    tmrOutput.Interval = 100
    tmrOutput.Enabled = False
    
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
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
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                '�v���Z�X�̏I���������s��
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '�u�ێ��ʃA�N�e�B�u�\���v����M�����ꍇ
                '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '�\������ʁi�ێ�f�[�^���W��ʁj���A�N�e�B�u�\������B
'                AppActivate frmRenewOutput.Caption, False  ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
                AppActivate frmTomasDataOut.Caption, False  ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
                pfFormActive (frmTomasDataOut.hwnd)         ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
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
'//  �֐�����  : tmrOutput_Timer
'//  �@�\����  : �o�͏������s�^�C�}
'//  �@�\�T�v  : �}�̏o�͂���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrOutput_Timer()

    On Error Resume Next
    
    tmrOutput.Enabled = False
    Call sOutput_Data
     
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sOutput_Data
'//  �@�\����  : �ݒ�l�o�͏���
'//  �@�\�T�v  : �ݒ�l��ҏW���Ĕ}�̏o�͂���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sOutput_Data()

    Dim bySyoAssort As Byte             '���O�p������
    Dim strFilePath As String           '�o�̓t�@�C���p�X
    Dim strCornerPath As String         '�ݒ�t�@�C���p�X
    Dim strStationNm As String          '�w��
    Dim strCornerNm As String           '�R�[�i��
    Dim intCount As Integer             '�J�E���^
    Dim intCount2 As Integer            '�J�E���^
    Dim intOutFile As Integer           '�o�̓t�@�C���ԍ�
    Dim intTgtFileNo As Integer         '�o�͑Ώېݒ�t�@�C���ԍ�
    Dim strTgtFileName As String        '�o�͑Ώېݒ�t�@�C��
    Dim strTargetFile() As String       '�o�͑Ώۃt�@�C��
    Dim intFileNum As Integer
    Dim strDefault As String
    Dim strRet As String * 32
    Dim lngRet As Long
    Dim strOutFileName As String
    Dim strFileName As String
    Dim strCabTarget As String
    Dim lngRetZip As Long
    Dim objFileObj As FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
    Const lngBufSize = 32
    
    On Error GoTo Err_Handler
    
    Set objFileObj = New FileSystemObject
    
    Select Case gintTomasDataDispDiv
    Case TOMAS_DISP_DIV.TOMAS_DATA_VERSION
        strFileName = TOMAS_FILE_VERINFO
        
    Case TOMAS_DISP_DIV.TOMAS_DATA_KIKI
        strFileName = TOMAS_FILE_KIKIINFO
        
    Case TOMAS_DISP_DIV.TOMAS_DATA_ERR
        strFileName = TOMAS_FILE_ERRINFO
    Case Else
    End Select
    
    strOutFileName = gstrOutPath & strFileName
    
    '�o�͑Ώېݒ�t�@�C�������݂��Ȃ��ꍇ�ُ͈�I��
    If objFileObj.FileExists(PATH_WORK & strFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strFileName, 0)
        GoTo Err_Handler
    End If
    
    'TOMAS�f�[�^�e�L�X�g�t�@�C�����R�s�[����
    Call objFileObj.CopyFile(PATH_WORK & strFileName, strOutFileName, True)
    
    Set objFileObj = Nothing
    
    lblMessage(0).Caption = "����I�����܂����B"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    DoEvents
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    Exit Sub
    
'�G���[����
Err_Handler:

    Set objFileObj = Nothing
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, bySyoAssort, KAKARISET_OUTPUT_ERR, 0)
    
    lblMessage(0).Caption = "�ُ�I�����܂����B"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
End Sub


