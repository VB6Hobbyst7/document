VERSION 5.00
Begin VB.Form frmRYTSyusyuOutPut 
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
Attribute VB_Name = "frmRYTSyusyuOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmRYTSyusyuOutPut.frm
'//  �p�b�P�[�W���F�q�x�s���O�f�[�^�o�͒����
'//
'//  �T�v�F�q�x�s���O�f�[�^�o�͒����
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�q�x�s���O�f�[�^�o�͒���ʒǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �q�x�s���O�f�[�^�o�͒����(�A�N�e�B�u��)
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
    
    On Error Resume Next
  
    '�o�͒��̃K�C�h��\������
    lblMessage(0) = "�q�x�s���O�f�[�^���o�͒��ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    tmrMail.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �q�x�s���O�f�[�^�o�͒����(���[�h��)
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
'//  �@�\����  : �q�x�s���O�f�[�^�o�͒����(�f�B�A�N�e�B�u��)
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
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    On Error Resume Next
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSyusyuOutPut.Caption, False
    End If
    
    '�o�̓t�@�C���쐬�������s���B
    sOutPutRYTLog_Data

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sOutPutRYTLog_Data
'//  �@�\����  : �ێ�f�[�^�R�s�[����
'//  �@�\�T�v  : �ێ�f�[�^�R�s�[���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@sOutPath�@[IN]�o�͐�p�X
'//              String  sFromPath [IN]�R�s�[���p�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function sOutPutRYTLog_Data()
   
   On Error GoTo Error_RYT_OutPut
 
   FileCopy m_sCopyMoto, m_sCopySaki
   
   sOutPutEnd (0)
  
  Exit Function
   
Error_RYT_OutPut:
 '�t�@�C���R�s�[�ُ�
 sOutPutEnd (1)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sOutPutEnd
'//  �@�\����  : �o�͌��ʕ\������
'//  �@�\�T�v  : �q�x�s���O�f�[�^�o�͌��ʂ̌��ʕ�����\������B
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
Private Sub sOutPutEnd(iEnd As Integer)
    Dim i As Integer       '�J�E���^
    Dim lngErrCode As Long '�G���[�R�[�h

    On Error Resume Next
        
    If iEnd = 0 Then
       '����I�����̕�����\������B
       lblMessage(0) = "����I�����܂����B"
       lblMessage(1) = ""
       '�u�q�x�s���O�f�[�^�o�͏�������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RYT_LOG_KANRI_GAMEN_OUTPUT_OK, 0)
    Else
       '���W���s���̕�����\������B
       lblMessage(0) = "�ُ�I�����܂����B"
       lblMessage(1) = ""
       '�u�q�x�s���O�f�[�^�o�͗��ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RYT_LOG_KANRI_GAMEN_OUTPUT_ERROR, lngErrCode)
    End If
    cmdOK.Visible = True
    cmdOK.Enabled = True
End Sub

