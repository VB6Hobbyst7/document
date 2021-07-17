VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�f�B���N�g���I�����"
   ClientHeight    =   2355
   ClientLeft      =   3960
   ClientTop       =   4425
   ClientWidth     =   5790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   2355
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTorikesi 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   1920
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "�m��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.DirListBox dirSelection 
      Height          =   1770
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.DriveListBox drvSelection 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmDir.frm
'//  �p�b�P�[�W���F�t�H���_(�f�B���N�g��)�I�����
'//
'//  �T�v�F�t�H���_(�f�B���N�g��)�I�����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �EEG10���A�t�H���_�I����ʗ��p�B
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �t�H���_(�f�B���N�g��)�I�����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}�N��
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
'//  �@�\����  : �t�H���_(�f�B���N�g��)�I�����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}��~
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
    '�^�C�}���~����
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �t�H���_(�f�B���N�g��)�I�����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
On Error Resume Next
    '�u�t�H���_�I����ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DIR_GAMEN_START, 0)

    '�\���ʒu��ݒ肷��
    Me.Move Screen.Width - Me.Width, 0

     ' �f�B���N�g���p�X�̐ݒ�B
    dirSelection.Path = Left(App.Path, 3)
    
    'V1.3.0.1 ADD START
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdKakutei_Click
'//  �@�\����  : �u�m�F�v�t����������
'//  �@�\�T�v  : �I�����ꂽ���e��ۑ����A��ʏ����B
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
Private Sub cmdKakutei_Click()
On Error Resume Next
    '�I���f�B���N�g����ݒ肷��B
    '�uc:\�v��ud:\�v�Ƃ��������[�g�h���C�u��ݒ肵���ꍇ�́��}�[�N���폜
    gstrMyPath = IIf(Len(dirSelection.Path) = 3, dirSelection.Path, _
                 dirSelection.Path + "\")
     '�u�t�H���_�I����ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DIR_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdTorikesi_Click
'//  �@�\����  : �u����v�v�t����������
'//  �@�\�T�v  : �I���f�B���N�g���Ȃ���Ԃ�ۑ����A��ʏ����B
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
Private Sub cmdTorikesi_Click()
On Error Resume Next
    '�I���f�B���N�g���Ȃ��A��ݒ肷��B
    gstrMyPath = ""
     '�u�t�H���_�I����ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DIR_GAMEN_END, 0)
    '����ʂ������B
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : drvSelection_Change
'//  �@�\����  : �h���C�u���X�g�{�b�N�X�̓��e�ύX������
'//  �@�\�T�v  : ���X�g�{�b�N�X�̓��e���X�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub drvSelection_Change()
On Error GoTo Drive_Error
     ' �f�B���N�g���p�X�̐ݒ�B
    dirSelection.Path = Left$(drvSelection.Drive, 2) & "\"
    Exit Sub
Drive_Error:
'    If Left$(drvSelection.Drive, 1) = "a" Then     'V1.12.0.1 DEL
    If Left$(drvSelection.Drive, 1) = "H" Then      'V1.12.0.1 ADD
    'a:�h���C�u���ُ�Ȃ�A�J�����g�h���C�u��\��������B
        drvSelection.Drive = Left$(App.Path, 2)
    End If
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
On Error Resume Next
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmDir.Caption, False
        pfFormActive (frmDir.hwnd)
    End If
End Sub
