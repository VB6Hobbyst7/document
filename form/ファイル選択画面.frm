VERSION 5.00
Begin VB.Form frmFil 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�t�@�C���I�����"
   ClientHeight    =   2715
   ClientLeft      =   3795
   ClientTop       =   4860
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   2715
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelected 
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
      Height          =   1095
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.FileListBox filSelection 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   2520
      Pattern         =   "*.exe;*.com;*.bat;*.cmd"
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelected 
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
      Height          =   1095
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.DirListBox dirSelection 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.DriveListBox drvSelection 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblFileSelection 
      Caption         =   "���s�t�@�C���I��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmFil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmFil.frm
'//  �p�b�P�[�W���F�t�@�C���I�����
'//
'//  �T�v�F�t�@�C���I�����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l 'V1.3.0.1 ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �t�@�C���I�����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
'//  �@�\����  : �t�@�C���I�����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}��~
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
'//  �@�\����  : �t�@�C���I�����(���[�h��)
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
   '�u�t�@�C���I����ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, FIR_GAMEN_START, 0)

    lblFileSelection.Caption = "����������������"
    Me.filSelection.Pattern = "*.XXX"

    '�\���ʒu��ݒ肷��
    Me.Move Screen.Width - Me.Width, 0
    
    dirSelection.Path = drvSelection.Drive
    
    'V1.3.0.1 ADD START
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdSelected_Click
'//  �@�\����  : �e�t����������
'//  �@�\�T�v  : �e�t���̂̏������s���B�u�m��v�u����v
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdSelected_Click(Index As Integer)
        
On Error Resume Next
    Select Case Index
    Case 0
    '�u�m��v�{�^�������̏ꍇ
        '�\���t�@�C���w��̃`�F�b�N���s��
        If filSelection.ListIndex = -1 Then
            '�G���[���b�Z�[�W��\������
            MsgBox "�t�@�C�����I������Ă��܂���B" _
                   & Chr(vbKeyReturn) & "�I�����Ă��������B", _
                   vbOKOnly + vbExclamation, _
                   "�t�@�C���I��"                  '���s�t�@�C�����t�@�C���B
            Exit Sub
        End If
        '�t�@�C�������O���[�o���G���A�ɃZ�b�g����
        gstrMyPath = IIf(Len(dirSelection.Path) = 3, dirSelection.Path _
                 & filSelection.List(filSelection.ListIndex), dirSelection.Path & "\" _
                 & filSelection.List(filSelection.ListIndex))
    Case 1
    '�u����v�{�^�������̏ꍇ
        '�t�@�C�����Ȃ����Z�b�g����B
        gstrMyPath = ""
    End Select
   
    '�u�t�@�C���I����ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, FIR_GAMEN_END, 0)

    '����ʂ������B
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : dirSelection_Change
'//  �@�\����  : ���X�g�{�b�N�X���e�X�V�����@
'//  �@�\�T�v  : ���X�g�{�b�N�X���̍X�V���s���B
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
Private Sub dirSelection_Change()
On Error Resume Next
    ' �t�@�C���p�X��ݒ肷��B
    filSelection.Path = dirSelection.Path
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : dirSelection_Change
'//  �@�\����  : ���X�g�{�b�N�X���e�X�V�����A
'//  �@�\�T�v  : ���X�g�{�b�N�X���̍X�V���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-17  REVISED BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub drvSelection_Change()
On Error GoTo Drive_Error
    ' �f�B���N�g���p�X��ݒ肷��B
    dirSelection.Path = Left$(drvSelection.Drive, 2) & "\"
    Exit Sub
Drive_Error:
'    If Left$(drvSelection.Drive, 1) = "a" Then         'V1.12.0.1 DEL
    If Left$(drvSelection.Drive, 1) = "H" Then          'V1.12.0.1 ADD
    'a:�h���C�u���ُ�Ȃ�A�J�����g�h���C�u��\��������B
        drvSelection.Drive = Left$(App.Path, 2)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : dirSelection_Change
'//  �@�\����  : ���[����M�p�^�C�}���^�C���A�b�v����
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
'///////////////////////////////////////////////////////////////////*
Private Sub tmrMail_Timer()
On Error Resume Next
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmFil.Caption, False
    End If
End Sub
