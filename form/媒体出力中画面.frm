VERSION 5.00
Begin VB.Form frmSyorityu 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�}�̏o�͒�"
   ClientHeight    =   2955
   ClientLeft      =   3420
   ClientTop       =   4800
   ClientWidth     =   6030
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
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   600
      Top             =   600
   End
   Begin VB.Label lblLogMessage 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�}�̏o�͒�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2130
      TabIndex        =   0
      Top             =   1200
      Width           =   1755
   End
End
Attribute VB_Name = "frmSyorityu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmSyorityu.frm
'//  �p�b�P�[�W���F�������̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'�_�C�A���O�\���ʒu
Private Const DIALOGTOP     As Integer = 3495
Private Const DIALOGLEFT    As Integer = 2985
Private Const DIALOGHEIGHT  As Integer = 3375
Private Const DIALOGWIDTH   As Integer = 6165

Private Const MN_MAIL_INTERVAL = 100       '�^�C�}�̃C���^�[�o���l
Private Const MN_MAIL_INTERVAL2 = 1000     '�^�C�}�̃C���^�[�o���l ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���������(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �����N���^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
    
    '�^�C�}���N������
    tmrMail.Enabled = True
    tmrMail2.Enabled = True     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ���������(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�A�^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    '�^�C�}���~����
    tmrMail.Enabled = False
    tmrMail2.Enabled = False     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : ���������(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    '�z�u�ݒ�
    Me.Top = DIALOGTOP
    Me.Left = DIALOGLEFT
    Me.Height = DIALOGHEIGHT
    Me.Width = DIALOGWIDTH
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    tmrMail2.Interval = MN_MAIL_INTERVAL2
    tmrMail2.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : �^�C�}�����i�^�C���A�b�v���F�C�x���g�v���V�[�W���j
'//  �@�\�T�v  : �^�C���A�E�g�������s��
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : Long�@ �@ �T�C�Y         ���[�����M�T�C�Y
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    Dim bRet As Boolean '�֐��߂�l

    On Error Resume Next
    
    '�^�C�}���~����
    tmrMail.Enabled = False
    
    If glShoriNo = SHORI_NO.NO_MEDIUM_OUT Then
        
        '�t�r�a���O������
        bRet = dllEjectUsbDevice(glErrsts)
    ElseIf glShoriNo = SHORI_NO.NO_INSTOL Then
    
        '�w�s�x�f�[�^�}�̃C���X�g�[��
        Call pfTgEkiDataInstol
    End If
        
    '����ʂ������B
    Unload Me

End Sub

' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail2_Timer
'//  �@�\����  : �^�C�}�����i�^�C���A�b�v���F�C�x���g�v���V�[�W���j
'//  �@�\�T�v  : �^�C���A�E�g�������s��
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V8.1.0.1) 2014-06-05  CODED   BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()

    On Error Resume Next

    ' �ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSyorityu.Caption, False
        pfFormActive (frmSyorityu.hwnd)
    End If
    
End Sub
' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD END
