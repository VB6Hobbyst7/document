VERSION 5.00
Begin VB.Form frmKikiDataMenu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�w�s�x�f�[�^�m�F"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   360
      Top             =   8280
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6360
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�w���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "    ���j���[     ��ʂ֖߂�"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�@��\���ݒ�"
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
Attribute VB_Name = "frmKikiDataMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �F�@��\���ݒ胁�j���[���.frm
'//  �p�b�P�[�W���F�@��\���ݒ胁�j���[��ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000       '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �@��\���ݒ胁�j���[���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �őO�O�\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    '�G���[���[�`����錾
    On Error Resume Next
    
    '����ʍőO�ʕ\���������s���B
    pfFormActive (hwnd)
    
    '�^�C�}���N������
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �@��\���ݒ胁�j���[���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�A�^�C�}��~
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
Private Sub Form_Deactivate()
On Error Resume Next
    '�^�C�}���~����
    tmrMail.Enabled = False
End Sub
'EG20 V2.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �@��\���ݒ胁�j���[���(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIKOUSEISETMENU_GAMEN_START, 0)
    
    '----------------------------------------------------
    '��ʏ����l�ݒ�
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�����i�^�C���A�b�v���F�C�x���g�v���V�[�W���j
'//  �@�\�T�v  : �ėp���C����M�������s��
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmRenewData.Caption, False
        pfFormActive (frmRenewData.hwnd)
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �e�t����������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : Integer�@ Index          �I��t�̃C���f�b�N�X
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    Select Case Index
        
        Case 0                                 '�w���
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKIINFO, 0)
            
            '��ʕ\��
            Load frmKikiData
            frmKikiData.Show 1

        Case 1                                 '����
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_GATE, 0)
            
            '��ʕ\��
            Load frmKikiDataGate
            frmKikiDataGate.Show 1
    
        Case Else
            '�����Ȃ�
            
    End Select

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIKOUSEISETMENU_GAMEN_END, 0)
    
    Unload Me

End Sub
