VERSION 5.00
Begin VB.Form frmKikiSettei 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�@����ݒ�"
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�w���@��ID�m�F"
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
      Index           =   5
      Left            =   6360
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8280
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�W���ݒ� �ۑ��^����"
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
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "LAN�J�[�h�ݒ�"
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
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�w�s�x�f�[�^�m�F"
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
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�@��\���ݒ�"
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
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�w�s�x�f�[�^�ݒ�"
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
      Caption         =   " �����e�i���X   ��ʂ֖߂�"
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
      Caption         =   "�@����ݒ�"
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
      TabIndex        =   6
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �F�@����ݒ胁�j���[���.frm
'//  �p�b�P�[�W���F�@����ݒ胁�j���[�̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�w���@��ID�m�F��ʁALAN�J�[�h��ʉ��ʐݒ��ʒǉ�
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//                 ��ʃ��b�N�����^��ʃ��b�N���������ǉ�
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ��ʍđO�ʕ\���C��(�s��C��)
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000       '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �@����ݒ胁�j���[���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
    
    '�^�C�}���N������
    tmrMail.Enabled = True

End Sub

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �@����ݒ胁�j���[���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �@����ݒ胁�j���[���(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'V1.4.0.1�@ADD START
    'IDU�k�ރ`�F�b�N
    psIDUCheck
    
    If pbIDUSts = 1 Then
     '�w���@��ID�m�F��\��
      cmdFixedExe(5).Visible = False
    End If
    'V1.4.0.1�@ADD END
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
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
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//                 �E�w�s�x�f�[�^�m�F���j���[��ʁ��w�s�x�f�[�^�m�F(�w���)���
'//                 �E�@��\���ݒ胁�j���[��ʁ��@��\���ݒ�(�w���)���
'//                 �ELAN�J�[�h��ʉ��ʐݒ�t�������A��ʕ\�������ǉ�
'//                 �E�w���@��ID�m�F��ʒǉ�
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)

    '�G���[���[�`����錾
    On Error Resume Next
    
'V1.12.0.1 ADD START
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    Select Case Index
        
        Case 0                                 '�w�s�x�f�[�^�ݒ�
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKITUDO_DATA_SETTEI, 0)
            
            '��ʕ\��
            Load frmEkisettei
            frmEkisettei.Show 1

        Case 1                                 '�w�s�x�f�[�^�m�F
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKITUDO_DATA_KAKUNIN, 0)
            'V1.4.0.1 DEL START
            '��ʕ\��
            'Load frmEkiDataGateMenu
            'frmEkiDataGateMenu.Show 1
            'V1.4.0.1 DEL END
            'V1.4.0.1 ADD START
            '��ʕ\��
            Load frmEkiData
            frmEkiData.Show 1
            'V1.4.0.1 ADD END
   
        Case 2                                 '�@��\���ݒ�
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_KIKI_KOUSEI_SETTEI, 0)
            'V1.4.0.1 DEL START
            '��ʕ\��
            'Load frmKikiDataMenu
            'frmKikiDataMenu.Show 1
            'V1.4.0.1 DEL END
            'V1.4.0.1 ADD START
            '��ʕ\��
            Load frmKikiData
            frmKikiData.Show 1
            'V1.4.0.1 ADD END
            
        Case 3                                 'LAN�J�[�h��ʉ��ʐݒ�
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_LAN_CARD_SETTEI, 0)

'V1.4.0.1 ADD START
            '��ʕ\��
            Load frmLanSettei
            frmLanSettei.Show 1
'V1.4.0.1 ADD END

        Case 4                                 '�W���ݒ�ێ�����
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_KAKARI_SAVE_RESTORE, 0)
            
            '��ʕ\��
            Load frmRenewData
            frmRenewData.Show 1
'V1.4.0.1 ADD START
        Case 5                                 '�w���@��ID
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKIMUKIKI_ID, 0)
            '��ʕ\��
            Load frmEkimKikiId
            frmEkimKikiId.Show 1
'V1.4.0.1 ADD END

        Case Else
            '�����Ȃ�
            
    End Select

'V1.12.0.1 ADD START
    '�S�{�^���������Ƃ���B
    Call SetEnableTrue
'V1.12.0.1 ADD END
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�����e�i���X��ʂ֖߂�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
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
Private Sub cmdReturn_Click()

    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_END, 0)
    
    '����ʏ���
    Unload Me

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
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ��ʍđO�ʕ\���C��(�s��C��)
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       'V1.17.0.1 DEL START
'        AppActivate frmRenewData.Caption, False
'        pfFormActive (frmRenewData.hwnd)
       'V1.17.0.1 DEL START
       'V1.17.0.1 ADD START
        AppActivate frmKikiSettei.Caption, False
        pfFormActive (frmKikiSettei.hwnd)
       'V1.17.0.1 ADD END
    End If

End Sub

'V1.12.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  �֐�����  : SetEnableFalse
'//  �@�\����  : ��ʃ��b�N����
'//  �@�\�T�v  : ��ʂ����b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������s�Ƃ���B
    cmdFixedExe(0).Enabled = False
    cmdFixedExe(1).Enabled = False
    cmdFixedExe(2).Enabled = False
    cmdFixedExe(3).Enabled = False
    cmdFixedExe(4).Enabled = False
    cmdFixedExe(5).Enabled = False
    cmdReturn.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  �֐�����  : SetEnableTrue
'//  �@�\����  : ��ʃ��b�N��������
'//  �@�\�T�v  : ��ʂ̃��b�N����������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������Ƃ���B
    cmdFixedExe(0).Enabled = True
    cmdFixedExe(1).Enabled = True
    cmdFixedExe(2).Enabled = True
    cmdFixedExe(3).Enabled = True
    cmdFixedExe(4).Enabled = True
    cmdFixedExe(5).Enabled = True
    cmdReturn.Enabled = True
    
End Sub
'V1.12.0.1 ADD END

