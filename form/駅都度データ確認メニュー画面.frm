VERSION 5.00
Begin VB.Form frmEkiDataGateMenu 
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
      Caption         =   "�}�̎�O"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�w�ݒ�e�L�X�g�o��"
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
      TabIndex        =   6
      Top             =   5040
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�w�ݒ����"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�w�ݒ�o��"
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
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
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
   Begin VB.CommandButton cmdCancel 
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
   Begin VB.Label Label7 
      Caption         =   "�I������Ă���w�̌��݂̉w�s�x�f�[�^�P�w�����e�L�X�g�\������B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Label Label6 
      Caption         =   "�E�E�E"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "�E�E�E"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "�w�s�x�f�[�^�P�w�����C���X�g�[������B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "�E�E�E"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "�I������Ă���w�̌��݂̉w�s�x�f�[�^�P�w�����o�͂���B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   2880
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�w�ݒ�m�F"
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
Attribute VB_Name = "frmEkiDataGateMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �F�w�s�x�f�[�^�m�F���j���[���.frm
'//  �p�b�P�[�W���F�w�s�x�f�[�^�m�F���j���[��ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �w�ݒ�t�@�C�������ݐ�f�B���N�g���ʒu�ύX
'//                 �f�B�X�N���擾�ʒu�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000       '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �w�s�x�f�[�^�m�F���j���[���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
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
'//  �@�\����  : �w�s�x�f�[�^�m�F���j���[���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �w�s�x�f�[�^�m�F���j���[���(���[�h���F�C�x���g�v���V�[�W��)
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_START, 0)
    
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
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKIINFO, 0)
            
            '��ʕ\��
            Load frmEkiData
            frmEkiData.Show 1
        Case 1                                 '����
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_GATE, 0)
            
            '��ʕ\��
            Load frmEkiDataGate
            frmEkiDataGate.Show 1
        Case 2                                  '�w�ݒ�o��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_OUTPUT, 0)
            
            '�w�ݒ�o�͏���
            Call sEkiSetteiOutPut
        
        Case 3                                  '�w�ݒ����
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_INPUT, 0)
            
            '�w�ݒ���͏���
            Call sInstolEkiSettei
        
        Case 4                                  '�w�ݒ�e�L�X�g�o��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_DISP_TEXT, 0)
            
            '�w�ݒ�e�L�X�g�o�͏���
            Call sDispTextEkiDataNow
        
        Case 5                                  '�}�̎�O
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
            
            '�}�̎�O����
            Call pfRemove(Me)
        
    End Select
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdCancel_Click
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
Private Sub cmdCancel_Click()

    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_END, 0)
    
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sEkiSetteiOutPut
'//  �@�\����  : �u�w�ݒ�o�́v�t����������
'//  �@�\�T�v  : ���݉w�ݒ�t�@�C�����O���}�̂ɏo�͂���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �w�ݒ�t�@�C�������ݐ�f�B���N�g���ʒu�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sEkiSetteiOutPut()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
    Dim iResponse            As Integer         'MsgBox�߂�l

    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '�ُ�I��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", _
                vbOKOnly + vbExclamation, _
                 "�f�[�^���x��"
        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '�}�̏o�͏���
    '----------------------------------------------------
'    sWriteDir = pfDirSelection("a:", "�w�ݒ�t�@�C�������ݐ�̃f�B���N�g���I��")   'V1.12.0.1 DEL
    sWriteDir = pfDirSelection("H:", "�w�ݒ�t�@�C�������ݐ�̃f�B���N�g���I��")    'V1.12.0.1 ADD
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
        FileCopy EKI_SETTI_FILE, sWriteDir & Dir(EKI_SETTI_FILE)
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
       '����I��
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�̏o�͌���")
    
    End If
  
  Exit Sub
 
COPY_ERROR:

    Select Case Err.Number
        Case 61 ' �}�̏o�͋󂫗e�ʕs��
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case 71 ' �}�̂Ȃ�
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_NOT_DISK, 0)
        Case Else
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

    iResponse = MsgBox("�ُ�I�����܂���", vbOKOnly, "�}�̏o�͌���")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sInstolEkiSettei
'//  �@�\����  : �u�w�ݒ���́v�t����������
'//  �@�\�T�v  : �O���}�̂��猻�݉w�ݒ�t�@�C���C���X�g�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �f�B�X�N���擾�ʒu�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sInstolEkiSettei()

    Dim iResponse           As Integer          'MsgBox�߂�l
    Dim bRet                As Boolean          '�֐��߂�l
    Dim lErrCode            As Long             '�G���[�R�[�h
    Dim strFileName         As String           '�}�̃t�@�C����
    
    Dim iRet                    As Integer      '���b�Z�[�W�{�b�N�X�߂�l
    Dim lSekuta                 As Long         '�Z�N�^�i�N���X�^����j
    Dim lByte                   As Long         '�o�C�g���i�Z�N�^����j
    Dim lKurasuta               As Long         '�t���[�N���X�^��
    Dim lDrive                  As Long         '�h���C�u�̃N���X�^���i���v�j
    Dim strDrive                As String       '�h���C�u
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    iResponse = MsgBox("�w�s�x�f�[�^�P�w�����C���X�g�[�����܂��B" & Chr(vbKeyReturn) & _
                        "��낵���ł����H", _
                        vbYesNo + vbExclamation, _
                        "�w�ݒ���͊m�F")
    
    If iResponse = vbNo Then Exit Sub
    
    '�f�B�X�N�����擾
'    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
    iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD

    If lDrive = 0 Then
        strDrive = "d:"
    Else
'        strDrive = "a:"        'V1.12.0.1 DEL
        strDrive = "H:"         'V1.12.0.1 ADD
    End If

    '�}�̃t�@�C�����擾
    strFileName = pfFileSelection(strDrive, "*.csv", "�w�ݒ�̧�ّI��")
    
    '�t�@�C�����݃`�F�b�N
    If strFileName <> "" Then

        '���݉w�ݒ�f�[�^�C���X�g�[������
        bRet = dllInstolEkiDataNow(strFileName, EKI_SETTI_FILE, lErrCode)
    
        If bRet = False Then
            
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)
            
            '�ُ�I��
            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ���͌���")
            
        Else
        
             '���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
            '����I��
            iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�w�ݒ���͌���")
            
        End If
    
    Else
        '�t�@�C���Ȃ�
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, ERROR_MEDIUM_OTHER_ERR, 0)
        
        '�ُ�I��
        iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ���͌���")
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispTextEkiDataNow
'//  �@�\����  : �u�w�ݒ�e�L�X�g�o�́v�t����������
'//  �@�\�T�v  : ���݉w�ݒ�t�@�C�����e�L�X�g�\������
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
Private Sub sDispTextEkiDataNow()

    Dim strFileName          As String          '�t�@�C����
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim lRetVal              As Long            '�߂�l
    Dim sCommand             As String          '�R�}���h������

    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '�ُ�I��
        MsgBox "�e�L�X�g�\������f�[�^������܂���B", _
                vbOKOnly + vbExclamation, _
                 "�f�[�^���x��"
        Exit Sub
        
    End If
    
    sCommand = MN_EXE_MEMO & EKI_SETTI_FILE         '���������s�R�}���h���쐬����
    lRetVal = Shell(sCommand, vbMaximizedFocus)     '�m�[�g�p�b�h���N������
    AppActivate lRetVal, True                       '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
    SendKeys "{LEFT}", True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfStartUpProc
'//  �@�\����  : �t�@�C���I����ʏ���
'//  �@�\�T�v  : �t�@�C���I����ʂ�\�����A�I�����ꂽ�t�@�C������Ԃ��B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sDrive�@�@[IN]�����\���h���C�u��
'//  �@�@      : String�@�@sPattern�@[IN]�I��Ώۃt�@�C���g���q
'//  �@�@      : String�@�@sTitle�@�@[IN]��ʕ\�����x��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :String�@�@�@�@�@�@�@ [OUT]�߂�l
'//                                      �I�����ꂽ�t�@�C���p�X:����@""�F�G���[
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function pfFileSelection(sDrive As String, _
                                sPattern As String, _
                                sTitle As String) As String
                                
    Dim sWorkDrive As String                    '���[�N�p�����\���h���C�u��

    '�h���C�u�ُ폈�����`����B
    On Error GoTo Drive_Error
    
    sWorkDrive = sDrive                         '�����\���h���C�u�������[�N�p�ɃZ�b�g����B
    frmFil.filSelection.Pattern = sPattern      '�I��Ώۊg���q���Z�b�g����B
    frmFil.lblFileSelection = sTitle            '�T�u�^�C�g�����Z�b�g����B

Retry:
    frmFil.drvSelection.Drive = sWorkDrive      '�h���C�u���Z�b�g����B
    frmFil.dirSelection.Path = sWorkDrive & "\" '�f�B���N�g�����Z�b�g����B
    
    '�t�@�C���I����ʂ�\������B
    frmFil.Show 1
    
    '�I�����ꂽ�t�@�C������Ԃ��B
    pfFileSelection = gstrMyPath
    
    Exit Function

'**�h���C�u�w��ُ폈��**
Drive_Error:

'    If Left$(sWorkDrive, 1) = "a" Then     'V1.12.0.1 DEL
    If Left$(sWorkDrive, 1) = "H" Then      'V1.12.0.1 ADD
        'a:�h���C�u���ُ�Ȃ�A�J�����g�h���C�u��\��������B
        sWorkDrive = Left$(App.Path, 2)
        GoTo Retry
    End If
    
    '���̑��̃h���C�u�Ȃ�A�t�@�C���I���Ȃ��Ŗ߂�B
    pfFileSelection = ""

End Function
