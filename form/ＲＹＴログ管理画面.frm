VERSION 5.00
Begin VB.Form frmRYTSyusyu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ғ��E�����e�f�[�^���W�i�����㎩�����D�@�j"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.CommandButton cmdInstall 
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
      Height          =   1095
      Left            =   8520
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   7200
      Top             =   5760
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
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdSyusyu 
      Caption         =   " ���W "
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
      Left            =   1080
      TabIndex        =   0
      Top             =   3315
      Width           =   2175
   End
   Begin VB.CommandButton cmdFDWrite 
      Caption         =   "�}�̏o��"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   3315
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�q�x�s���O�Ǘ�"
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
Attribute VB_Name = "frmRYTSyusyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmRYTSyusyu.frm
'//  �p�b�P�[�W���F�q�x�s���O�Ǘ����
'//
'//  �T�v�F�q�x�s���O�Ǘ����
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�q�x�s���O�Ǘ���ʒǉ�
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �q�x�s���O�Ǘ����(���[�h��)
'//  �@�\�T�v  : �����������s���B
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
Private Sub Form_Load()

    On Error Resume Next
    
    '���C����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
              
    '�u�q�x�s���O�Ǘ���ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_START, 0)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �q�x�s���O�Ǘ����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[���^�C�}���N������B
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
Private Sub Form_Activate()
  
    On Error Resume Next
    
    '���[����M�p�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �q�x�s���O�Ǘ����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[���^�C�}���~����B
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
Private Sub Form_Deactivate()
  
    On Error Resume Next
    
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False

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
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next
    
    '�u�q�x�s���O�Ǘ���ʁF�����v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_END, 0)

    '����ʂ������B
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdSyusyu_Click
'//  �@�\����  : �u���W�v�t����������
'//  �@�\�T�v  : �u���W�v�t�������������s���B
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
Private Sub cmdSyusyu_Click()
    Dim iResponse As Integer   'MsgBox�߂�l
    
    On Error Resume Next

    '�u���W�t�����v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_SYUSYU_BUTTOM, 0)

     '�q�x�s���O�f�[�^���W�|�b�v�A�b�v��ʕ\��
    iResponse = MsgBox("�q�x�s���O�f�[�^�����W���܂�����낵���ł����H", _
                       vbOKCancel, "�q�x�s���O�f�[�^�Ǘ�")
    If iResponse = vbOK Then
       'OK�t�������ꂽ��A�q�x�s���O�f�[�^���W����ʕ\��
       'RYT���O�f�[�^���W���t�H�[�����A���[�_���E�B���h�E�ŕ\������B
         frmRYTSyusyuCyu.Show vbModal
    Else
       '�L�����Z���t����
       '�u���W���������s�v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_MISHORI, 0)
       Exit Sub
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFDWrite_Click
'//  �@�\����  : �u�}�̏o�́v�t����������
'//  �@�\�T�v  : �u�}�̏o�́v�t�������������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFDWrite_Click()
    
   Dim fso         As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
   Dim sWriteDir As String
   
   On Error Resume Next
  
   '�u�}�̏o�͖t�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_OUTPUT_BUTTOM, 0)
  
    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
'    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")     'V1.12.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)      'V1.12.0.1 ADD

    '�w��t�H���_�Ȃ�
    If Len(sWriteDir) = 0 Then
       '�u�}�̏o�͏��������s�v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RYT_LOG_KANRI_GAMEN_OUTPUT_MISHORI, 0)
       Exit Sub
    End If
    
    Set fso = Nothing

    m_sCopySaki = sWriteDir & "\" & RYT_LOG_FILE

    '�R�s�[���p�X�쐬
    m_sCopyMoto = E_FIRMWARE_LOG & RYT_LOG_FILE
    
    frmRYTSyusyuOutPut.Show vbModal
          
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdInstall_Click
'//  �@�\����  : �u�}�̎�O�v�t����������
'//  �@�\�T�v  : �}�̂̎��O�����s���B
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
Private Sub cmdInstall_Click()
   
   On Error Resume Next
   
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : �^�C���A�b�v������
'//  �@�\�T�v  : ���[����M�^�C���A�b�v���������s���B
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
        AppActivate frmRYTSyusyu.Caption, False
    End If

End Sub

