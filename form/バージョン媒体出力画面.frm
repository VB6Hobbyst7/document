VERSION 5.00
Begin VB.Form frmVerOutput 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�o�[�W�����Ǘ�"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   14.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
   Begin VB.Timer tmrMail 
      Left            =   5880
      Top             =   4680
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�e�L�X�g�\��"
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
      Left            =   8400
      TabIndex        =   31
      Top             =   5760
      Width           =   3255
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
      Index           =   1
      Left            =   4440
      TabIndex        =   30
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
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
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "EG-R����"
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   27
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�o�[�W�����`�F�b�N�t�@�C���F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�\��2�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   19
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E���C��CPU-OS�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   2370
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E���C��CPU-Pro�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   17
         Top             =   705
         Width           =   2205
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�\���P�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   16
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�T�uCPU-Pro�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   15
         Top             =   1035
         Width           =   2175
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E����CPU-Pro�F "
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2190
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   13
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   12
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   11
         Top             =   705
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   10
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   9
         Top             =   1035
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraAllKansiVersion 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11535
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   20
         Left            =   8520
         TabIndex        =   34
         Top             =   650
         Width           =   2895
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   19
         Left            =   4500
         TabIndex        =   33
         Top             =   650
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   17
         Left            =   450
         TabIndex        =   32
         Top             =   650
         Width           =   2295
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�Ď��ՁF"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   350
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�h�c���p���j�b�g�F"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   3
         Top             =   350
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�k�c���[�e�B���e�B�F"
         Height          =   375
         Index           =   3
         Left            =   8355
         TabIndex        =   2
         Top             =   350
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " �@ ���j���[�@   ��ʂ֖߂�"
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
   Begin VB.Label lblVerName 
      Caption         =   "Z9"
      Height          =   375
      Index           =   29
      Left            =   7320
      TabIndex        =   36
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "�E���C�^���F"
      Height          =   495
      Index           =   21
      Left            =   4440
      TabIndex        =   35
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "Z9.Z9.Z9.Z9"
      Height          =   375
      Index           =   18
      Left            =   7320
      TabIndex        =   28
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   9
      Left            =   7320
      TabIndex        =   25
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   8
      Left            =   7320
      TabIndex        =   24
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "99"
      Height          =   375
      Index           =   7
      Left            =   7320
      TabIndex        =   23
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "�E�w�s�x�o�[�W�����F"
      Height          =   495
      Index           =   5
      Left            =   4440
      TabIndex        =   22
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "�ENEG�����F"
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   21
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "�E�h�b���ʉ^���F"
      Height          =   495
      Index           =   28
      Left            =   4440
      TabIndex        =   20
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�o�[�W�����}�̏o��"
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
   Begin VB.Label lblVerName 
      Caption         =   "�E�h�b�|�l�F"
      Height          =   375
      Index           =   6
      Left            =   4440
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "frmVerOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmVerOutPut.frm
'//  �p�b�P�[�W���F�o�[�W�����}�̏o�͉��
'//
'//  �T�v�F�o�[�W�����}�̏o�͉��
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͒ǉ�
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK�Ή�
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 �o�[�W�����t�@�C�������ݐ�f�B���N�g���ʒu�ύX
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
Private Const PtnEkiVersion = "000002"  '�w�o�[�W����
Dim sWriteDir As String                 '�}�̏o�͐�

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����}�̏o�͉��(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
   On Error Resume Next
    
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����}�̏o�͉��(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    On Error Resume Next
   
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����}�̏o�͉��(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
   Dim strWork         As String   '��ƃG���A
 
   On Error Resume Next
 
   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
           
   sWriteDir = ""
   
   'IDU�k�ރ`�F�b�N
   psIDUCheck
    
   '�o�[�W�����擾����
   psGetVersion
   
   '���[����M�p�̃^�C�}�l��ݒ肷��B
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
   
   '�u�o�[�W�����}�̏o�͉�� �\�������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_GAMEN_START, 0)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psGetVersion
'//  �@�\����  : �o�[�W�����擾����
'//  �@�\�T�v  : �o�[�W�����擾�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub psGetVersion()
  
  Dim sVersion  As String
  Dim sGetJikiVer As String     'V1.10.0.1 ADD
  
  On Error Resume Next

 '�Ď��ՁAEG-R�S�̃o�[�W�����擾
  psKansiGetVersion
 
 If pbIDUSts = 1 Then
    'IDU�o�[�W������\��
    lblVerName(2).Enabled = False
    lblVerName(19).Caption = ""
 Else
    '��k�ގ��͕\������
    psIDUGetVersion
 End If
 
 'LDU�S�̃o�[�W�����擾
  psLDUVersion

 'EG-R�����o�[�W�����擾
  '����CPU
  sVersion = psEGRJVersion(HANTEI_CPU)
  lblVerName(10).Caption = sVersion
  '���C��CPU
  sVersion = psEGRJVersion(MAIN_CPU)
  lblVerName(13).Caption = sVersion
 '�T�uCPU
  sVersion = psEGRJVersion(SUB_CPU)
  lblVerName(11).Caption = sVersion
 '���C��OS
  sVersion = psEGRJVersion(MAIN_OS)
  lblVerName(14).Caption = sVersion
 '�\���P
  sVersion = psEGRJVersion(YOBI1)
  lblVerName(12).Caption = sVersion
 '�\���Q
  sVersion = psEGRJVersion(YOBI2)
  lblVerName(15).Caption = sVersion
 '�o�[�W�����`�F�b�N
  sVersion = psEGRJVersion(VER_CHK)
  lblVerName(0).Caption = sVersion
  
 'NEG�����o�[�W�����擾
  sVersion = psNEGJVersion
  lblVerName(7).Caption = sVersion
 
 '����IC-M�o�[�W�����擾
 If pbIDUSts = 1 Then
    '����IC-M(IC-M)�o�[�W������\��
    lblVerName(6).Enabled = False
    lblVerName(8).Caption = ""
 Else
    '��k�ގ��͕\������
    sVersion = psICMGetVersion
    lblVerName(8).Caption = sVersion
 End If
 
 '���ʉ^���o�[�W�����擾
 If pbIDUSts = 1 Then
    '���ʉ^���o�[�W������\��
    lblVerName(28).Enabled = False
    lblVerName(9).Caption = ""
 Else
    '��k�ގ��͕\������
    sVersion = psICUnchinGetVersion
    lblVerName(9).Caption = sVersion
 End If
  
 '�w�s�x�o�[�W�����擾
 pfEkiVersion
 
'V1.10.0.1 ADD START
 '���C�^���ǂݍ���
 sGetJikiVer = psJikiUnchinVersion
 lblVerName(29).Caption = CStr(sGetJikiVer)
'V1.10.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psKansiGetVersion
'//  �@�\����  : �Ď����u�S�́A�Ď��Ճo�[�W�����擾����
'//  �@�\�T�v  : KansiVersion.ini���o�[�W�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function psKansiGetVersion()
    Dim lSts As Long                                       '�֐��߂�l
    Dim strKansiVersion As String * VERSION_GATE_SIZE      '�Ď��ՑS�̃o�[�W����
    Dim strKansiVersion2 As String * VERSION_GATE_SIZE     '�Ď����u�S�̃o�[�W����
    
    On Error Resume Next
    
    strKansiVersion = ""
    strKansiVersion2 = ""

    ' KansiVersion.ini����Ď����u�̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSISYSTEMVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSION_FILE)
    If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
        fraAllKansiVersion.Caption = "�S�̃o�[�W�����F " & Left$(strKansiVersion, lSts) & ""
    Else
        '�o�[�W�����ԍ��̎擾�ُ�̏ꍇ�A�u--,--,--,--�v��\��
        fraAllKansiVersion.Caption = "�S�̃o�[�W�����F--.--.--.-- "
    End If
 
    ' KansiVersion.ini����Ď��Ղ̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion2, _
                                   Len(strKansiVersion2), _
                                   KANSI_VERSION_FILE)
     If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
        lblVerName(17).Caption = Left$(strKansiVersion2, lSts)
    Else
        '�o�[�W�����ԍ��̎擾�ُ�̏ꍇ�A�u--,--,--,--�v��\��
        lblVerName(17).Caption = "--.--.--.-- "
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psIDUGetVersion
'//  �@�\����  : ID���p���j�b�g�o�[�W�����擾����
'//  �@�\�T�v  : ID���p���j�b�g�o�[�W�����Ǘ��t�@�C�����A
'//              ID���p���j�b�g�̃o�[�W�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function psIDUGetVersion()
    Dim strWork     As String       '��ƃG���A
    Dim iFileNumber As Integer      '���g�p�t�@�C���ԍ�
    
    On Error Resume Next
    
    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
        
   'ID���p���j�b�g�o�[�W�����Ǘ��t�@�C�����I�[�v���B
    Open PATH_IDU_APP & PATH_IDU_VERKANRI For Input As #iFileNumber

    '���s�o�[�W�������擾����B
    Input #iFileNumber, strWork
    If (Trim(strWork) = "") Then
        '�o�[�W�����ԍ��擾�ُ�̏ꍇ
        lblVerName(19).Caption = "--.--.--.--"
    Else
       '�S�̃o�[�W����������쐬
        lblVerName(19).Caption = Trim(strWork)
    End If
      
   'ID���p���j�b�g�o�[�W�����Ǘ��t�@�C�����N���[�Y�B
    Close #iFileNumber
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psLDUVersion
'//  �@�\����  : LD���[�e�B���e�B�o�[�W�����擾����
'//  �@�\�T�v  : LD���[�e�B���e�B�o�[�W�����Ǘ��t�@�C�����A
'//              LD���[�e�B���e�B�̃o�[�W�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function psLDUVersion()
    Dim strWork     As String       '��ƃG���A
    Dim iFileNumber As Integer      '���g�p�t�@�C���ԍ�
    
    On Error Resume Next
    
    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
    
   'LD���[�e�B���e�B�o�[�W�����Ǘ��t�@�C�����I�[�v���B
    Open PATH_LDU_APP & PATH_LDU_VERKANRI For Input As #iFileNumber

    '���s�o�[�W�������擾����B
    Input #iFileNumber, strWork
    If (Trim(strWork) = "") Then
        '�o�[�W�����ԍ��擾�ُ�̏ꍇ
        lblVerName(20).Caption = "--.--.--.--"
    Else
       '�S�̃o�[�W����������쐬
        lblVerName(20).Caption = Trim(strWork)
    End If
      
   'LD���[�e�B���e�B�o�[�W�����Ǘ��t�@�C�����N���[�Y�B
    Close #iFileNumber

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfEkiVersion
'//  �@�\����  : �w�s�x�o�[�W�����擾����
'//  �@�\�T�v  : ���݉w�ݒ�t�@�C�����A
'//              �w�s�x�o�[�W�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function pfEkiVersion()

   Dim intFileNo            As Integer  '�t�@�C���ԍ�
   Dim intBunrui_Dai        As Integer         '�啪��
   Dim intBunrui_Tyu        As Integer         '������
   Dim intBunrui_Sho        As Integer         '������
   Dim strData              As String          '�ݒ�l
   Dim strPtnNo             As String          '�p�^�[���ԍ�
   Dim strEkiVersion        As String          '�w�o�[�W����
   Dim iGetDataCount        As Integer         '�f�[�^�擾�J�E���^
   Dim strFileName          As String          '�t�@�C����
 
   On Error Resume Next
 
   strFileName = Dir(EKI_SETTI_FILE, vbNormal)
   
   If strFileName = "" Then
      lblVerName(18).Caption = "--.--.--.--"
      Exit Function
   End If
   
   '�t�@�C���ԍ����擾����B
   intFileNo = FreeFile
    
   '�t�@�C���I�[�v��
   On Error GoTo FileGetError
   Open EKI_SETTI_FILE For Input As #intFileNo

   Do While Not EOF(intFileNo)
      '�P �s�Âϐ��ǂݍ���
       Input #intFileNo, intBunrui_Dai, intBunrui_Tyu, intBunrui_Sho, strData
   
       '�p�^�[���ԍ��擾
        strPtnNo = Format(intBunrui_Dai, "00") & Format(intBunrui_Tyu, "00") & Format(intBunrui_Sho, "00")
           
        Select Case strPtnNo
             
            '�w�o�[�W�����擾
             Case PtnEkiVersion
                   strEkiVersion = strData & "  "
                   iGetDataCount = iGetDataCount + 1
             
             Case Else
                   '�����Ȃ�
                    
         End Select
            
         '�w�o�[�W�������擾�����烋�[�v�𔲂���
         If iGetDataCount = 1 Then Exit Do
   Loop

   '�t�@�C���N���[�Y
   Close #intFileNo

   If strEkiVersion = "" Then
      lblVerName(18).Caption = "--.--.--.--"
   Else
      lblVerName(18).Caption = strEkiVersion
   End If
   
   Exit Function
FileGetError:
   lblVerName(18).Caption = "--.--.--.--"

End Function
     
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �u�}�̏o�́v�u�}�̎�O�v�u�e�L�X�g�\���v���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 �o�[�W�����t�@�C�������ݐ�f�B���N�g���ʒu�ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    
    Dim bRet      As Boolean         '�߂�l
    Dim lRetVal   As Long            '�e�L�X�g�\�������߂�l
    Dim sCommand  As String          '�R�}���h������
    
    On Error Resume Next
 
    Select Case Index
        Case 0                                 '�u�}�̏o�́v�t
            '�u�}�̏o�͖t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_OUTPUT_BUTTOM, 0)
            ' ��o����f�B���N�g����I������
'            sWriteDir = pfDirSelection("a:", "�o�[�W�����t�@�C�������ݐ�f�B���N�g���I��")     'V1.12.0.1 DEL
            sWriteDir = pfDirSelection("H:", "�o�[�W�����t�@�C�������ݐ�f�B���N�g���I��")      'V1.12.0.1 ADD
            If sWriteDir <> "" Then
            '�f�B���N�g�����w�肳���΁A�o�[�W�����t�@�C������o��
                bRet = dllEGRCreateVersionFile(PATH_IDU_APP, PATH_LDU_APP)
                If bRet = False Then
                  '�u�t�@�C���쐬�ُ�v�|�b�v�A�b�v��ʕ\��
                    MsgBox "�t�@�C���̍쐬�Ɏ��s���܂����B", vbOKOnly + vbCritical, "�t�@�C���쐬�ُ�"
                  '�u�t�@�C���쐬�ُ�v���O�o��
                  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_FILE_CREATE_ERROR, 0)
                  
                  '�u�}�̏o�͏����ُ�v���O�o��
                   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_OUTPUT_ERROR, 0)
                   Exit Sub
                Else
                   '�t�@�C���R�s�[����
                   fMakeOutPutFile
                End If
                
            Else
                '�u�}�̏o�͏��������s�v���O�o��
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_OUTPUT_MISHORI, 0)
            End If

        Case 1                                 '�u�}�̎�O�v�t
            '�u�}�̎�O�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
            '�}�̎�O����
            Call pfRemove(Me)
        Case 2                                 '�u�e�L�X�g�\���v�t
            '�u�e�L�X�g�\���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_TEXT_BUTTOM, 0)
            bRet = dllEGRCreateVersionFile(PATH_IDU_APP, PATH_LDU_APP)
            If bRet = False Then
              '�u�t�@�C���쐬�ُ�v�|�b�v�A�b�v��ʕ\��
               MsgBox "�t�@�C���̍쐬�Ɏ��s���܂����B", vbOKOnly + vbCritical, "�t�@�C���쐬�ُ�"
               
               '�u�t�@�C���쐬�ُ�v���O�o��
                   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_FILE_CREATE_ERROR, 0)
               '�u�e�L�X�g�\�������ُ�v���O�o��
                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_TEXT_ERROR, 0)
                Exit Sub
            Else
                 '�e�L�X�g�t�@�C���\������
                sCommand = MN_EXE_MEMO & EGR_KANSI_VERSION_FILE_PATH '���������s�R�}���h���쐬
                '���������N������
                lRetVal = Shell(sCommand, vbMaximizedFocus)
                '���������A�N�e�B�u�i�O�ʕ\���j�ɂ���
                AppActivate lRetVal, True
                SendKeys "{LEFT}", True
               '�u�e�L�X�g�\����������v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_TEXT_OK, 0)
            End If
    
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�����e�i���X��ʂ֖߂�v�t��������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
     
    On Error Resume Next
    
    '�u�o�[�W�����}�̏o�͉�� �����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERSION_OUTPUT_GAMEN_END, 0)
 
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeOutPutFile
'//  �@�\����  : �}�̏o�͏������s���B
'//  �@�\�T�v  : �}�̏o�̓t�@�C���o�͂��s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeOutPutFile()
   Dim iResponse As Integer   'MsgBox�߂�l
   Dim lngErrCode As Long     '�G���[�R�[�h
   Dim fso         As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
   Dim strWriteDir As String               '�o�͐�t�H���_

   On Error GoTo COPY_ERROR

   '�t�@�C���R�s�[
   FileCopy EGR_KANSI_VERSION_FILE_PATH, sWriteDir & EGR_KANSI_VERSION_FILE
   
   '�u�}�̏o�͐���I���v�|�b�v�A�b�v��ʕ\��
   MsgBox "�}�̏o�͂͐���I�����܂����B", vbOKOnly + vbInformation, "�}�̏o�͌���"
                    
   '�u�o�[�W�����}�̏o�́F�}�̏o�͏�������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERSION_OUTPUT_OUTPUT_OK, 0)
  
   Exit Function
    
COPY_ERROR:
   '�����ُ�̏ꍇ�A�o�͌��ʃ|�b�v�A�b�v(�ُ�)�\��
    MsgBox "�}�̏o�ُ͈͂�I�����܂����B", vbCritical, "�}�̏o�͌���"
   '�u�o�[�W�����}�̏o�́F�}�̏o�͏����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERSION_OUTPUT_OUTPUT_ERROR, lngErrCode)
End Function

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
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͉�ʒǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  
    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmVerOutput.Caption, False
    End If

End Sub

