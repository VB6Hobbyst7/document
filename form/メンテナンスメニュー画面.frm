VERSION 5.00
Begin VB.Form frmHoshu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ێ�"
   ClientHeight    =   9000
   ClientLeft      =   555
   ClientTop       =   2325
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   9.75
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
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer tmrMail 
      Left            =   1200
      Top             =   7680
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   19
      Left            =   9000
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   18
      Left            =   9000
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   17
      Left            =   9000
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   16
      Left            =   9000
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   15
      Left            =   9000
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "15.�ғ�Ver�ꗗ�\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   14
      Left            =   6120
      TabIndex        =   16
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "14.TOMAS�f�[�^�Ǘ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   13
      Left            =   6120
      TabIndex        =   15
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "13.�V�X�e���ݒ�@ "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   12
      Left            =   6120
      TabIndex        =   14
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "12.�k�c�t�Ɩ�    "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   11
      Left            =   6120
      TabIndex        =   13
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "11.�h�c�t�Ɩ�    "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   10
      Left            =   6120
      TabIndex        =   12
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "9.�p�X���[�h�ݒ�  "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   3240
      TabIndex        =   11
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "10.�Ӱ�����ݽ    "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   3240
      TabIndex        =   10
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5.�ʐM�m�F�E�\�� "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8.�A�v���N���E�I�� "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   3240
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1.�o�[�W�����Ǘ� "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2.�@����ݒ�   "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6.�V�X�e���������@"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7.հè�è�N��    "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3.���O�Ǘ�       "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�Ď���ʂ֖߂�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      TabIndex        =   0
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton cmdMenue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4.�f�[�^���W�E�o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�����e�i���X���j���["
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
      TabIndex        =   7
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmHoshu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmHoshu.frm
'//  �p�b�P�[�W���F�����e�i���X���j���[���
'//
'//  �T�v�F�����e�i���X���j���[���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-12   REVISED   BY [TCC] C.Terui
'//                 �E�Ǘ��v���Z�X���N�����Ă��Ȃ��ꍇ��
'//                   �u�Ď���ʂɖ߂�v�{�^����������
'//                 �E�ێ��ʂ��A�����[�h���ꂽ���̃C�x���g�ǉ�
'//     REVISIONS :(1.6.0.1) 2009-04-11   REVISED   BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�Ď��Րݒ�t�ǉ�
'//     REVISIONS :(1.10.0.1)2009-09-25   REVISED   BY [TCC] T.Furuya
'//                 �EKK�Ή�
'//     REVISIONS :(2.2.0.1)  2010-09-11  REVISED BY [TCC] S.Terao
'//                 �d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

Dim iGamenType As Integer '�擾��ʃ^�C�v

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �����e�i���X���j���[���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.10.0.1)2009-09-25   REVISED   BY [TCC] T.Furuya
'//                 �EKK�Ή��@�߂�t�����ύX
'//     REVISIONS :(EG20 V6.2.0.1) 2012-06-15  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    Dim iKansiAplChk As Integer     '�A�v���N���`�F�b�N�߂�l�p�@'V1.10.0.1 ADD
    
    On Error Resume Next
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
    
' EG20 V6.2.0.1 �ǉ��J�n
    If pubGetTomasFunction() = True Then
        cmdMenue(13).Enabled = True
    Else
        cmdMenue(13).Enabled = False
    End If
' EG20 V6.2.0.1 �ǉ��I��
    
'V1.10.0.1 ADD START
    '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '�Ď��ՋN�����F�߂�t�̕����u�Ď���ʂ֖߂�v
        cmdReturn.Caption = "�Ď���ʂ֖߂�"
    Else
        '�Ď����N�����F�߂�t�̕����uWindows�ɖ߂�v
        cmdReturn.Caption = "Windows�ɖ߂�"
    End If
'V1.10.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �����e�i���X���j���[���(�f�B�A�N�e�B�u��)
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
Private Sub Form_Deactivate()
    On Error Resume Next
    
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �����e�i���X���j���[���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.2.0.1) 2010-09-13   REVISED   BY [TCC] S.Terao
'//                 �E�d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    On Error Resume Next
   
    '�u�����e�i���X���j���[��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_GAMEN_START, 0)

    pbAbortFlag = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
       
    'IDU/LDU�k�ރ`�F�b�N
    psIDUCheck
    psLDUCheck

    If pbIDUSts = 1 Then
      'IDU�Ɩ���\��
       cmdMenue(10).Visible = False
    End If

    If pbLDUSts = 1 Then
      'LDU�Ɩ���\��
       cmdMenue(11).Visible = False
    End If
'V2.2.0.1 ADD START
    psUnchin_Dll
    psEki_Type iGamenType
'V2.2.0.1 ADD END

    Call gsGetGateInfo      'EG20 V2.1.0.1 ADD
    
    '���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�Ď���ʂ֖߂�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                 �E�Ǘ����N�����Ă��Ȃ���Ԃ̏�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
   On Error Resume Next
  
   '�u�����e�i���X���j���[��ʁF�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_GAMEN_END, 0)
' V1.3.0.1 ADD START
    If CheckAppStart(PROC_KANRI) = 0 Then
    '�Ǘ��v���Z�X���N�����Ă��Ȃ��ꍇ
        psEndHoshuProc
    Else
' V1.3.0.1 ADD END
        '�I���������s��
        psEndProc
    End If          ' V1.3.0.1 ADD
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdMenue_Click
'//  �@�\����  : �e��ʑJ�ږt����������
'//  �@�\�T�v  : �t���̉�ʂɑJ�ڂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-04-11   REVISED   BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�Ď��Րݒ�t�ǉ�
'//     REVISIONS :(2.2.0.1) 2010-09-13   REVISED   BY [TCC] S.Terao
'//                 �E�d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'//     REVISIONS :(EG20 V4.1.0.1) 2011-12-27   REVISED   BY [TCC] M.Matsumoto
'//                 �E�y�t�F�[�Y�RTOMAS�Ή��z
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.58�C���Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdMenue_Click(Index As Integer)
    Dim bRet As Boolean            '���[�����M�����̖߂�l
    Dim iResponse As Integer       '���b�Z�[�W�̖߂�l
    Dim udtMail As ML_DISP_INF     '��ʕ\���v��
    Dim lngErrCode As Long         '�G���[�R�[�h
    
    On Error Resume Next
   
    Select Case Index
        Case 0                                 '�o�[�W�����Ǘ�
           '�u�����e�i���X���j���[��ʁF�o�[�W�����Ǘ��t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_VERSION_BUTTOM, 0)
            Load frmVersion
            frmVersion.Show 1
        Case 1                                 '�@����ݒ�
           '�u�����e�i���X���j���[��ʁF�@����ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_KIKIINFOSETTEI_BUTTOM, 0)
            Load frmKikiSettei
            frmKikiSettei.Show 1
        Case 2                                 '���O�Ǘ�
           '�u�����e�i���X���j���[��ʁF���O�Ǘ��t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_LOG_BUTTOM, 0)
            Load frmLogMenu
            frmLogMenu.Show 1
        Case 3                                 '�����ێ�f�[�^
           '�u�����e�i���X���j���[��ʁF�����ێ�f�[�^�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_JIKAIHOSHU_BUTTOM, 0)
            Load frmGateHoshu
            frmGateHoshu.Show 1
        Case 4                                 '�ʐM�m�F�E�\��
           '�u�����e�i���X���j���[��ʁF�ʐM�m�F�E�\���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_CONECT_BUTTOM, 0)
            Load frmTusinMenu
            frmTusinMenu.Show 1
        Case 5                                 '�V�X�e��������
           '�u�����e�i���X���j���[��ʁF�V�X�e���������t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_SYSFORMAT_BUTTOM, 0)
            Load frmSysformatMenu
            frmSysformatMenu.Show 1
        Case 6                                 'հè�è�N��
           '�u�����e�i���X���j���[��ʁFհè�è�N���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_UTILITY_BUTTOM, 0)
            If pbUserLevel = 0 Then
                Load frmUtilityUSR             '��ʃ����e�i���X
                frmUtilityUSR.Show 1
            Else
                Load frmUtility                '���������e�i���X
                frmUtility.Show 1
            End If
        Case 7                                 '�A�v���I��
           '�u�����e�i���X���j���[��ʁF�A�v���N���E�I���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_APLSTART_END_BUTTOM, 0)
            Load frmAppConfig
            frmAppConfig.Show 1
        Case 8                                 '�p�X���[�h�ݒ�
           '�u�����e�i���X���j���[��ʁF�p�X���[�h�ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_PASS_SETTEI_BUTTOM, 0)
            Load frmPassSet
            frmPassSet.Show 1
        Case 9                                 '�Ӱ�����ݽ
            '�u�����e�i���X���j���[��ʁF�Ӱ�����ݽ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_RMENTE_BUTTOM, 0)
            Load frmRmenteMenu
            frmRmenteMenu.Show 1
        Case 10                                '�h�c�t�Ɩ�
           '�u�����e�i���X���j���[��ʁFIDU�Ɩ��t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_IDU_BUTTOM, 0)
           '��ʕ\���v��(IDU�Ɩ����)��ID����ɑ��M����
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_IDU_GAMEN
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
               '�u�����e�i���X���j���[��ʁF��ʕ\���v�����[�����M�ُ�v���O�o��
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
' EG20 V5.2.0.1�y����TR-No.58�C���Ή��z�폜�J�n
'               '�N�����s�|�b�v�A�b�v�\��
'               iResponse = MsgBox("IDU�Ɩ��t�A��`�G���[�B" & _
'                                  Chr(vbKeyReturn) & _
'                                  "ID���p���j�b�g�Ɩ���ʂ��N���ł��܂���B", _
'                                  vbOKOnly, _
'                                  "��ʋN���G���[")
' EG20 V5.2.0.1�y����TR-No.58�C���Ή��z�폜�I��
' EG20 V5.2.0.1�y����TR-No.58�C���Ή��z�ǉ��J�n
               '�N�����s�|�b�v�A�b�v�\��
               iResponse = MsgBox("IDU�Ɩ��t�A��`�G���[�B" & _
                                  Chr(vbKeyReturn) & _
                                  "�h�c�t�Ɩ���ʂ��N���ł��܂���B", _
                                  vbOKOnly, _
                                  "��ʋN���G���[")
' EG20 V5.2.0.1�y����TR-No.58�C���Ή��z�ǉ��I��
               Exit Sub
            End If
            '�u�����e�i���X���j���[��ʁF��ʕ\���v�����[�����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        Case 11                                '�k�c�t�Ɩ�
            '�u�����e�i���X���j���[��ʁFLDU�Ɩ��t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_LDU_BUTTOM, 0)
            '��ʕ\���v��(LDU�Ɩ����)��LD����ɑ��M����
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_LDU_GAMEN
            bRet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
              '�u�����e�i���X���j���[��ʁF��ʕ\���v�����[�����M�ُ�v���O�o��
              lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
               '�N�����s�|�b�v�A�b�v�\��
                iResponse = MsgBox("LDU�Ɩ��t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "LD���[�e�B���e�B�Ɩ���ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
               Exit Sub
            End If
            '�u�����e�i���X���j���[��ʁF��ʕ\���v�����[�����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
'EG20 V2.1.0.1 ADD START
        Case 12                                 '�Ď��Րݒ�
            '�u�����e�i���X���j���[��ʁF�V�X�e���ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_SYSTEM_SETTEI_BUTTOM, 0)
            Load frmSystemSetteiMenu
            frmSystemSetteiMenu.Show 1
'EG20 V2.1.0.1 ADD END
'V1.6.0.1 ADD START
        Case 13                                 '�Ď��Րݒ�
        'EG20 V4.1.0.1 DEL START �y�t�F�[�Y�RTOMAS�Ή��z
'            '�u�����e�i���X���j���[��ʁF�Ď��Րݒ�t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_KANSIBAN_SETTEI_BUTTOM, 0)
'            Load frmKansiSettei
'            frmKansiSettei.Show 1
        'EG20 V4.1.0.1 DEL END
        'EG20 V4.1.0.1 ADD START �y�t�F�[�Y�RTOMAS�Ή��z
            '�u�����e�i���X���j���[��ʁFTOMAS�f�[�^�Ǘ��t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_TOMAS_DATA_BUTTOM, 0)
            Load frmTomasDataMng
            frmTomasDataMng.Show 1
        'EG20 V4.1.0.1 ADD END
'V1.6.0.1 ADD END
'V2.2.0.1 ADD START
        'EG20 V5.2.0.1 DEL START �y�ғ��o�[�W�����Ǘ���ʒǉ��z
'       Case 14
'           '�u�����e�i���X���j���[��ʁF�^���f�[�^DLL�t�����v���O�o��
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_MENU_UNCHINDATA_DLL_BUTTOM, 0)
'
'           If iGamenType = 1 Then
'              '���g�����
'              '�^���f�[�^DLL�t������
'              Load frmICUnkai_Type1  '�^���f�[�^DLL��ʂ����[�h����
'              '�^���f�[�^DLL��ʂ�\������
'              frmICUnkai_Type1.Show 1
'           End If
'           If iGamenType <> 1 Then
'              '���l�E���S���
'              '�^���f�[�^DLL�t������
'              Load frmICUnkai_Type2  '�^���f�[�^DLL��ʂ����[�h����
'              '�^���f�[�^DLL��ʂ�\������
'              frmICUnkai_Type2.Show 1
'           End If
        'EG20 V5.2.0.1 DEL END
'V2.2.0.1 ADD END
        'EG20 V5.2.0.1 ADD START �y�ғ��o�[�W�����Ǘ���ʒǉ��z
        Case 14
            '�u�ғ�Ver�ꗗ�\����ʖt�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_KADOVER_BUTTOM, 0)
            gStrCurrentForm = sFormName_KadoVerKanri
            Load frmKadoVerKanri
            frmKadoVerKanri.Show 1
        'EG20 V5.2.0.1 ADD END
    End Select
End Sub
' V1.3.0.1 ADD START
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'     �T�v      : �ێ��ʂ��A�����[�h���ꂽ���̃C�x���g�v���V�[�W��
'     ����      : �ێ��ʃv���Z�X���I������B
'
'     ORIGINAL  :(1.3.0.1) '09-03-12   CODED   BY [TCC] C.Terui
'     REVISIONS :(X.X.X.X) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    End   '�ێ��ʃv���Z�X���I��(Exit)����B
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v  : ���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmHoshu.Caption, False
        pfFormActive (frmHoshu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
'V2.2.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psUnchin_Dll
'//  �@�\����  : �^���f�[�^DLL�t�\���^��\�����[�U�`�F�b�N����
'//  �@�\�T�v  : HOSHU.INI���A���C�^���Ή����[�U�ł��邩�ǂ����`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(2.2.0.1) 2010-09-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub psUnchin_Dll()

    Dim iFlag As Integer '�擾���[�U�t���O
 
    On Error Resume Next
 
  ' HOSHU.INI���^���f�[�^�c�k�k�Ή����[�U�t���O���擾����B
  '(�f�t�H���g�͔�\��)
    iFlag = GetPrivateProfileInt(KANSI_UNCHIN_DLL_SEC, _
                                 KANSI_UNCHIN_DLL_KEY, _
                                 DEFAILT_Int, _
                                 HOSHU_FILE)
     If iFlag = 1 Then
      '�t���O��1�̏ꍇ�u���C�^���v�t�͕\��
      cmdMenue(14).Visible = True
     End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psEki_Type
'//  �@�\����  : �w�s�x�ɂ���ĉ^���f�[�^DLL��ʃ^�C�v���f����
'//  �@�\�T�v  : UnchinDLL_Sts.ini���A�^���f�[�^DLL��ʃ^�C�v�𔻒f����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(2.2.0.1) 2010-09-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub psEki_Type(iType As Integer)

    Dim iFlag As Integer '�擾�t���O
    
    On Error Resume Next
 
  ' UnchinDLL_Sts.ini���^���f�[�^�c�k�k��ʃ^�C�v�𔻒f����B
  '(�f�t�H���g�͕��l)
    iFlag = GetPrivateProfileInt(UNCHIN_DLL_STS_SEC, _
                                    UNCHIN_DLL_STS_KEY, _
                                    DEFAILT_Int, _
                                    UNCHIN_DLL_STS_FILE)

    If iFlag = 0 Then
      '�擾�ُ�ꍇ�u���l�^���f�[�^DLL�v��ʂ�\��
      iFlag = 2
      iType = iFlag
     Else
      iType = iFlag
     End If
End Sub
'V2.2.0.1 ADD END
