VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ێ���p�X���[�h����"
   ClientHeight    =   9000
   ClientLeft      =   2700
   ClientTop       =   2520
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   Picture         =   "�p�X���[�h���͉��.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.CommandButton cmdKakutei 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�m  ��"
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
      Left            =   8280
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Frame fraPass 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5655
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�b"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   11
         Left            =   2880
         TabIndex        =   12
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BS"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   1680
         TabIndex        =   11
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   9
         Left            =   2880
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   1680
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   2880
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   1680
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   480
         TabIndex        =   5
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   2880
         TabIndex        =   4
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   480
         Width           =   3375
      End
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
      Height          =   1455
      Left            =   8400
      TabIndex        =   14
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Timer tmrMail 
      Left            =   600
      Top             =   720
   End
   Begin VB.Label lblGuide 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H0000FFFF&
      Caption         =   "  ���̉�ʂ́A�ێ����p�ł��I       �ێ���ȊO�̕��́A "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   0
      Left            =   2520
      TabIndex        =   16
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label lblGuide 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   20
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label lblGuide 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H0000FFFF&
      Caption         =   "  ��ʉE���́u�Ď���ʂ֖߂�v�{�^���������߂��ĉ������B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   19
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�ێ���p�X���[�h����"
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
      TabIndex        =   18
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  '����
      Caption         =   "�p�X���[�h����͂��ĉ������B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   1560
      Width           =   4095
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmPass.frm
'//  �p�b�P�[�W���F�p�X���[�h���͉��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �EEG10�ێ���A�p�X���[�h���͉��(frmPass.frm)�𗬗p
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                 �E���W�X�g���擾�ُ펞�����ύX
'//                 �E�t�H�[���A�����[�h�����ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-23   REVISED BY [TCC] S.Terao
'//                 �E��ʕ\��/�����^�C�~���O�C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yoshimori
'//                 �p�X���[�h�s��v�̉�ʑJ�ڕύX
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 �P���b�`�����g�p�w�Ή�
'//     REVISIONS :(EG20 V30.3.0.1)  2014-09-18  REVISED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή�
'//                 �yHKRK_Kansi07_007_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �p�X���[�h���͉��(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
    
    Dim iKansiAplChk As Integer     '�A�v���N���`�F�b�N�߂�l�p�@' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    
    '�ő剻�\������B
    Me.WindowState = 2
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True

' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '�Ď��ՋN�����F�߂�t�̕����u�Ď���ʂ֖߂�v
        cmdReturn.Caption = "�Ď���ʂ֖߂�"
    Else
        '�Ď����N�����F�߂�t�̕����uWindows�ɖ߂�v
        cmdReturn.Caption = "Windows�ɖ߂�"
    End If
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �p�X���[�h���͉��(�f�B�A�N�e�B�u��)
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
    
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �p�X���[�h���͉��(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.2.0.1) 2009-02-26   REVISED BY [TCC] S.Terao
'//             �����グ�����A���W�X�g���擾�ُ펞�N�������Ȃ��C��
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//             ���W�X�g���̃f�t�H���g�l�擾�ɔ����C��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-09   REVISED BY [TCC] M.Matsumoto
'//             �y�t�F�[�Y�Q�Ή��z���O�o�̓t�H���_�ύX
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim bRet As Boolean     '�߂�l
    Dim lngErrCode As Long  '�֐��G���[�R�[�h
    Dim slogPath As String  '�ێ烍�O
    
    On Error Resume Next
   
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '�ێ烍�O�N���X�𐶐�
'    slogPath = PATH_LOG & HOSHULOG_FILE            'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    slogPath = PATH_HOSHULOG & HOSHULOG_FILE        'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
    bRet = dllHoshulogClass(slogPath, lngErrCode)
    
    '����N��
    iStaFlag = 0
    '�u�ێ�N���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HOSHU_PROCESS_START, 0)
    '�N����
    iStaFlag = 1
    
     '�u�p�X���[�h���͉�ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_START, 0)
    
    '�p�X�ݒ�
    ChDrive App.Path
    ChDir App.Path
     
    '�����グ�������s��
    bRet = pfStartUpProc
    If bRet = False Then
    '�����グ�����ُ펞�͋����I������B
        'pbAbortFlag = True 'V1.2.0.1 DEL
        'V1.2.0.1 ADD START
        pfAbortProc
        End
        'V1.2.0.1 ADD END
    End If
   '�p�X���[�h�t�@�C�������������s���B
    sPassFileInitialize
    
    'IDU/LDU�̃p�X�����W�X�g�����擾����B
    bRet = sGetRegIDU_LDU_Path
' V1.3.0.1 DEL START
'   If False = bRet Then
'       'Exit Sub 'V1.2.0.1 DEL
'       'V1.2.0.1 ADD START
'       pfAbortProc
'       End
'       'V1.2.0.1 ADD END
'    End If
' V1.3.0.1 DEL START
    '���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

    pfFormActive (frmPass.hwnd)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdKakutei_Click
'//  �@�\����  : �u�m��v�t����������
'//  �@�\�T�v  : ���̓p�X���[�h���`�F�b�N���A
'//              �G���[�F�u�p�X���[�h�G���[�v�|�b�v�A�b�v��\���B
'//              ���@��F�ێ��ʂ�\�� �B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                �E�t�H�[���A�����[�h�����ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-23   REVISED BY [TCC] S.Terao
'//                 �E��ʕ\��/�����^�C�~���O�C��
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yoshimori
'//                 �p�X���[�h�s��v�̉�ʑJ�ڕύX
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(EG20 V30.3.0.1)  2014-09-18  REVISED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή�
'//                 �yHKRK_Kansi07_007_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdKakutei_Click()
    Dim iResponse As Integer      '���b�Z�[�W�{�b�N�X�\���߂�l
    Dim bFlag As Boolean          '���̓p�X���[�h�`�F�b�N�t���O(True�F��v�BFalse:�s��v)
    Dim intPassFileNo As Integer  '�p�X���[�h�t�@�C���̃t�@�C���ԍ�
    Dim sPassword As String       '�p�X���[�h�t�@�C���̂P�s���̃f�[�^
    Dim sPassData As String       '�V�X�e�������B
    Dim sHoshuPass As String      '�A���p�X���[�h=�t�@�C����`�p�X+�V�X�e������
    Dim sPass As String           '�t�@�C����`�p�X
    
    '�u�p�X���[�h���͉�ʁF�m��t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKUTEI_BUTTOM, 0)
 
    '���̓p�X���[�h�`�F�b�N�t���O��s��v�ŏ���������B
    bFlag = False
    
    '�����͂ł́u�m��v�t������
    If txtPass = "" Or IsNull(txtPass) Then
       '�u�p�X���[�h���͉�ʁF�p�X���[�h�����ُ͈�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_NOT_KEY, 0)
        'V1.20.0.1 DEL START
        ''���[����M�^�C�}���~����
        'tmrMail.Enabled = False
        ''�u�p�X���[�h���͉�ʁF�p�X���[�h���͉�ʏ����v���O�o��
        ' Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
        ''�I���������s��
        'psEndProc
        ' '�p�X���[�h���͉�ʂ����B
        'Unload Me
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
        '�G���[���b�Z�[�W��\������B
        iResponse = MsgBox("�p�X���[�h���Ⴂ�܂��B", _
                            vbOKOnly, _
                            "�p�X���[�h�G���[")
        'V1.20.0.1 ADD END
        Exit Sub
    End If
    
    '�p�X���[�h�����`�F�b�N���s���B
    If (Len(txtPass) <= INPUT_PASSWORD_MAX) And (Len(txtPass) >= INPUT_PASSWORD_MIN) Then
       On Error GoTo FileError
       '���g�p�̃t�@�C���ԍ����擾����B
       intPassFileNo = FreeFile
       '�p�X���[�h�t�@�C�����J���B
       Open PASSWORD_FILE_FULLPASS For Input As #intPassFileNo
       '�V�X�e���������uMMDD�v�Ŏ擾����B
       sPassData = Format(Date, "mmdd")
       '�t�@�C���̏I�[�܂ŌJ��Ԃ��B
'       Do While Not EOF(1)                                 ' EG20 V3.3.0.1�폜
       Do While Not EOF(intPassFileNo)                      ' EG20 V3.3.0.1�ǉ�
         '�p�X���[�h�t�@�C���̐擪����P�s���Ǎ��ށB
         Line Input #intPassFileNo, sPassword
         '�p�X���[�h�̒�`�ݒ肪���邩�ǂ����`�F�b�N����B
         If sPassword <> "" Then
           '�p�X���[�h�t�@�C���ɂ����`�Ǝ擾������A��������B
           sPass = Mid(sPassword, 3, 8)
           If sPass <> "" Then
              sHoshuPass = sPass + sPassData
              '�A���p�X���[�h�ƁA���͒l���r����B
              If sHoshuPass = txtPass Then
                 '��v�����ꍇ�A�p�X���[�h�`�F�b�N�t���O����v�ɂ���B
                 bFlag = True
                 '��v�����p�X���[�h���̃��[�U���x�����擾����B
                 pbUserLevel = CInt(Left$(sPassword, 1))
                 'Exit Do       'EG20 V30.3.0.1 DEL �yHKRK_Kansi07_007_01�z
                 'EG20 V30.3.0.1 ADD START �yHKRK_Kansi07_007_01�z
                 '���[�U���x�����u��ʁv�u�����v�̏ꍇ�݈̂�v�����Ƃ��A���[�v�𔲂���B
                 If pbUserLevel = 0 Or pbUserLevel = 1 Then
                    Exit Do
                 Else
                    '����ȊO�̓��[�v�𔲂����ɏ����𑱂���B
                    '�p�X���[�h�G���[�����Ȃ̂�pbUserLevel�̓��Z�b�g����B
                    '�������Ȃ��ƁApbUserLevel��2���Z�b�g����Ă����ꍇ�A�߂�t�������Ƀv���Z�X�I���v���̕⏕�G���A��2���Z�b�g����Ă��܂��A
                    '�Ɩ��n�ێ��ʂ��\������Ă��܂��B
                    pbUserLevel = 0
                    bFlag = False
                 End If
                 'EG20 V30.3.0.1 DEL END �yHKRK_Kansi07_007_01�z
              End If
           End If
         End If
       Loop
     '�p�X���[�h�t�@�C�������B
     Close #intPassFileNo
    End If
    
    '�p�X���[�h�G���[�̏ꍇ
    If bFlag = False Then
        '�G���[���b�Z�[�W��\������B
        iResponse = MsgBox("�p�X���[�h���Ⴂ�܂��B", _
                            vbOKOnly, _
                            "�p�X���[�h�G���[")
        '�u�p�X���[�h���͉�ʁF���̓p�X���[�h�ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_KEY_ERROR, 0)
        'V1.20.0.1 DEL START
        ''���[����M�^�C�}���~����B
        'tmrMail.Enabled = False
        ''�I���������s���B
        'psEndProc
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
        '���̓p�X���[�h������������
        txtPass = ""
        '�p�X���[�h���͉�ʂ�����A�������I������
        Exit Sub
        'V1.20.0.1 ADD END
        
    '�p�X���[�h����̏ꍇ
    Else
        '�@��\���f�[�^�N���X��������
        Call MakeInitKikiClas
        
        '�p�X���[�h���̓��O��ێ瑀�샍�O�ɏo�͂���B
        If pbUserLevel = 0 Then  '���[�U���x��
          '�u�p�X���[�h���͉�ʁF���̓p�X���[�h����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_OK, 0)
        ElseIf pbUserLevel = 1 Then
         '�u�p�X���[�h���͉�ʁF���̓p�X���[�h����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_OK, 0)
        Else
            'EG20 V30.3.0.1 ADD START�yHKRK_Kansi07_007_01�z
            '���[�U���x����0�A�P�ȊO��bFlag��False�ɂ��Ă��܂��Ă���̂ŁA���������s����邱�Ƃ͂Ȃ����O�̂��߁A
            '�Ɩ��n�ێ炪�N�����Ȃ��悤�ɁA�����ł��p�X���[�h�s��v�����Ƃ���
            iResponse = MsgBox("�p�X���[�h���Ⴂ�܂��B", _
                                vbOKOnly, _
                                "�p�X���[�h�G���[")
            '�u�p�X���[�h���͉�ʁF���̓p�X���[�h�ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_KEY_ERROR, 0)
            '���̓p�X���[�h������������
            txtPass = ""
            '�p�X���[�h���͉�ʂ�����A�������I������
            Exit Sub
            'EG20 V30.3.0.1 ADD END�yHKRK_Kansi07_007_01�z
            'EG20 V30.3.0.1 DEL START �yHKRK_Kansi07_007_01�z
'         '�u�p�X���[�h���͉�ʁF���̓p�X���[�h����v���O�o��
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_OK, 0)
'           '���ꃆ�[�U�̏ꍇ�͋Ɩ��I�����A
'           '�Ǘ��Ƀv���Z�X�I���v��Ұ�(�⏕���հ�����ق�ݒ�)�𑗐M����B
'            tmrMail.Enabled = False
'            '�u�p�X���[�h���͉�ʁF�p�X���[�h���͉�ʏ����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
'            Call psEndProc
'            Unload Me
'            Exit Sub
            'EG20 V30.3.0.1 DEL END �yHKRK_Kansi07_007_01�z
        End If
        'V1.6.0.1 ADD START
        '�����e�i���X���j���[��ʂ�\������
        frmHoshu.Show  'V1.6.0.1 DEL
        'V1.6.0.1 ADD END
        '�p�X���[�h���͉�ʂ��A�N�e�B�u�\���ɂ���
        '�u�p�X���[�h���͉�ʁF�p�X���[�h���͉�ʏ����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
        Me.Hide
        '�����e�i���X���j���[��ʂ�\������
        'frmHoshu.Show  'V1.6.0.1 DEL
    End If
' V1.3.0.1 ADD START
        '�p�X���[�h���͉�ʂ����B
        Unload Me
' V1.3.0.1 ADD END
  Exit Sub

FileError:
   '�u�p�X���[�h���͉�ʁF�p�X���[�h���͉�ʏ����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
    '�p�X���[�h���͉�ʂ����B
   Unload Me
 End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�Ď���ʂ֖߂�v�t����������
'//  �@�\�T�v  : �v���Z�X�I�����������A��ʂ����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
  On Error Resume Next
  
  '���[����M�^�C�}���~����
  tmrMail.Enabled = False
  '�u�p�X���[�h���͉�ʁF�p�X���[�h���͉�ʏ����v���O�o��
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_KEY_GAMEN_END, 0)
  
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    If CheckAppStart(PROC_KANRI) = 0 Then
        '�Ǘ��v���Z�X���N�����Ă��Ȃ��ꍇ
        psEndHoshuProc
    Else
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
        '�I���������s��
        psEndProc
    End If          ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
  '�p�X���[�h���͉�ʂ����B
  Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdNumber_Click
'//  �@�\����  : �e���L�[�e�t����������
'//  �@�\�T�v  : �������ꂽ�L�[�ɏ]���āA�p�X���[�h���͗����X�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����L�[�̎��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdNumber_Click(Index As Integer)
    
    '�a�r�L�[������
    If Index = 10 Then
        '���͒l�L���`�F�b�N���s���B
        If (txtPass <> "") Then
            '���͒l���L�����ꍇ�A�������͒l���P�����폜����B
            txtPass = Left(txtPass, Len(txtPass) - 1)
        End If
        '�����I���B
        Exit Sub
    End If
    
    '�b�L�[������
    If Index = 11 Then
        '���͒l��S�č폜����B
        txtPass = ""
        '�����I���B
        Exit Sub
    End If
    
    '�O�`�X�܂ł̐����L�[���������A���͍ςݕ�����̖����ɒǉ�����B
    txtPass = txtPass & Index
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v����
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
    
    '�ėp���[����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmPass.Caption, False
        pfFormActive (frmPass.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sPassFileInitialize
'//  �@�\����  : �p�X���[�h�t�@�C����������
'//  �@�\�T�v  : �p�X���[�h�t�@�C�������݂��Ȃ���΍쐬���A
'//              �f�t�H���g�̃p�X���[�h���i�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sPassFileInitialize()
    Dim intPassFileNo As Integer  '�p�X���[�h�t�@�C���̃t�@�C���ԍ�
    Dim iLine As Integer          '�p�X���[�h�t�@�C���̍s�J�E���^
    Dim sPassword As String       '�p�X���[�h�t�@�C���̂P�s���̃f�[�^
    
    '�s�J�E���^��0�ɂď���������B
    iLine = 0
    On Error GoTo FileError
    '���g�p�̃t�@�C���ԍ����擾����B
    intPassFileNo = FreeFile
    '�p�X���[�h�t�@�C�����J���B
    Open PASSWORD_FILE_FULLPASS For Input As #intPassFileNo
    '�t�@�C���̏I�[�܂ŌJ��Ԃ��B
'    Do While Not EOF(1)                                    ' EG20 V3.3.0.1�폜
    Do While Not EOF(intPassFileNo)                         ' EG20 V3.3.0.1�ǉ�
      '�p�X���[�h�t�@�C���ɗL���Ȓ�`�ݒ肪�L�邩�`�F�b�N���邽�߁A1�s���ǂݍ��ށB
      Line Input #intPassFileNo, sPassword
        '��`�ݒ肪����ꍇ�A�s�J�E���^�[���J�E���g�A�b�v����B
        If sPassword <> "" Then
            iLine = iLine + 1
            Exit Do
        End If
    Loop

    '�p�X���[�h�t�@�C�������B
    Close #intPassFileNo
'Exit Sub  'V1.7.0.1 DEL

FileError:
    '�u�p�X���[�h���͉�ʁF�f�t�H���g�p�X���[�h�t�@�C���쐬�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_KEY_FILE_CREATE, 0)
    '�s�J�E���^�[��0�̏ꍇ(=��`�ݒ薳��)
    If iLine = 0 Then
        '�p�X���[�h�t�@�C�����J���B
        Open PASSWORD_FILE_FULLPASS For Output As #intPassFileNo
        '�f�t�H���g�̃p�X���[�h�u '�������[�U�p�F"�P"�v���������ށB
        Print #intPassFileNo, "1,1"
        '�p�X���[�h�t�@�C�������B
        Close #intPassFileNo
    End If
End Sub
 
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sGetRegIDU_LDU_Path
'//  �@�\����  : IDU/LDU�̃p�X�����W�X�g�����擾����B
'//  �@�\�T�v  : IDU/LDU�̃p�X�����W�X�g�����擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   REVISED BY [TCC] C.Terui
'//                 �E���W�X�g���̃f�t�H���g�l�ݒ�y�я����ύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sGetRegIDU_LDU_Path()
    
    On Error Resume Next
    
    sGetRegIDU_LDU_Path = False
    
    'IDU�F�A�v���p�X�擾
    PATH_IDU_APP = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "AplRoot")
    If PATH_IDU_APP = "" Then
       '�u���W�X�g�����擾�ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       '���W�X�g���̃f�t�H���g�l���擾
       PATH_IDU_APP = REG_IDU_APLROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If

    'IDU�FDB�p�X�擾
    PATH_IDU_DB = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "DataRoot")
    If PATH_IDU_DB = "" Then
       '�u���W�X�g�����擾�ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       '���W�X�g���̃f�t�H���g�l���擾
       PATH_IDU_DB = REG_IDU_DBROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If
    
    'IDU�F�o�b�N�A�b�v�p�X�擾
    PATH_BUC = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "BackupRoot")
    If PATH_BUC = "" Then
       '�u���W�X�g�����擾�ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       '���W�X�g���̃f�t�H���g�l���擾
       PATH_BUC = REG_IDU_BACKUPROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If

    'IDU�F���O�p�X�擾
    PATH_IDU_LOG = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\ID_RelayUnit", "LogRoot")
    If PATH_IDU_LOG = "" Then
       '�u���W�X�g�����擾�ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       '���W�X�g���̃f�t�H���g�l���擾
       PATH_IDU_LOG = REG_IDU_LOGROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If

'    LDU�A�v���p�X�擾
    PATH_LDU_APP = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\LD_Utility", "AplRoot")
    If PATH_LDU_APP = "" Then
      '�u���W�X�g�����擾�ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       '���W�X�g���̃f�t�H���g�l���擾
       PATH_LDU_APP = REG_LDU_APLROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If
    
'    LDU���O�p�X�擾
    PATH_LDU_LOG = pfGetReg(HKEY_LOCAL_MACHINE, "SOFTWARE\TOSHIBA\LD_Utility", "LogRoot")
    If PATH_LDU_LOG = "" Then
      '�u���W�X�g�����擾�ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, GET_REGDATA_ERROR, 0)
' V1.3.0.1 ADD START
       '���W�X�g���̃f�t�H���g�l���擾
       PATH_LDU_LOG = REG_LDU_LOGROOT
' V1.3.0.1 ADD END
'       Exit Function                       ' V1.3.0.1 DEL
    End If
    
    sGetRegIDU_LDU_Path = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : MakeInitKikiClas
'//  �@�\����  : �@��\���f�[�^�N���X��������
'//  �@�\�T�v  : �@��\���f�[�^�N���X�𐶐�����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.1.0.1) 2010-05-28   REVISED BY [TCC] S.Yoshimori
'//                 �P���b�`�����g�p�w�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub MakeInitKikiClas()

    Dim lErrCode As Long          '�G���[�R�[�h
    Dim bRet As Boolean           '�֐��߂�l
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V2.1.0.1 ADD

    '---------------------------------------------
    '�@��\���f�[�^�N���X����
    '---------------------------------------------
    bRet = dllInitKikiClass(lErrCode)
    
    '���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_CREATE_KIKICLASS, lErrCode)
    
    '---------------------------------------------
    '�w�s�x�f�[�^�R�t���t�@�C���������W�J
    '---------------------------------------------
    bRet = dllMemEkiDataChange(EKI_DATA_CHANGE_FILE, lErrCode)
    
    If bRet = False Then
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_CREATE_EKITUDOMAP_ERR, lErrCode)
    Else
        '���탍�O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_CREATE_EKITUDOMAP, 0)
    End If
    
    'V2.1.0.1 ADD START
    '---------------------------------------------
    '�w�^�C�v�s�x�f�[�^�R�t���t�@�C���������W�J
    '---------------------------------------------
    '�w�^�C�v�s�x�f�[�^�R�t���t�@�C�������݂���Ȃ珈�����s�Ȃ�
    If (objFso.FileExists(EKI_TYPE_DATA_CHANGE_FILE) = True) Then

        '�w�^�C�v�s�x�f�[�^�R�t���t�@�C���������W�J�֐�
        bRet = dllMemEkiTypeDataChange(EKI_TYPE_DATA_CHANGE_FILE, lErrCode)
        
        If (False = bRet) Then
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_CREATE_EKITYPE_TUDOMAP_ERR, lErrCode)
        Else
            '���탍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_CREATE_EKITYPE_TUDOMAP, 0)
        End If
    Else
        '�w�^�C�v�s�x�f�[�^�R�t���t�@�C���Ȃ�
        '���O�o��
        Call sLogTraceReq(LTYP_WARNING, L3AN_FILE, PASS_NOT_EKITYPE_TUDOFILE, 0)
    End If
    
    Set objFso = Nothing
    'V2.1.0.1 ADD END
End Sub

