VERSION 5.00
Begin VB.Form frmIDUSysformat 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "                                                                    �V�X�e���������@�\"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrLogTimer 
      Left            =   11400
      Top             =   6720
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   8640
      Top             =   7920
   End
   Begin VB.Timer tmrMail 
      Left            =   8640
      Top             =   6120
   End
   Begin VB.CommandButton cmdZikko 
      Caption         =   "���������s"
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
      Left            =   9120
      TabIndex        =   18
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ListBox LstStatus 
      Height          =   2985
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   8415
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   15000
      Width           =   2895
   End
   Begin VB.Frame frmSentaku 
      Caption         =   "���������ڎw��"
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   11775
      Begin VB.Frame FraKoumoku 
         Height          =   615
         Left            =   1200
         TabIndex        =   31
         Top             =   240
         Width           =   10455
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "�S�ď������i�c�k�k�f�[�^�܂ށj"
            Height          =   375
            Index           =   2
            Left            =   5160
            TabIndex        =   34
            Top             =   200
            Width           =   4215
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "���ڑI��"
            Height          =   375
            Index           =   1
            Left            =   2640
            TabIndex        =   33
            Top             =   200
            Width           =   1575
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "�o�׎�������"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   200
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   8
         Left            =   8160
         Style           =   1  '���̨���
         TabIndex        =   30
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   7
         Left            =   8160
         Style           =   1  '���̨���
         TabIndex        =   29
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   6
         Left            =   4080
         Style           =   1  '���̨���
         TabIndex        =   28
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   5
         Left            =   4080
         Style           =   1  '���̨���
         TabIndex        =   27
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   4
         Left            =   4080
         Style           =   1  '���̨���
         TabIndex        =   26
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   3
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   25
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   2
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   24
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   23
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   0
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.Frame frmKoumoku 
         Caption         =   "����"
         Height          =   3855
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   11535
         Begin VB.Frame FraShosai 
            Caption         =   "���ڏڍ�"
            Height          =   1695
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   11295
            Begin VB.Label LblShosai 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1330
               Left            =   100
               TabIndex        =   21
               Top             =   220
               Width           =   11050
            End
         End
         Begin VB.CheckBox chkSonota 
            Caption         =   "���̑��f�[�^"
            Height          =   255
            Left            =   8880
            TabIndex        =   15
            Top             =   1560
            Value           =   1  '����
            Width           =   2175
         End
         Begin VB.Frame frmDLL 
            Caption         =   "�c�k�k�f�[�^"
            Height          =   975
            Left            =   7920
            TabIndex        =   13
            Top             =   360
            Width           =   3135
            Begin VB.CheckBox chkDLL 
               Height          =   375
               Left            =   960
               TabIndex        =   14
               Top             =   360
               Width           =   2055
            End
         End
         Begin VB.Frame frmLog 
            Caption         =   "���O�f�[�^"
            Height          =   1575
            Left            =   3840
            TabIndex        =   9
            Top             =   360
            Width           =   4035
            Begin VB.CheckBox chkLog 
               Caption         =   "����h�b���W���[�����O"
               DataField       =   "3"
               Height          =   375
               Index           =   2
               Left            =   960
               TabIndex        =   12
               Top             =   1080
               Value           =   1  '����
               Width           =   2955
            End
            Begin VB.CheckBox chkLog 
               Caption         =   "�ێ�v���O�������O"
               DataField       =   "2"
               Height          =   375
               Index           =   1
               Left            =   960
               TabIndex        =   11
               Top             =   720
               Value           =   1  '����
               Width           =   2535
            End
            Begin VB.CheckBox chkLog 
               Caption         =   "�A�v���P�[�V�������O"
               DataField       =   "1"
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   10
               Top             =   360
               Value           =   1  '����
               Width           =   2715
            End
         End
         Begin VB.Frame frmMeisai 
            Caption         =   "�h�b�ꌏ����"
            Height          =   1575
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   3615
            Begin VB.CheckBox chkMeisai 
               Caption         =   "�đ��f�[�^"
               Height          =   375
               Index           =   2
               Left            =   960
               TabIndex        =   8
               Top             =   1080
               Value           =   1  '����
               Width           =   2535
            End
            Begin VB.CheckBox chkMeisai 
               Caption         =   "�o�b�N�A�b�v�f�[�^"
               Height          =   375
               Index           =   1
               Left            =   960
               TabIndex        =   7
               Top             =   720
               Value           =   1  '����
               Width           =   2535
            End
            Begin VB.CheckBox chkMeisai 
               Caption         =   "�c�a�f�[�^"
               DataField       =   "0"
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   6
               Top             =   360
               Value           =   1  '����
               Width           =   2535
            End
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�V�X�e��������  ��ʂ֖߂�"
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
      Left            =   9120
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C000&
      Caption         =   "IDU�A�v���P�[�V�����V�X�e��������"
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
      TabIndex        =   19
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label lblKekka 
      BorderStyle     =   1  '����
      Caption         =   "�������͐������܂����B"
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
      Left            =   8760
      TabIndex        =   17
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Caption         =   "����������"
      Height          =   255
      Left            =   8760
      TabIndex        =   16
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "frmIDUSysformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmIDUSysformat.frm
'//  �p�b�P�[�W���F�V�X�e��������(IDU)���
'/
'//  �T�v�F�V�X�e��������(IDU)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�ۑ��p�ݒ�t�@�C���쐬�����ǉ�
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 ���X�g�{�b�N�X�̃X�N���[�������ǉ�
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ��A�m�F�|�b�v�A�b�v�C��
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'Private bChk() As Boolean              'V1.5.0.1 DEL

'���������s�t���O
Private bSysFormat As Boolean

Private ShosaiMoji(0 To 8) As String '�ڍו����i�[�G���A
Private Const SYSMOJI_SIZE = 500
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000     '�A�v���N���^�C�}�f�t�H���g�l
Dim lngMAX_Time As Long                 'INI�擾�ݒ�l
Dim lngtime     As Long                 '���݃^�C�}�l
Private bChk(8) As Boolean
'V1.5.0.1 ADD END
'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000     '���O�N���^�C�}�f�t�H���g�l(30�b)
Dim lngLogMAX_Time As Long                'INI�擾�ݒ�l(���O�j
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �V�X�e��������(IDU)���(�A�N�e�B�u��)
'//  �@�\�T�v  : �őO�ʕ\�����s���B
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
Private Sub Form_Activate()
    pfFormActive (hwnd)
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �V�X�e��������(IDU)���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
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
Private Sub Form_Deactivate()
   On Error Resume Next
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �V�X�e��������(IDU)���(���[�h��)
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
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//             �@�@�t�F�[�Y�Q�Ή��@IDU�k�ދ@�\�`�F�b�N�ǉ�
'//     REVISIONS  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim ii  As Integer
    
    On Error Resume Next
    
    '�uID���p�Ưļ��я�������ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_SYSFORMAT_GAMEN_START, 0)
    
    gStrCurrentForm = sFormName_IDUSys
    
    '�u�ڍׁv�t���������擾����
    ShosaiMongonGet

    '������
    OptShosai(0).Value = True   '���������ڎw��F�ڍזt����
    LstStatus.Clear             '�폜�t�@�C���\�����N���A
    OptKoumoku(0).Value = True  '���������ڎw��F�u�o�׎��������v�w��L��I��
    chkMeisai(0).Value = 1      'IC�ꌏ���ׁFDB�f�[�^�`�F�b�N�L��
    chkMeisai(1).Value = 1      'IC�ꌏ���ׁF�o�b�N�A�b�v�f�[�^�`�F�b�N�L��
    chkMeisai(2).Value = 1      'IC�ꌏ���ׁF�đ��f�[�^�`�F�b�N�L��
    chkLog(0).Value = 1         '���O�f�[�^�F�A�v���P�[�V�������O
    chkLog(1).Value = 1         '���O�f�[�^�F�ێ�v���O�������O
    chkLog(2).Value = 1         '���O�f�[�^�F����IC���W���[�����O
    chkSonota.Value = 1         '���̑��f�[�^
    lblKekka.Caption = ""       '���������s�\�����N���A
    
    frmKoumoku.Enabled = False
    frmMeisai.Enabled = False
    frmLog.Enabled = False
    chkMeisai(0).Enabled = False 'IC�ꌏ���ׁFDB�f�[�^�I��s��
    chkMeisai(1).Enabled = False 'IC�ꌏ���ׁF�o�b�N�A�b�v�f�[�^�I��s��
    chkMeisai(2).Enabled = False 'IC�ꌏ���ׁF�đ��f�[�^�I��s��
    chkLog(0).Enabled = False    '���O�f�[�^�F�A�v���P�[�V�������O�I��s��
    chkLog(1).Enabled = False    '���O�f�[�^�F�ێ�v���O�������O�I��s��
    chkLog(2).Enabled = False    '���O�f�[�^�F����IC���W���[�����O�I��s��
    chkSonota.Enabled = False    '���̑��f�[�^�I��s��
            
    fraKoumoku.BorderStyle = 0
    OptShosai(0).Enabled = True  '���������ڕ��F�ڍזt�����\
    OptShosai(0).Value = True    '���������ڕ��F�ڍזt����
    For ii = 1 To 8
        OptShosai(ii).Enabled = False  '���ڕ��F�ڍזt�����\
    Next
    
    OptKoumoku(2).Enabled = False
    frmDLL.Enabled = False
    chkDLL.Enabled = False       'DLL�f�[�^�I��s��
    '���O�C�����[�U�`�F�b�N
    If pbUserLevel = 1 Then
        OptKoumoku(2).Enabled = True
        frmDLL.Enabled = True
        chkDLL.Enabled = False
        chkDLL.Value = 1
    Else
        OptKoumoku(2).Enabled = False
        frmDLL.Enabled = False
        chkDLL.Enabled = False
        chkDLL.Value = 0
    End If
    '���������s�t���OOFF
    bSysFormat = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
   'V1.3.0.1 ADD START
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
   
   'V1.5.0.1 ADD START
   'INI�t�@�C�����A�v���N���^�C�}�l���擾
   lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      APL_INTERVAL, HOSHU_FILE)
   '�擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
   If lngMAX_Time = 0 Then
      lngMAX_Time = APL_INTERVAL
   End If
   '�^�C�}�l�ݒ�
   tmrAplTimer.Interval = MN_MAIL_INTERVAL
   tmrAplTimer.Enabled = False
   'V1.5.0.1 ADD END

   'V1.20.0.1 ADD START
   'INI�t�@�C����胍�O�N���^�C�}�l���擾
   lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
   '�擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
   If lngLogMAX_Time = 0 Then
      lngLogMAX_Time = LOG_INTERVAL
   End If
   
   tmrLogTimer.Interval = MN_MAIL_INTERVAL
   tmrLogTimer.Enabled = False
   'V1.20.0.1 ADD END
   
   'V1.4.0.1 ADD START
   'IDU�k�ރ`�F�b�N
    psIDUCheck
   'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : OptKoumoku_Click
'//  �@�\����  : ���W�I�t����������
'//  �@�\�T�v  : ���������ڎw�蕔�F���W�I�t�������������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@ [IN]�������W�I�t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub OptKoumoku_Click(Index As Integer)
    Dim ii As Integer  '�J�E���^�[
    
    On Error Resume Next
   
    Select Case Index
        Case 1:        '���ڑI����
            OptShosai(0).Enabled = False  '���������ڎw��F�ڍזt�I��s��
            OptShosai(1).Value = True     '���ڎw��FDB�f�[�^�ڍזt����
            For ii = 1 To 7
                OptShosai(ii).Enabled = True  '���ڎw��F�ڍזt�I����
            Next
            
            frmKoumoku.Enabled = True
            frmMeisai.Enabled = True
            frmLog.Enabled = True
            chkMeisai(0).Enabled = True  'IC�ꌏ���ׁFDB�f�[�^�I���\
            chkMeisai(1).Enabled = True  'IC�ꌏ���ׁF�o�b�N�A�b�v�f�[�^�I���\
            chkMeisai(2).Enabled = True  'IC�ꌏ���ׁF�đ��f�[�^�I���\
            chkLog(0).Enabled = True     '���O�f�[�^�F�A�v���P�[�V�������O�I���\
            chkLog(1).Enabled = True     '���O�f�[�^�F�ێ�v���O�������O�I���\
            chkLog(2).Enabled = True     '���O�f�[�^�F����IC���W���[�����O�I���\
            chkSonota.Enabled = True     '���̑��f�[�^�I���\

            '���O�C�����[�U�`�F�b�N
            If pbUserLevel = 1 Then
                frmDLL.Enabled = True        'DLL�f�[�^�^�I���\
                chkDLL.Enabled = True
                OptShosai(8).Enabled = True  'DLL�f�[�^�ڍזt�����\
            End If
            '�uID���p�Ưļ��я�������ʁF���ڑI��I�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_KOUMOKU, 0)
        Case Else:
            OptShosai(0).Enabled = True          '���������ڎw��F�ڍזt�I���\
            OptShosai(0).Value = True            '���������ڎw��F�ڍזt����
            For ii = 1 To 7
                OptShosai(ii).Enabled = False    '���ڕ��F�ڍזt�I��s�\
                OptShosai(ii).Value = False      '���ڕ��F�ڍזt�I�𖢉���
            Next
            frmKoumoku.Enabled = False
            frmMeisai.Enabled = False
            frmLog.Enabled = False
            chkMeisai(0).Enabled = False         'IC�ꌏ���ׁFDB�f�[�^�I��s��
            chkMeisai(1).Enabled = False         'IC�ꌏ���ׁF�o�b�N�A�b�v�f�[�^�I��s��
            chkMeisai(2).Enabled = False         'IC�ꌏ���ׁF�đ��f�[�^�I��s��
            chkLog(0).Enabled = False            '���O�f�[�^�F�A�v���P�[�V�������O�I��s��
            chkLog(1).Enabled = False            '���O�f�[�^�F�ێ�v���O�������O�I��s��
            chkLog(2).Enabled = False            '���O�f�[�^�F����IC���W���[�����O�I��s��
            chkSonota.Enabled = False            '���̑��f�[�^�I��s��

            '���O�C�����[�U�`�F�b�N
            If pbUserLevel = 1 Then
                frmDLL.Enabled = False           'DLL�f�[�^�I��s��
                chkDLL.Enabled = False
                OptShosai(8).Enabled = False     'DLL�f�[�^�ڍזt�����s��
            End If
            If Index = 0 Then
               '�uID���p�Ưļ��я�������ʁF�o�׎��������I�����v���O�o��
               Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_SHUKKA, 0)
            Else
               '�uID���p�Ưļ��я�������ʁF�S�ď������I�����v���O�o��
               Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_ALL, 0)
            End If
        End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdZikko_Click
'//  �@�\����  : ���������s�t��������
'//  �@�\�T�v  : ���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�ۑ��p�ݒ�t�@�C���쐬�����ǉ�
'//     REVISIONS  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ��A�m�F�|�b�v�A�b�v�C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZikko_Click()
    Dim i As Integer
    Dim bRtn As Boolean
    Dim bSentaku As Boolean
    Dim iRet As Integer
    Dim sLine As String
    Dim lRetVal As Long
    Dim lExitCode As Long
    Dim sExecName As String
    Dim sDbInitCmd As String
    'ReDim bChk(8)                                'V1.5.0.1 DEL
    Dim bRtn1 As Boolean
    Dim bRtn2 As Boolean
    Dim uMail As MAIL_IDU_LDU_APLEND_CMD           'IDU�A�v���I���v��
    Dim iRetApp         As Integer                 'IDU�A�v���I���t���O
    Dim iRetLog         As Integer                 'IDU���O�I���t���O
    Dim uIduEndMail As MAIL_IDU_LDU_LOGEND_CMD     'IDU���O�v���Z�X�I���v��
    Dim lngErrCode As Long                      '�G���[�R�[�h
    Dim iTargetDB As Integer                       '�Ώ�DB�l
    Dim bDB_Code As Boolean
    'V1.5.0.1  ADD START
    Dim bIDUPCRet    As Boolean            'IDU�A�v����������
    Dim bIDULOGRet   As Boolean            'IDU���O��������
    
    bIDUPCRet = False
    bIDULOGRet = False
    'V1.5.0.1  ADD END
    On Error GoTo ERR_SPACE
    
    '�uID���p�Ưļ��я�������ʁF���������s�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_START_BUTTOM, 0)

    '�\���̏�����
    LstStatus.Clear
    lblKekka.Caption = ""

    '�o�׎��������I����
    If OptKoumoku(0).Value = True Then
        For i = 2 To 8
            bChk(i) = True
        Next
        bChk(1) = False
    End If

    '���ڑI��I����
    If OptKoumoku(1).Value = True Then
        bSentaku = False
        '�h�b�ꌏ����
        '�c�a�f�[�^
        If chkMeisai(0).Value = 1 Then
            bSentaku = True
            bChk(2) = True
        Else
            bChk(2) = False
        End If
        '�o�b�N�A�b�v�f�[�^
        If chkMeisai(1).Value = 1 Then
            bSentaku = True
            bChk(3) = True
        Else
            bChk(3) = False
        End If
        '�đ��f�[�^
        If chkMeisai(2).Value = 1 Then
            bSentaku = True
            bChk(4) = True
        Else
            bChk(4) = False
        End If

        '���O�f�[�^
        '�A�v���P�[�V�������O
        If chkLog(0).Value = 1 Then
            bSentaku = True
            bChk(5) = True
        Else
            bChk(5) = False
        End If
        '�ێ�v���O�������O
        If chkLog(1).Value = 1 Then
            bSentaku = True
            bChk(6) = True
        Else
            bChk(6) = False
        End If
        '����h�b���W���[�����O
        If chkLog(2).Value = 1 Then
            bSentaku = True
            bChk(7) = True
        Else
            bChk(7) = False
        End If

        '���̑��f�[�^
        If chkSonota.Value = 1 Then
            bSentaku = True
            bChk(8) = True
        Else
            bChk(8) = False
        End If

        '�c�k�k�f�[�^
        If chkDLL.Value = 1 Then
            bSentaku = True
            bChk(1) = True
        Else
            bChk(1) = False
        End If

        If bSentaku = False Then
            '�uID���p�Ưļ��я�������ʁF���������������s�v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
            MsgBox "����������f�[�^���I������Ă��܂���", vbExclamation, "�f�[�^���x��"
            Exit Sub
        End If
    End If

    '�S�ď������i�c�k�k�f�[�^�܂ށj�I����
    If OptKoumoku(2).Value = True Then
        For i = 1 To 8
            bChk(i) = True
        Next
    End If
    
    iRet = MsgBox("�������������s���܂��B��낵���ł����H", vbExclamation + vbOKCancel, "�������m�F")
    If iRet = vbOK Then
        OptKoumoku(0).Enabled = False
        OptKoumoku(1).Enabled = False
        cmdZikko.Enabled = False
        cmdCancel.Enabled = False
        '���O�C�����[�U�`�F�b�N
        If pbUserLevel = 1 Then
         OptKoumoku(2).Enabled = False
        End If
        
        On Error GoTo ERR_SPACE2
    
        '����ŏ�����
        iRetApp = 1
        iRetLog = 1

        '�A�v���N���`�F�b�N
        If CheckAppStart(PROCESS_IDU_PC) = 1 Then
          'V1.20.0.1 DEL START
'          iRet = MsgBox("ID���p���j�b�g�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F")
'          If iRet = vbOK Then
          'V1.20.0.1 DEL END
             'IDU�A�v���I���v����ID���ɑ��M����
              uMail.mlHeader.dwId = ML_ID_IDU_APLEND_CMD
              uMail.mlHeader.dwSize = MlSize.IDUAPLEND_REQ
              uMail.mlHeader.dwProid = RHOSHU_ID
              uMail.mlHeader.dwSubArea = 0
              uMail.dwEndType = ML_ENDTYPE_APLEND
              uMail.dwCMDLevel = ML_CMDLEVEL_TUJYO        'V1.5.0.1 ADD
              'V1.5.0.1 DEL START
              'bRtn = DssSendMail(MAIL_SLOT_IDSEI, Len(uMail), uMail.mlHeader)
              'If bRtn = 0 Then
              'V1.5.0.1 DEL END
              'V1.5.0.1 ADD START
              bIDUPCRet = DssSendMail(MAIL_SLOT_IDSEI, Len(uMail), uMail.mlHeader)
              If bIDUPCRet = 0 Then
              'V1.5.0.1 ADD END
                 '�uID���p�Ưļ��я�������ʁF���[�����M�ُ�v���O�o��
                 lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                 Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lngErrCode)
                 GoTo ERR_SPACE2:
              Else
                 '�uID���p�Ưļ��я�������ʁF���[�����M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                'iRetApp = CheckAppEndComplete(PROCESS_IDU_PC, lExitCode)    'V1.5.0.1 DEL
              End If
     'V1.20.0.1 DEL START
'               'IDU���O�I���v��CMD���M
'               'V1.5.0.1 DEL START
'                'bRtn = EndIDULog
'                'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'                bIDULOGRet = EndIDULog
'                If bIDULOGRet = False Then
'               'V1.5.0.1 ADD END
'                  '���M�ُ�
'                  lblKekka.ForeColor = SYSFORMAT_ERROR
'                  lblKekka.Caption = "�������Ɏ��s���܂���"
'                  OptKoumoku(0).Enabled = True
'                  OptKoumoku(1).Enabled = True
'                  cmdZikko.Enabled = True
'                  cmdCancel.Enabled = True
'                  '���O�C�����[�U�`�F�b�N
'                  If pbUserLevel = 1 Then
'                     OptKoumoku(2).Enabled = True
'                  End If
'                  '�����𔲂���
'                  Exit Sub
'                End If
'
'               'IDU���O�v���Z�X�I���m�F
'               'iRetLog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)  'V1.5.0.1 DEL
'          Else
'             '�uID���p�Ưļ��я�������ʁF���������������s�v���O�o��
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'              OptKoumoku(0).Enabled = True
'              OptKoumoku(1).Enabled = True
'              cmdZikko.Enabled = True
'              cmdCancel.Enabled = True
'              '���O�C�����[�U�`�F�b�N
'              If pbUserLevel = 1 Then
'                OptKoumoku(2).Enabled = True
'              End If
'              '�����𔲂���
'              Exit Sub
'          End If
     'V1.20.0.1 DEL END
       Else
       bIDUPCRet = True                                 'V1.5.0.1 ADD
       
         '���O�v���Z�X�N���`�F�b�N
          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
             'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
             'V1.20.0.1 DEL START
'             iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
'             If iRet = vbOK Then
             'V1.20.0.1 DEL END
                'IDU���O�I���v��CMD���M
                'V1.5.0.1 DEL START
                'bRtn = EndIDULog
                'If bRtn = False Then
                'V1.5.0.1 DEL END
                'V1.5.0.1 ADD START
                bIDULOGRet = EndIDULog
                If bIDULOGRet = False Then
                'V1.5.0.1 ADD END
                  '���M�ُ�
                  lblKekka.ForeColor = SYSFORMAT_ERROR
                  lblKekka.Caption = "�������Ɏ��s���܂���"
                  OptKoumoku(0).Enabled = True
                  OptKoumoku(1).Enabled = True
                  cmdZikko.Enabled = True
                  cmdCancel.Enabled = True
                  '���O�C�����[�U�`�F�b�N
                  If pbUserLevel = 1 Then
                     OptKoumoku(2).Enabled = True
                  End If
                  '�����𔲂���
                  Exit Sub
                End If
               
               'IDU���O�v���Z�X�I���m�F
               'iRetLog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)    V1.5.0.1 DEL
         'V1.20.0.1 DEL START
'             Else
'               '�uID���p�Ưļ��я�������ʁF���������������s�v���O�o��
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'                OptKoumoku(0).Enabled = True
'                OptKoumoku(1).Enabled = True
'                cmdZikko.Enabled = True
'                cmdCancel.Enabled = True
'                '���O�C�����[�U�`�F�b�N
'                If pbUserLevel = 1 Then
'                   OptKoumoku(2).Enabled = True
'                End If
'                '�����𔲂���
'                Exit Sub
'             End If
'           'V1.5.0.1 ADD�@START
        'V1.20.0.1 DEL END
           Else
            bIDULOGRet = True
           'V1.5.0.1 ADD END
           End If
        End If

       '���������s�t���OON
        bSysFormat = True
'V1.5.0.1 ADD START
         'IDU�A�v���AIDU���O�̃��[�����M�������S�Đ��킾�����ꍇ�̂݁A�A�v���N���^�C�}���N�������A
         '�A�v���N���`�F�b�N�ɂ��A�v���̋N��/���N���𔻒f����B
'         If (bIDUPCRet = True) And (bIDULOGRet = True) Then            'V1.20.0.1 DEL
         If (bIDUPCRet = True) Then                                     'V1.20.0.1 ADD
            lngtime = 0
            lngtime = MN_MAIL_INTERVAL
            tmrAplTimer.Enabled = True
         Else
           'IDU�A�v���AIDU���O�̃��[�����M�ɂĂЂƂł��ُ킪�������ꍇ�A�������������ُ�I���Ƃ���B
           '�uID���p�Ưļ��я�������ʁF���������������s�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
           OptKoumoku(0).Enabled = True
           OptKoumoku(1).Enabled = True
           cmdZikko.Enabled = True
           cmdCancel.Enabled = True
           '���O�C�����[�U�`�F�b�N
           If pbUserLevel = 1 Then
              OptKoumoku(2).Enabled = True
           End If
           '�����𔲂���
           Exit Sub
         End If
   End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'       '�A�v���܂��̓��O�v���Z�X�ŏI�������Ɏ��s�����ꍇ
'       If (iRetApp <> 1) Or (iRetLog <> 1) Then
'         '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
'         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'         lblKekka.ForeColor = SYSFORMAT_ERROR
'         lblKekka.Caption = "�������Ɏ��s���܂���"
'         OptKoumoku(0).Enabled = True
'         OptKoumoku(1).Enabled = True
'         cmdZikko.Enabled = True
'         cmdCancel.Enabled = True
'         '���O�C�����[�U�`�F�b�N
'         If pbUserLevel = 1 Then
'            OptKoumoku(2).Enabled = True
'         End If
'         '�����𔲂���
'         Exit Sub
'       End If
'      'V1.4.0.1 ADD START
'      '�o�׎��������I�����A�S�ď�����(DLL�f�[�^��)�I�����A���̑��f�[�^���������Ώێ�
'      If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
'        If sCreateShokiFile = False Then
'           '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
'           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "�������Ɏ��s���܂���"
'           OptKoumoku(0).Enabled = True
'           OptKoumoku(1).Enabled = True
'           cmdZikko.Enabled = True
'           cmdCancel.Enabled = True
'           '���O�C�����[�U�`�F�b�N
'           If pbUserLevel = 1 Then
'              OptKoumoku(2).Enabled = True
'           End If
'           '�����𔲂���
'           Exit Sub
'        End If
'      End If
'      'V1.4.0.1 ADD END
'
'      '�V�X�e���t�@�C���̍폜
'      If bChk(8) = True Then
'         bRtn1 = sSysFileDelete()
'         DoEvents
'      Else
'         bRtn1 = True
'      End If
'
'      '�t�H���_�A�t�@�C���̍폜
'      If bRtn1 = True Then
'
'         If sFileDelete() = True Then
'
'            bDB_Code = True
'
'            'DB����������
'            'DB�f�[�^�FIC�ꌏ����
'             If bChk(2) = True Then
'                iTargetDB = chkMeisai(0).DataField
'                Me.LstStatus.AddItem "DB������:" & chkMeisai(0).Caption
'                DoEvents
'                bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DB�f�[�^�F�A�v�����O
'             If bChk(5) = True And bDB_Code = True Then
'               iTargetDB = chkLog(0).DataField
'               Me.LstStatus.AddItem "DB������:" & chkLog(0).Caption
'               DoEvents
'               bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DB�f�[�^�F�ێ烍�O
'             If bChk(6) = True And bDB_Code = True Then
'               iTargetDB = chkLog(1).DataField
'               Me.LstStatus.AddItem "DB������:" & chkLog(1).Caption
'               DoEvents
'               bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DB�f�[�^�F����IC���W���[�����O
'             If bChk(7) = True And bDB_Code = True Then
'                iTargetDB = chkLog(2).DataField
'                Me.LstStatus.AddItem "DB������:" & chkLog(2).Caption
'                DoEvents
'                bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'             End If
'
'            'DB�f�[�^�F�l�K���X�g(�o�׎��������̂ݗL��)
'             If OptKoumoku(1).Value = False And bDB_Code = True Then
'                iTargetDB = stsIDUNega
'                bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'             End If
'
'             '�߂�l������
'              If bDB_Code = True Then
'                 '�uID���p�Ưļ��я�������ʁF�V�X�e����������������v���O�o��
'                 Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
'                 lblKekka.ForeColor = SYSFORMAT_OK
'                 lblKekka.Caption = "�������͐������܂���"
'              Else
'                 '�uID���p�Ưļ��я�������ʁFDB�����������ُ�v���O�o��
'                 Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
'                 lblKekka.ForeColor = SYSFORMAT_ERROR
'                 lblKekka.Caption = "�������Ɏ��s���܂���"
'              End If
'        Else
'          '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
'          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'          lblKekka.ForeColor = SYSFORMAT_ERROR
'          lblKekka.Caption = "�������Ɏ��s���܂���"
'        End If
'    Else
'       '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
'       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'       lblKekka.ForeColor = SYSFORMAT_ERROR
'       lblKekka.Caption = "�������Ɏ��s���܂���"
'    End If
'
'    '����������I�����̏���
'    OptKoumoku(0).Enabled = True
'    OptKoumoku(1).Enabled = True
'    cmdZikko.Enabled = True
'    cmdCancel.Enabled = True
'    '���O�C�����[�U�`�F�b�N
'    If pbUserLevel = 1 Then
'       OptKoumoku(2).Enabled = True
'    End If
' End If
'V1.5.0.1 DEL END
Exit Sub

ERR_SPACE2:
        '�G���[�������̏���
        OptKoumoku(0).Enabled = True    '���������ڎw��F�o�׎��������I���\
        OptKoumoku(1).Enabled = True    '���������ڎw��F���ڑI��I���\
        cmdZikko.Enabled = True         '�u���������s�v�t�����\
        cmdCancel.Enabled = True        '�u���j���[��ʂ֖߂�v�t�����\
        '���O�C�����[�U�`�F�b�N
        If pbUserLevel = 1 Then
            OptKoumoku(2).Enabled = True
        End If
        '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "�������Ɏ��s���܂���"

ERR_SPACE:

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
   On Error Resume Next

   '�uID���p���j�b�g�V�X�e���������F�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_SYSFORMAT_GAMEN_END, 0)
   frmSysformatMenu.ZOrder
      
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFileDelete
'//  �@�\����  : �t�@�C���E�t�H���_�폜����
'//  �@�\�T�v  : �폜�Ώۃt�@�C���A�폜�Ώۃt�H���_�̍폜���s���B
'//
'//              �^        ����        �Ӗ�
'//   ����     :�Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :�Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@��ʍX�V����
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 ���X�g�{�b�N�X�̃X�N���[�������ǉ�
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sFileDelete()
    Dim iFileNo As Integer
    Dim sFileData As String
    Dim iMozi, iKbn As Integer
    Dim sShubetu, sRoot, sPass, sKomoku As String
    Dim bSyori As Boolean
    Dim fs As Object
    Dim MyName As String
    Dim i As Integer
    Dim sChkPass As String
    Dim iRet As Integer
    Dim lngErrCode As Long       '�G���[�R�[�h

    sFileDelete = False

    On Error GoTo ERR_SPACE
    
    '�t�@�C���L���`�F�b�N
    MyName = Dir(PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE, vbNormal)
    If MyName = "" Then
        GoTo ERR_SPACE
    End If
  
    '���g�p�̃t�@�C���ԍ����擾����B
    iFileNo = FreeFile
    '�V�X�e���������ݒ�t�@�C�����J���B
    Open PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE For Input As #iFileNo
    ' �P�s�ڂ͑S�̃o�[�W�����Ȃ̂œǔ�΂��B
    Line Input #iFileNo, sFileData
    Do While Not EOF(iFileNo)
        '1�s���ǂݍ��ށB
        Line Input #iFileNo, sFileData
        sFileData = Trim(sFileData)
        '�f�[�^���Ȃ����
        If Len(sFileData) = 0 Then
            Exit Do
        End If

        '��Ɨp�ϐ��̏�����
        iMozi = 1
        iKbn = 1
        bSyori = False

        '�t�@�C�����e�擾
        Do
            If Mid(sFileData, iMozi, 1) = "," Or iMozi = Len(sFileData) Then
                Select Case iKbn
                    '���
                    Case 1
                        sShubetu = Trim(Left(sFileData, iMozi - 1))
                        If sShubetu <> "2" And sShubetu <> "3" Then
                            Exit Do
                        End If
                    '���[�g�t�H���_
                    Case 2
                         sRoot = Trim(Left(sFileData, iMozi - 1))
                    '�p�X
                    Case 3
                         sPass = Trim(Left(sFileData, iMozi - 1))
                    '����
                    Case 4
                        sKomoku = Trim(sFileData)
                        If bChk(Int(sKomoku)) = False Then
                           Exit Do
                        End If
                        bSyori = True
                        Exit Do
                End Select
                sFileData = Trim(Mid(sFileData, iMozi + 1))
                iMozi = 0
                iKbn = iKbn + 1
            End If
            iMozi = iMozi + 1
        Loop

        '�擾�f�[�^�̏����̗L��
        If bSyori = True Then
            '�p�X�̎擾
            Select Case sRoot
                Case 1  '�A�v�����[�g
                    sPass = PATH_IDU_APP & "\\" & sPass
                Case 2  '�o�b�N�A�b�v
                    sPass = PATH_BUC & "\\" & sPass
                Case 3      '���g�p
'                    sPass = PATH_DAT & sPass
                Case 4  '���O���[�g
                    sPass = PATH_IDU_LOG & "\\" & sPass
                'EG20 V2.0.1.1 ADD START
                Case 5  '�c�a���[�g
                    sPass = PATH_IDU_DB & "\\" & sPass
                'EG20 V2.0.1.1 ADD START
            End Select

            '�t�@�C���L���`�F�b�N
            If sShubetu = 3 Then
                MyName = Dir(sPass, vbDirectory)
            Else
                MyName = Dir(sPass, vbNormal)
            End If

            '�������s
            If MyName <> "" Then
                Set fs = CreateObject("Scripting.FileSystemObject")
                  Select Case sShubetu
                      '�t�@�C���폜
                      Case 2:
                           iRet = fs.DeleteFile(sPass)
                          If iRet <> 0 Then
                              GoTo ERR_SPACE
                          End If
                          LstStatus.AddItem "�폜�����t�@�C�� - " & sPass
                          DoEvents          'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                      '�t�H���_�̍폜�^�쐬
                      Case 3:
                          fs.DeleteFolder (sPass), True
                          fs.CreateFolder (sPass)
                          LstStatus.AddItem "�폜�^�쐬�����t�H���_ - " & sPass
                          DoEvents          'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                  End Select
                '�I�u�W�F�N�g���
                Set fs = Nothing
            Else
                '�w��o�`�r�r�i�V
                Select Case sShubetu
                   Case 2:
                       LstStatus.AddItem "�w��t�@�C���Ȃ� - " & sPass
                       DoEvents          'V1.5.0.1 ADD
                       LstStatus.Selected(LstStatus.ListCount - 1) = True           'V1.12.0.1 ADD
                   Case 3:
                       Set fs = CreateObject("Scripting.FileSystemObject")
                       '�t�@�C���L���`�F�b�N
                       For i = 0 To Len(sPass)
                           If Mid(sPass, Len(sPass) - i, 1) = "\" Then
                               sChkPass = Left(sPass, Len(sPass) - i - 1)
                               Exit For
                           End If
                       Next
                       MyName = Dir(sChkPass, vbDirectory)
                       If MyName = "" Then
                           LstStatus.AddItem "�t�H���_�쐬���s - " & sPass
                           DoEvents          'V1.5.0.1 ADD
                           LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
                       Else
                           fs.CreateFolder (sPass)
                           LstStatus.AddItem "�쐬�����t�H���_ - " & sPass
                           DoEvents          'V1.5.0.1 ADD
                           LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
                       End If
                       '�I�u�W�F�N�g���
                       Set fs = Nothing
                End Select
            End If
        End If
    Loop
    Close #iFileNo

    sFileDelete = True
    
     Exit Function

ERR_SPACE:
    'V1.21.0.1 ADD  START
    If iFileNo > 0 Then
        Close #iFileNo
    End If
    'V1.21.0.1 ADD  END
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�uIDU�V�X�e����������ʁF�t�@�C���E�t�H���_�������ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TARGET_FILE_FOLDER_DELETE_ERROR, lngErrCode)
   '�I�u�W�F�N�g���
    Set fs = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSysFileDelete
'//  �@�\����  : �V�X�e���t�@�C���폜����
'//  �@�\�T�v  : �C�x���g���O�A���g�\�����O�A�������_���v�t�@�C�����폜����
'//
'//              �^        ����        �Ӗ�
'//   ����     :�Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :�Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 ���X�g�{�b�N�X�̃X�N���[�������ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sSysFileDelete()
    Dim iRet As Integer          '�폜�����߂�l
    Dim NameChk As String        '�t�@�C���L���`�F�b�N�߂�l
    Dim lhEventLog As Long       '�C�x���g���O�̃n���h���B
    Dim lReturn As Long          '�֐��߂�l
    Dim fs As Object
    Dim lngErrCode As Long       '�G���[�R�[�h
   
    sSysFileDelete = False
    
    On Err GoTo ERR_SPACE
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    '/////////////////////////////
    '�������_���v�t�@�C���̍폜
    '/////////////////////////////
    '�t�@�C���L���`�F�b�N
    NameChk = Dir(PATH_INS & MEMORYLOG, vbNormal)
    If NameChk <> "" Then
       iRet = fs.DeleteFile(PATH_INS & MEMORYLOG)
       If iRet <> 0 Then
           GoTo ERR_SPACE
       End If
       LstStatus.AddItem "�폜�����t�@�C�� - " & PATH_INS & MEMORYLOG
       LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
    End If
    
    '/////////////////////////////
    '���g�\�����O�t�@�C���̍폜
    '/////////////////////////////
    '�t�@�C���L���`�F�b�N
    NameChk = Dir(SYSDRWATSON_LOG, vbNormal)
    If NameChk <> "" Then
       iRet = fs.DeleteFile(SYSDRWATSON_LOG)
       If iRet <> 0 Then
          GoTo ERR_SPACE
       End If
       LstStatus.AddItem "�폜�����t�@�C�� - " & SYSDRWATSON_LOG
       LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
    End If
    '���
    Set fs = Nothing
    
    '/////////////////////////////
    '�C�x���g���O�̃N���A
    '/////////////////////////////
    ' �C�x���g���O�i�A�v���P�[�V�����j���N���A����B
    lhEventLog = OpenEventLog(vbNullString, "Application")
    lReturn = ClearEventLog(lhEventLog, vbNullString)
    lReturn = CloseEventLog(lhEventLog)

    ' �C�x���g���O�i�V�X�e���j���N���A����B
    lhEventLog = OpenEventLog(vbNullString, "System")
    lReturn = ClearEventLog(lhEventLog, vbNullString)
    lReturn = CloseEventLog(lhEventLog)

    ' �C�x���g���O�i�Z�L�����e�B�j���N���A����B
    lhEventLog = OpenEventLog(vbNullString, "Security")
    lReturn = ClearEventLog(lhEventLog, vbNullString)
    lReturn = CloseEventLog(lhEventLog)

    sSysFileDelete = True

    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�uIDU�V�X�e����������ʁF�V�X�e���t�@�C���폜�ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SYSFILE_DELETE_ERROR, lngErrCode)
    Set fs = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : OptShosai_Click
'//  �@�\����  : �u�ڍׁv�t����������
'//  �@�\�T�v  : �e�f�[�^�ɑ΂���ڍזt�������������s���B
'//
'//              �^        ����        �Ӗ�
'//   ����     :Integer�@�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :�Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub OptShosai_Click(Index As Integer)
   
   '�uID���p�Ưļ��я�������ʁF�ڍזt�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYS_INFO_BUTTOM, 0)
  
    LblShosai.Caption = ShosaiMoji(Index)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : ShosaiMongonGet
'//  �@�\����  : �u�ڍׁv�t�����\�������擾����
'//  �@�\�T�v  : �u�ڍׁv�t�����ɂĕ\�����镶�����t�@�C�����擾����B
'//
'//              �^        ����        �Ӗ�
'//   ����     :�Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :�Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub ShosaiMongonGet()
   Dim sWork As String                      '��ƃG���A
   Dim iKey As String                       '�L�[��
   Dim lSts As Long                         '�߂�l
   Dim lngRet As Long          '�֐��̕Ԃ�l
   Dim iGate As Integer        '����INDEX
   Dim j As Integer            '���[�NINDEX
   Dim cWork As Byte           '���[�N�G���A
   Dim sGateData As String * SYSMOJI_SIZE    '�P�s���t�@�C�����e�擾�p
   Dim iFCnt As Integer
   Dim iFLoop As Integer
   Dim iFLoop2 As Integer
   Dim MyName As String
   Dim i As Integer
    
   '�t�@�C���L���`�F�b�N
   MyName = Dir(PATH_SYSFORMAT_SHOUSAI_FILE, vbNormal)
   If MyName = "" Then
       sWork = ""
       For i = 0 To 8
        ShosaiMoji(i) = sWork
       Next
       Exit Sub
   End If
    
   For iGate = CNT_MIN To 8
      ' SysFormatShousai.ini��蕶�����擾����B
       sGateData = ""
       iKey = SYS_KEY_NAME & iGate
       lSts = GetPrivateProfileString(SYS_IDU_SECTION_NAME, _
                                      iKey, _
                                      DEFAILT, _
                                      sGateData, _
                                      Len(sGateData), _
                                      PATH_SYSFORMAT_SHOUSAI_FILE)
      If lSts = 0 Or sGateData = "" Then
         '��`�Ȃ���΋�
         ShosaiMoji(iGate) = sWork
      ElseIf Len(sGateData) <> 0 Then
         '�f�[�^�̎擾
          ReDim sFData(6)
           iFCnt = 1

           For iFLoop = 1 To Len(sGateData)
               If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
                 iFLoop2 = iFLoop
                 Do
                   iFLoop2 = iFLoop2 + 1
                   If iFLoop2 > Len(sGateData) Then
                      sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                      iFCnt = iFCnt + 1
                      If iFCnt >= 16 Then
                          Exit For
                      End If

                      iFLoop = iFLoop2
                      Exit Do
                  End If

                  If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
                     sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                     iFCnt = iFCnt + 1
                     If iFCnt >= 16 Then
                           Exit For
                     End If

                     iFLoop = iFLoop2
                     Exit Do
                  End If
                 Loop
               End If
           Next
           
           For i = 0 To 5
             If i = 0 Then
                 ShosaiMoji(iGate) = sFData(i + 1)
                 ShosaiMoji(iGate) = ShosaiMoji(iGate) & vbCrLf
             Else
                 ShosaiMoji(iGate) = ShosaiMoji(iGate) & sFData(i + 1)
                 ShosaiMoji(iGate) = ShosaiMoji(iGate) & vbCrLf
             End If
           Next
       End If
  Next
End Sub

'V1.3.0.1 ADD START
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
        AppActivate frmIDUSysformat.Caption, False
        pfFormActive (frmIDUSysformat.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
'V1.4.0.1�@ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCreateShokiFile
'//  �@�\����  : �ۑ��t�@�C�����쐬����B
'//  �@�\�T�v  : �e�ݒ�t�@�C���̕ۑ��p���쐬����B
'//
'//              �^        ����        �Ӗ�
'//   ����     :�Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :�Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sCreateShokiFile() As Boolean

   Dim NameChk As String        '�t�@�C���L���`�F�b�N�߂�l
   Dim lngErrCode As Long       '�G���[�R�[�h
    
    sCreateShokiFile = False
    
    On Error GoTo ERR_SPACE
           
    '///////////////////////////////////////////////////////////
    'IDU�k�ރ`�F�b�N��IDU�t�@�C���֘A�̕ۑ��p�t�@�C�����쐬����B
    '///////////////////////////////////////////////////////////
    '�t�@�C���L���`�F�b�N
    If pbIDUSts = 1 Then
       sCreateShokiFile = True
       '�uIDU�V�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
       Exit Function
    End If
    
    'IC_M�ݒ�t�@�C���L���`�F�b�N
    NameChk = Dir(PATH_IDU_APP & PATH_ICM_SETTEI, vbNormal)
    If NameChk <> "" Then
       FileCopy PATH_IDU_APP & PATH_ICM_SETTEI, PATH_IDU_APP & PATH_SHOKI_ICM_SETTEI
    End If
    
    'ID���p���j�b�g�ݒ�t�@�C���L���`�F�b�N
    NameChk = Dir(PATH_IDU_APP & PATH_IDU_SETTEI, vbNormal)
    If NameChk <> "" Then
       FileCopy PATH_IDU_APP & PATH_IDU_SETTEI, PATH_IDU_APP & PATH_SHOKI_IDU_SETTEI
    End If

    sCreateShokiFile = True
    '�uIDU�V�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
    
    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�uIDU�V�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬�ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SHOKI_CREATE_ERROR, lngErrCode)
    sCreateShokiFile = False
End Function
'V1.4.0.1�@ADD END

'V1.5.0.1�@ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrAplTimer_Timer
'//  �@�\����  : �A�v���N���`�F�b�N�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v  : �^�C���A�b�v���ɃA�v���N����Ԃ��`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplTimer_Timer()
  
  Dim bIDURet As Boolean  'IDU���O�t���O 'V1.20.0.1 ADD

  On Error Resume Next
 
  '�҂����Ԃ�INI��`�𒴂������ǂ����`�F�b�N
  If lngtime <= lngMAX_Time Then
    '�A�v���N���`�F�b�N���s���BIDU(�A�v���A���O)���I�������Ƃ��̂݁A�������������s���B
    'If CheckAppStart(PROCESS_IDU_PC) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 Then  'V1.20.0.1 DEL
    If CheckAppStart(PROCESS_IDU_PC) = 0 Then   'V1.20.0.1 ADD
      '�A�v���N���`�F�b�N�^�C�}���~����B
      tmrAplTimer.Enabled = False
     'V1.20.0.1 DEL START
'      '����������
'      DeleteFile_Folder
     'V1.20.0.1 DEL END
     'V1.20.0.1 ADD START
     If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
        bIDURet = EndIDULog 'IDU���O�N������IDU���O�ɑ΂��ă��O�I���v��CMD���M
     Else
        bIDURet = True
     End If
     
     If bIDURet = True Then
        lngtime = 0
        lngtime = MN_MAIL_INTERVAL
        tmrLogTimer.Enabled = True
     Else
       '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "�������Ɏ��s���܂���"
        OptKoumoku(0).Enabled = True
        OptKoumoku(1).Enabled = True
        cmdZikko.Enabled = True
        cmdCancel.Enabled = True
        '���O�C�����[�U�`�F�b�N
        If pbUserLevel = 1 Then
           OptKoumoku(2).Enabled = True
        End If        '�A�v���N���`�F�b�N�^�C�}���~����B
        Exit Sub
     End If
     'V1.20.0.1 ADD END
    Else
    '�N���A�v���L��̏ꍇ�A�^�C�}�𒣂蒼��
      tmrAplTimer.Interval = MN_MAIL_INTERVAL
    '���v�o�ߑ҂����Ԃ��A�b�v
     lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI��`�l�𒴂����ꍇ�A�����������ُ�Ƃ���B
    '�uID���p�Ưļ��я�������ʁF���������������s�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "�������Ɏ��s���܂���"
    OptKoumoku(0).Enabled = True
    OptKoumoku(1).Enabled = True
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
    '���O�C�����[�U�`�F�b�N
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True
     End If
    '�A�v���N���`�F�b�N�^�C�}���~����B
    tmrAplTimer.Enabled = False
  End If
End Sub

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrLogTimer_Timer
'//  �@�\����  : ���O�N���`�F�b�N�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v  : �^�C���A�b�v���Ƀ��O�N����Ԃ��`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ��A�m�F�|�b�v�A�b�v�C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()
     
   On Error Resume Next

   '�҂����Ԃ�INI��`�𒴂������ǂ����`�F�b�N
  If lngtime <= lngLogMAX_Time Then
    '���O�N���`�F�b�N���s���B�S�ďI�������Ƃ��̂݁A�������������s���B
    If CheckAppStart(PROCESS_IDU_LOG) = 0 Then
      '���O�N���`�F�b�N�^�C�}���~����B
      tmrLogTimer.Enabled = False
      '����������
      DeleteFile_Folder
    Else
    '�N�����O�L��L��̏ꍇ�A�^�C�}�𒣂蒼��
      tmrLogTimer.Interval = MN_MAIL_INTERVAL
    '���v�o�ߑ҂����Ԃ��A�b�v
     lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI��`�l�𒴂����ꍇ�A�����������ُ�Ƃ���B
    '�uID���p�Ưļ��я�������ʁF���������������s�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "�������Ɏ��s���܂���"
    OptKoumoku(0).Enabled = True
    OptKoumoku(1).Enabled = True
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
    '���O�C�����[�U�`�F�b�N
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True
    End If
    '���O�N���`�F�b�N�^�C�}���~����B
    tmrLogTimer.Enabled = False
  End If
End Sub
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : DeleteFile_Folder
'//  �@�\����  : �t�@�C���A�t�H���_�ADB����������
'//  �@�\�T�v  : �������������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 ���X�g�{�b�N�X�̃X�N���[�������ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub DeleteFile_Folder()
    
    Dim bRtn As Boolean
    Dim lExitCode As Long
    Dim bRtn1 As Boolean
    Dim lngErrCode As Long                      '�G���[�R�[�h
    Dim iTargetDB As Integer                       '�Ώ�DB�l
    Dim bDB_Code As Boolean
   
    On Error GoTo ERR_SPACE
   
    '�o�׎��������I�����A�S�ď�����(DLL�f�[�^��)�I�����A���̑��f�[�^���������Ώێ�
    If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
       If sCreateShokiFile = False Then
          '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "�������Ɏ��s���܂���"
          OptKoumoku(0).Enabled = True
          OptKoumoku(1).Enabled = True
          cmdZikko.Enabled = True
          cmdCancel.Enabled = True
          '���O�C�����[�U�`�F�b�N
          If pbUserLevel = 1 Then
             OptKoumoku(2).Enabled = True
          End If
          '�����𔲂���
          Exit Sub
       End If
    End If
     
    '�V�X�e���t�@�C���̍폜
    If bChk(8) = True Then
       bRtn1 = sSysFileDelete()
       DoEvents
    Else
       bRtn1 = True
    End If

    '�t�H���_�A�t�@�C���̍폜
    If bRtn1 = True Then

       If sFileDelete() = True Then

          bDB_Code = True

          'DB����������
          'DB�f�[�^�FIC�ꌏ����
          If bChk(2) = True Then
             iTargetDB = chkMeisai(0).DataField
             Me.LstStatus.AddItem "DB������:" & chkMeisai(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
         End If

' EG20 V2.0.1.1 ADD START
          'DB�f�[�^�F�o�b�N�A�b�v�f�[�^
          If bChk(3) = True And bDB_Code = True Then
             iTargetDB = chkLog(0).DataField
             Me.LstStatus.AddItem "DB������:" & chkLog(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If

          'DB�f�[�^�F�đ��f�[�^
          If bChk(4) = True And bDB_Code = True Then
             iTargetDB = chkLog(0).DataField
             Me.LstStatus.AddItem "DB������:" & chkLog(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If
' EG20 V2.0.1.1 ADD START

          'DB�f�[�^�F�A�v�����O
          If bChk(5) = True And bDB_Code = True Then
             iTargetDB = chkLog(0).DataField
             Me.LstStatus.AddItem "DB������:" & chkLog(0).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If
          
          'DB�f�[�^�F�ێ烍�O
          If bChk(6) = True And bDB_Code = True Then
             iTargetDB = chkLog(1).DataField
             Me.LstStatus.AddItem "DB������:" & chkLog(1).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If

          'DB�f�[�^�F����IC���W���[�����O
          If bChk(7) = True And bDB_Code = True Then
             iTargetDB = chkLog(2).DataField
             Me.LstStatus.AddItem "DB������:" & chkLog(2).Caption
             DoEvents
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If

          'DB�f�[�^�F�l�K���X�g(�o�׎��������̂ݗL��)
          If OptKoumoku(1).Value = False And bDB_Code = True Then
             iTargetDB = stsIDUNega
             bDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
             DoEvents
             LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
          End If
          
          '�߂�l������
          If bDB_Code = True Then
             '�uID���p�Ưļ��я�������ʁF�V�X�e����������������v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
             lblKekka.ForeColor = SYSFORMAT_OK
             lblKekka.Caption = "�������͐������܂���"
          Else
             '�uID���p�Ưļ��я�������ʁFDB�����������ُ�v���O�o��
             Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
             lblKekka.ForeColor = SYSFORMAT_ERROR
             lblKekka.Caption = "�������Ɏ��s���܂���"
          End If
     Else
        '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "�������Ɏ��s���܂���"
     End If
  Else
    '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "�������Ɏ��s���܂���"
  End If

 '�������I�����̏���
 OptKoumoku(0).Enabled = True
 OptKoumoku(1).Enabled = True
 cmdZikko.Enabled = True
 cmdCancel.Enabled = True
 '���O�C�����[�U�`�F�b�N
 If pbUserLevel = 1 Then
    OptKoumoku(2).Enabled = True
 End If
 
Exit Sub

ERR_SPACE2:
        '�G���[�������̏���
        OptKoumoku(0).Enabled = True    '���������ڎw��F�o�׎��������I���\
        OptKoumoku(1).Enabled = True    '���������ڎw��F���ڑI��I���\
        cmdZikko.Enabled = True         '�u���������s�v�t�����\
        cmdCancel.Enabled = True        '�u���j���[��ʂ֖߂�v�t�����\
        '���O�C�����[�U�`�F�b�N
        If pbUserLevel = 1 Then
            OptKoumoku(2).Enabled = True
        End If
        '�uID���p�Ưļ��я�������ʁF�V�X�e�������������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "�������Ɏ��s���܂���"

ERR_SPACE:
End Sub
'V1.5.0.1 ADD�@END
