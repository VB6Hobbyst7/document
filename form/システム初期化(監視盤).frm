VERSION 5.00
Begin VB.Form frmKansiSysformat 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�V�X�e���������@�\"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chksolog 
      Caption         =   "����샍�O�f�[�^"
      Height          =   495
      Left            =   5280
      TabIndex        =   23
      Top             =   1800
      Value           =   1  '����
      Width           =   2535
   End
   Begin VB.Timer tmrLogTimer 
      Left            =   11400
      Top             =   6480
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   8640
      Top             =   7800
   End
   Begin VB.Timer tmrMail 
      Left            =   8640
      Top             =   6000
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
      TabIndex        =   9
      Top             =   5400
      Width           =   2415
   End
   Begin VB.ListBox LstStatus 
      Height          =   3210
      Left            =   120
      TabIndex        =   8
      Top             =   5400
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
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   11775
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   4
         Left            =   4200
         Style           =   1  '���̨���
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   7
         Left            =   7680
         Style           =   1  '���̨���
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   5
         Left            =   4200
         Style           =   1  '���̨���
         TabIndex        =   27
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   3
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   2
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   25
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   24
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   6
         Left            =   7560
         Style           =   1  '���̨���
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton OptShosai 
         Caption         =   "�ڍ�"
         Height          =   375
         Index           =   0
         Left            =   360
         Style           =   1  '���̨���
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Frame frmKoumoku 
         Caption         =   "����"
         Height          =   3615
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   11295
         Begin VB.Frame FraShosai 
            Caption         =   "���ڏڍ�"
            Height          =   1725
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   11100
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
               Height          =   1380
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   10845
            End
         End
         Begin VB.CheckBox chkDLL 
            Caption         =   "�v���O��������f�[�^"
            Height          =   375
            Left            =   8160
            TabIndex        =   5
            Top             =   360
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.CheckBox chkIC 
            Caption         =   "IC�֘A�f�[�^"
            Height          =   375
            Left            =   8280
            TabIndex        =   4
            Top             =   840
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.CheckBox chkSonota 
            Caption         =   "���̑��f�[�^"
            Height          =   375
            Left            =   5040
            TabIndex        =   3
            Top             =   840
            Value           =   1  '����
            Width           =   3000
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "�����Ď��Ճ��O�f�[�^"
            Height          =   375
            Left            =   1200
            TabIndex        =   13
            Top             =   1320
            Value           =   1  '����
            Width           =   3000
         End
         Begin VB.CheckBox chkBackUp 
            Caption         =   "�o�b�N�A�b�v�f�[�^  �@���O���p�@�]���f�[�^"
            Height          =   495
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Value           =   1  '����
            Width           =   3000
         End
         Begin VB.CheckBox chkMeisai 
            Caption         =   "�W�v�֘A�f�[�^"
            Height          =   375
            Left            =   1200
            TabIndex        =   1
            Top             =   360
            Value           =   1  '����
            Width           =   3000
         End
      End
      Begin VB.Frame FraKomoku 
         Height          =   620
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   10455
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "�o�׎�������"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   225
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "���ڑI��"
            Height          =   300
            Index           =   1
            Left            =   2640
            TabIndex        =   19
            Top             =   225
            Width           =   1575
         End
         Begin VB.OptionButton OptKoumoku 
            Caption         =   "�S�ď������i�v���O��������f�[�^�܂ށj"
            Height          =   300
            Index           =   2
            Left            =   5160
            TabIndex        =   18
            Top             =   225
            Visible         =   0   'False
            Width           =   4935
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
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�����Ď��ՃV�X�e��������"
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
      TabIndex        =   14
      Top             =   0
      Width           =   12015
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
      TabIndex        =   11
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Caption         =   "����������"
      Height          =   255
      Left            =   8760
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
End
Attribute VB_Name = "frmKansiSysformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmKansiSysformat.frm
'//  �p�b�P�[�W���F�V�X�e��������(�Ď���)���
'/
'//  �T�v�F�V�X�e��������(�Ď���)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�ۑ��p�ݒ�t�@�C�������ǉ�
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                         �ێ瑍�_�����ʏC��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 ���X�g�{�b�N�X�̃X�N���[�������ǉ�
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ��A�m�F�|�b�v�A�b�v�C��
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-240�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'Private bChk() As Boolean           'V1.5.0.1 DEL

'���������s�t���O
Private bSysFormat As Boolean

Private ShosaiMoji(0 To 7) As String '�ڍו����i�[�G���A
Private Const SYSMOJI_SIZE = 500
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000     '�A�v���N���^�C�}�f�t�H���g�l
Dim lngMAX_Time As Long                    'INI�擾�ݒ�l
Dim lngtime     As Long                    '���݃^�C�}�l
Private bChk(8) As Boolean
'V1.5.0.1 ADD END
'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        '���O�N���^�C�}�f�t�H���g�l(30�b)
Dim lngLogMAX_Time As Long                'INI�擾�ݒ�l(���O�j
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �V�X�e��������(�Ď���)���(�A�N�e�B�u��)
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
'//  �@�\����  : �V�X�e��������(�Ď���)���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �V�X�e��������(�Ď���)���(���[�h��)
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
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//     REVISIONS :(EG20 v2.0.1.1) 2011-11-24  REVISED BY [TCC] T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�z
'//                �E���������ڒǉ��i����샍�O�f�[�^�j
'//                �E���ڏڍ׍폜(�h�b�֘A�f�[�^,�v���O�����E����f�[�^�j
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer    '�J�E���^�[
   
    On Error Resume Next

    '�u�Ď��ՃV�X�e����������ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SYSFORMAT_GAMEN_START, 0)

    '�u�ڍׁv�t���������擾����
    ShosaiMongonGet

   '������
    OptShosai(0).Value = True   '���������ڎw��F�ڍזt����
    LstStatus.Clear             '�폜�t�@�C���\�����N���A
    OptKoumoku(0).Value = True  '���������ڎw��u�o�׎��������v�w��L��I��
    chkMeisai.Value = 1         '�W�v�֘A�f�[�^�F�`�F�b�N�L��
    chkMeisai.Enabled = False   '�W�v�֘A�f�[�^�F�I��s��
    chkBackUp.Value = 1         '�o�b�N�A�b�v�f�[�^�F�`�F�b�N�L��
    chkBackUp.Enabled = False   '�o�b�N�A�b�v�f�[�^�F�I��s��
    chkLog.Value = 1            '���O�f�[�^�F�`�F�b�N�L��i�����Ď��Ճ��O�f�[�^�j
    chkLog.Enabled = False      '���O�f�[�^�F�I��s�i�����Ď��Ճ��O�f�[�^�j
' EG20 V2.0.1.1�y�c����54�zADD START
    chksolog.Value = 1          '����샍�O�f�[�^�F�`�F�b�N�L��
    chksolog.Enabled = False    '����샍�O�f�[�^�F�I��s��
' EG20 V2.0.1.1�y�c����54�zADD START
    chkSonota.Value = 1         '���̑��f�[�^�F�`�F�b�N�L��
    chkSonota.Enabled = False   '���̑��f�[�^�F�I��s��
' EG20 V2.0.1.1�y�c����54�zDEL START
'    chkDLL.Value = 0            '�v���O��������f�[�^�F�`�F�b�N����
'    chkDLL.Enabled = False      '�v���O��������f�[�^�F�I��s��
'    chkIC.Value = 1             'IC�֘A�f�[�^�F�`�F�b�N�L��
'    chkIC.Enabled = False       'IC�֘A�f�[�^�F�I��s��
' EG20 V2.0.1.1�y�c����54�zDEL�@END
    lblKekka.Caption = ""       '���������s�\�����N���A
    frmKoumoku.Enabled = False  '���ڕ������s��

    OptKoumoku(2).Enabled = False
    
    '���O�C�����[�U�`�F�b�N
    If pbUserLevel = 1 Then
        OptKoumoku(2).Enabled = True
        chkDLL.Value = 1
        chkDLL.Enabled = False   '�v���O��������f�[�^�F�I����
    Else
        OptKoumoku(2).Enabled = False
    End If
    
    '���������s�t���OOFF
    bSysFormat = False
    
    OptShosai(0).Enabled = True '���������ڕ��F�ڍזt�����\
    OptShosai(0).Value = True   '���������ڕ��F�ڍזt����
    For i = 1 To 6
        OptShosai(i).Enabled = False '���ڕ��F�ڍזt�����s��
    Next

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
   
   'V1.20.0.1 ADD START
   'INI�t�@�C����胍�O�N���^�C�}�l���擾
   lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
   '�擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
   If lngLogMAX_Time = 0 Then
      lngLogMAX_Time = LOG_INTERVAL
   End If
   'V1.20.0.1 ADD END
   
   '�^�C�}�l�ݒ�
   tmrAplTimer.Interval = MN_MAIL_INTERVAL
   tmrAplTimer.Enabled = False
   'V1.5.0.1 ADD END
   
   'V1.20.0.1 ADD START
   tmrLogTimer.Interval = MN_MAIL_INTERVAL
   tmrLogTimer.Enabled = False
   'V1.20.0.1 ADD END
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
'//     REVISIONS :(Eg20 V2.0.1.1) 2011-11-24  REVISED BY [TCC] T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub OptKoumoku_Click(Index As Integer)
    Dim i As Integer    '�J�E���^�[

    On Error Resume Next
     
     Select Case Index
          Case 0  '�o�׎��������I����
            frmKoumoku.Enabled = False       '���ڃt���[��
            chkMeisai.Enabled = False        '�W�v�֘A�f�[�^
            chkBackUp.Enabled = False        '�o�b�N�A�b�v�f�[�^
            chkLog.Enabled = False           '���O�f�[�^
            chkSonota.Enabled = False        '���̑��f�[�^
' EG20 V2.0.1.1�y�c����54�zADD START
            chksolog.Enabled = False         '����샍�O�f�[�^
' EG20 V2.0.1.1�y�c����54�zADD END
' EG20 V2.0.1.1�y�c����54�zDEL START
'            chkIC.Enabled = False            'IC�֘A�f�[�^
'
'            '���O�C�����[�U�`�F�b�N
'            If pbUserLevel = 1 Then
'                chkDLL.Enabled = False       '�v���O��������f�[�^
'            End If
' EG20 V2.0.1.1�y�c����54�zDEL END
            
            OptShosai(0).Enabled = True      '���������ڕ��F�ڍזt�����\
            OptShosai(0).Value = True        '���������ڕ��F�ڍזt����
            For i = 1 To 6
                OptShosai(i).Enabled = False '���������ڕ��F�ڍזt�����s��
            Next
            '�u�Ď��ռ��я�������ʁF�o�׎��������I�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_SHUKKA, 0)
        Case 1   '���ڑI����
            frmKoumoku.Enabled = True        '���ڃt���[��
            chkMeisai.Enabled = True         '�W�v�֘A�f�[�^
            chkBackUp.Enabled = True         '�o�b�N�A�b�v�f�[�^
            chkLog.Enabled = True            '���O�f�[�^
            chkSonota.Enabled = True         '���̑��f�[�^
' EG20 V2.0.1.1�y�c����54�zADD START
            chksolog.Enabled = True          '����샍�O�f�[�^
' EG20 V2.0.1.1�y�c����54�zADD END
' EG20 V2.0.1.1�y�c����54�zDEL START
'            chkIC.Enabled = True             'IC�֘A�f�[�^
'
'            '���O�C�����[�U�`�F�b�N
'            If pbUserLevel = 1 Then
'                chkDLL.Enabled = True        '�v���O��������f�[�^
'                OptShosai(6).Enabled = True  '�v���O��������f�[�^�ڍזt�����\
'            End If
' EG20 V2.0.1.1�y�c����54�zDEL END
            OptShosai(0).Enabled = False     '���������ڎw��F�ڍזt�I��s��
            OptShosai(1).Value = True        '���������ڎw��F�ڍזt��s��
            For i = 1 To 5
                OptShosai(i).Enabled = True  '���ڎw��F�ڍזt�I���\
            Next
            
            '�u�Ď��ռ��я�������ʁF���ڑI��I�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_KOUMOKU, 0)
        Case Else:
            frmKoumoku.Enabled = False       '���ڃt���[��
            chkMeisai.Enabled = False        '�W�v�֘A�f�[�^
            chkBackUp.Enabled = False        '�o�b�N�A�b�v�f�[�^
            chkLog.Enabled = False           '���O�f�[�^
            chkSonota.Enabled = False        '���̑��f�[�^
' EG20 V2.0.1.1�y�c����54�zADD START
            chksolog.Enabled = True          '����샍�O�f�[�^
' EG20 V2.0.1.1�y�c����54�zADD END
' EG20 V2.0.1.1�y�c����54�zDEL START
'            chkIC.Enabled = False            'IC�֘A�f�[�^
'
'            '���O�C�����[�U�`�F�b�N
'            If pbUserLevel = 1 Then
'                chkDLL.Enabled = False       '�v���O��������f�[�^
'            End If
' EG20 V2.0.1.1�y�c����54�zDEL END
            OptShosai(0).Enabled = True      '���������ڎw��F�ڍזt�I���\
            OptShosai(0).Value = True        '���������ڎw��F�ڍזt����
            For i = 1 To 6
                OptShosai(i).Enabled = False '���ڎw��F�ڍזt�I��s��
            Next
            '�u�Ď��ռ��я�������ʁF�S�ď������I�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTYPE_ALL, 0)
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
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                         �ێ瑍�_�����ʏC��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ��A�m�F�|�b�v�A�b�v�C��
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-240�z
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
    Dim bRtn1 As Boolean
    Dim bRtn2 As Boolean
    Dim iRetApp         As Integer
    Dim iRetLog         As Integer
    Dim uMail As ML_KYOTU_INF           '���[��
    'ReDim bChk(8)                      'V1.5.0.1 DEL
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim iTargetDB As Integer            '�Ώ�DB�l
    Dim bDB_Code As Boolean
    Dim iRetIDULog As Integer           'IDU���O�N���t���O
    Dim iRetLDULog As Integer           'IDU���O�N���t���O
    Dim bRet As Boolean
    'V1.5.0.1  ADD START
    Dim bKansiRet As Boolean            '�Ď��ՃA�v����������
    Dim bIDURet   As Boolean            'IDU�A�v����������
    Dim bLDURet   As Boolean            'LDU�A�v����������
   
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1  ADD END
    On Error GoTo ERR_SPACE

    '�u�Ď��ռ��я�������ʁF���������s�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_START_BUTTOM, 0)

    '�\���̏�����
    LstStatus.Clear
    lblKekka.Caption = ""
    iRetIDULog = 0
    iRetLDULog = 0
 
    '�o�׎��������I����
    If OptKoumoku(0).Value = True Then
        For i = 1 To 6
           bChk(i) = True
        Next
           '�v���O��������f�[�^�̓`�F�b�N����
'           bChk(5) = False                     ' EG20 V3.3.0.1�y����TR-240�z�폜
           bChk(6) = False                      ' EG20 V3.3.0.1�y����TR-240�z�ǉ�
    End If

    '���ڑI��I����
    If OptKoumoku(1).Value = True Then
        bSentaku = False
        '�W�v�֘A�f�[�^
        If chkMeisai.Value = 1 Then
            bSentaku = True
            bChk(1) = True
        Else
            bChk(1) = False
        End If

        '�o�b�N�A�b�v�f�[�^
        If chkBackUp.Value = 1 Then
            bSentaku = True
            bChk(2) = True
        Else
            bChk(2) = False
        End If

        '���O�f�[�^
        If chkLog.Value = 1 Then
            bSentaku = True
            bChk(3) = True
        Else
            bChk(3) = False
        End If
' EG20 V 2.0.1.1�y�c����54�zDEL START
'        '���̑��f�[�^
'        If chkSonota.Value = 1 Then
'           bSentaku = True
'           bChk(4) = True
'        Else
'           bChk(4) = False
'        End If
' EG20 V 2.0.1.1�y�c����54�zDEL END
' EG20 V 2.0.1.1�y�c����54�zADD START
        '����샍�O�f�[�^
        If chksolog.Value = 1 Then
           bSentaku = True
           bChk(4) = True
        Else
           bChk(4) = False
        End If

        '���̑��f�[�^
        If chkSonota.Value = 1 Then
           bSentaku = True
           bChk(5) = True
        Else
           bChk(5) = False
        End If
' EG20 V 2.0.1.1�y�c����54�zADD END
' EG20 V 2.0.1.1�y�c����54�zDEL START
'        '�c�k�k�f�[�^
'        If chkDLL.Value = 1 Then
'            bSentaku = True
'            bChk(5) = True
'        Else
'            bChk(5) = False
'        End If
'
'        'IC�֘A�f�[�^
'        If chkIC.Value = 1 Then
'            bSentaku = True
'            bChk(6) = True
'        Else
'            bChk(6) = False
'        End If
' EG20 V 2.0.1.1�y�c����54�zDEL END
        bChk(6) = False                      ' EG20 V3.3.0.1�y����TR-240�z�ǉ�

        If bSentaku = False Then
            '�u�Ď��ռ��я�������ʁF���������������s�v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
            MsgBox "����������f�[�^���I������Ă��܂���", vbExclamation, "�f�[�^���x��"
            Exit Sub
        End If
    End If

' EG20 V 2.0.1.1�y�c����54�zDEL START
'    '�S�ď������i�c�k�k�f�[�^�܂ށj�I����
'    If OptKoumoku(2).Value = True Then
'        For i = 1 To 6
'            '�S�đI���`�F�b�N�L��
'            bChk(i) = True
'        Next
'    End If
' EG20 V 2.0.1.1�y�c����54�zDEL END
    
    iRet = MsgBox("�������������s���܂��B��낵���ł����H", vbExclamation + vbOKCancel, "�������m�F")
    If iRet = vbOK Then
        '����������I�����̏���
         OptKoumoku(0).Enabled = False      '�u�o�׎��������v���W�I�t�I��s��
         OptKoumoku(1).Enabled = False      '�u���ڑI���v���W�I�t�I��s��
' EG20 V 2.0.1.1�y�c����54�zDEL START
'        '���O�C�����[�U�`�F�b�N
'        If pbUserLevel = 1 Then
'           OptKoumoku(2).Enabled = False  '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
'        End If
' EG20 V 2.0.1.1�y�c����54�zDEL END
        cmdZikko.Enabled = False          '�u���������s�v�t�����s��
        cmdCancel.Enabled = False         '�u���j���[��ʂ֖߂�v�t�����s��
    
        On Error GoTo ERR_SPACE2

        '�Ď���(�Ǘ��v���Z�X)���N�����Ă��邩�ǂ����`�F�b�N����B
        If CheckAppStart(PROC_KANRI) <> 0 Then
          'V1.20.0.1 DEL START
          ' iRet = MsgBox("�Ď��ՃA�v���P�[�V�������I�����܂��B" & Chr(vbKeyReturn) & _
          '               "��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F")
          'If iRet = vbOK Then
          'V1.20.0.1 DEL END
              '�A�v���I���v�����Ǘ��ɑ��M����
               uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
               uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
               uMail.udtlHeader.dwProid = RHOSHU_ID
               uMail.udtlHeader.dwSubArea = 0
               'V1.5.0.1 DEL START
               'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               'If bRtn = 0 Then
               'V1.5.0.1 DEL END
               'V1.5.0.1 ADD START
               bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               If bKansiRet = 0 Then
               'V1.5.0.1 ADD END
                 '�u�Ď��ռ��я�������ʁF���[�����M�ُ�v���O�o��
                 lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                 Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lngErrCode)
                 GoTo ERR_SPACE2:
               Else
                 '�u�Ď��ռ��я�������ʁF���[�����M����v���O�o��
                 Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                 '�A�v���I���m�F
                  'iRetApp = CheckAppEndComplete(PROC_KANRI, lExitCode)   'V1.5.0.1 DEL
               End If
         'V1.20.0.1 DEL START
'              'IDU���O�v���Z�X�N���`�F�b�N
'              If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'
'                 'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
'                 iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
'
'                 If iRet = vbOK Then
'                   'IDU���O�I���v��CMD���M
'                   'V1.5.0.1 DEL START
'                   'bRet = EndIDULog
'                   'If bRtn = False Then
'                   'V1.5.0.1 DEL END
'                   'V1.5.0.1 ADD START
'                   bIDURet = EndIDULog
'                   If bIDURet = False Then
'                   'V1.5.0.1 ADD END
'                     '���M�ُ폈��
'                     lblKekka.ForeColor = SYSFORMAT_ERROR
'                     lblKekka.Caption = "�������Ɏ��s���܂���"
'                     OptKoumoku(0).Enabled = True
'                     OptKoumoku(1).Enabled = True
'                     '���O�C�����[�U�`�F�b�N
'                     If pbUserLevel = 1 Then
'                        OptKoumoku(2).Enabled = True
'                     End If
'                     cmdZikko.Enabled = True
'                     cmdCancel.Enabled = True
'                     Exit Sub
'                  End If
'                  'IDU���O�v���Z�X�I���m�F
'                  'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)  'V1.5.0.1 DEL
'                'V1.7.0.1 ADD START
'                Else
'                 GoTo ERR_SPACE3
'                'V1.7.0.1 ADD END
'                End If
'              'V1.5.0.1 ADD START
'              Else
'              bIDURet = True
'              'V1.5.0.1 ADD END
'              End If
'              'LDU���O�v���Z�X�N���`�F�b�N
'              'If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then 'V1.7.0.1 DEL
'              If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then  'V1.7.0.1 ADD
'
'                 'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
'                 iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
'
'                 If iRet = vbOK Then
'                   'IDU���O�I���v��CMD���M
'                   'V1.5.0.1 DEL START
'                   'bRet = EndLDULog
'                   'If bRtn = False Then
'                   'V1.5.0.1 DEL END
'                   'V1.5.0.1 ADD START
'                   bLDURet = EndLDULog
'                   If bLDURet = False Then
'                   'V1.5.0.1 ADD END
'                     '���M�ُ폈��
'                     lblKekka.ForeColor = SYSFORMAT_ERROR
'                     lblKekka.Caption = "�������Ɏ��s���܂���"
'                     OptKoumoku(0).Enabled = True
'                     OptKoumoku(1).Enabled = True
'                     '���O�C�����[�U�`�F�b�N
'                     If pbUserLevel = 1 Then
'                        OptKoumoku(2).Enabled = True
'                     End If
'                     cmdZikko.Enabled = True
'                     cmdCancel.Enabled = True
'                     Exit Sub
'                  End If
'                  'LDU���O�v���Z�X�I���m�F
'                  'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)  'V1.5.0.1 DEL
'                'V1.7.0.1 ADD START
'                Else
'                 GoTo ERR_SPACE3
'                'V1.7.0.1 ADD END
'                End If
'              'V1.5.0.1 ADD START
'              Else
'              bLDURet = True
'              'V1.5.0.1 ADD END
'              End If
'          Else
'             '�u�L�����Z���t�����v
'             '�u�Ď��ռ��я�������ʁF���������������s�v���O�o��
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'              OptKoumoku(0).Enabled = True    '�u�o�׎��������v���W�I�t�I��s��
'              OptKoumoku(1).Enabled = True    '�u���ڑI���v���W�I�t�I��s��
'              '���O�C�����[�U�`�F�b�N
'              If pbUserLevel = 1 Then
'                OptKoumoku(2).Enabled = True  '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
'              End If
'              cmdZikko.Enabled = True         '�u���������s�v�t�����s��
'              cmdCancel.Enabled = True        '�u���j���[��ʂ֖߂�v�t�����s��
'              Exit Sub
'          End If
         'V1.20.0.1 DEL END
      Else
        iRetApp = 1
        bKansiRet = True    'V1.5.0.1 ADD
      'End If               'V1.5.0.1 DEL
      
       If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
          
          'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
        'V1.20.0.1 DEL START
'          iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
'
'           If iRet = vbOK Then
        'V1.20.0.1 DEL END
              'IDU���O�I���v��CMD���M
               'V1.5.0.1 DEL START
               'bRet = EndIDULog
               'If bRtn = False Then
               'V1.5.0.1 DEL END
               'V1.5.0.1 ADD START
               bIDURet = EndIDULog
               If bIDURet = False Then
               'V1.5.0.1 ADD END
                 '���M�ُ폈��
                  lblKekka.ForeColor = SYSFORMAT_ERROR
                  lblKekka.Caption = "�������Ɏ��s���܂���"
                  OptKoumoku(0).Enabled = True
                  OptKoumoku(1).Enabled = True
                  '���O�C�����[�U�`�F�b�N
                   If pbUserLevel = 1 Then
                      OptKoumoku(2).Enabled = True
                   End If
                   cmdZikko.Enabled = True
                   cmdCancel.Enabled = True
                   Exit Sub
               End If
               'IDU���O�v���Z�X�I���m�F
               'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode)    'V1.5.0.1 DEL
           'V1.7.0.1 ADD START
        'V1.20.0.1 DEL START
'           Else
'              GoTo ERR_SPACE3
'           'V1.7.0.1 ADD END
'           End If
        'V1.20.0.1 DEL END
      Else
        iRetIDULog = 1
        bIDURet = True 'V1.5.0.1 ADD
      End If
       
      If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
         
         'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
        'V1.20.0.1 DEL START
'         iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
'
'         If iRet = vbOK Then
        'V1.20.0.1 DEL END
           'IDU���O�I���v��CMD���M
            'V1.5.0.1 DEL START
            'bRet = EndLDULog
            'If bRtn = False Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 DEL START
             bLDURet = EndLDULog
             If bLDURet = False Then
            'V1.5.0.1 DEL END
                '���M�ُ폈��
                 lblKekka.ForeColor = SYSFORMAT_ERROR
                 lblKekka.Caption = "�������Ɏ��s���܂���"
                 OptKoumoku(0).Enabled = True
                 OptKoumoku(1).Enabled = True
                 '���O�C�����[�U�`�F�b�N
                 If pbUserLevel = 1 Then
                    OptKoumoku(2).Enabled = True
                 End If
                 cmdZikko.Enabled = True
                 cmdCancel.Enabled = True
                 Exit Sub
              End If
            'LDU���O�v���Z�X�I���m�F
            'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)  'V1.5.0.1 DEL
         'V1.7.0.1 ADD START
       'V1.20.0.1 DEL START
'         Else
'            GoTo ERR_SPACE3
'         'V1.7.0.1 ADD END
'         End If
       'V1.20.0.1 DEL END
      Else
         iRetLDULog = 1
         bLDURet = True 'V1.5.0.1 ADD
      End If
     End If             'V1.5.0.1 ADD
'V1.5.0.1 ADD START
     '�Ď��ՁAIDU�ALDU�A�v���̃��[�����M�������S�Đ��킾�����ꍇ�̂݁A�A�v���N���^�C�}���N�������A
     '�A�v���N���`�F�b�N�ɂ��A�v���̋N��/���N���𔻒f����B
     'If (bKansiRet = True) And (bIDURet = True) And (bLDURet = True) Then         'V1.20.0.1 DEL
     If (bKansiRet = True) Then                                                    'V1.20.0.1 ADD
         lngtime = 0
         lngtime = MN_MAIL_INTERVAL
         tmrAplTimer.Enabled = True
     Else
         '�Ď��ՁAIDU�ALDU�A�v���̃��[�����M�ɂĂЂƂł��ُ킪�������ꍇ�A�������������ُ�I���Ƃ���B
         '�u�Ď��ՃV�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "�������Ɏ��s���܂���"
         '����������I�����̏���
         OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
         OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
         '���O�C�����[�U�`�F�b�N
          If pbUserLevel = 1 Then
             OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
          End If
          cmdZikko.Enabled = True        '�u���������s�v�t�����s��
          cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
          '�����𔲂���
           Exit Sub
      End If
   End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'       '�A�v���܂��̓��O�v���Z�X�ŏI�������Ɏ��s�����ꍇ
'       If (iRetApp <> 1) Or (iRetIDULog <> 1) Or (iRetLDULog <> 1) Then
'           '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "�������Ɏ��s���܂���"
'           '����������I�����̏���
'            OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
'            OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
'            '���O�C�����[�U�`�F�b�N
'            If pbUserLevel = 1 Then
'               OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
'            End If
'            cmdZikko.Enabled = True        '�u���������s�v�t�����s��
'            cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
'            '�����𔲂���
'            Exit Sub
'       End If
'
'      '���������s�t���OON
'      bSysFormat = True
'
'      'V1.4.0.1 ADD START
'      '�o�׎��������I�����A�S�ď�����(DLL�f�[�^��)�I�����A���̑��f�[�^���������Ώێ�
'      If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
'        If sCreateShokiFile = False Then
'          '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'          lblKekka.ForeColor = SYSFORMAT_ERROR
'          lblKekka.Caption = "�������Ɏ��s���܂���"
'          '����������I�����̏���
'           OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
'           OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
'           '���O�C�����[�U�`�F�b�N
'           If pbUserLevel = 1 Then
'              OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
'           End If
'           cmdZikko.Enabled = True        '�u���������s�v�t�����s��
'           cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
'           '�����𔲂���
'           Exit Sub
'        End If
'      End If
'      'V1.4.0.1 ADD END
'
'      '�V�X�e���t�@�C���̍폜
'      If bChk(4) = True Then
'           bRtn1 = sSysFileDelete()
'           DoEvents
'      Else
'           bRtn1 = True
'      End If
'
'      '�t�H���_�A�t�@�C���̍폜
'      If bRtn1 = True Then
'
'        If sFileDelete() = True Then
'
'           bDB_Code = True
'
'          If bChk(1) = True Then
'             Me.LstStatus.AddItem "DB������:" & chkMeisai.Caption
'             DoEvents
'             Me.AutoRedraw = True
'
'             '�Ď��ՁF�ꌏ����
'             iTargetDB = stsKansiMeisai
'             bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'             DoEvents
'             Me.AutoRedraw = True
'             If bDB_Code = True Then
'                '�Ď��ՁF�ʏW�D
'                iTargetDB = stsKansiBetu
'                'DB����������
'                bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'                DoEvents
'                Me.AutoRedraw = True
'             End If
'          End If
'
'          If bDB_Code = True Then
'             '�u�Ď��ռ��я�������ʁF�V�X�e����������������v���O�o��
'             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
'             lblKekka.ForeColor = SYSFORMAT_OK
'             lblKekka.Caption = "�������͐������܂���"
'          Else
'             '�u�Ď��ռ��я�������ʁFDB�����������ُ�v���O�o��
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, DBFORMAT_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "�������Ɏ��s���܂���"
'          End If
'        Else
'          '�u�Ď��ռ��я�������ʁF�V�X�e�������������ُ�v���O�o��
'          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'          lblKekka.ForeColor = SYSFORMAT_ERROR
'          lblKekka.Caption = "�������Ɏ��s���܂���"
'        End If
'    Else
'       '�u�Ď��ռ��я�������ʁF�V�X�e�������������ُ�v���O�o��
'       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'       lblKekka.ForeColor = SYSFORMAT_ERROR
'       lblKekka.Caption = "�������Ɏ��s���܂���"
'    End If
'End If
'    '����������I�����̏���
'    OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
'    OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
'    '���O�C�����[�U�`�F�b�N
'    If pbUserLevel = 1 Then
'       OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
'    End If
'     cmdZikko.Enabled = True        '�u���������s�v�t�����s��
'     cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
'V1.5.0.1 DEL END
Exit Sub

'V1.7.0.1 ADD START
ERR_SPACE3:
'�u�L�����Z���t�����v
'�u�Ď��ռ��я�������ʁF���������������s�v���O�o��
Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
OptKoumoku(0).Enabled = True    '�u�o�׎��������v���W�I�t�I��s��
OptKoumoku(1).Enabled = True    '�u���ڑI���v���W�I�t�I��s��
'���O�C�����[�U�`�F�b�N
If pbUserLevel = 1 Then
   OptKoumoku(2).Enabled = True  '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
End If
cmdZikko.Enabled = True         '�u���������s�v�t�����s��
cmdCancel.Enabled = True        '�u���j���[��ʂ֖߂�v�t�����s��
Exit Sub
'V1.7.0.1 ADD END

ERR_SPACE2:
        '�G���[�������̏���
        OptKoumoku(0).Enabled = True    '�u�o�׎��������v���W�I�t�I��s��
        OptKoumoku(1).Enabled = True    '�u���ڑI���v���W�I�t�I��s��
        '���O�C�����[�U�`�F�b�N
        If pbUserLevel = 1 Then
           OptKoumoku(2).Enabled = True '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
        End If
        cmdZikko.Enabled = True         '�u���������s�v�t�����s��
        cmdCancel.Enabled = True        '�u���j���[��ʂ֖߂�v�t�����s��
        '�u�Ď��ռ��я�������ʁF�V�X�e�������������ُ�v���O�o��
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
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    On Error Resume Next

     '�u�Ď��ՃV�X�e���������F�����v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SYSFORMAT_GAMEN_END, 0)
     'frmALLSysformat.ZOrder 'V1.5.0.1 DEL
     frmSysformatMenu.ZOrder 'V1.5.0.1 ADD
     Unload Me

End Sub

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
   '�u�Ď��ՃV�X�e����������ʁF�V�X�e���t�@�C���폜�ُ�v���O�o��
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SYSFILE_DELETE_ERROR, lngErrCode)
   Set fs = Nothing
End Function

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
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//             �@�@�t�F�[�Y�P�s��Ή��@�uDoEvents�v�ɂĉ�ʂ̕`�ʂ��s���B
'//     REVISIONS :(1.12.0.1) 2009-11-12  REVISED BY [TCC] C.Terui
'//                 ���X�g�{�b�N�X�̃X�N���[�������ǉ�
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-19  REVISED BY [TCC] M.Matsumoto
'//                 �y��-313�Ή��z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-240�z
'//     REVISIONS :(EG20 V5.3.0.1) 2012-03-16  CODED BY  [TCC] H.Sugimoto
'//                 EG20�y5002P2 TR-19�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sFileDelete()
    Dim iFileNo As Integer
    Dim sFileData As String
    Dim iMozi, iKbn As Integer
    Dim sShubetu As String
    Dim sRoot As String
    Dim sPass As String
    Dim sKomoku As String
    Dim bSyori As Boolean
    Dim fs As Object
    Dim MyName As String
    Dim i As Integer
    Dim sChkPass As String
    Dim iRet As Integer
    Dim lngErrCode As Long       '�G���[�R�[�h
    Dim lBool As Boolean         ' EG20 V2.0.1.1�y����TR-240�z�ǉ�

    sFileDelete = False

    On Error GoTo ERR_SPACE
        
    '�t�@�C���L���`�F�b�N
    MyName = Dir(KANSI_SYSTEMFILE, vbNormal)
    If MyName = "" Then
        GoTo ERR_SPACE
    End If

' EG20 V3.3.0.1�y����TR-240�z�ǉ��J�n�i�ʒu�ړ��j
    ' �ێ烍�O�t�@�C��CLOSE
    lBool = dllCloseHoshuLogFile()
' EG20 V3.3.0.1�y����TR-240�z�ǉ��I���i�ʒu�ړ��j

    iFileNo = FreeFile                                           '���g�p�̃t�@�C���ԍ����擾����B
    Open KANSI_SYSTEMFILE For Input As #iFileNo                  '�V�X�e���������ݒ�t�@�C�����J���B
    Line Input #iFileNo, sFileData                               ' �P�s�ڂ͑S�̃o�[�W�����Ȃ̂œǔ�΂��B
    Do While Not EOF(iFileNo)
    Line Input #iFileNo, sFileData                               ' �P�s���Ǎ��ށB
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
                    sPass = PATH_KANSI & sPass
                Case 2  '�o�b�N�A�b�v���[�g
                    If sPass = "" Then
                       sPass = Mid(PATH_FKANSI, 1, Len(PATH_FKANSI) - 2)
                    Else
                       sPass = PATH_FKANSI & sPass
                    End If
                Case 4  '���O���[�g
                    sPass = PATH_EKANSI & sPass
' EG20 V5.3.0.1�ǉ��J�n
                Case 5  ' �p�X�w�薳���i�t���p�X�j
                    ' �p�X��ʂ̖����� sPass = sPass
' EG20 V5.3.0.1�ǉ��I��
            End Select
                    
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
                          DoEvents  'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                      '�t�H���_�̍폜�^�쐬
                      Case 3:
                          fs.DeleteFolder (sPass), True
                          fs.CreateFolder (sPass)
                          LstStatus.AddItem "�폜�^�쐬�����t�H���_ - " & sPass
                          DoEvents  'V1.5.0.1 ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                  End Select
                '�I�u�W�F�N�g���
                Set fs = Nothing
            Else
                '�w��o�`�r�r�i�V
                Select Case sShubetu
                   Case 2:
                       LstStatus.AddItem "�w��t�@�C���Ȃ� - " & sPass
                       DoEvents  'V1.5.0.1 ADD
                       LstStatus.Selected(LstStatus.ListCount - 1) = True           'V1.12.0.1 ADD
                   Case 3:
                       Set fs = CreateObject("Scripting.FileSystemObject")
                       '�t�@�C���L���`�F�b�N
'                       For i = 0 To Len(sPass)         'EG20 V2.1.0.1 DEL �y��-313�Ή��z
                       For i = 0 To Len(sPass) - 1      'EG20 V2.1.0.1 ADD �y��-313�Ή��z
                           If Mid(sPass, Len(sPass) - i, 1) = "\" Then
                               sChkPass = Left(sPass, Len(sPass) - i - 1)
                               Exit For
                           End If
                       Next
                       MyName = Dir(sChkPass, vbDirectory)
                       If MyName = "" Then
                           LstStatus.AddItem "�t�H���_�쐬���s - " & sPass
                           DoEvents  'V1.5.0.1 ADD
                           LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
                       Else
                           fs.CreateFolder (sPass)
                           LstStatus.AddItem "�쐬�����t�H���_ - " & sPass
                           DoEvents  'V1.5.0.1 ADD
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
   '�u�Ď��ՃV�X�e����������ʁF�t�@�C���E�t�H���_�������ُ�v���O�o��
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TARGET_FILE_FOLDER_DELETE_ERROR, lngErrCode)
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
   
   '�u�Ď��ռ��я�������ʁF�ڍזt�����v���O�o��
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
       For i = 0 To 7
        ShosaiMoji(i) = sWork
       Next
       Exit Sub
   End If
    
   For iGate = CNT_MIN To 7
      ' SysFormatShousai.ini��蕶�����擾����B
       sGateData = ""
       iKey = SYS_KEY_NAME & iGate
       lSts = GetPrivateProfileString(SYS_KANSI_SECTION_NAME, _
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmKansiSysformat.Caption, False
        pfFormActive (frmKansiSysformat.hwnd)
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
        
    '//////////////////////////////////////////////
    '�����ݒ�A�Ď��ݒ�̕ۑ��p�t�@�C�����쐬����B
    '//////////////////////////////////////////////
    '�����ݒ�t�@�C���L���`�F�b�N
    NameChk = Dir(G_SETTEI_FILE, vbNormal)
    If NameChk <> "" Then
       FileCopy G_SETTEI_FILE, SHOKI_G_SETTEI_FILE
    End If
    
    '�Ď��ݒ�t�@�C���L���`�F�b�N
    NameChk = Dir(K_SETTEI_FILE, vbNormal)
    If NameChk <> "" Then
       FileCopy K_SETTEI_FILE, SHOKI_K_SETTEI_FILE
    End If
    
    sCreateShokiFile = True
    '�u�Ď��ՃV�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
    
    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�u�Ď��ՃV�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬�ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SHOKI_CREATE_ERROR, lngErrCode)
    sCreateShokiFile = False
End Function
'V1.4.0.1�@ADD END

'V1.5.0.1 ADD START
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
  'V1.20.0.1 ADD START
  Dim bLDURet As Boolean  'LDU���O�t���O
  Dim bIDURet As Boolean  'IDU���O�t���O
  'V1.20.0.1 ADD END
  
   On Error Resume Next

  '�҂����Ԃ�INI��`�𒴂������ǂ����`�F�b�N
  If lngtime <= lngMAX_Time Then
    '�A�v���N���`�F�b�N���s���B�S�A�v�����I�������Ƃ��̂݁A�������������s���B
    'If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then 'V1.20.0.1 DEL
    If CheckAppStart(PROC_KANRI) = 0 Then 'V1.20.0.1 ADD
      '�A�v���N���`�F�b�N�^�C�}���~����B
      tmrAplTimer.Enabled = False
      'V1.20.0.1 DEL START
'      '����������
'      DeleteFile_Folder
      'V1.20.0.1 DEL END
      'V1.20.0.1  ADD START
      If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
         bIDURet = EndIDULog 'IDU���O�N������IDU���O�ɑ΂��ă��O�I���v��CMD���M
      Else
         bIDURet = True
      End If
      If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
         bLDURet = EndLDULog  'LDU���O�N������LDU���O�ɑ΂��ă��O�I���v��CMD���M
      Else
         bLDURet = True
      End If
      
      If bIDURet = True And bLDURet = True Then
         lngtime = 0
         lngtime = MN_MAIL_INTERVAL
         tmrLogTimer.Enabled = True
      Else
         '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "�������Ɏ��s���܂���"
         '����������I�����̏���
          OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
          OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
          '���O�C�����[�U�`�F�b�N
          If pbUserLevel = 1 Then
             OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
          End If
          cmdZikko.Enabled = True        '�u���������s�v�t�����s��
          cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
          Exit Sub
      End If
      'V1.20.0.1  ADD END
    Else
    '�N���A�v���L��̏ꍇ�A�^�C�}�𒣂蒼��
      tmrAplTimer.Interval = MN_MAIL_INTERVAL
    '���v�o�ߑ҂����Ԃ��A�b�v
     lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI��`�l�𒴂����ꍇ�A�����������ُ�Ƃ���B
    '�u�Ď��ՃV�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "�������Ɏ��s���܂���"
    '����������I�����̏���
    OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
    OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
    '���O�C�����[�U�`�F�b�N
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
    End If
    cmdZikko.Enabled = True        '�u���������s�v�t�����s��
    cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
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
    If CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
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
    '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "�������Ɏ��s���܂���"
    '����������I�����̏���
    OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
    OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
    '���O�C�����[�U�`�F�b�N
    If pbUserLevel = 1 Then
       OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
    End If
    cmdZikko.Enabled = True        '�u���������s�v�t�����s��
    cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
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
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-23  REVISED BY [TCC] T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�z
'//                �E�ێ烍�O�t�@�C���b�k�n�r�d�����ǉ�
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-240�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub DeleteFile_Folder()

    Dim i As Integer
    Dim bRtn As Boolean
    Dim bSentaku As Boolean
    Dim iRet As Integer
    Dim lExitCode As Long
    Dim bRtn1 As Boolean
    Dim bRtn2 As Boolean
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim iTargetDB As Integer            '�Ώ�DB�l
    Dim bDB_Code As Boolean
    Dim iRetIDULog As Integer           'IDU���O�N���t���O
    Dim iRetLDULog As Integer           'IDU���O�N���t���O
    Dim bRet As Boolean
  
    Dim lBool As Boolean                ' EG20 V2.0.1.1�y�c����54�zADD
 
    'EG20 V2.1.0.1 ADD START �y��-313�Ή��z
    Dim intLoop As Integer
    Dim lSts As Long
    'EG20 V2.1.0.1 ADD END
    
    On Error GoTo ERR_SPACE
  
  '�o�׎��������I�����A�S�ď�����(DLL�f�[�^��)�I�����A���̑��f�[�^���������Ώێ�
  If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkSonota.Value = 1 Then
     
' EG20 V3.3.0.1�y����TR-240�z�폜�J�n�i�ʒu�ړ��j
'     ' EG20 V2.0.1.1�y�c����54�zADD START
'     If OptKoumoku(0).Value = True Or OptKoumoku(2).Value = True Or chkLog.Value = 1 Then
'
'        ' �ێ烍�O�t�@�C��CLOSE
'         lBool = dllCloseHoshuLogFile()
'      End If
'     ' EG20 V2.0.1.1�y�c����54�zADD START
' EG20 V3.3.0.1�y����TR-240�z�폜�I���i�ʒu�ړ��j
      
     If sCreateShokiFile = False Then
        '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "�������Ɏ��s���܂���"
        '����������I�����̏���
        OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
        OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
        '���O�C�����[�U�`�F�b�N
        If pbUserLevel = 1 Then
           OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
        End If
        cmdZikko.Enabled = True        '�u���������s�v�t�����s��
        cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��
        '�����𔲂���
        Exit Sub
      End If
   End If

   '�V�X�e���t�@�C���̍폜
'   If bChk(4) = True Then              ' EG20 V3.3.0.1�y����TR-240�z�폜
   If bChk(5) = True Then               ' EG20 V3.3.0.1�y����TR-240�z�ǉ�
      bRtn1 = sSysFileDelete()
   Else
      bRtn1 = True
   End If

   '�t�H���_�A�t�@�C���̍폜
   If bRtn1 = True Then

      If sFileDelete() = True Then

         bDB_Code = True
         
         If bChk(1) = True Then
            Me.LstStatus.AddItem "DB������:" & chkMeisai.Caption
            DoEvents
            LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
            
            '�Ď��ՁF�ꌏ����
            Me.LstStatus.AddItem "�ꌏ���׃R�[�i�P�@DB�������J�n"
            DoEvents
            iTargetDB = stsKansiMeisai
            bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
            Me.LstStatus.AddItem "�ꌏ���׃R�[�i�P�@DB�������I��"
            DoEvents
            
            If bDB_Code = True Then
               '�Ď��ՁF�ꌏ���ׁi�R�[�i�Q�j
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�Q�@DB�������J�n"
               DoEvents
               iTargetDB = stsKansiMeisai2
               'DB����������
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�Q�@DB�������I��"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '�Ď��ՁF�ꌏ���ׁi�R�[�i�R�j
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�R�@DB�������J�n"
               DoEvents
               iTargetDB = stsKansiMeisai3
               'DB����������
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�R�@DB�������I��"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '�Ď��ՁF�ꌏ���ׁi�R�[�i�S�j
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�S�@DB�������J�n"
               DoEvents
               iTargetDB = stsKansiMeisai4
               'DB����������
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�S�@DB�������I��"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '�Ď��ՁF�ꌏ���ׁi�R�[�i�T�j
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�T�@DB�������J�n"
               DoEvents
               iTargetDB = stsKansiMeisai5
               'DB����������
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�T�@DB�������I��"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '�Ď��ՁF�ꌏ���ׁi�R�[�i�U�j
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�U�@DB�������J�n"
               DoEvents
               iTargetDB = stsKansiMeisai6
               'DB����������
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "�ꌏ���׃R�[�i�U�@DB�������I��"
               DoEvents
            End If
            
            If bDB_Code = True Then
               '�Ď��ՁF�ʏW�D
               Me.LstStatus.AddItem "�ʏW�D�@DB�������J�n"
               DoEvents
               iTargetDB = stsKansiBetu
               'DB����������
               bDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
               Me.LstStatus.AddItem "�ʏW�D�@DB�������I��"
               DoEvents
            End If
         
            'EG20 V2.1.0.1 ADD START �y��-313 START�z
            For intLoop = 1 To 6
                If intLoop = 1 Then
                    lSts = WritePrivateProfileString(SHKEI_EKITUDO_INI_SECTION, _
                           SHKEI_EKITUDO_INI_CNGFLG_KEY, "1", SHUKEI_EKITUDO_FILE)
                Else
                    lSts = WritePrivateProfileString(SHKEI_EKITUDO_INI_SECTION & CStr(intLoop), _
                           SHKEI_EKITUDO_INI_CNGFLG_KEY, "1", SHUKEI_EKITUDO_FILE)
                End If
            Next intLoop
            'EG20 V2.1.0.1 ADD END
         End If

         If bDB_Code = True Then
            '�u�Ď��ռ��я�������ʁF�V�X�e����������������v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
            lblKekka.ForeColor = SYSFORMAT_OK
            lblKekka.Caption = "�������͐������܂���"
         Else
            '�u�Ď��ռ��я�������ʁFDB�����������ُ�v���O�o��
             Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
             lblKekka.ForeColor = SYSFORMAT_ERROR
             lblKekka.Caption = "�������Ɏ��s���܂���"
         End If
      Else
        '�u�Ď��ռ��я�������ʁF�V�X�e�������������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "�������Ɏ��s���܂���"
      End If
   Else
      '�u�Ď��ռ��я�������ʁF�V�X�e�������������ُ�v���O�o��
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
      lblKekka.ForeColor = SYSFORMAT_ERROR
      lblKekka.Caption = "�������Ɏ��s���܂���"
   End If
 
 '����������I�����̏���
 OptKoumoku(0).Enabled = True      '�u�o�׎��������v���W�I�t�I��s��
 OptKoumoku(1).Enabled = True      '�u���ڑI���v���W�I�t�I��s��
 '���O�C�����[�U�`�F�b�N
 If pbUserLevel = 1 Then
    OptKoumoku(2).Enabled = True   '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
 End If
 cmdZikko.Enabled = True        '�u���������s�v�t�����s��
 cmdCancel.Enabled = True       '�u���j���[��ʂ֖߂�v�t�����s��

Exit Sub

ERR_SPACE2:
  '�G���[�������̏���
  OptKoumoku(0).Enabled = True    '�u�o�׎��������v���W�I�t�I��s��
  OptKoumoku(1).Enabled = True    '�u���ڑI���v���W�I�t�I��s��
  '���O�C�����[�U�`�F�b�N
  If pbUserLevel = 1 Then
     OptKoumoku(2).Enabled = True '�u�S�ď�����(�v���O��������f�[�^�܂�)�v���W�I�t�I��s��
  End If
  cmdZikko.Enabled = True         '�u���������s�v�t�����s��
  cmdCancel.Enabled = True        '�u���j���[��ʂ֖߂�v�t�����s��
  '�u�Ď��ռ��я�������ʁF�V�X�e�������������ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
  lblKekka.ForeColor = SYSFORMAT_ERROR
  lblKekka.Caption = "�������Ɏ��s���܂���"
ERR_SPACE:
End Sub
