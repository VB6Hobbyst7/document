VERSION 5.00
Begin VB.Form frmAppConfig 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ꊇ�N���E�I��"
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
   ScaleHeight     =   8625
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrLogTimer 
      Left            =   3600
      Top             =   8040
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   1080
      Top             =   7920
   End
   Begin VB.Timer tmrMail 
      Left            =   480
      Top             =   7920
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   9793
      TabIndex        =   13
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "�V���b�g�_�E���E���u�[�g"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   242
      TabIndex        =   8
      Top             =   4440
      Width           =   11535
      Begin VB.CommandButton cmdShoutDown 
         Caption         =   "�V���b�g�_�E��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   480
         TabIndex        =   10
         Top             =   330
         Width           =   2145
      End
      Begin VB.CommandButton cmdReboot 
         Caption         =   "���u�[�g"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   480
         TabIndex        =   9
         Top             =   1030
         Width           =   2145
      End
      Begin VB.Label lblAllEndApl 
         Caption         =   "�A�v���N�����̏ꍇ�͑S�ẴA�v���P�[�V�������I�����A�ċN������B"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   12
         Top             =   1200
         Width           =   7215
      End
      Begin VB.Label lblAllEndApl 
         Caption         =   "�A�v���N�����̏ꍇ�S�ẴA�v���P�[�V�������I�����A�d����؂�B"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   7215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�N���E�I���w��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   242
      TabIndex        =   3
      Top             =   600
      Width           =   11535
      Begin VB.Frame Frame2 
         Caption         =   "�A�v���P�[�V�����I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   10815
         Begin VB.CommandButton cmdAppEnd 
            Caption         =   "�A�v���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   19
            Top             =   330
            Width           =   2145
         End
         Begin VB.CommandButton cmdAppAllEnd 
            Caption         =   "�A�v�����S�I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   18
            Top             =   1030
            Width           =   2145
         End
         Begin VB.Label lblAllEndApl 
            Caption         =   "�S�ẴA�v���P�[�V�������A�����Ď��Ղ̕ێ�̂ݎc���ďI������B"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   21
            Top             =   480
            Width           =   7215
         End
         Begin VB.Label lblAllEndApl 
            Caption         =   "�S�ẴA�v���P�[�V�������I�����AWindows�݂̂̏�Ԃɂ���B"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   20
            Top             =   1200
            Width           =   7215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "�A�v���P�[�V�����N��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   939
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   10815
         Begin VB.CommandButton cmdAppStart 
            Caption         =   "�A�v���N��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   15
            Top             =   330
            Width           =   2145
         End
         Begin VB.Label lblAllEndApl 
            Caption         =   "�S�ẴA�v���P�[�V�������N������B"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   16
            Top             =   480
            Width           =   7215
         End
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "LDU"
         Height          =   375
         Index           =   3
         Left            =   7920
         TabIndex        =   7
         Top             =   320
         Width           =   1215
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "IDU"
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   6
         Top             =   320
         Width           =   1335
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "�����Ď���"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   5
         Top             =   320
         Width           =   1695
      End
      Begin VB.OptionButton Koumoku 
         Caption         =   "�S�A�v���ꊇ"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   320
         Value           =   -1  'True
         Width           =   1935
      End
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
      Top             =   15420
      Width           =   2895
   End
   Begin VB.ListBox LstStatus 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   242
      TabIndex        =   1
      Top             =   6360
      Width           =   11535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�A�v���P�[�V�����N���E�I��"
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
      TabIndex        =   2
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmAppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmAppConfig.frm
'//  �p�b�P�[�W���F�A�v���N���E�I�����
'//
'//  �T�v�F�A�v���N���E�I�����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R�Ď��Ձ@������Ή�
'//     REVISIONS :(2.4.0.1) 2010-10-27   REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@������Ή� �s��C���i���W�I�t�j
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Ή��y�c����54�z
'//                 �E�|�b�v�A�b�v�\�����b�Z�[�W�ύX
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.10�C���Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private iTimeCnt As Integer
Private iChoseAplEndSta As Integer  '�I�����W�I�t
Private Const AllApl = 0
Private Const KANSIApl = 1
Private Const IDUApl = 2
Private Const LDUApl = 3
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000        '�A�v���N���^�C�}�f�t�H���g�l
Dim lngMAX_Time As Long                    'INI�擾�ݒ�l
Dim lngtime     As Long                    '���݃^�C�}�l
Private Const APL_END = 4                  '�A�v���I���t����
Private Const APL_SHOUT_DOWN = 5           '�V���b�g�_�E���t����
Private Const APL_REBOOT = 6               '���u�[�g�t����
'V1.5.0.1 ADD END
'V1.7.0.1 ADD START
Private iChoseEnd As Integer  '�I���I������
Private Const NotEnd = -1
'V1.7.0.1 ADD END
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        '���O�N���^�C�}�f�t�H���g�l(30�b)
Dim lngLogMAX_Time As Long                'INI�擾�ݒ�l(���O�j
'V1.20.0.1 ADD END
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �A�v���N���E�I�����(�A�N�e�B�u��)
'//  �@�\�T�v  : �őO�ʕ\�����s���B
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
Private Sub Form_Activate()
    pfFormActive (hwnd)
    tmrMail.Enabled = True  'V1.3.0.1 ADD    '���[����M�^�C�}���N������B
End Sub

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �A�v���N���E�I�����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
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
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �A�v���N���E�I�����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    On Error Resume Next
 
    '�u�A�v���N���E�I����ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_END_GAMEN_START, 0)
    
    '������
    LstStatus.Clear

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    Koumoku(0).Value = True
    iChoseAplEndSta = AllApl
    iChoseEnd = NotEnd         'V1.7.0.1 ADD
    '�k�ރ`�F�b�N
    psIDUCheck
   
    If pbIDUSts = 1 Then
      'IDU�Ɩ���\��
       Koumoku(2).Enabled = False
    End If
      
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
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdCancel_Click
'//  �@�\����  : �u�����e�i���X��ʂ֖߂�v�t����������
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
   
   '�u�A�v���N���E�I����ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_END_GAMEN_END, 0)
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Koumoku_Click
'//  �@�\����  : �N���I���I��t����������
'//  �@�\�T�v  : �������W�I�t��Ԃ̉�ʂ�\������B
'//              [�S�A�v���ꊇ][�Ď���][IDU][LDU]
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@[IN]�������W�I�t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Koumoku_Click(Index As Integer)
    
    Select Case Index
        Case 0
          lblAllEndApl(0).Caption = "�S�ẴA�v���P�[�V�������N������B"
'          lblAllEndApl(1).Caption = "�S�ẴA�v���P�[�V�������A�Ď��Ղ̕ێ�̂ݎc���ďI������B"       'EG20 V2.1.0.1 DEL �yMainte_03_01�z
          lblAllEndApl(1).Caption = "�S�ẴA�v���P�[�V�������A�����Ď��Ղ̕ێ�̂ݎc���ďI������B"    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
          lblAllEndApl(2).Caption = "�S�ẴA�v���P�[�V�������I�����AWindows�݂̂̏�Ԃɂ���B"
          cmdAppEnd.Enabled = True
          cmdAppAllEnd.Enabled = True
          iChoseAplEndSta = AllApl
        Case 1
'          lblAllEndApl(0).Caption = "�Ď��ՃA�v���P�[�V�������N������B"           'EG20 V2.1.0.1 DEL �yMainte_03_01�z
          lblAllEndApl(0).Caption = "�����Ď��ՃA�v���P�[�V�������N������B"        'EG20 V2.1.0.1 ADD �yMainte_03_01�z
          lblAllEndApl(1).Caption = ""
          lblAllEndApl(2).Caption = ""
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
          iChoseAplEndSta = KANSIApl
        Case 2
          lblAllEndApl(0).Caption = "IDU�A�v���P�[�V�������N������B"
          lblAllEndApl(1).Caption = ""
          lblAllEndApl(2).Caption = "IDU�A�v���P�[�V�������I������B"
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = True
          iChoseAplEndSta = IDUApl
        Case 3
          lblAllEndApl(0).Caption = "LDU�A�v���P�[�V�������N������B"
          lblAllEndApl(1).Caption = ""
          lblAllEndApl(2).Caption = "LDU�A�v���P�[�V�������I������B"
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = True
          iChoseAplEndSta = LDUApl
    End Select
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdAppStart_Click
'//  �@�\����  : �A�v���N���t����������
'//  �@�\�T�v  : �ΏۃA�v���P�[�V�������N������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R�Ď��Ձ@������Ή�
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Ή��y�c����54�z
'//                 �E�|�b�v�A�b�v�\�����b�Z�[�W�ύX
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��i�v���O���X�o�[�N���Ή��j
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub cmdAppStart_Click()
    Dim iRet As Integer '�߂�l
    
    On Error Resume Next
   
    '�u�A�v���N���E�I����ʁF�A�v���N���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_BUTTOM, 0)
      
    Select Case iChoseAplEndSta
       Case AllApl   '�S�A�v���ꊇ�N��
           iRet = 1
           
           If CheckAppStart(PROC_KANRI) <> 0 Then
               '2�d�N��
               iRet = 0
           ElseIf CheckAppStart(PROCESS_IDU_PC) <> 0 Then
               '2�d�N��
               iRet = 0
           ElseIf CheckAppStart(PROCESS_LDU_PC) <> 0 Then
               '2�d�N��
               iRet = 0
           End If
           
           If iRet = False Then
              '2�d�x���N���|�b�v�A�b�v�\��
'               iRet = MsgBox("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")       'EG20 V2.1.0.1 DEL �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL START
'               iRet = MsgBox("�����Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
               iRet = MsgBox("�����Ď��ՁA�h�c�t�A�k�c�t�A�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zADD END
               Exit Sub
           End If
           '�N���m�F�|�b�v�A�b�v�\��
'           iRet = MsgBox("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL START
'           iRet = MsgBox("�����Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")           'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
           iRet = MsgBox("�����Ď��ՁA�h�c�t�A�k�c�t�A�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")           'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zADD END
           If iRet = vbCancel Then
             '[�L�����Z��]�t�����Ȃ�I��
             Exit Sub
           End If
            
           '��ʂ����b�N����B
           SetEnableFalse
           '�u�A�v���N���E�I����ʁF�S�A�v���ꊇ�N���v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_ALL, 0)

' EG20 V3.0.0.2 �ǉ��J�n
            ' �v���O���X�o�[�N��
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 �ǉ��I��

           '�S�A�v�����ꊇ�N������B
           '�Ǘ��N��
             iRet = CheckAppStartComplete(FLD_KPROGNOW & "\\" & PROC_KANRI, 1)
           'IDU�N��
            If pbIDUSts = 0 Then 'V2.3.0.1 ADD
             iRet = CheckAppStartComplete(PATH_IDU_APP & PATH_IDU_PROG & PROCESS_LUNCHER, 1)
             Sleep (10000)
            End If 'V2.3.0.1 ADD
           'LDU�N��
             iRet = CheckAppStartComplete(PATH_LDU_APP & PATH_LDU_PROG & PROCESS_LDU_LUNCHER, 1)
             Sleep (10000)
           '�S�A�v���̋N���`�F�b�N���s���B
           '�Ǘ��`�F�b�N
           If CheckAppStart(PROC_KANRI) = 0 Then
              '�u�A�v���N���E�I����ʁF�A�v���N�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'              LstStatus.AddItem ("�Ď��ՃA�v���P�[�V�����̋N���Ɏ��s���܂����B")       'EG20 V2.1.0.1 DEL �yMainte_03_01�z
              LstStatus.AddItem ("�����Ď��ՃA�v���P�[�V�����̋N���Ɏ��s���܂����B")    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
              LstStatus.ListIndex = LstStatus.ListCount - 1
'          ElseIf CheckAppStart(PROCESS_IDU_PC) = 0 Then 'V2.3.0.1 DEL
           ElseIf CheckAppStart(PROCESS_IDU_PC) = 0 And pbIDUSts = 0 Then 'V2.3.0.1 ADD
             '�u�A�v���N���E�I����ʁF�A�v���N�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1�y�c����54�zDEL START
'              LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
              LstStatus.AddItem ("�h�c�t�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
            ElseIf CheckAppStart(PROCESS_LDU_PC) = 0 Then
             '�u�A�v���N���E�I����ʁF�A�v���N�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1�y�c����54�zDEL START
'              LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
              LstStatus.AddItem ("�k�c�t�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
            Else
             '�u�A�v���N���E�I����ʁF�A�v���N����������v���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'              LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͐���ɋN�����܂����B")     'EG20 V2.1.0.1 DEL �yMainte_03_01�z
              LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͐���ɋN�����܂����B")  'EG20 V2.1.0.1 ADD �yMainte_03_01�z
              LstStatus.ListIndex = LstStatus.ListCount - 1
           End If
            
            '��ʂ����b�N����������B
            SetEnableTrue
       
       Case KANSIApl '�Ď��ՋN��
            If CheckAppStart(PROC_KANRI) <> 0 Then
               '2�d�x���N���|�b�v�A�b�v�\��
'               iRet = MsgBox("�Ď��ՃA�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")     'EG20 V2.1.0.1 DEL �yMainte_03_01�z
               iRet = MsgBox("�����Ď��ՃA�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")  'EG20 V2.1.0.1 ADD �yMainte_03_01�z
               Exit Sub
            End If
            
            '�N���m�F�|�b�v�A�b�v�\��
'            iRet = MsgBox("�Ď��ՃA�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")     'EG20 V2.1.0.1 DEL �yMainte_03_01�z
            iRet = MsgBox("�����Ď��ՃA�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")  'EG20 V2.1.0.1 ADD �yMainte_03_01�z
            If iRet = vbCancel Then
              '[�L�����Z��]�t�����Ȃ�I��
              Exit Sub
            End If
            
            '��ʂ����b�N����B
             SetEnableFalse
            '�u�A�v���N���E�I����ʁF�Ď��ՃA�v���N���v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_KANSI, 0)

' EG20 V3.0.0.2 �ǉ��J�n
            ' �v���O���X�o�[�N��
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 �ǉ��I��

            '�Ǘ��N��
             iRet = CheckAppStartComplete(FLD_KPROGNOW & "\\" & PROC_KANRI, 1)
            
            '�Ǘ��`�F�b�N
            If CheckAppStart(PROC_KANRI) = 0 Then
              '�u�A�v���N���E�I����ʁF�A�v���N�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'              LstStatus.AddItem ("�Ď��ՃA�v���P�[�V�����̋N���Ɏ��s���܂����B")       'EG20 V2.1.0.1 DEL �yMainte_03_01�z
              LstStatus.AddItem ("�����Ď��ՃA�v���P�[�V�����̋N���Ɏ��s���܂����B")    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
              LstStatus.ListIndex = LstStatus.ListCount - 1
           Else
             '�u�A�v���N���E�I����ʁF�A�v���N����������v���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'              LstStatus.AddItem ("�Ď��ՃA�v���P�[�V�����͐���ɋN�����܂����B")       'EG20 V2.1.0.1 DEL �yMainte_03_01�z
              LstStatus.AddItem ("�����Ď��ՃA�v���P�[�V�����͐���ɋN�����܂����B")    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
              LstStatus.ListIndex = LstStatus.ListCount - 1
            End If
             '��ʂ����b�N����������B
             SetEnableTrue
            
             cmdAppEnd.Enabled = False
             cmdAppAllEnd.Enabled = False
        
       Case IDUApl   'IDU�A�v���N��
           If CheckAppStart(PROCESS_IDU_PC) <> 0 Then
               '2�d�x���N���|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y�c����54�zDEL START
'               iRet = MsgBox("ID���p���j�b�g�A�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
               iRet = MsgBox("�h�c�t�A�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")
'EG20 V2.0.1.1�y�c����54�zADD END
              Exit Sub
            End If
            
            '�N���m�F�|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y�c����54�zDEL START
'            iRet = MsgBox("ID���p���j�b�g�A�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
            iRet = MsgBox("�h�c�t�A�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")
'EG20 V2.0.1.1�y�c����54�zADD END
            If iRet = vbCancel Then
              '[�L�����Z��]�t�����Ȃ�I��
              Exit Sub
            End If
            
            '��ʂ����b�N����B
             SetEnableFalse
            '�u�A�v���N���E�I����ʁFIDU�A�v���N���v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_IDU, 0)
  
' EG20 V3.0.0.2 �ǉ��J�n
            ' �v���O���X�o�[�N��
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 �ǉ��I��

            'IDU�N��
             iRet = CheckAppStartComplete(PATH_IDU_APP & PATH_IDU_PROG & PROCESS_LUNCHER, 1)
             Sleep (10000)
             DoEvents
            'IDU�`�F�b�N
            If CheckAppStart(PROCESS_IDU_PC) = 0 Then
             '�u�A�v���N���E�I����ʁF�A�v���N�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1�y�c����54�zDEL START
'              LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
              LstStatus.AddItem ("�h�c�t�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
           Else
             '�u�A�v���N���E�I����ʁF�A�v���N����������v���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'EG20 V2.0.1.1�y�c����54�zDEL START
'              LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����͐���ɋN�����܂����B")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
              LstStatus.AddItem ("�h�c�t�A�v���P�[�V�����͐���ɋN�����܂����B")
'EG20 V2.0.1.1�y�c����54�zADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
            End If
             '��ʂ����b�N����������B
             SetEnableTrue
             cmdAppEnd.Enabled = False
                    
       Case LDUApl   'LDU�A�v���N��
            If CheckAppStart(PROCESS_LDU_PC) <> 0 Then
               '2�d�x���N���|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y�c����54�zDEL START
'               iRet = MsgBox("LD���[�e�B���e�B�A�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
               iRet = MsgBox("�k�c�t�A�v���P�[�V�����͊��ɋN�����Ă��܂��B", vbOKOnly + vbExclamation, "�Q�d�N���x��")
'EG20 V2.0.1.1�y�c����54�zADD END
               Exit Sub
            End If
            
            '�N���m�F�|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y�c����54�zDEL START
'            iRet = MsgBox("LD���[�e�B���e�B�A�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
            iRet = MsgBox("�k�c�t�A�v���P�[�V�������N�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�N���m�F")
'EG20 V2.0.1.1�y�c����54�zADD END
            If iRet = vbCancel Then
              '[�L�����Z��]�t�����Ȃ�I��
              Exit Sub
            End If
            
            '��ʂ����b�N����B
             SetEnableFalse
            '�u�A�v���N���E�I����ʁFLDU�A�v���N���v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_LDU, 0)
           
' EG20 V3.0.0.2 �ǉ��J�n
            ' �v���O���X�o�[�N��
            Call psfuncStartupProgressBar
' EG20 V3.0.0.2 �ǉ��I��
           
            'LDU�N��
             iRet = CheckAppStartComplete(PATH_LDU_APP & PATH_LDU_PROG & PROCESS_LDU_LUNCHER, 1)
             Sleep (10000)

            'LDU�`�F�b�N
            If CheckAppStart(PROCESS_LDU_PC) = 0 Then
             '�u�A�v���N���E�I����ʁF�A�v���N�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_START_ERROR, 0)
'EG20 V2.0.1.1�y�c����54�zDEL START
'              LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
              LstStatus.AddItem ("�k�c�t�A�v���P�[�V�����̋N���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y�c����54�zADD END
              LstStatus.ListIndex = LstStatus.ListCount - 1
           Else
             '�u�A�v���N���E�I����ʁF�A�v���N����������v���O�o��
               Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_START_OK, 0)
'EG20 V2.0.1.1�y�c����54�zDEL START
'               LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����͐���ɋN�����܂����B")
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
               LstStatus.AddItem ("�k�c�t�A�v���P�[�V�����͐���ɋN�����܂����B")
'EG20 V2.0.1.1�y�c����54�zADD END
               LstStatus.ListIndex = LstStatus.ListCount - 1
            End If
             '��ʂ����b�N����������B
             SetEnableTrue
             cmdAppEnd.Enabled = False
             
   End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdAppAllEnd_Click
'//  �@�\����  : �A�v�����S�I���t����������
'//  �@�\�T�v  : �ΏۃA�v���P�[�V���������S�I������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Ή��y�c����54�z
'//                 �E�|�b�v�A�b�v�\�����b�Z�[�W�ύX
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.10�C���Ή��z
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub cmdAppAllEnd_Click()
    Dim iRet As Integer '�߂�l
    Dim uMail As ML_KYOTU_INF                        '���[��
    Dim udtMail As MAIL_IDU_LDU_APLEND_CMD           '���[��
    Dim bRtn As Boolean
    Dim lExitCode As Long
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '�Ď��ՃA�v����������
    Dim bIDURet As Boolean                          'IDU�A�v����������
    Dim bLDURet As Boolean                          'LDU�A�v����������
    Dim bIDULOGRet As Boolean                       'IDU���O��������
    Dim bLDULOGRet As Boolean                       'LDU���O��������
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    bIDULOGRet = False
    bLDULOGRet = False
    'V1.5.0.1 ADD END

    On Error Resume Next
    
    '�u�A�v���N���E�I����ʁF�A�v�����S�I���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_ALLEND_BUTTOM, 0)

    Select Case iChoseAplEndSta
       Case AllApl   '�S�A�v���ꊇ�I��
           If CheckAppStart(PROC_KANRI) = 0 Then
               '�I���ϊm�F�|�b�v�A�b�v�\��
'               iRet = MsgBox("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")         'EG20 V2.1.0.1 DEL �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL START
'               iRet = MsgBox("�����Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")      'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
               iRet = MsgBox("�����Ď��ՁA�h�c�t�A�k�c�t�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")      'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zADD END
               '�A�v���N���c�[���N��
                Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
           
               '�I������
                psEndHoshuProc
          
               '�ێ�v���Z�X�I��
                End
           End If
           
           '�I���m�F�|�b�v�A�b�v�\��
'           iRet = MsgBox("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL START
'           iRet = MsgBox("�����Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zDEL END
'EG20 V2.0.1.1�y�c����54�zADD START
           iRet = MsgBox("�����Ď��ՁA�h�c�t�A�k�c�t�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y�c����54�zADD START
           If iRet = vbCancel Then
             '[�L�����Z��]�t�����Ȃ�I��
             Exit Sub
           End If
            
            '��ʂ����b�N����B
            SetEnableFalse
            '�u�A�v���N���E�I����ʁF�S�A�v���ꊇ�I���v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_ALL, 0)
            '�A�v���I���v�����Ǘ��ɑ��M����
            uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
            uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
            uMail.udtlHeader.dwProid = RHOSHU_ID
            uMail.udtlHeader.dwSubArea = 0
            'V1.5.0.1 DEL START
            'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
            'If bRtn <> 0 Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 ADD START
            bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
            If bKansiRet <> 0 Then
            'V1.5.0.1 ADD END
              ' �u�A�v���N���E�I����ʁF���[�����M���팋�ʁv���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
            Else
              ' �u�A�v���N���E�I����ʁF���[�����M�ُ팋�ʁv���O�o��
              lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
              '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
              SetEnableTrue
              Exit Sub
            End If
     'V1.20.0.1 DEL START
'            'IDU/LDU���O�I���v��CMD���M
'            If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'              'IDU���O�I���v��CMD���M
'               'V1.5.0.1 DEL START
'               'bRtn = EndIDULog
'               'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'               bIDULOGRet = EndIDULog
'               If bIDULOGRet = False Then
'                 LstStatus.AddItem ("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'                 LstStatus.ListIndex = LstStatus.ListCount - 1
'               'V1.5.0.1 ADD END
'                SetEnableTrue
'                Exit Sub
'               End If
'            'V1.5.0.1 ADD START
'            Else
'               bIDULOGRet = True
'            'V1.5.0.1 ADD END
'            End If
'
'            If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'              'LDU���O�I���v��CMD���M
'               'V1.5.0.1 DEL START
'               'bRtn = EndLDULog
'               'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'               bLDULOGRet = EndLDULog
'               If bLDULOGRet = False Then
'                  LstStatus.AddItem ("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'                  LstStatus.ListIndex = LstStatus.ListCount - 1
'               'V1.5.0.1 ADD END
'                  SetEnableTrue
'                  Exit Sub
'               End If
'            'V1.5.0.1 ADD START
'            Else
'             bLDULOGRet = True
'            'V1.5.0.1 ADD END
'            End If
     'V1.20.0.1 DEL END
'V1.5.0.1 ADD START
            '�Ǘ��AIDU���O�ALDU���O�ւ̃��[�����M���펞�̂݁A�A�v���N���`�F�b�N�^�C�}���N�����A
            'INI�t�@�C�����擾�������Ԃ܂ŃA�v���N���`�F�b�N���s���B
            'If bKansiRet = True And bIDULOGRet = True And bLDULOGRet = True Then       'V1.20.0.1 DEL
             If bKansiRet = True Then                                                   'V1.20.0.1 ADD
               lngtime = 0
               lngtime = MN_MAIL_INTERVAL
               tmrAplTimer.Enabled = True
               iChoseEnd = AllApl 'V1.7.0.1 ADD
            End If
'V1.5.0.1 ADD END
'           V1.5.0.1 DEL START
'           If CheckAppEndComplete(PROC_KANRI, lExitCode) = 0 _
'            And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 _
'            And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'              '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
'              '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'              LstStatus.AddItem ("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'              SetEnableTrue
'              Exit Sub
'           End If
'           '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'
'           '�A�v���N���c�[���N��
'           Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
'
'           '�I������
'           psEndHoshuProc
'
'           '�ێ�v���Z�X�I��
'           End
'V1.5.0.1 DEL END
       Case IDUApl   'IDU�A�v��
           If CheckAppStart(PROCESS_IDU_PC) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 Then
               '�I���όx���|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y��54�zDEL START
'               iRet = MsgBox("ID���p���j�b�g�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
'               iRet = MsgBox("�h�c�t�͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")                   ' EG20 V3.6.0.1�폜
               iRet = MsgBox("�h�c�t�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")    ' EG20 V3.6.0.1�ǉ�
'EG20 V2.0.1.1�y��54�zADD END
              Exit Sub
           End If
            
           '�I���m�F�|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y��54�zDEL START
'           iRet = MsgBox("ID���p���j�b�g�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
'           iRet = MsgBox("�h�c�t���I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")                  ' EG20 V3.6.0.1�폜
           iRet = MsgBox("�h�c�t�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")   ' EG20 V3.6.0.1�ǉ�
'EG20 V2.0.1.1�y��54�zADD END
           If iRet = vbCancel Then
             '[�L�����Z��]�t�����Ȃ�I��
             Exit Sub
           End If
           
           '��ʂ����b�N����B
            SetEnableFalse
           '�u�A�v���N���E�I����ʁFIDU�A�v�����S�I���v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_IDU, 0)
           'ID���ɃA�v���I���v���𑗐M����B
              udtMail.mlHeader.dwId = ML_ID_IDU_APLEND_CMD
              udtMail.mlHeader.dwSize = MlSize.IDUAPLEND_REQ
              udtMail.mlHeader.dwProid = RHOSHU_ID
              udtMail.mlHeader.dwSubArea = 0
              udtMail.dwEndType = ML_ENDTYPE_APLEND
              udtMail.dwCMDLevel = ML_CMDLEVEL_TUJYO        'V1.5.0.1 ADD
            'V1.5.0.1 DEL START
              'bRtn = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
            'If bRtn <> 0 Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 ADD START
            bIDURet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
            If bIDURet <> 0 Then
            'V1.5.0.1 ADD END
              ' �u�A�v���N���E�I����ʁF���[�����M���팋�ʁv���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
            Else
              ' �u�A�v���N���E�I����ʁF���[�����M�ُ팋�ʁv���O�o��
             lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
              '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
              SetEnableTrue
              cmdAppEnd.Enabled = False
              Exit Sub
            End If
   'V1.20.0.1 DEL START
'            'IDU/LDU���O�I���v��CMD���M
'            If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'              'IDU���O�I���v��CMD���M
'               'V1.5.0.1 DEL START
'               'bRtn = EndIDULog
'               'If bRtn = False Then
'               'V1.5.0.1 DEL END
'               'V1.5.0.1 ADD START
'               bIDULOGRet = EndIDULog
'               If bIDULOGRet = False Then
'                  LstStatus.AddItem ("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'                  LstStatus.ListIndex = LstStatus.ListCount - 1
'               'V1.5.0.1 ADD END
'                  SetEnableTrue
'                  cmdAppEnd.Enabled = False
'                  Exit Sub
'               End If
'            End If
   'V1.20.0.1 DEL END
'V1.5.0.1 ADD START
            'IDU�A�v��(ID��)�AIDU���O���O�ւ̃��[�����M���펞�̂݁A�A�v���N���`�F�b�N�^�C�}���N�����A
            'INI�t�@�C�����擾�������Ԃ܂ŃA�v���N���`�F�b�N���s���B
            'If bIDURet = True And bIDULOGRet = True Then    'V1.20.0.1 DEL
            If bIDURet = True Then                           'V1.20.0.1 ADD
               lngtime = 0
               lngtime = MN_MAIL_INTERVAL
               tmrAplTimer.Enabled = True
               iChoseEnd = IDUApl 'V1.7.0.1 ADD
            End If
'V1.5.0.1 ADD END
'V1.5.0.1�@DEL�@START
'           'IDU�`�F�b�N
'           If CheckAppEndComplete(PROCESS_IDU_PC, lExitCode) = 0 And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 Then
'              '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'              LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           Else
'             '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'             LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̏I���ɐ������܂����B")
'             LstStatus.ListIndex = LstStatus.ListCount - 1
'           End If
'            '��ʂ����b�N����������B
'            SetEnableTrue
'            cmdAppEnd.Enabled = False
'V1.5.0.1�@DEL�@END
       Case LDUApl   'LDU�A�v���N��
           If CheckAppStart(PROCESS_LDU_PC) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
              '�I���όx���|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y��54�zDEL START
'              iRet = MsgBox("LD���[�e�B���e�B�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
              iRet = MsgBox("�k�c�t�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��")
'EG20 V2.0.1.1�y��54�zADD END
              Exit Sub
           End If
            
           '�I���m�F�|�b�v�A�b�v�\��
'EG20 V2.0.1.1�y��54�zDEL START
'           iRet = MsgBox("LD���[�e�B���e�B�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
           iRet = MsgBox("�k�c�t�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")
'EG20 V2.0.1.1�y��54�zADD END
           If iRet = vbCancel Then
             '[�L�����Z��]�t�����Ȃ�I��
             Exit Sub
           End If
            
           '��ʂ����b�N����B
            SetEnableFalse
           '�u�A�v���N���E�I����ʁFLDU�A�v�����S�I���v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_LDU, 0)
           'LD���ɃA�v���I���v���𑗐M����B
            udtMail.mlHeader.dwId = ML_ID_LDU_APLEND_CMD
            udtMail.mlHeader.dwSize = MlSize.LDUAPLEND_REQ
            udtMail.mlHeader.dwProid = RHOSHU_ID
            udtMail.mlHeader.dwSubArea = 0
            udtMail.dwEndType = ML_ENDTYPE_APLEND
            udtMail.dwCMDLevel = ML_CMDLEVEL_TUJYO        'V1.5.0.1 ADD
            'V1.5.0.1 DEL START
            'bRtn = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.mlHeader)
            'If bRtn <> 0 Then
            'V1.5.0.1 DEL END
            'V1.5.0.1 ADD START
            bLDURet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.mlHeader)
            If bLDURet <> 0 Then
            'V1.5.0.1 ADD END
              ' �u�A�v���N���E�I����ʁF���[�����M���팋�ʁv���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
            Else
              ' �u�A�v���N���E�I����ʁF���[�����M�ُ팋�ʁv���O�o��
             lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
              Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
              '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
              SetEnableTrue
              cmdAppEnd.Enabled = False
              Exit Sub
            End If
 'V1.20.0.1 DEL START
'            If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'              'LDU���O�I���v��CMD���M
'              'V1.5.0.1 DEL START
'               'bRtn = EndLDULog
'               'If bRtn = False Then
'              'V1.5.0.1 DEL END
'              'V1.5.0.1 ADD START
'              bLDULOGRet = EndLDULog
'              If bLDULOGRet = False Then
'                 LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'                 LstStatus.ListIndex = LstStatus.ListCount - 1
'              'V1.5.0.1 ADD END
'                  SetEnableTrue
'                  cmdAppEnd.Enabled = False
'                  Exit Sub
'               End If
'            End If
 'V1.20.0.1 DEL END
'V1.5.0.1 ADD START
            'LDU�A�v��(LD��)�ALDU���O���O�ւ̃��[�����M���펞�̂݁A�A�v���N���`�F�b�N�^�C�}���N�����A
            'INI�t�@�C�����擾�������Ԃ܂ŃA�v���N���`�F�b�N���s���B
            'If bLDURet = True And bLDULOGRet = True Then   'V1.20.0.1 DEL
            If bLDURet = True Then     'V1.20.0.1 ADD
               lngtime = 0
               lngtime = MN_MAIL_INTERVAL
               tmrAplTimer.Enabled = True
               iChoseEnd = LDUApl 'V1.7.0.1 ADD
            End If
'V1.5.0.1 ADD END
'V1.5.0.1�@DEL�@START
'           'LDU�`�F�b�N
'           If CheckAppEndComplete(PROCESS_LDU_PC, lExitCode) = 0 And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'              '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'              LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           Else
'             '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'              LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���ɐ������܂����B")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           End If
'            '��ʂ����b�N����������B
'            SetEnableTrue
'            cmdAppEnd.Enabled = False
'V1.5.0.1�@DEL�@END
   End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdAppEnd_Click
'//  �@�\����  : �A�v���I���t����������
'//  �@�\�T�v  : �ΏۃA�v���P�[�V�������I������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub cmdAppEnd_Click()
    Dim uMail As ML_KYOTU_INF           '���[��
    Dim bRtn As Boolean                 '���[���̖߂�l
    Dim iRetApp As Integer              '�Ď��ՏI���m�F�߂�l
    Dim iRetIDUApp As Integer           'IDU�I���m�F�߂�l
    Dim iRetLDUApp As Integer           'LDU�I���m�F�߂�l
    Dim iRet As Integer                 '���b�Z�[�W�{�b�N�X�߂�l
    Dim lExitCode As Long
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '�Ď��ՃA�v����������
    Dim bIDURet As Boolean                          'IDU�A�v����������
    Dim bLDURet As Boolean                          'LDU�A�v����������
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1 ADD END
    
    On Error Resume Next
  
'    iChoseAplEndSta = APL_END           'V1.5.0.1 ADD 'V1.7.0.1 DEL
   
    '�u�A�v���N���E�I����ʁF�A�v���I���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_BUTTOM, 0)

    If CheckAppStart(PROC_KANRI) = 0 Then
       '�I���όx���|�b�v�A�b�v�\��
       'iRet = MsgBox("�Ď��ՃA�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��") 'V1.8.0.1 DEL
'       iRet = MsgBox("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��") 'V1.8.0.1 ADD       'EG20 V2.1.0.1 DEL �yMainte_03_01�z
'EG20 V2.0.1.1�y��54�zDEL START
'       iRet = MsgBox("�����Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��") 'V1.8.0.1 ADD    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
       iRet = MsgBox("�����Ď��ՁA�h�c�t�A�k�c�t�A�v���P�[�V�����͊��ɏI�����Ă��܂��B", vbOKOnly + vbExclamation, "�I���όx��") 'V1.8.0.1 ADD    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y��54�zADD END
       Exit Sub
    End If
    
    '�I���m�F�|�b�v�A�b�v�\��
     'iRet = MsgBox("�Ď��ՃA�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")�@ �fV1.8.0.1�@DEL
'     iRet = MsgBox("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")    'V1.8.0.1�@ADD        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
'EG20 V2.0.1.1�y��54�zDEL START
'     iRet = MsgBox("�����Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")    'V1.8.0.1�@ADD     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
     iRet = MsgBox("�����Ď��ՁA�h�c�t�A�k�c�t�A�v���P�[�V�������I�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")    'V1.8.0.1�@ADD     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
'EG20 V2.0.1.1�y��54�zADD END
     If iRet = vbCancel Then
        '[�L�����Z��]�t�����Ȃ�I��
        Exit Sub
     End If
           
     '��ʂ����b�N����B
     SetEnableFalse
     
     '�u�A�v���N���E�I����ʁF�Ď��ՃA�v���I���v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_KANSI, 0)
     '�A�v���I���v�����Ǘ��ɑ��M����
     uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
     uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
     uMail.udtlHeader.dwProid = RHOSHU_ID
     uMail.udtlHeader.dwSubArea = 0
     'V1.5.0.1 DEL START
     'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
     'If bRtn <> 0 Then
     'V1.5.0.1 DEL END
     'V1.5.0.1 ADD START
     bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
     If bKansiRet <> 0 Then
     'V1.5.0.1 ADD END
        '�u�A�v���N���E�I����ʁF���[�����M���팋�ʁv���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
        '�A�v���I���m�F
        'iRetApp = CheckAppEndComplete(PROC_KANRI, lExitCode) 'V1.5.0.1 DEL
  'V1.20.0.1 DEL START
'        'IDU���O�m�F
'        If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'           'IDU���O�I���v��CMD���M
'           'V1.5.0.1 DEL START
'           'bRtn = EndIDULog
'           'If bRtn = False Then
'           'V1.5.0.1 DEL END
'           'V1.5.0.1 ADD START
'           bIDURet = EndIDULog
'           If bIDURet = False Then
'              LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")
'              LstStatus.ListIndex = LstStatus.ListCount - 1
'           'V1.5.0.1 ADD END
'              SetEnableTrue
'              Exit Sub
'           End If
'          'IDU���O�v���Z�X�I���m�F
'          'iRetIDUApp = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) 'V1.5.0.1 DEL
'        Else
'           iRetIDUApp = 1
'           bIDURet = True       'V1.5.0.1 ADD
'        End If
'        'LDU���O�m�F
'        If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'           'LDU���O�I���v��CMD���M
'            'V1.5.0.1 DEL START
'            'bRtn = EndLDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bLDURet = EndLDULog
'            If bLDURet = False Then
'               LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'           'LDU���O�v���Z�X�I���m�F
'            'iRetLDUApp = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)   'V1.5.0.1 DEL
'        Else
'           iRetLDUApp = 1
'           bLDURet = True    'V1.5.0.1 ADD
'        End If
  'V1.20.0.1 DEL END
     Else
        '�u�A�v���N���E�I����ʁF���[�����M�ُ팋�ʁv���O�o��
        lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
        '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
     End If

'V1.5.0.1 ADD START
     'If bKansiRet = True And bIDURet = True And bLDURet = True Then   'V1.20.0.1 DEL
     If bKansiRet = True Then                                          'V1.20.0.1 ADD
        lngtime = 0
        lngtime = MN_MAIL_INTERVAL
        tmrAplTimer.Enabled = True
        iChoseEnd = APL_END         'V1.7.0.1 ADD
     Else
        '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'        LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")       'EG20 V2.1.0.1 DEL �yMainte_03_01�z
        LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
        LstStatus.ListIndex = LstStatus.ListCount - 1
        '��ʂ����b�N����������B
        SetEnableTrue
     End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'     If iRetApp = 1 And iRetIDUApp = 1 And iRetLDUApp = 1 Then
'        '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'        LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���ɐ������܂����B")
'        LstStatus.ListIndex = LstStatus.ListCount - 1
'     Else
'        '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
'        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'        LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")
'        LstStatus.ListIndex = LstStatus.ListCount - 1
'     End If
'
'     '��ʂ����b�N����������B
'     SetEnableTrue
'V1.5.0.1 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdShoutDown_Click
'//  �@�\����  : �u�V���b�g�_�E���v�t����������
'//  �@�\�T�v  : OS���I������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub cmdShoutDown_Click()
    Dim bRtn As Boolean                 '���[���̖߂�l
    Dim iRet As Integer                 '���b�Z�[�W�{�b�N�X�߂�l
    Dim uMail As ML_KYOTU_INF           '���[��
    Dim lExitCode As Long               '�G���[�R�[�h
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '�Ď��ՃA�v����������
    Dim bIDURet As Boolean                          'IDU�A�v����������
    Dim bLDURet As Boolean                          'LDU�A�v����������
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1 ADD END
    
    On Error Resume Next
  
    'iChoseAplEndSta = APL_SHOUT_DOWN           'V1.5.0.1 ADD 'V1.7.0.1 DEL

    '�u�A�v���N���E�I����ʁF�V���b�g�_�E���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_SHOUT_DOWN_BUTTOM, 0)
 
    '�u�V���b�g�_�E���m�F�v�|�b�v�A�b�v�\��
    iRet = MsgBox("�R���s���[�^���V���b�g�_�E�����܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")
    If iRet = vbCancel Then
      '[�L�����Z��]�t�����Ȃ�I��
      Exit Sub
    End If
           
    '��ʂ����b�N����B
     SetEnableFalse
     
    If CheckAppStart(PROC_KANRI) <> 0 Then
       '�A�v���I���v�����Ǘ��ɑ��M����
       uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
       uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
       uMail.udtlHeader.dwProid = RHOSHU_ID
       uMail.udtlHeader.dwSubArea = 0
       'V1.5.0.1 DEL START
       'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       'If bRtn <> 0 Then
       'V1.5.0.1 DEL END
       'V1.5.0.1 ADD START
       bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       If bKansiRet <> 0 Then
       'V1.5.0.1 ADD END
          '�u�A�v���N���E�I����ʁF���[�����M���팋�ʁv���O�o��
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
  'V1.20.0.1 DEL START
'          'IDU���O�m�F
'          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'            'IDU���O�I���v��CMD���M
'            'V1.5.0.1 DEL START
'            'bRtn = EndIDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bIDURet = EndIDULog
'            If bIDURet = False Then
'               LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'         'V1.5.0.1 ADD START
'         Else
'          bIDURet = True
'         'V1.5.0.1 ADD END
'         End If
'         'LDU���O�m�F
'         If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'           'LDU���O�I���v��CMD���M
'            'V1.5.0.1 DEL START
'            'bRtn = EndLDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bLDURet = EndLDULog
'            If bLDURet = False Then
'               LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'         'V1.5.0.1 ADD START
'         Else
'          bLDURet = True
'        'V1.5.0.1 ADD END
'        End If
 'V1.20.0.1 DEL END
       Else
          '�u�A�v���N���E�I����ʁF���[�����M�ُ팋�ʁv���O�o��
          lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
          Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
          '�u�A�v���N���E�I����ʁF�V���b�g�_�E�����������ُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_SHOUT_DOWN_ERROR, 0)
          SetEnableTrue
          Exit Sub
       End If
'V1.5.0.1 ADD START
       'If bKansiRet = True And bIDURet = True And bLDURet = True Then 'V1.20.0.1 DEL
       If bKansiRet = True Then                                        'V1.20.0.1 ADD
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplTimer.Enabled = True
          iChoseEnd = APL_SHOUT_DOWN         'V1.7.0.1 ADD
       Else
           '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'           LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
           LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
           LstStatus.ListIndex = LstStatus.ListCount - 1
           '��ʃ��b�N����
           'SetEnableTrue     'V1.7.0.1 DEL
           'V1.7.0.1 ADD START
           If iChoseAplEndSta = AllApl Then
              '���W�I�t�F�S�A�v���ꊇ
              SetEnableTrue
           ElseIf iChoseAplEndSta = KANSIApl Then
              '���W�I�t�F�Ď���
              SetEnableTrue
              cmdAppEnd.Enabled = False
              cmdAppAllEnd.Enabled = False
           ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
              '���W�I�t�FIDU����LDU
              SetEnableTrue
              cmdAppEnd.Enabled = False
           End If
           'V1.7.0.1 ADD END
       End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'      If CheckAppEndComplete(PROC_KANRI, lExitCode) = 0 _
'          And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 _
'          And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'       End If
'
'       '�I������
'       psEndHoshuProc
'       '�V���b�g�_�E������
'       dllAPLEndDwon
'V1.5.0.1 DEL END
    Else
     '�I������
     psEndHoshuProc
     '�V���b�g�_�E������
     dllAPLEndDwon
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReboot_Click
'//  �@�\����  : �u���u�[�g�v�t����������
'//  �@�\�T�v  : OS���ċN������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReboot_Click()
    Dim bRtn As Boolean                 '���[���̖߂�l
    Dim iRet As Integer                 '���b�Z�[�W�{�b�N�X�߂�l
    Dim uMail As ML_KYOTU_INF           '���[��
    Dim lExitCode As Long                 '�G���[�R�[�h
    'V1.5.0.1 ADD START
    Dim bKansiRet As Boolean                        '�Ď��ՃA�v����������
    Dim bIDURet As Boolean                          'IDU�A�v����������
    Dim bLDURet As Boolean                          'LDU�A�v����������
    
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1 ADD END
    On Error Resume Next
   
'    iChoseAplEndSta = APL_REBOOT           'V1.5.0.1 ADD 'V1.7.0.1 DEL
     
    '�u�A�v���N���E�I����ʁF���u�[�g�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_RBOOT_BUTTOM, 0)
    
    '�u���u�[�g�m�F�v�|�b�v�A�b�v�\��
    iRet = MsgBox("�R���s���[�^�����u�[�g���܂��B��낵���ł����H", vbOKCancel + vbQuestion, "�I���m�F")
    If iRet = vbCancel Then
      '[�L�����Z��]�t�����Ȃ�I��
      Exit Sub
    End If
           
    '��ʂ����b�N����B
    SetEnableFalse
  
    If CheckAppStart(PROC_KANRI) <> 0 Then
       '�A�v���I���v�����Ǘ��ɑ��M����
       uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
       uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
       uMail.udtlHeader.dwProid = RHOSHU_ID
       uMail.udtlHeader.dwSubArea = 0
       'V1.5.0.1 DEL START
       'bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       'If bRtn <> 0 Then
       'V1.5.0.1 DEL END
       'V1.5.0.1 ADD START
       bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
       If bKansiRet <> 0 Then
       'V1.5.0.1 ADD END
          '�u�A�v���N���E�I����ʁF���[�����M���팋�ʁv���O�o��
          Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
 'V1.20.0.1 DEL START
'          'IDU���O�m�F
'          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'            'IDU���O�I���v��CMD���M
'            'V1.5.0.1 DEL START
'            'bRtn = EndIDULog
'            'If bRtn = False Then
'            'V1.5.0.1 DEL END
'            'V1.5.0.1 ADD START
'            bIDURet = EndIDULog
'            If bIDURet = False Then
'               LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")
'               LstStatus.ListIndex = LstStatus.ListCount - 1
'            'V1.5.0.1 ADD END
'               SetEnableTrue
'               Exit Sub
'            End If
'          'V1.5.0.1 ADD START
'          Else
'           bIDURet = True
'          'V1.5.0.1 ADD END
'          End If
'          'LDU���O�m�F
'          If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'            'LDU���O�I���v��CMD���M
'             'V1.5.0.1 DEL START
'             'bRtn = EndLDULog
'             'If bRtn = False Then
'             'V1.5.0.1 DEL END
'             'V1.5.0.1 ADDL START
'             bLDURet = EndLDULog
'             If bLDURet = False Then
'                LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")
'                LstStatus.ListIndex = LstStatus.ListCount - 1
'             'V1.5.0.1 ADD END
'                SetEnableTrue
'                Exit Sub
'             End If
'           'V1.5.0.1 ADD START
'           Else
'            bLDURet = True
'           'V1.5.0.1 ADD END
'          End If
 'V1.20.0.1 DEL END
      Else
          '�u�A�v���N���E�I����ʁF���[�����M�ُ팋�ʁv���O�o��
          lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
          Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
          '�u�A�v���N���E�I����ʁF���u�[�g���������ُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_SHOUT_RBOOT_ERROR, 0)
          SetEnableTrue
          Exit Sub
      End If
'V1.5.0.1 ADD START
       'If bKansiRet = True And bIDURet = True And bLDURet = True Then 'V1.20.0.1 DEL
       If bKansiRet = True Then  'V1.20.0.1 ADD
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplTimer.Enabled = True
          iChoseEnd = APL_REBOOT         'V1.7.0.1 ADD
       Else
           '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'           LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
           LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
           LstStatus.ListIndex = LstStatus.ListCount - 1
           '��ʃ��b�N����
           'SetEnableTrue     'V1.7.0.1 DEL
           'V1.7.0.1 ADD START
           If iChoseAplEndSta = AllApl Then
              '���W�I�t�F�S�A�v���ꊇ
              SetEnableTrue
           ElseIf iChoseAplEndSta = KANSIApl Then
              '���W�I�t�F�Ď���
              SetEnableTrue
              cmdAppEnd.Enabled = False
              cmdAppAllEnd.Enabled = False
           ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
              '���W�I�t�FIDU����LDU
              SetEnableTrue
              cmdAppEnd.Enabled = False
           End If
           'V1.7.0.1 ADD END
       End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'     If CheckAppEndComplete(PROC_KANRI, lExitCode) = 0 _
'        And CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) = 0 _
'        And CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) = 0 Then
'     End If
'
'     '�I������
'     psEndHoshuProc
'     '���u�[�g����
'     dllAPLEndReboot
'V1.5.0.1 DEL END
   Else
    '�I������
    psEndHoshuProc
    '���u�[�g����
    dllAPLEndReboot
  End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.4.0.1) 2010-10-27   REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@������Ή� �s��C���i���W�I�t�j
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
  Dim iCnt As Integer '�J�E���^�[
  
  cmdAppEnd.Enabled = True
  cmdAppAllEnd.Enabled = True
  cmdShoutDown.Enabled = True
  cmdReboot.Enabled = True
  cmdCancel.Enabled = True
  cmdAppStart.Enabled = True
  
  For iCnt = 0 To 3
   Koumoku(iCnt).Enabled = True
   'V2.4.0.1 ADD START
   If pbIDUSts = 1 Then
      'IDU�Ɩ���\��
       Koumoku(2).Enabled = False
   End If
  'V2.4.0.1 ADD END
  Next
  
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
  Dim iCnt As Integer '�J�E���^�[
  
  cmdAppEnd.Enabled = False
  cmdAppAllEnd.Enabled = False
  cmdShoutDown.Enabled = False
  cmdReboot.Enabled = False
  cmdCancel.Enabled = False
  cmdAppStart.Enabled = False
  
  For iCnt = 0 To 3
   Koumoku(iCnt).Enabled = False
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
        AppActivate frmAppConfig.Caption, False
        pfFormActive (frmAppConfig.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

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
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplTimer_Timer()
 
  On Error Resume Next

'  Select Case iChoseAplEndSta 'V1.7.0.1 DEL
   Select Case iChoseEnd 'V1.7.0.1 ADD
    '�S�A�v���ꊇ�F���S�I��
    Case AllApl
         '�S�A�v���ꊇ���S�I������
         ALL_APLEND
    'IDU�A�v���F���S�I��
    Case IDUApl
         'IDU�A�v�����S�I������
         IDU_APLEND
       
    'LDU�A�v���F���S�I��
    Case LDUApl
         'LDU�A�v�����S�I������
         LDU_APLEND

    '�Ď��ՃA�v���F�A�v���I��
    Case APL_END
         '�Ď��ՃA�v���F�A�v���I������
         APL_APLEND
    
    '�Ď��ՁAIDU�ALDU�A�v���F�V���b�g�_�E��
    Case APL_SHOUT_DOWN
         '�Ď��ՁAIDU�ALDU�A�v���F�V���b�g�_�E���I������
         APL_SHOUT_DOWN_END
    
    '�Ď��ՁAIDU�ALDU�A�v���F���u�[�g
    Case APL_REBOOT
         '�Ď��ՁAIDU�ALDU�A�v���F���u�[�g�I������
         APL_REBOOT_END
 End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : ALL_APLEND
'//  �@�\����  : �S�A�v���ꊇ���S�I������
'//  �@�\�T�v  : �S�A�v���ꊇ���S�I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub ALL_APLEND()
 
 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 On Error Resume Next
 
'V1.20.0.1 DEL START
' If CheckAppStart(PROC_KANRI) <> 0 _
'    Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
'    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'V1.20.0.1 DEL END
 If CheckAppStart(PROC_KANRI) <> 0 Then  'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       SetEnableTrue
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
       Exit Sub
    Else
       '�^�C�}���蒼��
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
      Exit Sub
   Else
      '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
      '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'      LstStatus.AddItem ("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")         'EG20 V2.1.0.1 DEL
      LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")      'EG20 V2.1.0.1 ADD
      LstStatus.ListIndex = LstStatus.ListCount - 1
      SetEnableTrue
      iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
   'V1.20.0.1 DEL START
'   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   '�A�v���N���c�[���N��
'   Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
'   '�I������
'    psEndHoshuProc
'   '�ێ�v���Z�X�I��
'    End
   'V1.20.0.1 DEL END
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : IDU_APLEND
'//  �@�\����  : IDU�A�v�����S�I������
'//  �@�\�T�v  : IDU�A�v�����S�I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub IDU_APLEND()
  
  Dim bIDURet As Boolean  'V1.20.0.1 ADD
 
  On Error Resume Next
 
  'If CheckAppStart(PROCESS_IDU_PC) <> 0 Or CheckAppStart(PROCESS_IDU_LOG) <> 0 Then   'V1.20.0.1 DEL
  If CheckAppStart(PROCESS_IDU_PC) <> 0 Then                                           'V1.20.0.1 ADD
     If lngtime >= lngMAX_Time Then
        tmrAplTimer.Enabled = False
        '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1�y��54�zDEL START
'        LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
        LstStatus.AddItem ("�h�c�t�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zADD END
LstStatus.ListIndex = LstStatus.ListCount - 1
        iChoseEnd = NotEnd         'V1.7.0.1 ADD
     Else
        '�^�C�}���蒼��
        tmrAplTimer.Interval = MN_MAIL_INTERVAL
        lngtime = lngtime + MN_MAIL_INTERVAL
        Exit Sub
     End If
  Else
     tmrAplTimer.Enabled = False
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
        Exit Sub
     Else
      '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1�y��54�zDEL START
'        LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
        LstStatus.AddItem ("�h�c�t�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zADD END
        LstStatus.ListIndex = LstStatus.ListCount - 1
        iChoseEnd = NotEnd
     End If
     'V1.20.0.1 ADD END
     'V1.20.0.1 DEL START
     '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'     LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̏I���ɐ������܂����B")
'     LstStatus.ListIndex = LstStatus.ListCount - 1
'     iChoseEnd = NotEnd         'V1.7.0.1 ADD
     'V1.20.0.1 DEL END
  End If
  
  '��ʂ����b�N����������B
  SetEnableTrue
  cmdAppEnd.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : LDU_APLEND
'//  �@�\����  : LDU�A�v�����S�I������
'//  �@�\�T�v  : LDU�A�v�����S�I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub LDU_APLEND()
 
 Dim bLDURet As Boolean   'V1.20.0.1 ADD END

 On Error Resume Next
 
' If CheckAppStart(PROCESS_LDU_PC) <> 0 And CheckAppStart(PROCESS_LDU_LOG) <> 0 Then  'V1.20.0.1 DEL
If CheckAppStart(PROCESS_LDU_PC) <> 0 Then   'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1�y��54�zDEL START
'       LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
       LstStatus.AddItem ("�k�c�t�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zADD END
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
    Else
       '�^�C�}���蒼��
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
    tmrAplTimer.Enabled = False
    'V1.20.0.1 ADD START
    If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
       bLDURet = EndLDULog  'LDU���O�N������LDU���O�ɑ΂��ă��O�I���v��CMD���M
    Else
       bLDURet = True
    End If
    
    If bLDURet = True Then
      lngtime = 0
      lngtime = MN_MAIL_INTERVAL
      tmrLogTimer.Enabled = True
      Exit Sub
   Else
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1�y��54�zDEL START
'       LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
       LstStatus.AddItem ("�k�c�t�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zADD END
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
    'V1.20.0.1 DEL START
'    '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'    LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���ɐ������܂����B")
'    LstStatus.ListIndex = LstStatus.ListCount - 1
'    iChoseEnd = NotEnd         'V1.7.0.1 ADD
    'V1.20.0.1 DEL END
 End If
 
 '��ʂ����b�N����������B
 SetEnableTrue
 cmdAppEnd.Enabled = False
 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : APL_APLEND
'//  �@�\����  : �Ď��ՃA�v���A�A�v���I������
'//  �@�\�T�v  : �Ď��ՃA�v���A�A�v���I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub APL_APLEND()

 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 
 On Error Resume Next

'V1.20.0.1 DEL START
' If CheckAppStart(PROC_KANRI) <> 0 _
'    Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
'    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'V1.20.0.1 DEL END
 If CheckAppStart(PROC_KANRI) <> 0 Then 'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")            'EG20 V2.1.0.1 DEL �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")         'EG20 V2.1.0.1 ADD �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
    Else
       '�^�C�}���蒼��
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
      Exit Sub
   Else
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
   'V1.20.0.1 DEL START
'   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���ɐ������܂����B")
'   LstStatus.ListIndex = LstStatus.ListCount - 1
'   iChoseEnd = NotEnd         'V1.7.0.1 ADD
   'V1.20.0.1 DEL END
 End If
 
 '��ʂ����b�N����������B
 SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : APL_SHOUT_DOWN_END
'//  �@�\����  : �V���b�g�_�E������
'//  �@�\�T�v  : �V���b�g�_�E���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub APL_SHOUT_DOWN_END()
 
 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 
 On Error Resume Next

'V1.20.0.1 ADD START
' If CheckAppStart(PROC_KANRI) <> 0 _
'    Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
'    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'V1.20.0.1 ADD END
 If CheckAppStart(PROC_KANRI) <> 0 Then 'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '��ʃ��b�N����
       'SetEnableTrue     'V1.7.0.1 DEL
       'V1.7.0.1 ADD START
       If iChoseAplEndSta = AllApl Then
          '���W�I�t�F�S�A�v���ꊇ
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          '���W�I�t�F�Ď���
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          '���W�I�t�FIDU����LDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       'V1.7.0.1 ADD END
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
       Exit Sub
    Else
       '�^�C�}���蒼��
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
      '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
      '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'      LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")         'EG20 V2.1.0.1 DEL �yMainte_03_01�z
      LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")      'EG20 V2.1.0.1 ADD �yMainte_03_01�z
      LstStatus.ListIndex = LstStatus.ListCount - 1
      '��ʃ��b�N����
      If iChoseAplEndSta = AllApl Then
         '���W�I�t�F�S�A�v���ꊇ
         SetEnableTrue
      ElseIf iChoseAplEndSta = KANSIApl Then
         '���W�I�t�F�Ď���
         SetEnableTrue
         cmdAppEnd.Enabled = False
         cmdAppAllEnd.Enabled = False
      ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
         '���W�I�t�FIDU����LDU
         SetEnableTrue
         cmdAppEnd.Enabled = False
      End If
      iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
   'V1.20.0.1 DEL START
'   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   '�I������
'   psEndHoshuProc
'   '�V���b�g�_�E������
'   dllAPLEndDwon
   'V1.20.0.1 DEL END
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : APL_REBOOT_END
'//  �@�\����  : ���u�[�g����
'//  �@�\�T�v  : ���u�[�g�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.5.0.1) 2009-05-08   CODED   BY [TCC] S.Terao
'//                �t�F�[�Y�P�s��Ή��@�A�v���N���`�F�b�N�����������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub APL_REBOOT_END()
 
 'V1.20.0.1 ADD START
 Dim bIDURet As Boolean
 Dim bLDURet As Boolean
 'V1.20.0.1 ADD END
 
 On Error Resume Next
 
 'V1.20.0.1 DEL START
 'If CheckAppStart(PROC_KANRI) <> 0 _
 '   Or CheckAppStart(PROCESS_IDU_LOG) <> 0 _
 '   Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
 'V1.20.0.1 DEL END
 If CheckAppStart(PROC_KANRI) <> 0 Then  'V1.20.0.1 ADD
    If lngtime >= lngMAX_Time Then
       tmrAplTimer.Enabled = False
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '��ʃ��b�N����
       'SetEnableTrue     'V1.7.0.1 DEL
       'V1.7.0.1 ADD START
       If iChoseAplEndSta = AllApl Then
          '���W�I�t�F�S�A�v���ꊇ
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          '���W�I�t�F�Ď���
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          '���W�I�t�FIDU����LDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       'V1.7.0.1 ADD END
       iChoseEnd = NotEnd         'V1.7.0.1 ADD
       Exit Sub
    Else
       '�^�C�}���蒼��
       tmrAplTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrAplTimer.Enabled = False
   'V1.20.0.1 ADD START
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
     '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
     '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'     LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")      'EG20 V2.1.0.1 DEL �yMainte_03_01�z
     LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")   'EG20 V2.1.0.1 ADD �yMainte_03_01�z
     LstStatus.ListIndex = LstStatus.ListCount - 1
     '��ʃ��b�N����
     If iChoseAplEndSta = AllApl Then
        '���W�I�t�F�S�A�v���ꊇ
        SetEnableTrue
     ElseIf iChoseAplEndSta = KANSIApl Then
        '���W�I�t�F�Ď���
        SetEnableTrue
        cmdAppEnd.Enabled = False
        cmdAppAllEnd.Enabled = False
     ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
        '���W�I�t�FIDU����LDU
        SetEnableTrue
        cmdAppEnd.Enabled = False
     End If
     iChoseEnd = NotEnd
   End If
   'V1.20.0.1 ADD END
'V1.20.0.1 DEL START
'   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
'   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   '�I������
'   psEndHoshuProc
'   '���u�[�g����
'   dllAPLEndReboot
'V1.20.0.1 DEL END
 End If
End Sub
'V1.5.0.1 ADD END

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrLogTimer_Timer
'//  �@�\����  : ���O�N���`�F�b�N�^�C�}����
'//  �@�\�T�v  : ���O�N���`�F�b�N�^�C�}�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()
    
   On Error Resume Next

   Select Case iChoseEnd
        '�S�A�v���ꊇ�F���S�I��
        Case AllApl
             '�S�A�v���ꊇ���S�I������
             ALL_APLEND_LOG
        'IDU�A�v���F���S�I��
        Case IDUApl
             'IDU�A�v�����S�I������
             IDU_APLEND_LOG
           
        'LDU�A�v���F���S�I��
        Case LDUApl
             'LDU�A�v�����S�I������
             LDU_APLEND_LOG
    
        '�Ď��ՃA�v���F�A�v���I��
        Case APL_END
             '�Ď��ՃA�v���F�A�v���I������
             APL_APLEND_LOG
        
        '�Ď��ՁAIDU�ALDU�A�v���F�V���b�g�_�E��
        Case APL_SHOUT_DOWN
             '�Ď��ՁAIDU�ALDU�A�v���F�V���b�g�_�E���I������
             APL_SHOUT_DOWN_END_LOG
        
        '�Ď��ՁAIDU�ALDU�A�v���F���u�[�g
        Case APL_REBOOT
             '�Ď��ՁAIDU�ALDU�A�v���F���u�[�g�I������
             APL_REBOOT_END_LOG
      End Select
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : ALL_APLEND
'//  �@�\����  : �S�A�v���ꊇ���S�I������
'//  �@�\�T�v  : �S�A�v���ꊇ���S�I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub ALL_APLEND_LOG()
 
 On Error Resume Next
 
 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       '���O�N���`�F�b�N�^�C�}���~����B
       tmrLogTimer.Enabled = False
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAID���p���j�b�g�ALD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       SetEnableTrue
       iChoseEnd = NotEnd
       Exit Sub
    Else
       '�^�C�}���蒼��
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrLogTimer.Enabled = False
   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
   '�A�v���N���c�[���N��
   Call Shell(EXEC_APP_TOOL & EXEC_APP_NAME, vbNormalFocus)
   '�I������
    psEndHoshuProc
   '�ێ�v���Z�X�I��
    End
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : IDU_APLEND
'//  �@�\����  : IDU�A�v�����S�I������
'//  �@�\�T�v  : IDU�A�v�����S�I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Ή��y�c����54�z
'//                 �E�|�b�v�A�b�v�\�����b�Z�[�W�ύX
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub IDU_APLEND_LOG()
 
  On Error Resume Next
 
  If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
     If lngtime >= lngLogMAX_Time Then
        tmrLogTimer.Enabled = False
        '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'        LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̏I���Ɏ��s���܂����B")     'EG20 V2.0.1.1 DEL
        LstStatus.AddItem ("�h�c�t�A�v���P�[�V�����̏I���Ɏ��s���܂����B")              'EG20 V2.0.1.1 ADD
        LstStatus.ListIndex = LstStatus.ListCount - 1
        iChoseEnd = NotEnd
     Else
        '�^�C�}���蒼��
        tmrLogTimer.Interval = MN_MAIL_INTERVAL
        lngtime = lngtime + MN_MAIL_INTERVAL
        Exit Sub
     End If
  Else
     tmrLogTimer.Enabled = False
     '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'     LstStatus.AddItem ("ID���p���j�b�g�A�v���P�[�V�����̏I���ɐ������܂����B")        'EG20 V2.0.1.1 DEL
     LstStatus.AddItem ("�h�c�t�A�v���P�[�V�����̏I���ɐ������܂����B")                 'EG20 V2.0.1.1 ADD
     LstStatus.ListIndex = LstStatus.ListCount - 1
     iChoseEnd = NotEnd
  End If
  
  '��ʂ����b�N����������B
  SetEnableTrue
  cmdAppEnd.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : LDU_APLEND
'//  �@�\����  : LDU�A�v�����S�I������
'//  �@�\�T�v  : LDU�A�v�����S�I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub LDU_APLEND_LOG()
 
 On Error Resume Next
 
 If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'EG20 V2.0.1.1�y��54�zDEL START
'       LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
       LstStatus.AddItem ("�k�c�t�A�v���P�[�V�����̏I���Ɏ��s���܂����B")
'EG20 V2.0.1.1�y��54�zADD END
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
    Else
       '�^�C�}���蒼��
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
    tmrLogTimer.Enabled = False
    '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'EG20 V2.0.1.1�y��54�zDEL START
'    LstStatus.AddItem ("LD���[�e�B���e�B�A�v���P�[�V�����̏I���ɐ������܂����B")
'EG20 V2.0.1.1�y��54�zDEL END
'EG20 V2.0.1.1�y��54�zADD START
    LstStatus.AddItem ("�k�c�t�A�v���P�[�V�����̏I���ɐ������܂����B")
'EG20 V2.0.1.1�y��54�zADD END
    LstStatus.ListIndex = LstStatus.ListCount - 1
    iChoseEnd = NotEnd
 End If
 
 '��ʂ����b�N����������B
 SetEnableTrue
 cmdAppEnd.Enabled = False
 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : APL_APLEND
'//  �@�\����  : �Ď��ՃA�v���A�A�v���I������
'//  �@�\�T�v  : �Ď��ՃA�v���A�A�v���I���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub APL_APLEND_LOG()
 
 On Error Resume Next

 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL  �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD  �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       iChoseEnd = NotEnd
    Else
       '�^�C�}���蒼��
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
       Exit Sub
    End If
 Else
   tmrLogTimer.Enabled = False
   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
'   LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���ɐ������܂����B")        'EG20 V2.1.0.1 DEL  �yMainte_03_01�z
   LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���ɐ������܂����B")     'EG20 V2.1.0.1 ADD  �yMainte_03_01�z
   LstStatus.ListIndex = LstStatus.ListCount - 1
   iChoseEnd = NotEnd
 End If
 
 '��ʂ����b�N����������B
 SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : APL_SHOUT_DOWN_END
'//  �@�\����  : �V���b�g�_�E������
'//  �@�\�T�v  : �V���b�g�_�E���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub APL_SHOUT_DOWN_END_LOG()
 
 On Error Resume Next

 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL  �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD  �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '��ʃ��b�N����
       If iChoseAplEndSta = AllApl Then
          '���W�I�t�F�S�A�v���ꊇ
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          '���W�I�t�F�Ď���
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          '���W�I�t�FIDU����LDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       iChoseEnd = NotEnd
       Exit Sub
    Else
       '�^�C�}���蒼��
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrLogTimer.Enabled = False
   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
   '�I������
   psEndHoshuProc
   '�V���b�g�_�E������
   dllAPLEndDwon
 End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : APL_REBOOT_END
'//  �@�\����  : ���u�[�g����
'//  �@�\�T�v  : ���u�[�g�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.20.0.1) 2010-03-11  CODED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub APL_REBOOT_END_LOG()
 
 On Error Resume Next

 If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
    Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
    If lngtime >= lngLogMAX_Time Then
       tmrLogTimer.Enabled = False
       '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
       '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
'       LstStatus.AddItem ("�Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")        'EG20 V2.1.0.1 DEL  �yMainte_03_01�z
       LstStatus.AddItem ("�����Ď��ՁAIDU�ALDU�A�v���P�[�V�����͏I���Ɏ��s���܂����B")     'EG20 V2.1.0.1 ADD  �yMainte_03_01�z
       LstStatus.ListIndex = LstStatus.ListCount - 1
       '��ʃ��b�N����
       If iChoseAplEndSta = AllApl Then
          '���W�I�t�F�S�A�v���ꊇ
           SetEnableTrue
       ElseIf iChoseAplEndSta = KANSIApl Then
          '���W�I�t�F�Ď���
          SetEnableTrue
          cmdAppEnd.Enabled = False
          cmdAppAllEnd.Enabled = False
       ElseIf iChoseAplEndSta = IDUApl Or iChoseAplEndSta = LDUApl Then
          '���W�I�t�FIDU����LDU
          SetEnableTrue
          cmdAppEnd.Enabled = False
       End If
       iChoseEnd = NotEnd
       Exit Sub
    Else
       '�^�C�}���蒼��
       tmrLogTimer.Interval = MN_MAIL_INTERVAL
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
 Else
   tmrLogTimer.Enabled = False
   '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)
   '�I������
   psEndHoshuProc
   '���u�[�g����
   dllAPLEndReboot
 End If
End Sub
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : psfuncStartupProgressBar
'//  �@�\����  : �v���O���X�o�[�N������
'//  �@�\�T�v  : �v���O���X�o�[�̋N�������s����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��i�v���O���X�o�[�N���Ή��j
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psfuncStartupProgressBar()

    Dim iRet As Integer     ' �߂�l

    On Error Resume Next

    If CheckAppStart(PROCESS_TOOL_PROGRESSBAR) = 0 Then
        ' �v���O���X�o�[�N��
        iRet = CheckAppStartComplete(FILEPATH_PROGRESSTOOL & PROCESS_TOOL_PROGRESSBAR, 1)
        If iRet <> 0 Then
            '�u�A�v���N���E�I����ʁF�v���O���X�o�[�N�������v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_STARTOK_PROGRESSBAR, 0)
        Else
            '�u�A�v���N���E�I����ʁF�v���O���X�o�[�N���ُ�v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_STARTERR_PROGRESSBAR, 0)
        End If
    End If
End Sub
