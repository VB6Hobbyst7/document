VERSION 5.00
Begin VB.Form frmPing 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ʐM�m�F"
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
   Begin VB.Timer tmrMail 
      Left            =   10800
      Top             =   6360
   End
   Begin VB.Frame frmKekka 
      Caption         =   "�o�h�m�f����"
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   8535
      Begin VB.CommandButton cmdZikko 
         Caption         =   "�o�h�m�f���s"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1905
      End
      Begin VB.ListBox LstStatus 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   6015
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
      Top             =   15000
      Width           =   2895
   End
   Begin VB.Frame frmkiki 
      Caption         =   "�@��I��"
      Enabled         =   0   'False
      Height          =   4455
      Left            =   6120
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
      Begin VB.ListBox LstKiki 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         ItemData        =   "�ʐM�m�F���.frx":0000
         Left            =   240
         List            =   "�ʐM�m�F���.frx":0002
         TabIndex        =   37
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   32
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP2 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblIP20 
         Caption         =   "�D"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   36
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP21 
         Caption         =   "�D"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   35
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP22 
         Caption         =   "�D"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   34
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame frmTe 
      Caption         =   "�h�o�A�h���X�����"
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
      Begin VB.CommandButton cmdC 
         Caption         =   "�N���A"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   360
         TabIndex        =   26
         Top             =   960
         Width           =   1875
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   9
         Left            =   4080
         TabIndex        =   25
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   8
         Left            =   3240
         TabIndex        =   24
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   7
         Left            =   2400
         TabIndex        =   23
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   4
         Left            =   2400
         TabIndex        =   22
         Top             =   1800
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   5
         Left            =   3240
         TabIndex        =   21
         Top             =   1800
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   6
         Left            =   4080
         TabIndex        =   20
         Top             =   1800
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   1
         Left            =   2400
         TabIndex        =   19
         Top             =   2640
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   2
         Left            =   3240
         TabIndex        =   18
         Top             =   2640
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   3
         Left            =   4080
         TabIndex        =   17
         Top             =   2640
         Width           =   800
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   800
         Index           =   0
         Left            =   2400
         TabIndex        =   16
         Top             =   3480
         Width           =   800
      End
      Begin VB.CommandButton cmdOct 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3240
         TabIndex        =   15
         Top             =   3480
         Width           =   800
      End
      Begin VB.CommandButton cmdBs 
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
         Height          =   800
         Left            =   4080
         TabIndex        =   14
         Top             =   3480
         Width           =   800
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Index           =   3
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Index           =   2
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Index           =   1
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtIP1 
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Index           =   0
         Left            =   360
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblIP10 
         Caption         =   "�D"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP11 
         Caption         =   "�D"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblIP12 
         Caption         =   "�D"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame frmSentaku 
      Caption         =   "�h�o�A�h���X���͕��@�I��"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   11535
      Begin VB.OptionButton OptTe 
         Caption         =   "�����"
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptKiki 
         Caption         =   "�@��I��"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   $"�ʐM�m�F���.frx":0004
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
      TabIndex        =   3
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�ʐM�m�F"
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
      TabIndex        =   38
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmPing.frm
'//  �p�b�P�[�W���F�ʐM�m�F���
'//
'//  �T�v�F�ʐM�m�F���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//               EG10���A�ʐM�m�F(frmPing.frm)��ʗ��p
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//     REVISIONS :(EG30 31.2.0.1) 2015-07-17   REVISED BY [TCC] T.Nakajima
'//                 ping�̃X�e�[�^�X��\���ł���悤�C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private bIP0 As Boolean
Private bIP1 As Boolean
Private bIP2 As Boolean

Private Const MAXIPKIKIINFO = 96            '�@��\�����ő�

'Private sKikiIP(45) As String              'IP�A�h���X�i�[�G���A   ' EG20 V3.4.0.1�폜
Private sKikiIP(MAXIPKIKIINFO) As String    'IP�A�h���X�i�[�G���A   ' EG20 V3.4.0.1�ǉ�
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l '1.3.0.1 ADD

' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��J�n
' ��ʋ@��ݒ�\��
Private Type TRANSKIKI_INFO
    bStatus As Boolean              ' �ݒ�L���iTRUE:�L��,FALSE:�����j
    sGetInf As String               ' ��ʕ\���p����
    iAreaID As Integer              ' �ΏۊO���@���ʋ@��ʐM��ԃG���AID
    nIniListNo As Integer           ' �O���@�탊�X�g�ԍ�
    nCorner As Integer              ' �R�[�i�ԍ�
End Type
Private gTransKikiInfo(1 To CONECT_KIKI_INI_MAX) As TRANSKIKI_INFO

' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��I��

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �ʐM�m�F���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�A�^�C�}�N��
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
On Error Resume Next
    '�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �ʐM�m�F���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�A�^�C�}��~
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
    '�^�C�}���~����
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �ʐM�m�F���(���[�h��)
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
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    Dim sKeyName As String
    Dim sGateData As String * 128    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim sIP As String * 15
    Dim sCPUReg As String           'LDU_APLROOT���W�X�g���擾�p
    Dim sCPUData As String * 128    '�P�s���t�@�C�����e�擾�p
    
    '�z�u�ݒ�
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '������
    LstStatus.Clear
    LstKiki.Clear
    
    txtIP1(0).Text = ""
    txtIP1(1).Text = ""
    txtIP1(2).Text = ""
    txtIP1(3).Text = ""
    txtIP2(0).Text = ""
    txtIP2(1).Text = ""
    txtIP2(2).Text = ""
    txtIP2(3).Text = ""

    OptTe.Value = True
    frmTe.Enabled = True
    txtIP1(0).Enabled = True
    txtIP1(1).Enabled = True
    txtIP1(2).Enabled = True
    txtIP1(3).Enabled = True
    lblIP10.Enabled = True
    lblIP11.Enabled = True
    lblIP12.Enabled = True
    cmdNum(0).Enabled = True
    cmdNum(1).Enabled = True
    cmdNum(2).Enabled = True
    cmdNum(3).Enabled = True
    cmdNum(4).Enabled = True
    cmdNum(5).Enabled = True
    cmdNum(6).Enabled = True
    cmdNum(7).Enabled = True
    cmdNum(8).Enabled = True
    cmdNum(9).Enabled = True
    cmdOct.Enabled = True
    cmdBs.Enabled = True
    cmdC.Enabled = True
    
    frmkiki.Enabled = False
    txtIP2(0).Enabled = False
    txtIP2(1).Enabled = False
    txtIP2(2).Enabled = False
    txtIP2(3).Enabled = False
    lblIP20.Enabled = False
    lblIP21.Enabled = False
    lblIP22.Enabled = False
    LstKiki.Enabled = False
    
    bIP0 = False
    bIP1 = False
    bIP2 = False
    
' EG20 V3.4.0.1�ǉ��J�n
    '���@���擾
    Call gsGetGateInfo
    ' �R�[�i���̐ݒ菈��
    Call gsGetCornerName
' EG20 V3.4.0.1�ǉ��I��
    
    'V1.8.0.1 ADD START
'    For i = 0 To 45                    ' EG20 V3.4.0.1�폜
    For i = 0 To MAXIPKIKIINFO          ' EG20 V3.4.0.1�ǉ�
     sKikiIP(i) = ""
    Next
    'V1.8.0.1 ADD END
    
    'V1.3.0.1 ADD START
    '���C����M�p�̃��C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
    
    On Error GoTo FileError
   
    '�O���@����(OUTKIKI_LIST.ini)�擾�\��
    OverKikiPing
    '�������D�@(Gate.ini)����莩�����擾�\��
    GatePing
    '�������D�@(Gate.ini)����蔻��ICM���擾���\��
    ICMPing
    
    ' �������\��
    OperatePing
    
    '�u�ʐM�m�F��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_START, 0)
    
FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : OverKikiPing
'//  �@�\����  : �ʐM�m�F���(���[�h��)
'//  �@�\�T�v  : �O���@������擾�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-17   REVISED BY [TCC] C.Terui
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub OverKikiPing()
  Dim sIniFilPath As String              '�ΏۊO���@���pINI�t�@�C��
  Dim iCnt As Integer                    '�J�E���^�[
  Dim sKey As String                     '�L�[��
  Dim sGetInf As String * PING_SIZE      '�擾���(�\������)
  Dim sFilePath As String * PING_SIZE    '�擾���(�ΏۊO���@��INI�p�X)
  Dim sSectionName As String * PING_SIZE '�擾���(�Z�N�V������)
  Dim sKeyName As String * PING_SIZE     '�擾���(�L�[��)
  Dim lSts As Long                       'INI�擾�����߂�l
'  Dim sIP As String * PING_IP_SIZE       '�擾IP�A�h���X           ' EG20 V6.1.0.1�폜
  Dim sIP As String                      ' �擾IP�A�h���X           ' EG20 V6.1.0.1�ǉ�
  Dim iType As Integer                   '�擾���(�@��^�C�v)
  Dim sTargetPath As String              '�ΏۊO���@��INI�t�@�C���p�X
  Dim iAreaID As Integer                 '�擾���(�G���AID)        ' EG20 V3.4.0.1�ǉ�
    
  Dim sGetString As String * 128         ' INI�擾������
  Dim nNullIndex As Integer              ' ���������[�N
    
  On Error Resume Next
 
  '������
  sIP = ""
  
'  For iCnt = 1 To 10                               ' EG20 V3.4.0.1�폜

' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��J�n
  For iCnt = 1 To CONECT_KIKI_INI_MAX
    gTransKikiInfo(iCnt).bStatus = False               ' �ݒ�L���iTRUE:�L��,FALSE:�����j
    gTransKikiInfo(iCnt).sGetInf = ""                  ' ��ʕ\���p����
    gTransKikiInfo(iCnt).iAreaID = 0                   ' �ΏۊO���@���ʋ@��ʐM��ԃG���AID
    gTransKikiInfo(iCnt).nIniListNo = 0                ' �O���@�탊�X�g�ԍ�
    gTransKikiInfo(iCnt).nCorner = 0                   ' �R�[�i�ԍ�

    ' OUTKIKI_LIST.ini�����ʒʐM�G���AID���擾����B
    sKey = ""
    sKey = PROFILE_KEY_KIKIAREA_NAME & Format(iCnt, "00")
    iAreaID = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                sKey, _
                                DEFAILT_Int, _
                                OUTKIKI_LIST_FILE)

' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��I��

    ' OUTKIKI_LIST.ini����\���p�O���@�햼�̂��擾����B
    sGetInf = ""
    sKey = ""
    sKey = PROFILE_KEY_KIKINAME & Format(iCnt, "00")
    lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                  sKey, _
                                  DEFAILT, _
                                  sGetInf, _
                                  Len(sGetInf), _
                                  OUTKIKI_LIST_FILE)
     If lSts = False Then
        '�������Ȃ�
     Else
'       LstKiki.AddItem sGetInf                             ' EG20 V3.4.0.1�폜
        Call psAddKikiCornerName(sGetInf, iAreaID, iCnt)    ' EG20 V3.4.0.1�ǉ�
     End If

    If gTransKikiInfo(iCnt).bStatus = True Then             ' EG20 V3.4.0.1�ǉ�

        sKey = ""
        sFilePath = ""
        ' OUTKIKI_LIST.ini����\���ΏۊO���@��INI�t�@�C���p�X���擾����B
        sKey = PROFILE_KEY_KIKIPATH & Format(iCnt, "00")
        lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                       sKey, _
                                       DEFAILT, _
                                       sFilePath, _
                                       Len(sFilePath), _
                                       OUTKIKI_LIST_FILE)
                                   
        sKey = ""
        ' OUTKIKI_LIST.ini����@��^�C�v(�Ď���orIDUorLDU)���擾����B
        sKey = PROFILE_KEY_TYPE & Format(iCnt, "00")
        iType = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                     sKey, _
                                     DEFAILT_Int, _
                                     OUTKIKI_LIST_FILE)
        
        sTargetPath = ""
        If iType = 1 Then  '�@��^�C�v���Ď��Ղ̏ꍇ
           sTargetPath = PATH_KANSI & sFilePath
        End If
        If iType = 2 Then  '�@��^�C�v��IDU�̏ꍇ
           sTargetPath = PATH_IDU_APP & "\\" & sFilePath
        End If
        If iType = 3 Then  '�@��^�C�v��LDU�̏ꍇ
           sTargetPath = PATH_LDU_APP & "\\" & sFilePath
        End If
        sKey = ""

        ' OUTKIKI_LIST.ini����ΏۊO���@��INI�t�@�C���̃Z�N�V���������擾����B
        sKey = PROFILE_KEY_SECTION_NAME & Format(iCnt, "00")
        lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                       sKey, _
                                       DEFAILT, _
                                       sSectionName, _
                                       Len(sSectionName), _
                                       OUTKIKI_LIST_FILE)

         sKey = ""

        ' OUTKIKI_LIST.ini����ΏۊO���@��INI�t�@�C���̃L�[�����擾����B
        sKey = PROFILE_KEY_KEY_NAME & Format(iCnt, "00")
        lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                       sKey, _
                                       DEFAILT, _
                                       sKeyName, _
                                       Len(sKeyName), _
                                       OUTKIKI_LIST_FILE)
        sKey = ""
        sIP = "" 'V1.8.0.1 ADD
        ' �ΏۊO���@��INI�t�@�C������IP�A�h���X���擾����B
        lSts = GetPrivateProfileString(sSectionName, _
                                       sKeyName, _
                                       DEFAILT, _
                                       sGetString, _
                                       Len(sGetString), _
                                       sTargetPath)
        If lSts > 0 Then                             ' V1.3.0.1 ADD
            LstKiki.AddItem gTransKikiInfo(iCnt).sGetInf    ' EG20 V3.4.0.1�ǉ�
            
            nNullIndex = InStr(sGetString, Chr(0))
            If nNullIndex <> 0 Then
                sIP = Left(sGetString, nNullIndex - 1)
            Else
                sIP = sGetString
            End If
            sKikiIP(LstKiki.ListCount - 1) = Trim(sIP)
        End If                                       ' V1.3.0.1 ADD
    End If                                                  ' EG20 V3.4.0.1�ǉ�
  Next iCnt
  
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : GatePing
'//  �@�\����  : �ʐM�m�F���(���[�h��)
'//  �@�\�T�v  : �������D�@�����擾�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-12  CODED BY  [TCC] H.Sugimoto
'//                 �y�\�����@���P�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub GatePing()
    Dim sKeyName As String
    Dim sGateData As String * PING_GATE_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim sIP As String * PING_IP_SIZE
    Dim nCorner As Integer                      ' �R�[�i�ԍ�    ' EG20 V6.1.0.1�ǉ�

    On Error Resume Next

   '�������D�@���擾
    For i = 1 To MAX_GATE_NO
        sKeyName = "gate" & Format(i, "00")
        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                       sKeyName, _
                                       DEFAILT, sGateData, Len(sGateData), _
                                       PATH_GATE_FILE)
        If iRet = 0 Then
            '�u�ʐM�m�F��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
            Exit Sub
        End If
          If Len(sGateData) <> 0 Then
            '�f�[�^�̎擾
            ReDim sFData(15)
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
            
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�J�n
'            '�@��^�C�v�ɂ���ĕ\�����s���B
'            'E�FEG-R�������D�@���F�\�����Ȃ��B
'            If Trim(sFData(4)) = EGR Then
'                LstKiki.AddItem "EG-R�������D�@" & "#" & i
'                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(5))
'            End If
'            If Trim(sFData(4)) = MISETI Then
'               '�������s��Ȃ��B
'            End If
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�I��
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��J�n
            ' EG20���D�@�ł���Ε\��
            If Trim(sFData(GATE_IDX.IDX_KISHU)) = EG20 Then
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�폜�J�n
'                LstKiki.AddItem "�������D�@" & "#" & i
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�폜�I��
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�ǉ��J�n
                nCorner = CInt(Trim(sFData(GATE_IDX.IDX_RONRI_CORNER)))
                LstKiki.AddItem "�������D�@" & "#" & Trim(sFData(GATE_IDX.IDX_DISP_GOKI)) & _
                                    "(" & Format(nCorner, "00") & ")"
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�ǉ��I��
                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(GATE_IDX.IDX_ADDRESS))
            End If
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��I��
       End If
    Next
    
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�J�n
'      '�������D�@���擾(NEG)
'    For i = 1 To MAX_GATE_NO
'        sKeyName = "gate" & Format(i, "00")
'        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
'                                       sKeyName, _
'                                       DEFAILT, sGateData, Len(sGateData), _
'                                       PATH_GATE_FILE)
'        If iRet = 0 Then
'            '�u�ʐM�m�F��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
'            Exit Sub
'        End If
'
'        If Len(sGateData) <> 0 Then
'            '�f�[�^�̎擾
'            ReDim sFData(15)
'            iFCnt = 1
'
'            For iFLoop = 1 To Len(sGateData)
'                If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
'                    iFLoop2 = iFLoop
'                    Do
'                        iFLoop2 = iFLoop2 + 1
'                        If iFLoop2 > Len(sGateData) Then
'                            sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
'                            iFCnt = iFCnt + 1
'                            If iFCnt >= 16 Then
'                                Exit For
'                            End If
'                            iFLoop = iFLoop2
'                            Exit Do
'                        End If
'
'                        If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
'                            sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
'                            iFCnt = iFCnt + 1
'                            If iFCnt >= 16 Then
'                                Exit For
'                            End If
'                            iFLoop = iFLoop2
'                            Exit Do
'                        End If
'                    Loop
'                End If
'            Next
'
'            '�@��^�C�v�ɂ���ĕ\�����s���B
'            'N�FNEG�������D�@�B���F�\�����Ȃ��B
'            If Trim(sFData(4)) = NEG Then
'                LstKiki.AddItem "NEG�������D�@" & "#" & i
'                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(5))
'            End If
'            If Trim(sFData(4)) = MISETI Then
'               '�������s��Ȃ��B
'            End If
'       End If
'    Next
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�I��

FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : GatePing
'//  �@�\����  : �ʐM�m�F���(���[�h��)
'//  �@�\�T�v  : �������D�@�����擾�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-12  CODED BY  [TCC] H.Sugimoto
'//                 �y�\�����@���P�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub ICMPing()
    Dim sKeyName As String
    Dim sGateData As String * PING_GATE_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim i As Integer
    Dim iRet As Integer
    Dim sIP As String * PING_IP_SIZE
    Dim szIniFilePath As String     ' INI�t�@�C���p�X   ' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ�
    Dim nCorner As Integer                      ' �R�[�i�ԍ�    ' EG20 V6.1.0.1�ǉ�

    On Error Resume Next

   '�������D�@���擾
    For i = 1 To MAX_GATE_NO
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�J�n
'        sKeyName = "gate" & Format(i, "00")
'        iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
'                                       sKeyName, _
'                                       DEFAILT, sGateData, Len(sGateData), _
'                                       PATH_GATE_FILE)
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�I��
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��J�n
        ' IDU��ICM.INI������D�@�����擾
        szIniFilePath = PATH_IDU_APP & IDU_ICM_FILE
        sKeyName = "icm" & Format(i, "00")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ICM, _
                                    sKeyName, _
                                    DEFAILT, sGateData, Len(sGateData), _
                                    szIniFilePath)
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��I��
        If iRet = 0 Then
            '�u�ʐM�m�F��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
            Exit Sub
        End If
        
        If Len(sGateData) <> 0 Then
            '�f�[�^�̎擾
            ReDim sFData(15)
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
                       
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�J�n
'            If Trim(sFData(4)) <> MISETI Then   'V1.8.0.1 ADD
'            '����IC-M�̃A�h���X�`�F�b�N���s���B
'             LstKiki.AddItem "����IC-M���W���[��" & "#" & i
'             sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(14))
'             End If  'V1.8.0.1 ADD
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�폜�I��
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��J�n
            If Trim(sFData(5)) <> MISETI Then   'V1.8.0.1 ADD
                '����IC-M�̃A�h���X�`�F�b�N���s���B
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�폜�J�n
'                LstKiki.AddItem "�h�b�l" & "#" & i
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�폜�I��
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�ǉ��J�n
                nCorner = CInt(Trim(sFData(3)))
                LstKiki.AddItem "�h�b�l" & "#" & Trim(sFData(1)) & _
                                    "(" & Format(nCorner, "00") & ")"
' EG20 V6.1.0.1�y�\�����@���P�Ή��z�ǉ��I��
                sKikiIP(LstKiki.ListCount - 1) = Trim(sFData(7))
             End If  'V1.8.0.1 ADD
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��I��
       End If
    Next
        
FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : OptTe_Click
'//  �@�\����  : ���W�I�t�F����͑I��������
'//  �@�\�T�v  : ��ʂ��X�V����B
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
Private Sub OptTe_Click()
    
    On Error Resume Next
   
    '�u�ʐM�m�F��ʁF����͑I���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_HAND_PING, 0)
 
    frmTe.Enabled = True
    txtIP1(0).Enabled = True
    txtIP1(1).Enabled = True
    txtIP1(2).Enabled = True
    txtIP1(3).Enabled = True
    lblIP10.Enabled = True
    lblIP11.Enabled = True
    lblIP12.Enabled = True
    cmdNum(0).Enabled = True
    cmdNum(1).Enabled = True
    cmdNum(2).Enabled = True
    cmdNum(3).Enabled = True
    cmdNum(4).Enabled = True
    cmdNum(5).Enabled = True
    cmdNum(6).Enabled = True
    cmdNum(7).Enabled = True
    cmdNum(8).Enabled = True
    cmdNum(9).Enabled = True
    cmdOct.Enabled = True
    cmdBs.Enabled = True
    cmdC.Enabled = True
    
    frmkiki.Enabled = False
    txtIP2(0).Enabled = False
    txtIP2(1).Enabled = False
    txtIP2(2).Enabled = False
    txtIP2(3).Enabled = False
    lblIP20.Enabled = False
    lblIP21.Enabled = False
    lblIP22.Enabled = False
    LstKiki.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : OptKiki_Click
'//  �@�\����  : ���W�I�t�F�@��I��I��������
'//  �@�\�T�v  : ��ʂ��X�V����B
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
Private Sub OptKiki_Click()
    
    On Error Resume Next
    
    '�u�ʐM�m�F��ʁF�@��I���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_KIKI_PING, 0)
    
    frmTe.Enabled = False
    txtIP1(0).Enabled = False
    txtIP1(1).Enabled = False
    txtIP1(2).Enabled = False
    txtIP1(3).Enabled = False
    lblIP10.Enabled = False
    lblIP11.Enabled = False
    lblIP12.Enabled = False
    cmdNum(0).Enabled = False
    cmdNum(1).Enabled = False
    cmdNum(2).Enabled = False
    cmdNum(3).Enabled = False
    cmdNum(4).Enabled = False
    cmdNum(5).Enabled = False
    cmdNum(6).Enabled = False
    cmdNum(7).Enabled = False
    cmdNum(8).Enabled = False
    cmdNum(9).Enabled = False
    cmdOct.Enabled = False
    cmdBs.Enabled = False
    cmdC.Enabled = False
    
    frmkiki.Enabled = True
    txtIP2(0).Enabled = True
    txtIP2(1).Enabled = True
    txtIP2(2).Enabled = True
    txtIP2(3).Enabled = True
    lblIP20.Enabled = True
    lblIP21.Enabled = True
    lblIP22.Enabled = True
    LstKiki.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdNum_Click
'//  �@�\����  : �e�����t����������
'//  �@�\�T�v  : �e�L�X�g�{�b�N�X��IP�\��
'//
'//              �^        ����      �Ӗ�
'//  ����      :Integer�@�@Index�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdNum_Click(Index As Integer)

    If Len(txtIP1(0).Text) <> 3 And bIP0 = False Then
        txtIP1(0).Text = txtIP1(0).Text & Trim(str(Index))
        Exit Sub
    End If
    If Len(txtIP1(1).Text) <> 3 And bIP1 = False Then
        txtIP1(1).Text = txtIP1(1).Text & Trim(str(Index))
        Exit Sub
    End If
    If Len(txtIP1(2).Text) <> 3 And bIP2 = False Then
        txtIP1(2).Text = txtIP1(2).Text & Trim(str(Index))
        Exit Sub
    End If
    If Len(txtIP1(3).Text) <> 3 Then
        txtIP1(3).Text = txtIP1(3).Text & Trim(str(Index))
        Exit Sub
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdOct_Click
'//  �@�\����  : �I�N�e�b�h(�u.�v)�t����������
'//  �@�\�T�v  :
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
Private Sub cmdOct_Click()

    If Len(txtIP1(0).Text) <> 3 And bIP0 = False And Len(txtIP1(0)) <> 0 Then
        bIP0 = True
        Exit Sub
    End If
    If Len(txtIP1(1).Text) <> 3 And bIP1 = False And Len(txtIP1(1)) <> 0 Then
        bIP1 = True
        Exit Sub
    End If
    If Len(txtIP1(2).Text) <> 3 And bIP2 = False And Len(txtIP1(2)) <> 0 Then
        bIP2 = True
        Exit Sub
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdBs_Click
'//  �@�\����  : �uBS�v�t����������
'//  �@�\�T�v  :
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
Private Sub cmdBs_Click()

    If Len(txtIP1(3).Text) <> 0 Then
        txtIP1(3).Text = Left(txtIP1(3).Text, Len(txtIP1(3).Text) - 1)
        Exit Sub
    End If
    
    If bIP2 = True Then
        bIP2 = False
    End If

    If Len(txtIP1(2).Text) <> 0 Then
        txtIP1(2).Text = Left(txtIP1(2).Text, Len(txtIP1(2).Text) - 1)
        Exit Sub
    End If
    
    If bIP1 = True Then
        bIP1 = False
    End If

    If Len(txtIP1(1).Text) <> 0 Then
        txtIP1(1).Text = Left(txtIP1(1).Text, Len(txtIP1(1).Text) - 1)
        Exit Sub
    End If
    
    If bIP0 = True Then
        bIP0 = False
    End If

    If Len(txtIP1(0).Text) <> 0 Then
        txtIP1(0).Text = Left(txtIP1(0).Text, Len(txtIP1(0).Text) - 1)
        Exit Sub
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdC_Click
'//  �@�\����  : �u�N���A�v�t����������
'//  �@�\�T�v  : IP�e�L�X�g�{�b�N�X���N���A����B
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
Private Sub cmdC_Click()

    txtIP1(0).Text = ""
    txtIP1(1).Text = ""
    txtIP1(2).Text = ""
    txtIP1(3).Text = ""
    
    bIP0 = False
    bIP1 = False
    bIP2 = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtIP1_KeyPress
'//  �@�\����  : �e�L�X�g�{�b�N�X����͏���
'//  �@�\�T�v  :
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
Private Sub txtIP1_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 46 Then
        If Index <> 3 Then
            txtIP1(Index + 1).SetFocus
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtIP1_Change
'//  �@�\����  : �e�L�X�g�{�b�N�X����͏���
'//  �@�\�T�v  :
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtIP1_Change(Index As Integer)
    If InStr(txtIP1(Index).Text, ".") <> 0 Then
        txtIP1(Index).Text = Replace(txtIP1(Index).Text, ".", "")
        Select Case Index
            Case 0:
                bIP0 = True
            Case 1:
                bIP1 = True
            Case 2:
                bIP2 = True
        End Select
    End If
    If Len(txtIP1(Index).Text) = 3 Then
        If Index <> 3 Then
            txtIP1(Index + 1).SetFocus
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : LstKiki_Click
'//  �@�\����  : �@�탊�X�g�{�b�N�X����������
'//  �@�\�T�v  : �@����̓e�L�X�g�{�b�N�X��IP�A�h���X�\�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub LstKiki_Click()
    Dim sText As String
    Dim i As Integer
    Dim iTop As Integer
    Dim iText As Integer
    
    For i = 0 To 3
      txtIP2(i).Text = " "
    Next
    
    sText = sKikiIP(LstKiki.ListIndex)
    iTop = 1
    iText = 0
    
    'sText���Ȃ��ꍇ�A�e�L�X�g�Ƀu�����N���Z�b�g���A�����I��
    If Len(sText) = 0 Then
        For i = 0 To 3
            txtIP2(i).Text = ""
        Next
        Exit Sub
    End If
        
    'IP�A�h���X���e�L�X�g�ɃZ�b�g
    For i = 1 To Len(sText)
        If Mid(sText, i, 1) = "." Then
            txtIP2(iText).Text = Mid(sText, iTop, i - iTop)
            iTop = i + 1
            iText = iText + 1
        End If
    Next
       
    txtIP2(iText).Text = Right(sText, Len(sText) - iTop + 1)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdZikko_Click
'//  �@�\����  : �e�L�X�g�{�b�N�X����͏���
'//  �@�\�T�v  :
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG30 31.2.0.1) 2015-07-17   REVISED BY [TCC] T.Nakajima
'//                 ping�̃X�e�[�^�X��\���ł���悤�C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZikko_Click()
  Dim host_path As String
  Dim wsaDD As wsaDATA
  Dim i As Integer
  Dim VerReq As Integer
  Dim rc As Long
  Dim HostAddress As Long
  Dim IcmpHandle As Long
  Dim RepryBuffer As ICMP_REPRY_BUFFER  'ICMP������M�o�b�t�@

    On Error GoTo ERR_SPACE
    
    '�u�ʐM�m�F��ʁFPING���s�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_PING_BUTTOM, 0)
    
    LstStatus.Clear
    DoEvents

    Me.Enabled = False
    cmdZikko.Enabled = False
    If OptTe.Value = True Then
        host_path = Trim(txtIP1(0).Text) & "." & _
                    Trim(txtIP1(1).Text) & "." & _
                    Trim(txtIP1(2).Text) & "." & _
                    Trim(txtIP1(3).Text)
        If Len(Trim(txtIP1(0).Text)) = 0 Or _
            Len(Trim(txtIP1(1).Text)) = 0 Or _
            Len(Trim(txtIP1(2).Text)) = 0 Or _
            Len(Trim(txtIP1(3).Text)) = 0 Then
            
            LstStatus.AddItem "Unknown host " & host_path
            Me.Enabled = True
            cmdZikko.Enabled = True
            Exit Sub
        End If
    Else
        host_path = txtIP2(0).Text & "." & _
                    txtIP2(1).Text & "." & _
                    txtIP2(2).Text & "." & _
                    txtIP2(3).Text
        If Len(txtIP2(0).Text) = 0 Or _
            Len(txtIP2(1).Text) = 0 Or _
            Len(txtIP2(2).Text) = 0 Or _
            Len(txtIP2(3).Text) = 0 Then

            LstStatus.AddItem "Unknown host " & host_path
            Me.Enabled = True
            cmdZikko.Enabled = True
            Exit Sub
        End If
    End If
    DoEvents
    
    '�@WinSockAPI�̏�����
    VerReq = MakeInteger(1, 1)                      'WinSock1.1��v��
    rc = WSAStartup(VerReq, wsaDD)
    If rc <> 0 Then
        'Winsock���\�[�X�̊m�ۂɎ��s
        Me.Enabled = True
        cmdZikko.Enabled = True
        Exit Sub
    End If
    
    '�A���M���IP�A�h���X�̎擾
    HostAddress = inet_addr(host_path)              'IP�A�h���X�ɕϊ�(���l�̏ꍇ ex:127.0.0.1)
    
    Call WSACleanup                                 'WinSock�̃N���[�Y

    '�BICMP���g���ăG�R�[�𑗂�
    If HostAddress <> INADDR_NONE Then
        'ICMP����n���h���擾
        IcmpHandle = IcmpCreateFile()
        
        LstStatus.AddItem "Pinging " & host_path & " with 32 bytes of data:"
        DoEvents
        
        '�G�R�[���S�񑗂�
        For i = 1 To 4
            'EG30 V31.2.0.1 DEL START
            'rc = IcmpSendEcho(IcmpHandle, HostAddress, 8, 0, 0, RepryBuffer, Len(RepryBuffer), 300)
            'If rc = 0 Then
            '    LstStatus.AddItem "Request timed out."
            'Else
            '    LstStatus.AddItem "Reply From " & host_path & _
            '                     ": bytes=32 " & _
            '                     "time=" & RepryBuffer.EchoRepry.RoundTripTime & "ms " & _
            '                     "TTL=" & CByte("128")
            'End If
            'EG30 V31.2.0.1 DEL END
            'EG30 V31.2.0.1 ADD START
            call IcmpSendEcho(IcmpHandle, HostAddress, 8, 0, 0, RepryBuffer, Len(RepryBuffer), 300)
            If RepryBuffer.EchoRepry.Status = ICMP_SUCCESS Then
                LstStatus.AddItem "Reply From " & host_path & _
                                 ": bytes=32 " & _
                                 "time=" & RepryBuffer.EchoRepry.RoundTripTime & "ms " & _
                                 "TTL=" & CByte("128")
            Else
                LstStatus.AddItem EvaluatePingResponse(RepryBuffer.EchoRepry.Status)
            End If
            'EG30 V31.2.0.1 ADD END
            DoEvents
        Next
        
        'ICMP����n���h���N���[�Y
        rc = IcmpCloseHandle(IcmpHandle)
    Else
        LstStatus.AddItem "Unknown host " & host_path
    End If
    Me.Enabled = True
    cmdZikko.Enabled = True
ERR_SPACE:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdCancel_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����������
'//  �@�\�T�v  :�@����ʂ���������
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
    
    '�u�ʐM�m�F��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, CONECT_GAMEN_END, 0)
    Unload Me
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
        AppActivate frmPing.Caption, False
        pfFormActive (frmPing.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : psAddKikiCornerName
'//  �@�\����  : ��ʋ@��R�[�i���̒ǉ�����
'//  �@�\�T�v  : ��ʋ@�햼�̂ɑ΂��ăR�[�i���̂�t������K�v������Βǉ�����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String �@ sName     [IN]��ʋ@�햼��
'//  ����      : Integer�@ iAreaID   [IN]��ʋ@��ʐM��ԃG���AID
'//  ����      : Integer�@ nIndex    [IN]��ʋ@��ݒ�\��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V3.4.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psAddKikiCornerName(sName As String, iAreaID As Integer, nIndex As Integer)

    Dim nCorner As Integer                  ' �R�[�i�C���f�b�N�X
    Dim szCornerName As String              ' �R�[�i����
    Dim szResultName As String              ' �o�͖���

    szResultName = ""
    nCorner = 0                                         ' �R�[�i�ݒ�s�v
    gTransKikiInfo(nIndex).nIniListNo = nIndex          ' �O���@�탊�X�g�ԍ�
    gTransKikiInfo(nIndex).iAreaID = iAreaID            ' �Q�ƃG���A
    gTransKikiInfo(nIndex).nCorner = nCorner
    gTransKikiInfo(nIndex).bStatus = True
    ' 1.�ΏۊO���@���ʋ@��ʐM��ԃG���AID���`�F�b�N����
    '   �ڑ��Ώۂ�I�ʂ���B
    Select Case iAreaID
    Case IdKikiComSts.ID_DESYU_COM                                       ' 1:�f�W�ʐM���
        nCorner = 1
    Case IdKikiComSts.ID_DESYU2_COM                                      ' 9:�f�W2�ʐM���
        nCorner = 2
    Case IdKikiComSts.ID_DESYU3_COM                                      ' 10:�f�W3�ʐM���
        nCorner = 3
    Case IdKikiComSts.ID_DESYU4_COM                                      ' 11:�f�W4�ʐM���
        nCorner = 4
    Case IdKikiComSts.ID_DESYU5_COM                                      ' 12:�f�W5�ʐM���
        nCorner = 5
    Case IdKikiComSts.ID_DESYU6_COM                                      ' 13:�f�W6�ʐM���
        nCorner = 6
    Case IdKikiComSts.ID_ENKAKU_COM                                      ' 2:���u�ʐM���
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 1
    Case IdKikiComSts.ID_ENKAKU2_COM                                     ' 21:���u2�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 2
    Case IdKikiComSts.ID_ENKAKU3_COM                                     ' 22:���u3�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 3
    Case IdKikiComSts.ID_ENKAKU4_COM                                     ' 23:���u4�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 4
    Case IdKikiComSts.ID_ENKAKU5_COM                                     ' 24:���u5�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 5
    Case IdKikiComSts.ID_ENKAKU6_COM                                     ' 25:���u6�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).bStatus = False
        nCorner = 6
    Case Else
    End Select

    If nCorner <> 0 Then
        If gblnCornerSet(nCorner - 1) <> True Then
            gTransKikiInfo(nIndex).bStatus = False
        End If
        szCornerName = "(" & Format(nCorner, "00") & ")"
    End If
    szResultName = Left(sName, InStr(sName, Chr(0)) - 1)
    szResultName = szResultName + szCornerName
    gTransKikiInfo(nIndex).sGetInf = szResultName

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : OperatePing
'//  �@�\����  : �����ݒ�쐬
'//  �@�\�T�v  : ���������擾�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V6.1.0.1) 2012-06-08  CODED BY  [TCC] H.Sugimoto
'//                 �y�����o�h�m�f�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub OperatePing()
  
    Dim iLoop As Integer                ' ���[�v
    Dim lSts As Long                    ' INI�擾�����߂�l
    Dim ikousei As Integer              ' �ݒu�\��
    Dim szSection As String             ' �Z�N�V����
    Dim sIP As String                   ' �擾IP�A�h���X
    Dim sGetString As String * 128      ' INI�擾������
    Dim szDispName As String            ' �\����
    Dim nNullIndex As Integer           ' ���������[�N
  
    On Error Resume Next
  
  
    ' �����R�[�i�������[�v
    For iLoop = 1 To 6
        szSection = "KOUSEI" & Format(iLoop, "0") & "_INFO"
        szDispName = "�����" & "(" & Format(iLoop, "00") & ")"
        
        ' �R�[�i�ݒ�L��
        ' OPERATE.INI�t�@�C������u�ڑ��L��(0:�ݒu�Ȃ��A1:�ݒu����j�v���擾����B
        ikousei = GetPrivateProfileInt(szSection, "kousei", _
                                           0, OPERATEINI_FILE)
        
        If ikousei = 1 Then
            ' OPERATE.INI�t�@�C������u�h�o�A�h���X�v���擾����B
            lSts = GetPrivateProfileString(szSection, _
                                           "ip_address", _
                                           "0.0.0.0", _
                                           sGetString, _
                                           Len(sGetString), _
                                           OPERATEINI_FILE)
        
            LstKiki.AddItem szDispName
            
            nNullIndex = InStr(sGetString, Chr(0))
            If nNullIndex <> 0 Then
                sIP = Left(sGetString, nNullIndex - 1)
            Else
                sIP = sGetString
            End If
            sKikiIP(LstKiki.ListCount - 1) = Trim(sIP)
        End If
    Next iLoop
  
End Sub


