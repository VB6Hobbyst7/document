VERSION 5.00
Begin VB.Form frmPassSet 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�p�X���[�h�ݒ�"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ClipControls    =   0   'False
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
   Begin VB.Timer tmrMail 
      Left            =   960
      Top             =   960
   End
   Begin VB.Frame fraPassWord 
      Caption         =   "�p�X���[�h(�Ɩ����[�U)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   5055
      Index           =   2
      Left            =   8040
      TabIndex        =   24
      Top             =   1980
      Width           =   3675
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   29
         Left            =   1920
         TabIndex        =   34
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   28
         Left            =   1920
         TabIndex        =   33
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   27
         Left            =   1920
         TabIndex        =   32
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   26
         Left            =   1920
         TabIndex        =   31
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   25
         Left            =   1920
         TabIndex        =   30
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   180
         TabIndex        =   29
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   180
         TabIndex        =   28
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   180
         TabIndex        =   27
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   180
         TabIndex        =   26
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   180
         TabIndex        =   25
         Top             =   720
         Width           =   1600
      End
   End
   Begin VB.Frame fraPassWord 
      Caption         =   "�p�X���[�h(�������[�U)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5055
      Index           =   1
      Left            =   4080
      TabIndex        =   13
      Top             =   1980
      Width           =   3735
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   1920
         TabIndex        =   23
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   1920
         TabIndex        =   22
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   1920
         TabIndex        =   21
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   1920
         TabIndex        =   20
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   1920
         TabIndex        =   19
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   180
         TabIndex        =   18
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   180
         TabIndex        =   17
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   180
         TabIndex        =   16
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   180
         TabIndex        =   15
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1600
      End
   End
   Begin VB.Frame fraPassWord 
      Caption         =   "�p�X���[�h(��ʃ��[�U)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   1980
      Width           =   3675
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   1920
         TabIndex        =   12
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1920
         TabIndex        =   11
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1920
         TabIndex        =   10
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   1920
         TabIndex        =   9
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   1920
         TabIndex        =   8
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Top             =   4080
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   3240
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   2400
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   1560
         Width           =   1600
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   1600
      End
   End
   Begin VB.CommandButton cmdSettei 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�����e�i���X  ��ʂ֖߂�"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�p�X���[�h�ݒ�"
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
      TabIndex        =   35
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPassSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmPassSet.frm
'//  �p�b�P�[�W���F�p�X���[�h�ݒ���
'//
'//  �T�v�F�p�X���[�h�ݒ���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21  REVISED BY [TCC]  T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή��y�c��54�z
'//                �E�������[�U���̋Ɩ����[�U�p�X���[�h���͕��폜
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private bExchanged As Boolean  '�ύX�f�[�^�L��^�����i��True�^False�j
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �p�X���[�h�ݒ���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}�A�N��
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
    '���C����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �p�X���[�h�ݒ���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}�A��~
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
    '���C����M�p�̃^�C�}�����~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �p�X���[�h�ݒ���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21  REVISED BY [TCC]  T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή��y�c��54�z
'//                �E�������[�U���̋Ɩ����[�U�p�X���[�h���͕��폜
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim intPassFileNo As Integer  '�p�X���[�h�t�@�C���̃t�@�C���ԍ�
    Dim sPassword As String       '�p�X���[�h�t�@�C���̂P�s���̃f�[�^
    Dim iLineSSB As Integer       '���ꃆ�[�U�e�L�X�g�{�b�N�X��INDEX
    Dim iLineTSB As Integer       '�������[�U�e�L�X�g�{�b�N�X��INDEX
    Dim iLineUSR As Integer       '��ʕێ烆�[�U�e�L�X�g�{�b�N�X��INDEX

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

 Select Case pbUserLevel
        '��ʃ��[�U
        Case 0
            fraPassWord(0).Caption = "�p�X���[�h"
            fraPassWord(0).Left = 4080
            fraPassWord(0).Visible = True
            fraPassWord(1).Visible = False
            fraPassWord(2).Visible = False
        '���ꃆ�[�U
        Case 1, 2
'EG20 V2.0.1.1 DEL START
'            fraPassWord(0).Left = 180
'            fraPassWord(1).Left = 4080
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
            fraPassWord(0).Left = 1980
            fraPassWord(1).Left = 5880
'EG20 V2.0.1.1 ADD END
            fraPassWord(2).Left = 8040
            fraPassWord(0).Visible = True
            fraPassWord(1).Visible = True
'EG20 V2.0.1.1 DEL START
'            fraPassWord(2).Visible = True
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
            fraPassWord(2).Visible = False
'EG20 V2.0.1.1 ADD END
        Case Else
    End Select
    
    On Error GoTo FileError
    iLineUSR = 0
    iLineTSB = 10
    iLineSSB = 20
    
    '�ێ���p�X���[�h�t�@�C���̐擪����P�s���Ǎ��݁A�e�L�X�g�{�b�N�X�ɕ\������B
    intPassFileNo = FreeFile        ' ���g�p�̃t�@�C���ԍ����擾����B
    Open PASSWORD_FILE_FULLPASS For Input As #intPassFileNo     ' �p�X���[�h�t�@�C�����J���B
'    Do While Not EOF(1)             ' �t�@�C���̏I�[�܂ŌJ��Ԃ��B             ' EG20 V3.3.0.1�폜
    Do While Not EOF(intPassFileNo)  ' �t�@�C���̏I�[�܂ŌJ��Ԃ��B             ' EG20 V3.3.0.1�ǉ�
        Line Input #intPassFileNo, sPassword  ' �P�s���Ǎ��ށB
        If sPassword <> "" Then               '������̋L�q������B
            If Left(sPassword, 1) = "0" Then        '��ʕێ烆�[�U�p�X���[�h�ł���B
                txtPassWord(iLineUSR) = Mid(sPassword, 3, 8)
                iLineUSR = iLineUSR + 1
            ElseIf Left(sPassword, 1) = "1" Then    '�������[�U�p�X���[�h�ł���B
                txtPassWord(iLineTSB) = Mid(sPassword, 3, 8)
                iLineTSB = iLineTSB + 1
            Else                                    '���ꃆ�[�U�p�X���[�h�ł���B
                txtPassWord(iLineSSB) = Mid(sPassword, 3, 8)
                iLineSSB = iLineSSB + 1
            End If
        End If
    Loop
    Close #intPassFileNo             ' �t�@�C�������B
    bExchanged = False               ' �ύX�f�[�^�����Ƃ��Ă����B

    '���C����M�p�̃��C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '�u�p�X���[�h�ݒ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_START, 0)

    Exit Sub
FileError:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdSettei_Click
'//  �@�\����  : �u�ݒ�v�t����������
'//  �@�\�T�v  : �ݒ肳�ꂽ�p�X���[�h�̍X�V���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
 Private Sub cmdSettei_Click()
  
   Dim iResponse As Integer     '�{�^���R�[�h
   Dim iLine As Integer         '�e�L�X�g�{�b�N�XINDEX
   Dim iLineMax As Integer      '�e�L�X�g�{�b�N�X�̌�
   Dim iLineTSB As Integer      '�������[�U�e�L�X�g�{�b�N�X�̐擪INDEX
   Dim sPassword As String      '�e�L�X�g�{�b�N�X�̕\�����e
   Dim intPassFileNo As Integer '�p�X���[�h�t�@�C���̃t�@�C���ԍ�
   Dim bRet As Boolean          '�֐��߂�l
   Dim lngErrCode As Long       '�G���[�R�[�h
    
    '�u�p�X���[�h�ݒ��ʁF�ݒ�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_SETTEI_BUTTOM, 0)

   '�X�V�m�F���b�Z�[�W��\������B
   iResponse = MsgBox("�\�����̃p�X���[�h���A�o�^���܂��B" _
                       & Chr(vbKeyReturn) & " ��낵���ł����H", _
                       vbYesNo + vbExclamation, _
                       "�p�X���[�h�̍X�V")
   If iResponse = vbYes Then
   ' [�͂�] �{�^����I�������ꍇ
       On Error GoTo FileError

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       
       '�\�����̃p�X���[�h���A�s���߂ĕێ���p�X���[�h�t�@�C���֏����ށB
       intPassFileNo = FreeFile        ' ���g�p�̃t�@�C���ԍ����擾����B
       Open PASSWORD_FILE_FULLPASS For Output As #intPassFileNo
       iLineMax = txtPassWord.UBound '�e�L�X�g�{�b�N�X�̌�
       iLineTSB = (iLineMax + 1) / 2 '�������[�U�e�L�X�g�{�b�N�X�̐擪INDEX
       For iLine = 0 To iLineMax
           sPassword = txtPassWord(iLine)
           If sPassword <> "" Then
               If iLine < 10 Then
                   sPassword = "0," & sPassword '��ʕێ烆�[�U�p
                ElseIf iLine < 20 Then
                    sPassword = "1," & sPassword '�������[�U�p
                Else
                    sPassword = "2," & sPassword '���ꃆ�[�U�p
                End If
                               
                Print #intPassFileNo, sPassword
            End If
        Next
        Close #intPassFileNo
        bExchanged = False     ' �ύX�f�[�^�����ɖ߂��B
        '�u�p�X���[�h�ݒ��ʁF�ݒ�X�V����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_SETTEI_OK, 0)
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    Else
    ' [������] �{�^����I�������ꍇ
        '�������Ȃ��B
        Exit Sub
    End If
    Exit Sub

FileError:   '�p�X���[�h�t�@�C���A�N�Z�X�G���[�������[�`��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    MsgBox "�p�X���[�h�t�@�C���A�N�Z�X�G���[�F" & _
            vbCrLf & Error(Err.Number)
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�u�p�X���[�h�ݒ��ʁF�ݒ�X�V�ُ�v���O�o��
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, PASS_SET_GAMEN_SETTEI_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtPassWord_DblClick
'//  �@�\����  : �e�L�X�g�{�b�N�X�A�_�u���N���b�N������
'//  �@�\�T�v  : �[���e���L�[��ʂ�\�����A�p�X���[�h�ݒ���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@[IN]�_�u���N���b�N���ꂽ�e�L�X�g�{�b�N�X�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtPassWord_DblClick(Index As Integer)
    
    gstrTenKeyData = txtPassWord(Index)  ' ���݂̍s�ʒu�̕\��������n��
    gstrTenKeySize = 8                   '���͉\���������w�肷��B
    ' �[���e���L�[��ʂ�\������B
    frmTenKey.Show 1
    If gstrTenKeyData <> txtPassWord(Index) Then
    '���e���X�V����Ă���΁A
        '�ݒ肳�ꂽ���ŕ\���X�V����
        txtPassWord(Index) = gstrTenKeyData
        bExchanged = True    '�ύX�f�[�^�L��Ƃ���B
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtPassWord_KeyPress
'//  �@�\����  : �e�L�X�g�{�b�N�X�A�L�[���͏���
'//  �@�\�T�v  : �f�[�^�ύX���L�^����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@[IN]
'//  �@�@      : Integer�@KeyAscii [IN]
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtPassWord_KeyPress(Index As Integer, KeyAscii As Integer)
    bExchanged = True '�ύX�f�[�^�L��Ƃ���B
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�����e�i���X��ʂ֖߂�v�t����������
'//  �@�\�T�v  : �ݒ肳�ꂽ�p�X���[�h�̍X�V�L���ƁA����ʂ���������B
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
Private Sub cmdReturn_Click()
      
    Dim iResponse As Integer     '�{�^���R�[�h
   
    On Error Resume Next

    If bExchanged = True Then
    '��ʕ\�����̕ύX���o�^ ����Ă��Ȃ��Ƃ��A�m�F���b�Z�[�W��\������B
     iResponse = MsgBox("��ʕ\�����ɐݒ肳�ꂽ�f�[�^�������܂��B" _
                        & Chr(vbKeyReturn) & "��낵���ł����H", _
                        vbYesNo + vbExclamation, _
                        "�ݒ�f�[�^�̃L�����Z���m�F")
       If iResponse = vbYes Then
         ' [�͂�] �{�^����I�������ꍇ�A
         '�p�X���[�h�ݒ��ʂ����B
         '�u�p�X���[�h�ݒ��ʁF�����v���O�o��
          Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_END, 0)
          Unload Me
       Else
       ' [������] �{�^����I�������ꍇ�A
       '�������Ȃ��B
         '�u�p�X���[�h�ݒ��ʁF�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_END, 0)
         Exit Sub
       End If
    Else
    '����ȊO�́A
        '�p�X���[�h�ݒ��ʂ����B
        '�u�p�X���[�h�ݒ��ʁF�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, PASS_SET_GAMEN_END, 0)
        Unload Me
    End If
    Unload Me
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmPassSet.Caption, False
        pfFormActive (frmPassSet.hwnd)
    End If
End Sub

