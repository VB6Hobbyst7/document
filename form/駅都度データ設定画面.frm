VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEkisettei 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�w�s�x�f�[�^�ݒ�"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
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
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   11400
      Top             =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "  �@����ݒ�    ��ʂ֖߂�"
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
      Left            =   9480
      TabIndex        =   17
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdOut 
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
      Left            =   6600
      TabIndex        =   15
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox txtDummy 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   10000
      Width           =   975
   End
   Begin VB.CommandButton cmdDataHanei 
      Caption         =   "�ݒu�w�ύX"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "�w�s�x�f�[�^�}�� �C���X�g�[��"
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
      Left            =   360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   2
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdPageUp 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPageDown 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   5
      Top             =   6180
      Width           =   1215
   End
   Begin VB.ListBox LstStation 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2225
      Width           =   9135
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�w�s�x�f�[�^�ݒ�"
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
      TabIndex        =   16
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblNo 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "No."
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label lblStation 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�w��"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   1860
      Width           =   6135
   End
   Begin VB.Label lblVer 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   " �o�[�W����"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   1860
      Width           =   2295
   End
   Begin VB.Label lblZen 
      Caption         =   "Z9.Z9.Z9.Z9"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblZenTop 
      Caption         =   "�S�̃o�[�W����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblNow 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   9225
   End
   Begin VB.Label lblNowTop 
      Caption         =   "���݂̐ݒu�w"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmEkisettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �F�w�s�x�f�[�^�ݒ���.frm
'//  �p�b�P�[�W���F�w�s�x�f�[�^�ݒ��ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//                 �f�B�X�N���擾�ʒu�ύX
'//                 ��ʃ��b�N�^��ʃ��b�N���������ǉ�
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ��ʍđO�ʕ\���C��(�s��C��)
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 �P���b�`�����g�p�w�Ή�
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000       '���C���^�C�}�̃C���^�[�o���l

'�p�^�[���ԍ���`
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ύX�J�n
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'Private Const PtnZenVersion = "000000"      '�S�̃o�[�W����
'Private Const PtnEkiName = "000001"         '�w��
'Private Const PtnEkiVersion = "000002"      '�w�o�[�W����
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
' �����ނ�3���ɏC��
Private Const PtnZenVersion = "0000000"      '�S�̃o�[�W����
Private Const PtnEkiName = "0000001"         '�w��
Private Const PtnEkiVersion = "0000002"      '�w�o�[�W����
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ύX�I��
Private gstrFileName        As String                       ' �o�̓t�@�C����    ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �w�s�x�f�[�^�ݒ���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
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
'//  �@�\����  : �w�s�x�f�[�^�ݒ��ʁi�G���R�[�h�j���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �w�s�x�f�[�^�ݒ���(���[�h���F�C�x���g�v���V�[�W��)
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

    Dim bRet As Boolean
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_GAMEN_START, 0)
    
    '----------------------------------------------------
    '��ʏ����l�ݒ�
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��

    '������ʕ\��
    bRet = sDisp
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdUp_Click
'//  �@�\����  : �u���v�t����������
'//  �@�\�T�v  : ���X�g�{�b�N�X�̃C���f�b�N�X�𓮂����B
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
Private Sub cmdUp_Click()
    
    '�G���[���[�`����錾
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex <= 0 Then
            LstStation.ListIndex = 1
            LstStation.ListIndex = 0
        Else
            LstStation.ListIndex = LstStation.ListIndex - 1
        End If
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdPageUp_Click
'//  �@�\����  : �u�����v�t����������
'//  �@�\�T�v  : ���X�g�{�b�N�X�̃C���f�b�N�X�𓮂����B
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
Private Sub cmdPageUp_Click()

    '�G���[���[�`����錾
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex <= 18 Then
            LstStation.ListIndex = 1
            LstStation.ListIndex = 0
        Else
            LstStation.ListIndex = LstStation.ListIndex - 18
        End If
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdPageDown_Click
'//  �@�\����  : �u�����v�t����������
'//  �@�\�T�v  : ���X�g�{�b�N�X�̃C���f�b�N�X�𓮂����B
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
Private Sub cmdPageDown_Click()

    Dim iCnt As Integer
    
    '�G���[���[�`����錾
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex >= LstStation.ListCount - 19 Or LstStation.ListIndex = -1 Then
            LstStation.ListIndex = LstStation.ListCount - 2
            LstStation.ListIndex = LstStation.ListCount - 1
        Else
            iCnt = LstStation.ListIndex
            LstStation.ListIndex = LstStation.ListCount - 1
            LstStation.ListIndex = iCnt + 18
        End If
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdDown_Click
'//  �@�\����  : �u���v�t����������
'//  �@�\�T�v  : ���X�g�{�b�N�X�̃C���f�b�N�X�𓮂����B
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
Private Sub cmdDown_Click()

    '�G���[���[�`����錾
    On Error Resume Next

    If LstStation.ListCount <> 0 Then
        If LstStation.ListIndex < LstStation.ListCount - 1 Then
            If LstStation.ListIndex = -1 Then
                LstStation.ListIndex = LstStation.ListCount - 2
                LstStation.ListIndex = LstStation.ListCount - 1
            Else
                LstStation.ListIndex = LstStation.ListIndex + 1
            End If
        Else
            LstStation.ListIndex = LstStation.ListCount - 2
            LstStation.ListIndex = LstStation.ListCount - 1
        End If
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdInstall_Click
'//  �@�\����  : �u�w�s�x�f�[�^�}�̃C���X�g�[���v�t����������
'//  �@�\�T�v  : �w�s�x�f�[�^�}�̂��A�C���X�g�[�����w����\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//                 �f�B�X�N���擾�ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdInstall_Click()

    Dim strFileName             As String       '�}�̃t�@�C����
    Dim strZenVersion           As String       '�S�̃o�[�W����
    Dim iRet                    As Integer      '���b�Z�[�W�{�b�N�X�߂�l
    Dim lSekuta                 As Long         '�Z�N�^�i�N���X�^����j
    Dim lByte                   As Long         '�o�C�g���i�Z�N�^����j
    Dim lKurasuta               As Long         '�t���[�N���X�^��
    Dim lDrive                  As Long         '�h���C�u�̃N���X�^���i���v�j
    Dim strDrive                As String       '�h���C�u
    Dim bFrmShow                As Boolean
    Dim bRet                    As Boolean
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g  'V1.20.0.1 ADD
    
    '�G���[���[�`����錾
    On Error Resume Next

'V1.12.0.1 ADD START
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_INSTOL, 0)
    
    'V1.20.0.1 DEL START
    ''�f�B�X�N�����擾
''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD
    '
    'If lDrive = 0 Then
    '    strDrive = "d:"
    'Else
''        strDrive = "a:"    'V1.12.0.1 DEL
    '    strDrive = "H:"     'V1.12.0.1 ADD
    'End If
    '
    ''�}�̃t�@�C�����擾
    'strFileName = pfFileSelection(strDrive, "*.csv", "�w�s�x̧�ّI��")
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    '�擾�t�@�C������������
    CommonDialog1.FileName = ""
    '�����f�B���N�g����ݒ�
    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    '�t�H���_�I����ʃf�t�H���g�p�X�P�����݂��邩
        '���݂��邽�߁A�f�t�H���g�p�X�P�iH:�j��ݒ�
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
    Else
        '���݂��Ȃ����߁A�f�t�H���g�p�X�Q�iC:�j��ݒ�
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
    End If
    Set objFso = Nothing
    '�g���q��ݒ�
    CommonDialog1.Filter = "�b�r�u�i�J���}��؂�j(*.csv)|*.csv|"
    '�t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    '�I�������t�@�C�������擾
    strFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
    
    Call ChDrive("D")  'V2.5.0.1 ADD

    '�t�@�C�����݃`�F�b�N
    If strFileName <> "" Then
        
        '�����t�@�C���G���[�̃g���b�v
        On Error GoTo Err_LOG

' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL Start
        '���[�N�t�H���_�ɔ}�̃t�@�C�����R�s�[����
'        Call FileCopy(strFileName, PATH_WORK_EKI_DATA_FILE)
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL End
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
        '�ꎞ�ۑ��t�H���_�Ƀf�[�^���R�s�[���ǎ��p����������
        If pfChangeAttrNormal(strFileName, PATH_HOSHUTMP_EKI_DATA, PATH_WORK_EKI_DATA_FILE) = False Then
            Goto Err_LOG
        End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
        '�S�̃o�[�W�����擾�i���[�N�j
        strZenVersion = sGetZenVersion
    
        iRet = MsgBox("�S�̃o�[�W�����u����" & strZenVersion & "�̓����w�s�x�f�[�^��" & vbCrLf & _
                      "�C���X�g�[�����܂�����낵���ł����H", _
                      vbOKCancel + vbQuestion, _
                      "�}�̃C���X�g�[���m�F")
        
        If iRet = vbOK Then
        
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
            '�v���O���X�o�[��\������
            Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
        
            '�G���[���[�`����錾
            On Error Resume Next
            
            '�����ԍ��i�[�i�}�̃C���X�g�[�����j
            glShoriNo = SHORI_NO.NO_INSTOL
        
            '�}�̃C���X�g�[�����|�b�v�A�b�v��ʕ\��
            Load frmSyorityu
            frmSyorityu.lblLogMessage.Caption = "�}�̃C���X�g�[����"
            frmSyorityu.Caption = "�}�̃C���X�g�[����"
            frmSyorityu.Show vbModal
        
            '�����w�s�x�f�[�^�C���X�g�[������
            If gTgEkiData = False Then GoTo Err_LOG
            
            '������ʕ\��
            bRet = sDisp
            If bRet = False Then
                '�����w�s�x�f�[�^�t�@�C�������ɖ߂�
                Kill EKI_DATA_FILE
                Name EKI_DATA_RENAME_FILE As EKI_DATA_FILE
                bRet = sDisp
                GoTo Err_LOG
            End If
    
            '�����w�s�x�f�[�^�o�b�N�A�b�v�t�@�C���폜
            Kill EKI_DATA_RENAME_FILE
            
            '���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
            '�}�̃C���X�g�[�����ʃ|�b�v�A�b�v��ʕ\��
            iRet = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�̃C���X�g�[������")
            
        End If
                
        '���[�N�t�H���_���̓����w�s�x�f�[�^�t�@�C�����폜
        iRet = DeleteFile(PATH_WORK_EKI_DATA_FILE)
        
    End If
    
'V1.12.0.1 ADD START
    '�S�{�^���������Ƃ���B
    Call SetEnableTrue
'V1.12.0.1 ADD END
    
    Exit Sub

Err_LOG:

    '���[�N�t�H���_���̓����w�s�x�f�[�^�t�@�C�����폜
    iRet = DeleteFile(PATH_WORK_EKI_DATA_FILE)

' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
    '�ꎞ�ۑ��t�H���_���폜����
    psDeleteFolder PATH_HOSHUTMP
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End

    '���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_NG, 0)

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��

    '�}�̃C���X�g�[�����ʃ|�b�v�A�b�v��ʕ\��
    'iRet = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbExclamation, "�}�̃C���X�g�[������")  'V1.8.0.1 DEL
     iRet = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̃C���X�g�[������") 'V1.8.0.1 ADD

'V1.12.0.1 ADD START
    '�S�{�^���������Ƃ���B
    Call SetEnableTrue
'V1.12.0.1 ADD END

    Set objFso = Nothing    'V1.20.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdDataHanei_Click
'//  �@�\����  : �u�ݒu�w�f�[�^���f�v�t����������
'//  �@�\�T�v  : �w��w����INI�t�@�C���ɔ��f����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 �P���b�`�����g�p�w�Ή�
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdDataHanei_Click()

    Dim iRet As Integer         '���b�Z�[�W�{�b�N�X�߂�l
    Dim bRet As Boolean         '�֐��߂�l
    Dim lErrCode As Long        '�G���[�R�[�h
    Dim bSysChange As Boolean   '�V�X�e���ݒ菈���߂�l�@'V1.8.0.1 ADD
    Dim bInstolType As Boolean  '�����w�^�C�v�s�x�f�[�^�C���X�g�[�������ς݃t���O
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim lResult             As Long             ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Dim iResponse           As Integer          ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�

    '������
    bInstolType = False  '�����w�^�C�v�s�x�f�[�^�C���X�g�[�������ς݃t���O

    '�G���[���[�`����錾
    On Error Resume Next
       
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_UPDATE, 0)
    
    bRet = False
    
    iRet = MsgBox("�ݒu�w����ύX���܂��B��낵���ł����H" & vbCrLf & "���f�͍ċN����ɂȂ�܂��B", _
                  vbOKCancel + vbInformation, _
                  "�����w�s�x�f�[�^���f")
             
    If iRet = vbCancel Then
        Call SetEnableTrue
        Set objFso = Nothing  'V2.1.0.1 ADD
        Exit Sub
    End If

' EG20 V3.3.0.1 �ǉ��J�n
    ' ���X�g�ɂP�����f�[�^���Ȃ��ꍇ�ُ͈�I��
    If LstStation.ListCount = 0 Then
        GoTo ErrorHandler
    End If
' EG20 V3.3.0.1 �ǉ��I��

    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
        
    ' /////////////////////////////////////////////////////////
    ' // �����w�s�x�f�[�^����I�������w�f�[�^��؂肾��
    ' // [IN]�����w�s�x�f�[�^�t�@�C����
    ' // [IN]�w�f�[�^�t�@�C���̕ۑ��t�@�C����
    ' // [IN]�I�����ꂽ�����w�s�x�f�[�^�t�@�C���̃C���f�b�N�X
    ' // [out] �G���[�R�[�h
    bRet = dllCreateFile_ChooseEkiData(EKI_DATA_FILE, EKI_DATA_CHOOSE_FILE, LstStation.ListIndex, lErrCode)

    '��������
    If bRet = False Then
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, lErrCode)
        GoTo ErrorHandler               ' EG20 V3.3.0.1�ǉ�
    Else
        gstrFileName = EKI_DATA_CHOOSE_FILE
        ' //////////////////////////////////////////////
        ' // �����v���O��������
        ' //////////////////////////////////////////////
        lResult = pubfuncTakuProgramData(2, gstrFileName)
        If lResult = 0 Then
            GoTo ErrorHandler           ' EG20 V3.3.0.1�ǉ�
' EG20 V3.3.0.1 �폜�J�n
'           '�v���O���X�o�[����������
'           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'           ' �ُ�I��
'           iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�����w�s�x�f�[�^���f")
'           Set objFso = Nothing  'V2.1.0.1 ADD
'           Call SetEnableTrue
'           Exit Sub
' EG20 V3.3.0.1 �폜�I��
        ElseIf lResult = 1 Then
           ' ���[�����M��
           ' ���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
           Set objFso = Nothing  'V2.1.0.1 ADD
            
           Exit Sub
        End If
    
        ' //////////////////////////////////////////////
        ' // �����Ď��Ք񓮍쒆�̂��߃��[��������҂�����
        ' // �����X�V
        ' //////////////////////////////////////////////
        bRet = pfuncInstallEkiSettei
    
    End If

    Exit Sub                            ' EG20 V3.3.0.1 �ǉ�
' EG20 V3.3.0.1 �ǉ��J�n�i�G���[�������܂Ƃ߂�j
ErrorHandler:
    Set objFso = Nothing
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    MsgBox "�����w�s�x�f�[�^���f�������ُ�I�����܂����B", _
            vbOKOnly + vbCritical, _
             "�����w�s�x�f�[�^���f����"
    Call SetEnableTrue
    Exit Sub
' EG20 V3.3.0.1 �ǉ��I���i�G���[�������܂Ƃ߂�j
End Sub

' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�폜�J�n�i�S�̌������j
'Private Sub cmdDataHanei_Click()
'
'    Dim iRet As Integer         '���b�Z�[�W�{�b�N�X�߂�l
'    Dim bRet As Boolean         '�֐��߂�l
'    Dim lErrCode As Long        '�G���[�R�[�h
'    Dim bSysChange As Boolean   '�V�X�e���ݒ菈���߂�l�@'V1.8.0.1 ADD
''V2.1.0.1 ADD START
'    Dim bInstolType As Boolean  '�����w�^�C�v�s�x�f�[�^�C���X�g�[�������ς݃t���O
'    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
'    Dim lResult             As Long             ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'    Dim iResponse           As Integer          ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'
'    '������
'    bInstolType = False  '�����w�^�C�v�s�x�f�[�^�C���X�g�[�������ς݃t���O
''V2.1.0.1 ADD END
'
'    '�G���[���[�`����錾
'    On Error Resume Next
'
''V1.12.0.1 ADD START
'    '�S�{�^���������s�Ƃ���B
'    Call SetEnableFalse
''V1.12.0.1 ADD END
'
'    '��ʑ��샍�O�o��
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_UPDATE, 0)
'
'    bRet = False
'
'    iRet = MsgBox("�ݒu�w����ύX���܂��B��낵���ł����H" & vbCrLf & "���f�͍ċN����ɂȂ�܂��B", _
'                  vbOKCancel + vbInformation, _
'                  "�����w�s�x�f�[�^���f")
'
''    If iRet = vbCancel Then Exit Sub   'V1.12.0.1 DEL
''V1.12.0.1 ADD START
'    If iRet = vbCancel Then
'        Call SetEnableTrue
'        Set objFso = Nothing  'V2.1.0.1 ADD
'        Exit Sub
'    End If
''V1.12.0.1 ADD END
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'    '�v���O���X�o�[��\������
'    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'    '�����w�s�x�f�[�^�C���X�g�[������
'    bRet = dllInstolEkiData(EKI_DATA_FILE, EKI_NAME_FILE, EKI_SETTI_FILE, LstStation.ListIndex, lErrCode)
'
'    '��������
'    If bRet = False Then
'        '�ُ탍�O�o��
''       Call pfOutPutErrLog(lErrCode)    'V2.1.0.1 DEL
'        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, lErrCode)    'V2.1.0.1 ADD
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'        '�v���O���X�o�[����������
'        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'        '�ُ�I��
'        'V1.8.0.1 DEL START
'        'MsgBox "�ُ�I�����܂����B", _
'        '        vbOKOnly + vbExclamation, _
'        '         "�����w�s�x�f�[�^���f����"
'        'V1.8.0.1 DEL END
''V2.1.0.1 DEL START
''       'V1.8.0.1 ADD START
''       MsgBox "�ُ�I�����܂����B", _
''               vbOKOnly + vbCritical, _
''                "�����w�s�x�f�[�^���f����"
''       'V1.8.0.1 ADD END
''V2.1.0.1 DEL END
''V2.1.0.1 ADD START
'        MsgBox "�����w�s�x�f�[�^���f�������ُ�I�����܂����B", _
'                vbOKOnly + vbCritical, _
'                 "�����w�s�x�f�[�^���f����"
''V2.1.0.1 ADD END
'    Else
''V2.1.0.1 ADD START
'        '�����w�^�C�v�s�x�f�[�^�t�@�C�������݂���H
'        If (objFso.FileExists(EKI_TYPE_DATA_FILE) = True) Then
'
'            '�����w�^�C�v�s�x�f�[�^�C���X�g�[�������֐�
'            bRet = dllInstolEkiTypeData(EKI_TYPE_DATA_FILE, lErrCode)
'            '�����w�^�C�v�s�x�f�[�^�C���X�g�[�������ς݃t���O�𗧂Ă�
'            bInstolType = True
'
'        End If
''V2.1.0.1 ADD END
'        '----------------------------------------------------
'        '���݂̐ݒu�w���x���X�V
'        '----------------------------------------------------
'        Call sDispNowEkiLabel
'
'        'V1.8.0.1 START ADD
'        '----------------------------------------------------
'        '�R���s���[�^���A�l�b�g���[�N�ύX����
'        '----------------------------------------------------
'        'Call pfNetWorkChng(Me)
'        bSysChange = pfNetWorkChng(Me)
'        'V1.8.0.1 START END
'
''V2.1.0.1 DEL START
''       '���O�o��
''       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
''V2.1.0.1 DEL END
'
'        If bSysChange = True Then 'V1.8.0.1 ADD
'        'V2.1.0.1 DEL START
'        ''����I��
'        'MsgBox "����I�����܂����B", _
'        '        vbOKOnly + vbInformation, _
'        '         "�����w�s�x�f�[�^���f����"
'        ''V2.1.0.1 DEL END
'
'' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'            ' //////////////////////////////////////////////
'            ' // �����v���O��������
'            ' //////////////////////////////////////////////
'             lResult = pubfuncTakuProgramData(2)
'             If lResult = 0 Then
'                '�v���O���X�o�[����������
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'                ' �ُ�I��
'                iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�ݒ蔽�f����")
'                Set objFso = Nothing  'V2.1.0.1 ADD
'                Call SetEnableTrue
'                Exit Sub
'             ElseIf lResult = 1 Then
'                ' ���[�����M��
'                ' ���O�o��
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'                Set objFso = Nothing  'V2.1.0.1 ADD
'
'                Exit Sub
'             End If
'' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'
'            'V2.1.0.1 ADD START
'            '���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET, 0)
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'            '�v���O���X�o�[����������
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'            '����I��
'            MsgBox "�����w�s�x�f�[�^���f����������I�����܂����B", _
'                    vbOKOnly + vbInformation, _
'                     "�����w�s�x�f�[�^���f����"
'            'V2.1.0.1 ADD END
''V1.12.0.1 ADD START
'        Else
'            'V2.1.0.1 DEL START
'            ''�ُ�I��
'            'MsgBox "�ُ�I�����܂����B", _
'            '        vbOKOnly + vbCritical, _
'            '         "�����w�s�x�f�[�^���f����"
'            'V2.1.0.1 DEL END
'            'V2.1.0.1 ADD START
'            '�ُ탍�O�o��
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, 0)
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'            '�v���O���X�o�[����������
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'            '�ُ�I��
'            MsgBox "�����w�s�x�f�[�^���f�������ُ�I�����܂����B", _
'                    vbOKOnly + vbCritical, _
'                     "�����w�s�x�f�[�^���f����"
'            'V2.1.0.1 ADD END
''V1.12.0.1 ADD END
'       End If                    'V1.8.0.1 ADD
''V2.1.0.1 ADD START
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'        '�v���O���X�o�[����������
'        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'        '�����w�^�C�v�s�x�f�[�^�C���X�g�[�����������s�������H
'        If (True = bInstolType) Then
'            '�����w�^�C�v�s�x�f�[�^�C���X�g�[����������
'            If ((False = bRet) And (ERR_EKITYPE_NO_TYPE = lErrCode)) Then
'                '�Y���w�^�C�v�s�x�f�[�^�Ȃ��I��
'                '���O�o��
'                Call sLogTraceReq(LTYP_WARNING, L3AN_ETC, EKITUDODATASET_NO_EKITYPE_DATA, 0)
'                MsgBox "�Y������w�^�C�v�s�x�f�[�^�����݂��܂���ł����B", _
'                        vbOKOnly + vbExclamation, _
'                        "�����w�^�C�v�s�x�f�[�^���f����"
'            ElseIf (False = bRet) Then
'                '�ُ�I��
'                '�ُ탍�O�o��
'                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKITYPE_TUDO_SET_NG, lErrCode)
'                MsgBox "�����w�^�C�v�s�x�f�[�^���f�������ُ�I�����܂����B", _
'                        vbOKOnly + vbCritical, _
'                        "�����w�^�C�v�s�x�f�[�^���f����"
'            Else
'                '����I��
'                '���O�o��
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_EKITYPE_TUDO_SET, 0)
'                MsgBox "�����w�^�C�v�s�x�f�[�^���f����������I�����܂����B", _
'                        vbOKOnly + vbInformation, _
'                        "�����w�^�C�v�s�x�f�[�^���f����"
'            End If
'        Else
'            '�����w�^�C�v�s�x�f�[�^�C���X�g�[�����������s
'            '���O�o��
'            Call sLogTraceReq(LTYP_WARNING, L3AN_FILE, EKITUDODATASET_NO_EKITYPE_FILE, 0)
'        End If
''V2.1.0.1 ADD END
'    End If
'
'    Set objFso = Nothing  'V2.1.0.1 ADD
''V1.12.0.1 ADD START
'    '�S�{�^���������Ƃ���B
'    Call SetEnableTrue
''V1.12.0.1 ADD END
'
'End Sub
' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�폜�I���i�S�̌������j

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdOut_Click
'//  �@�\����  : �u�}�̎��O���v�t����������
'//  �@�\�T�v  : USB���O���������s��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOut_Click()

    '�G���[���[�`����錾
    On Error Resume Next
    
'V1.12.0.1 ADD START
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
    
    '�}�̎�O����
    Call pfRemove(Me)

'V1.12.0.1 ADD START
    '�S�{�^���������Ƃ���B
    Call SetEnableTrue
'V1.12.0.1 ADD END
        
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_GAMEN_END, 0)
    
    '����ʏ���
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////

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
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ��ʍđO�ʕ\���C��(�s��C��)
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  '���[����M�G���A
    Dim lngLength As Long            '��M���[���o�C�g�T�C�Y
    Dim intStatus As Integer         '��M���[���`�F�b�N����
    Dim iResponse As Integer
    
    On Error Resume Next
    
    '���[������M����B
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
    '��M���[��������΁A���[���h�c���̏���������B
        Select Case udtReadMail.udtlHeader.dwId        '���[���h�c
            Case ML_ID_PROEND_ORD
                '�u�v���Z�X�I���w���v����M�����ꍇ�A
                '�u�v���Z�X�I���w����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
                '�v���Z�X�̏I���������s��
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '�u�ێ��ʃA�N�e�B�u�\���v����M�����ꍇ
                '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '�\������ʁi�ێ�f�[�^���W��ʁj���A�N�e�B�u�\������B
'                AppActivate frmInputMstData.Caption, False ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
                AppActivate frmEkisettei.Caption, False     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
                pfFormActive (frmEkisettei.hwnd)            ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '�u�ێ瑀���v���O�������M�v���v����M�����ꍇ
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
                    '�v���O���X�o�[����������
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                    MsgBox "�����w�s�x�f�[�^���f�������ُ�I�����܂����B", _
                           vbOKOnly + vbCritical, _
                           "�����w�s�x�f�[�^���f����"
                    Call SetEnableTrue
                Else
                    Call pfuncInstallEkiSettei
                End If
            Case Else
                 '���̑��̃��[������M�����ꍇ
                 '�u���[��ID�s���v���O�o��
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDisp
'//  �@�\����  : ��ʍĕ`�揈��
'//  �@�\�T�v  : ��ʂ��ĕ`�悷��
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
Private Function sDisp() As Boolean

    '�G���[���[�`����錾
    On Error Resume Next
    
    Dim bRet                 As Boolean         '�֐��߂�l
    
    sDisp = False
    
    '----------------------------------------------------
    '�t�����l�ݒ�
    '----------------------------------------------------
    cmdUp.Enabled = False                       '���t
    cmdPageUp.Enabled = False                   '�����t
    cmdPageDown.Enabled = False                 '���t
    cmdDown.Enabled = False                     '�����t
    cmdInstall.Enabled = True                   '�w�s�x�f�[�^�}�� �C���X�g�[���t
    cmdDataHanei.Enabled = False                '�ݒu�w�f�[�^���f�t
    cmdOut.Enabled = True                       '�}�̎�O�t
    
    '----------------------------------------------------
    '�����l�ݒ�
    '----------------------------------------------------
    lblNow.Caption = ""
    lblZen.Caption = ""
    LstStation.Clear

    '----------------------------------------------------
    '���݂̐ݒu�w���x���X�V
    '----------------------------------------------------
    Call sDispNowEkiLabel
    
    '----------------------------------------------------
    '�w���X�V
    '----------------------------------------------------
    bRet = sDispEkiData
    
    '�t�����ݒ�
    If bRet = True Then
        cmdDataHanei.Enabled = True             '�ݒu�w�f�[�^���f�t
        cmdUp.Enabled = True                    '���t
        cmdPageUp.Enabled = True                '�����t
        cmdPageDown.Enabled = True              '���t
        cmdDown.Enabled = True                  '�����t
        
        '�w���R���{�{�b�N�X�̃C���f�b�N�X�ݒ�
        LstStation.ListIndex = 0
        
        sDisp = True
    
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispNowEkiLabel
'//  �@�\����  : ���݂̐ݒu�w���x���X�V����
'//  �@�\�T�v  : ���݂̐ݒu�w���x���ɉw���A�w�o�[�W������ݒ肷��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 �P���b�`�����g�p�w�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispNowEkiLabel()

    Dim strFileName          As String          '�t�@�C����
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    '�t�@�C�������݂����ꍇ
    If strFileName <> "" Then
    
        '�w�o�[�W�����擾
'       lblNow.Caption = pfGetEkiNameInfo               'V2.1.0.1 DEL
        lblNow.Caption = pfGetEkiNameInfo(SetEkiVer)    'V2.1.0.1 ADD
    
    Else
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
    
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispEkiData
'//  �@�\����  : �w���X�V����
'//  �@�\�T�v  : �S�̃o�[�W�������x�����o�[�W������ݒ肵�A
'//              �w���R���{�{�b�N�X�ɉw����ݒ肷��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@ TURE      ����I��
'//                        FALSE     �ُ�I���i�R���{�{�b�N�X�f�[�^�f�[�^�����݂��Ȃ��ꍇ�j
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sDispEkiData() As Boolean

    Dim LOG_Event            As String          '���O�̃C�x���g��
    Dim LOG_FukaData         As String          '�t���f�[�^��
    Dim ECOD3                As Integer         '������
    
    Dim intLoopCount         As Integer         '���[�v�J�E���^
    Dim intFileNumber        As Integer
    Dim strFileName          As String          '�t�@�C����
    
    Dim bRet                 As Boolean         '�֐��߂�l
    Dim lErrCode             As Long            '�G���[�R�[�h
    Dim strData              As String          '�t�@�C���Ǎ��f�[�^
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    sDispEkiData = False

    '----------------------------------------------------
    '�����w�s�x�f�[�^�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_DATA_FILE)
    
    '�t�@�C�������݂����ꍇ
    If strFileName <> "" Then
    
        ' �����w�s�x�f�[�^�w���t�@�C���쐬
        bRet = dllCreateEkiNameFile(EKI_DATA_FILE, EKI_NAME_FILE, lErrCode)
        If bRet = False Then
            '�����w�s�x�f�[�^�w���t�@�C���폜
            Kill EKI_NAME_FILE
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)
            Exit Function
        End If
        
        '�w���t�@�C������
        strFileName = Dir(EKI_NAME_FILE)
        
        '�t�@�C�������݂����ꍇ
        If strFileName <> "" Then
        
            '�����t�@�C���G���[�̃g���b�v
            On Error GoTo Err_LOG
        
            '���g�p�̃t�@�C���ԍ��擾
            intFileNumber = FreeFile
            
            '���݉w�ݒ�t�@�C�����I�[�v������B
            Open EKI_NAME_FILE For Input As #intFileNumber
            
            intLoopCount = 0
            Do While Not EOF(intFileNumber)
                '�P �s�ǂݍ���
                Input #intFileNumber, strData
                
                '�擪��s�ڂ͑S�̃o�[�W�������擾���A���x�����X�V����
                If intLoopCount = 0 Then
                  '  lblZen.Caption = "V" & strData 'V1.8.0.1 DEL
                     lblZen.Caption = strData       'V1.8.0.1 ADD
                    intLoopCount = 1
                Else
                    '���X�g�{�b�N�X�ɉw������ǉ�
                    LstStation.AddItem strData
                End If
            Loop
            
            '�t�@�C�����N���[�Y����B
            Close #intFileNumber
        
        Else
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKINAME, 0)
            
            '�����w�s�x�f�[�^�w���t�@�C���Ȃ�
            sDispEkiData = False
        End If
        
    Else
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_TG_EKITUDO, 0)
        
        '�����w�s�x�f�[�^�t�@�C���Ȃ�
        sDispEkiData = False
    End If
    
    sDispEkiData = True
    
    Exit Function
    
'�G���[����
Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfFileSelection
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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sGetZenVersion
'//  �@�\����  : �S�̃o�[�W�����擾
'//  �@�\�T�v  : ���[�N�t�H���_���̓����w�s�x�f�[�^�t�@�C������S�̃o�[�W�������擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sGetZenVersion() As String

    Dim LOG_Event            As String          '���O�̃C�x���g��
    Dim LOG_FukaData         As String          '�t���f�[�^��
    Dim ECOD3                As Integer         '������
    
    Dim intFileNumber        As Integer
    Dim strFileName          As String          '�t�@�C����
    
    Dim intBunrui_Dai        As Integer         '�啪��
    Dim intBunrui_Tyu        As Integer         '������
    Dim intBunrui_Sho        As Integer         '������
    Dim strData              As String          '�ݒ�l
    
    Dim strPtnNo             As String          '�p�^�[���ԍ�
    Dim strZenVersion        As String          '�S�̃o�[�W����
    
    Dim iGetDataCount        As Integer         '�f�[�^�擾�J�E���^
    Dim intBunrui_Corner     As Integer         '������
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    sGetZenVersion = ""
    strZenVersion = ""
    iGetDataCount = 0

    '----------------------------------------------------
    '�����w�s�x�f�[�^�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(PATH_WORK_EKI_DATA_FILE)

    '�t�@�C�������݂����ꍇ
    If strFileName <> "" Then
    
        '���g�p�̃t�@�C���ԍ��擾
        intFileNumber = FreeFile
    
        '�����t�@�C���G���[�̃g���b�v
        On Error GoTo Err_LOG
        
        '���݉w�ݒ�t�@�C�����I�[�v������B
        Open PATH_WORK_EKI_DATA_FILE For Input As #intFileNumber
    
        Do While Not EOF(intFileNumber)
            '�P �s�Âϐ��ǂݍ���
'            Input #intFileNumber, intBunrui_Dai, intBunrui_Tyu, intBunrui_Sho, strData                     ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
            Input #intFileNumber, intBunrui_Dai, intBunrui_Tyu, intBunrui_Sho, intBunrui_Corner, strData    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
            '�p�^�[���ԍ��擾
'            strPtnNo = Format(intBunrui_Dai, "00") & Format(intBunrui_Tyu, "00") & Format(intBunrui_Sho, "00") ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
            strPtnNo = Format(intBunrui_Dai, "00") & Format(intBunrui_Tyu, "00") & Format(intBunrui_Sho, "000") ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
            
            Select Case strPtnNo
                
                '�S�̃o�[�W�����擾
                Case PtnZenVersion
                    strZenVersion = strData
                    iGetDataCount = iGetDataCount + 1
                
                Case Else
                    '�����Ȃ�
            End Select
            
            '�S�̃o�[�W�������擾�����烋�[�v�𔲂���
            If iGetDataCount > 0 Then Exit Do

        Loop
        
        '�t�@�C�����N���[�Y����B
        Close #intFileNumber
        
        '�߂�l�ݒ�i�S�̃o�[�W�����j
        sGetZenVersion = strZenVersion
    
    Else
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_TG_EKITUDO, 0)
    End If
    
    Exit Function
    
'�G���[����
Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
    
End Function

'V1.12.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
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
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������s�Ƃ���B
    cmdInstall.Enabled = False
    cmdDataHanei.Enabled = False
    cmdOut.Enabled = False
    cmdCancel.Enabled = False
    cmdUp.Enabled = False
    cmdPageUp.Enabled = False
    cmdPageDown.Enabled = False
    cmdDown.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
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
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������Ƃ���B
    cmdInstall.Enabled = True
    cmdDataHanei.Enabled = True
    cmdOut.Enabled = True
    cmdCancel.Enabled = True
    
    '���X�g�{�b�N�X�ɍ��ڂ��Ȃ��ꍇ�A�u���v�u�����v�u���v�u�����v��False�̂܂܂Ƃ���B
    If LstStation.ListCount <> 0 Then
        cmdUp.Enabled = True
        cmdPageUp.Enabled = True
        cmdPageDown.Enabled = True
        cmdDown.Enabled = True
    End If
    
End Sub
'V1.12.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : pfuncInstallEkiSettei
'//  �@�\����  : �w�ݒ�C���X�g�[������
'//  �@�\�T�v  :
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfuncInstallEkiSettei() As Boolean

    Dim bRet As Boolean                  ' �֐��߂�l
    Dim lErrCode As Long                 ' �G���[�R�[�h
    Dim bSysChange As Boolean            ' �V�X�e���ݒ菈���߂�l
    Dim objFso As New FileSystemObject   ' �t�@�C���V�X�e���I�u�W�F�N�g

    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse

    '���݉w�ݒ�f�[�^�C���X�g�[������
    bRet = dllInstolEkiDataNow(gstrFileName, EKI_SETTI_FILE, lErrCode)

    If bRet = False Then

        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)

        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)

        pfuncInstallEkiSettei = False

        '�ُ�I��
        MsgBox "�����w�s�x�f�[�^���f�������ُ�I�����܂����B", _
                vbOKOnly + vbCritical, _
                "�����w�s�x�f�[�^���f����"

    Else

        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)

        '----------------------------------------------------
        '���݂̐ݒu�w���x���X�V
        '----------------------------------------------------
        Call sDispNowEkiLabel
        
        '----------------------------------------------------
        '�R���s���[�^���A�l�b�g���[�N�ύX����
        '----------------------------------------------------
        bSysChange = pfNetWorkChng(Me)

        If bSysChange = True Then

            '���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET, 0)

            '����I��
            MsgBox "�����w�s�x�f�[�^���f����������I�����܂����B", _
                    vbOKOnly + vbInformation, _
                     "�����w�s�x�f�[�^���f����"
        Else

            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, EKITUDODATASET_EKI_TUDO_SET_NG, 0)

            '�ُ�I��
            MsgBox "�����w�s�x�f�[�^���f�������ُ�I�����܂����B", _
                    vbOKOnly + vbCritical, _
                     "�����w�s�x�f�[�^���f����"
        End If

    End If

    gstrFileName = ""
    Call SetEnableTrue

End Function


