VERSION 5.00
Begin VB.Form frmRenewData 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�e�ݒ�l�ۑ��E���f"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
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
   Begin VB.CommandButton cmdOutput 
      Caption         =   $"�W���ݒ�ۑ��������.frx":0000
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
      TabIndex        =   19
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame Frame7 
      Caption         =   "�R�[�i�I��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   8775
      Begin VB.CheckBox RenewChk 
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   870
         Value           =   1  '����
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Value           =   1  '����
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Value           =   1  '����
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Value           =   1  '����
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Value           =   1  '����
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CheckBox RenewChk 
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   4440
         Value           =   1  '����
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "�ۑ��ςݐݒ�t�@�C���쐬��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9�N Z9�� Z9�� Z9�� Z9�� Z9�b"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   17
         Top             =   915
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9�N Z9�� Z9�� Z9�� Z9�� Z9�b"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   16
         Top             =   1605
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9�N Z9�� Z9�� Z9�� Z9�� Z9�b"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   15
         Top             =   2325
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9�N Z9�� Z9�� Z9�� Z9�� Z9�b"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   14
         Top             =   3045
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9�N Z9�� Z9�� Z9�� Z9�� Z9�b"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   13
         Top             =   3765
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label lblSetteDate 
         Caption         =   "  ZZZ9�N Z9�� Z9�� Z9�� Z9�� Z9�b"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4080
         TabIndex        =   12
         Top             =   4485
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.Label Label2 
         Caption         =   "�R�[�i��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   405
         TabIndex        =   11
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  �@����ݒ�    ��ʂ֖߂�"
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
      Left            =   9500
      TabIndex        =   3
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8040
   End
   Begin VB.CommandButton cmdRenew 
      Caption         =   "����"
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
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdKeep 
      Caption         =   "�ۑ�"
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
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�W���ݒ� �ۑ��^����"
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
Attribute VB_Name = "frmRenewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �F�W���ݒ�ۑ��������.frm
'//  �p�b�P�[�W���F�W���ݒ�ۑ�������ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//                 �E��}�ێ���A�W���ݒ�ۑ�����(frmRenewData.frm)�𗬗p
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const K_SETTEI = 0                  '�Ď��ݒ�
Private Const G_SETTEI = 1                  '�����ݒ�
Private Const MSG_NORMAL = "����"           '����I�����\������
Private Const MSG_ERROR = "�ُ�"            '�ُ�I�����\������
Private Const RET_ERROR = -1                '�ُ�
Private Const RET_NASI = 0                  '�ύX����
Private Const RET_ARI = 1                   '�ύX�L��

Private Const INVALID_HANDLE_VALUE = -1     '�n���h���G���[
Private Const MN_MAIL_INTERVAL = 1000       '���C���^�C�}�̃C���^�[�o���l

Private glbSaveFoldePath    As String       '�ۑ��t�@�C���i�[�p�t�H���_�p�X
Private udtIniGate          As INI_GATE     '�@���񎩓����D�@�G���A           'EG20 V2.1.0.1 DEL �yMainte_03_01�z
Private udtIniGateFile      As INI_GATE     '�@���񎩓����D�@�G���A�ۑ��t�@�C��

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdOutput_Click
'//  �@�\����  : �ۑ��f�[�^�}�̏o�͖t����������
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 5.4.0.1) 2012-03-25  REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()

    Dim iResponse       As Integer      'MsgBox�߂�l
    Dim blnChecked      As Boolean      '�ΏۃR�[�i�`�F�b�N�L��
    Dim blnExistFile    As Boolean      '�ۑ��t�@�C���L���`�F�b�N
    Dim intCount        As Integer      '���[�v�J�E���^
    
    On Error Resume Next

    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_RENEW, 0)
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    '�����ΏۑI��L���`�F�b�N
    blnChecked = False
    blnExistFile = True
    Erase glngTergetCorner
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnChecked = True
            glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
            If lblSetteDate(intCount).Caption = "    �N   ��   ��   ��   ��   �b" Then
                blnExistFile = False
            End If
        End If
    Next intCount
    '�I���R�[�i�Ȃ��̏ꍇ�A���b�Z�[�W��\�����ďI��
    If blnChecked = False Then
' EG20 5.4.0.1 �폜�J�n
'        iResponse = MsgBox("�ΏۃR�[�i���I������Ă��܂���B" & vbCrLf & _
'                            "�I�����Ă��������B", vbOKOnly + vbExclamation, "�R�[�i���I��")
' EG20 5.4.0.1 �폜�I��
' EG20 5.4.0.1 �ǉ��J�n
        iResponse = MsgBox("�ΏۃR�[�i���I������Ă��܂���B" & vbCrLf & _
                            "�I�����Ă��������B", vbOKOnly + vbCritical, "�R�[�i���I��")
' EG20 5.4.0.1 �ǉ��I��
        Exit Sub
    End If
    '�ۑ��t�@�C�����݂Ȃ��̏ꍇ�A���b�Z�[�W��\�����ďI��
    If blnExistFile = False Then
' EG20 5.4.0.1 �폜�J�n
'        iResponse = MsgBox("�}�̏o�͂Ɏg�p����f�[�^������܂���B", vbOKOnly + vbExclamation, "�f�[�^�Ȃ�")
' EG20 5.4.0.1 �폜�I��
' EG20 5.4.0.1 �ǉ��J�n
        iResponse = MsgBox("�}�̏o�͂Ɏg�p����f�[�^������܂���B", vbOKOnly + vbCritical, "�f�[�^�Ȃ�")
' EG20 5.4.0.1 �ǉ��I��
        Exit Sub
    End If
    'EG20 V2.1.0.1 ADD END
    
    iResponse = MsgBox("�ۑ��f�[�^���O���}�̂ɏo�͂��܂��B" & vbCrLf & _
                        "���s���Ă���낵���ł����H", vbOKCancel + vbExclamation, "���s�m�F")

    '�u�L�����Z���v�{�^�����������͏������I������
    If iResponse = vbCancel Then Exit Sub

    frmRenewOutput.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �W���ݒ�ۑ��������(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
    '�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �W���ݒ�ۑ��������(�f�B�A�N�e�B�u��:�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
Private Sub Form_Deactivate()
    '�^�C�}���~����
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �W���ݒ�ۑ��������(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-02  CODED BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim lSts As Long
    Dim strPath As String * 128
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intCount As Integer         '���[�v�J�E���^
    Dim intIndex As Integer         '�`�F�b�N�{�b�N�XIndex
    Dim strSaveFile As String       '�ۑ��t�@�C���p�X
    'EG20 V2.1.0.1 ADD END
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    glbSaveFoldePath = ""
    strPath = ""
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    
    intIndex = 0
    gsGetGateInfo
    
    Call gsGetGateInfo
    Call gsGetCornerName
    Call gsGetCornerType        '�R�[�i��ʂ��擾   EG20 V30.1.0.1 ADD
    
    For intCount = 0 To UBound(gblnCornerSet)
    
        '�ݒ肠��̃R�[�i
        If gblnCornerSet(intCount) = True Then
            '�R�[�i�[���̕\��
            RenewChk(intIndex).Caption = gstrCornerName(intCount)
            '�R�[�iIndex���L�^
            RenewChk(intIndex).Tag = CStr(intCount)
            RenewChk(intIndex).Visible = True
            lblSetteDate(intIndex).Visible = True
            intIndex = intIndex + 1
        End If
    
    Next intCount
    'EG20 V2.1.0.1 ADD END
    
    ' RENEWDATAINFO.INI����ۑ���t�H���_�p�X���擾����
    lSts = GetPrivateProfileString(RENEWDATA_SECTION_NAME, _
                                   FOLDER_PATH_KEY_NAME, _
                                   "", _
                                   strPath, _
                                   Len(strPath), _
                                   PATH_RENEWDATAINFO_FILE)
    'INI�t�@�C�����擾����
    If strPath = "" Then
        'INI���擾�ُ펞�A�ݒ�l�ۑ������t�H���_���f�t�H���g�ݒ�Ƃ���
        glbSaveFoldePath = PATH_HOSHU_RENEW_DATA
    Else
        'INI����ݒ�
        glbSaveFoldePath = Left$(strPath, lSts) & "\\"
    End If
    
    '��ʐݒ菈��
    Call sFromInitialize
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdKeep_Click
'//  �@�\����  : �u�ۑ��v�t����������
'//  �@�\�T�v  : �m�F���b�Z�[�W�\����A�����E�Ď��ݒ�t�@�C����
'//              �����E�Ď��ݒ�ۑ��t�@�C���ɃR�s�[����
'//              �@���񎩓����D�@�G���A�̏����擾���A�ۑ��t�@�C���ɏ�������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 5.4.0.1) 2012-03-25  REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdKeep_Click()
    
    Dim iResponse       As Integer      'MsgBox�߂�l
    Dim strMessage      As String       'MsgBox����
    Dim bRet            As Boolean      '�߂�l
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intCount        As Integer
    Dim udtSendData     As MAIL_KAKARIIN_SETTEI
    Dim lngRet          As Long
    Dim blnIsSelected   As Boolean
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SAVE, 0)
    
    '�Ď��ݒ�f�[�^�ۑ��t�@�C���A�����ݒ�f�[�^�ۑ��t�@�C���̑��݃`�F�b�N
'    If Dir(glbSaveFoldePath & H_K_SETTEI_FILE) <> "" Or Dir(glbSaveFoldePath & H_G_SETTEI_FILE) <> "" Then 'EG20 V2.1.0.1 DEL �yMainte_03_01�z

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Erase glngTergetCorner
    blnIsSelected = False
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnIsSelected = True
            If RenewChk(intCount).Visible = True Then
                glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
            End If
        End If
    Next intCount
    '�I���R�[�i�Ȃ��̏ꍇ�A���b�Z�[�W��\�����ďI��
    If blnIsSelected = False Then
' EG20 5.4.0.1 �폜�J�n
'        iResponse = MsgBox("�ΏۃR�[�i���I������Ă��܂���B" & vbCrLf & _
'                            "�I�����Ă��������B", vbOKOnly + vbExclamation, "�R�[�i���I��")
' EG20 5.4.0.1 �폜�I��
' EG20 5.4.0.1 �ǉ��J�n
        iResponse = MsgBox("�ΏۃR�[�i���I������Ă��܂���B" & vbCrLf & _
                            "�I�����Ă��������B", vbOKOnly + vbCritical, "�R�[�i���I��")
' EG20 5.4.0.1 �ǉ��I��
        Exit Sub
    End If
    
    For intCount = 0 To lblSetteDate.UBound
        If lblSetteDate(intCount).Visible = True And RenewChk(intCount).Value = 1 And _
           lblSetteDate(intCount).Caption <> "    �N   ��   ��   ��   ��   �b" Then
    'EG20 V2.1.0.1 ADD END
        
            '�m�F���b�Z�[�W�{�b�N�X��\������B
            iResponse = MsgBox("�ݒ�t�@�C�����㏑�����܂�����낵���ł����H", _
                                vbOKCancel + vbCritical, "�㏑���ۑ��x��")
    
            '�u�L�����Z���v�{�^�����������͏������I������
            If iResponse = vbCancel Then Exit Sub
            Exit For        'EG20 V2.1.0.1 ADD �yMainte_03_01�z
        End If
    Next intCount           'EG20 V2.1.0.1 ADD �yMainte_03_01�z
    
    '�����ݒ�
    strMessage = MSG_NORMAL '����
    bRet = False            '�ُ�

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    frmRenewSave.Show vbModal
    'EG20 V2.1.0.1 ADD END
    
    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
    '�Ď��ݒ�f�[�^�t�@�C�����Ď��ݒ�f�[�^�ۑ��t�@�C���Ƃ��ăR�s�[
'    bRet = fCopySetteiFile(K_SETTEI_FILE, glbSaveFoldePath & H_K_SETTEI_FILE, MU_KSETTEI)
'    If bRet = False Then GoTo ErrorHandler          '�ُ�̏ꍇ�A�R�s�[�������I��
'
'    '�����ݒ�f�[�^�t�@�C���������ݒ�f�[�^�ۑ��t�@�C���Ƃ��ăR�s�[
'    bRet = fCopySetteiFile(G_SETTEI_FILE, glbSaveFoldePath & H_G_SETTEI_FILE, MU_GSETTEI)
'    If bRet = False Then GoTo ErrorHandler          '�ُ�̏ꍇ�A�R�s�[�������I��
'
'    '�@���񎩓����D�@�G���A�ۑ��t�@�C���쐬����
'    bRet = fKeepGateIniInf

'ErrorHandler:
'
'    '�����͐���ɏI���������H
'    If bRet = False Then
'        '�ُ폈��
'        Call fDeleteKeepFile        '�ۑ��t�@�C�����폜
'        strMessage = MSG_ERROR      '�ُ핶����ݒ�
'    End If
'
'    '�������ʃ��b�Z�[�W�{�b�N�X��\������B
'    iResponse = MsgBox("    " & strMessage & "�I�����܂����B    ", vbOKOnly, "�ۑ���������")

    'EG20 V2.1.0.1 DEL END
    
    '��ʕ\�����X�V
    Call sFromInitialize
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdRenew_Click
'//  �@�\����  : �u�����v�t����������
'//  �@�\�T�v  : �@���񎩓����D�@�G���A�Ƌ@���񎩓����D�@�G���A�ۑ��t�@�C�����r
'//              �����E�Ď��ݒ�f�[�^�ۑ��t�@�C���������E�Ď��ݒ�t�@�C���ɍX�V����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(EG20 5.4.0.1) 2012-03-25  REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdRenew_Click()
    Dim iResponse       As Integer      'MsgBox�߂�l
    Dim lngChgFlg       As Long         '��r��������
    Dim bRet            As Boolean      '���J�o����������
    Dim lngRet          As Long         '���[�����M��������
    Dim strKSetMessage  As String       '�Ď��ݒ�t�@�C���X�V��������
    Dim strGSetMessage  As String       '�����ݒ�t�@�C���X�V��������
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim blnChecked      As Boolean      '�ΏۃR�[�i�`�F�b�N�L��
    Dim blnExistFile    As Boolean      '�ۑ��t�@�C���L���`�F�b�N
    Dim intCount        As Integer      '���[�v�J�E���^
    Dim udtSendData     As MAIL_KAKARIIN_SETTEI
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_RENEW, 0)
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    '�����ΏۑI��L���`�F�b�N
    blnChecked = False
    blnExistFile = True
    Erase glngTergetCorner
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnChecked = True
            glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
            If lblSetteDate(intCount).Caption = "    �N   ��   ��   ��   ��   �b" Then
                blnExistFile = False
            End If
        End If
    Next intCount
    '�I���R�[�i�Ȃ��̏ꍇ�A���b�Z�[�W��\�����ďI��
    If blnChecked = False Then
' EG20 5.4.0.1 �폜�J�n
'        iResponse = MsgBox("�ΏۃR�[�i���I������Ă��܂���B" & vbCrLf & _
'                            "�I�����Ă��������B", vbOKOnly + vbExclamation, "�R�[�i���I��")
' EG20 5.4.0.1 �폜�I��
' EG20 5.4.0.1 �ǉ��J�n
        iResponse = MsgBox("�ΏۃR�[�i���I������Ă��܂���B" & vbCrLf & _
                            "�I�����Ă��������B", vbOKOnly + vbCritical, "�R�[�i���I��")
' EG20 5.4.0.1 �ǉ��I��
        Exit Sub
    End If
    '�ۑ��t�@�C�����݂Ȃ��̏ꍇ�A���b�Z�[�W��\�����ďI��
    If blnExistFile = False Then
' EG20 5.4.0.1 �폜�J�n
'        iResponse = MsgBox("�����Ɏg�p����f�[�^������܂���B", vbOKOnly + vbExclamation, "�f�[�^�Ȃ�")
' EG20 5.4.0.1 �폜�I��
' EG20 5.4.0.1 �ǉ��J�n
        iResponse = MsgBox("�����Ɏg�p����f�[�^������܂���B", vbOKOnly + vbCritical, "�f�[�^�Ȃ�")
' EG20 5.4.0.1 �ǉ��I��
        Exit Sub
    End If
    'EG20 V2.1.0.1 ADD END
    
    iResponse = MsgBox("�Ď��Ղɕۑ��ς݂̐ݒ�l�𔽉f�����܂��B" & vbCrLf & _
                        "���s���Ă���낵���ł����H", vbOKCancel + vbExclamation, "���s�m�F")

    '�u�L�����Z���v�{�^�����������͏������I������
    If iResponse = vbCancel Then Exit Sub

    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'    strKSetMessage = MSG_ERROR  '�X�V�����ُ�I��
'    strGSetMessage = MSG_ERROR  '�X�V�����ُ�I��
'
'    On Error GoTo ErrorHandler
'
'    lngChgFlg = RET_ERROR                           '�����ݒ�i-1�F�ُ�j
'
'    lngChgFlg = fCompareGateIniInf                              '�@���񎩓����D�@�G���A��r����
'    If lngChgFlg <> RET_NASI Then GoTo ErrorHandler             '�ύX�������̏ꍇ�A�X�V�������s��
'
'    bRet = False                                    '�����ݒ�iFalse�F�ُ�j
'    lngRet = INVALID_HANDLE_VALUE                   '�����ݒ�i-1�F�ُ�j
'
'    bRet = dllK_Settei_File_Recovery                            '�Ď����u�ݒ�f�[�^�t�@�C�����J�o������
'    If bRet = False Then GoTo ErrorHandler                      '�ُ�̏ꍇ�A�X�V�������I��
'
'    lngRet = fKansiSetteiMailSend                               '�Ď��ݒ�w�����[�����M
'    If lngRet = INVALID_HANDLE_VALUE Then GoTo ErrorHandler     '�ُ�̏ꍇ�A�X�V�������I��
'
'    '*****�Ď����u�ݒ�f�[�^�X�V��������I��
'    strKSetMessage = MSG_NORMAL                     '����I���̃��b�Z�[�W��ݒ�

'    bRet = dllG_Settei_File_Recovery                            '�����ݒ�f�[�^�t�@�C�����J�o������
'    If bRet = False Then GoTo ErrorHandler                      '�ُ�̏ꍇ�A�X�V�������I��

'    lngRet = fGateSetteiMailSend                                '�W���ݒ�ۑ��v�����[�����M
    'EG20 V2.1.0.1 DEL END
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    frmRenewCyu.Show vbModal
    'EG20 V2.1.0.1 ADD END
    
    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'    If lngRet = INVALID_HANDLE_VALUE Then GoTo ErrorHandler     '�ُ�̏ꍇ�A�X�V�������I��
'
'    '*****�����ݒ�f�[�^�X�V��������I��
'    strGSetMessage = MSG_NORMAL                     '����I���̃��b�Z�[�W��ݒ�
'
'ErrorHandler:
'
'    '�e�ۑ��t�@�C�����폜
'    Call fDeleteKeepFile
'
'    '�������ʃ��b�Z�[�W��\��
'    If lngChgFlg = RET_ARI Then     '�ύX�L��̏ꍇ
'        '�����\���ύX�L�胁�b�Z�[�W�{�b�N�X��\������B
'        iResponse = MsgBox("�����\�����ύX���ꂽ����" & vbCrLf & _
'                            "�X�V�����͂ł��܂���B", vbOKOnly + vbExclamation, "���f��������")
'    Else
'        '�������ʃ��b�Z�[�W�{�b�N�X��\������B
'        iResponse = MsgBox("    �����ݒ�̍X�V��" & strGSetMessage & "�I�����܂����B    " & vbCrLf & _
'                           "    �Ď��ݒ�̍X�V��" & strKSetMessage & "�I�����܂����B    ", vbOKOnly, "���f��������")
'    End If
    'EG20 V2.1.0.1 DEL END

    '��ʕ\�����X�V
    Call sFromInitialize
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_END, 0)
    
    '����ʂ������B
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFromInitialize
'//  �@�\����  : ��ʐݒ菈��
'//  �@�\�T�v  : �t�@�C���̍쐬�������擾���A�쐬���t�\�����ɕ\��
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
Private Sub sFromInitialize()
    Dim strFileName(0 To 1)     As String           '�쐬����
    Dim intCnt                  As Integer          '�J�E���^
    Dim lngHandle               As Long             '�n���h��

    Dim lpCreatTime             As FILETIME         '�쐬����
    Dim lpAccessTime            As FILETIME         '�ŏI�A�N�Z�X����
    Dim lpLastwTime             As FILETIME         '�X�V����
    Dim lpLocalTime             As FILETIME         '���[�J������
    Dim lpSystemTime            As SYSTEMTIME       '�V�X�e������
    Dim bRet                    As Boolean          '�߂�l
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim blnExistFile            As Boolean          '�ۑ��t�@�C���L��
    Dim strSaveFile             As String
    Dim intIndex                As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

                
                
    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'    '�Ď��ݒ�f�[�^�ۑ��t�@�C���A�����ݒ�f�[�^�ۑ��t�@�C���A�@���񎩓����D�@�G���A�ۑ��t�@�C���̑��݃`�F�b�N
'    '��ł����݂��Ȃ��ꍇ�A�u�����N��\��
'    If Dir(glbSaveFoldePath & H_K_SETTEI_FILE) = "" Or _
'       Dir(glbSaveFoldePath & H_G_SETTEI_FILE) = "" Or _
'       Dir(glbSaveFoldePath & H_G_INFO_FILE) = "" Then GoTo ErrorHandler
'
'    strFileName(K_SETTEI) = glbSaveFoldePath & H_K_SETTEI_FILE     '�Ď��ݒ�f�[�^�t�@�C��
'    strFileName(G_SETTEI) = glbSaveFoldePath & H_G_SETTEI_FILE     '�����ݒ�f�[�^�t�@�C��
'
'    '�t�@�C���̍쐬�������擾���A�쐬���t�\�����ɕ\��
'    For intCnt = 0 To lblSetteDate.ubound
'
'        '�t�@�C�����I�[�v��
'        lngHandle = CreateFile(strFileName(intCnt), GENERIC_READ, FILE_SHARE_READ, _
'                                0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    blnExistFile = False
    intIndex = 0
    For intCnt = 0 To UBound(gudtSettiCorner)
        If gblnCornerSet(intCnt) = True Then
            '�ۑ��t�@�C���̓��t���擾
            strSaveFile = PATH_OPERATE & "CORNER" & CStr(intCnt + 1) & "\\SETTEI\\" & CONDENSE_FILE
            If Dir(strSaveFile) = "" Then
                lblSetteDate(intIndex).Caption = "    �N   ��   ��   ��   ��   �b"
            Else
                '�t�@�C�����I�[�v��
                lngHandle = CreateFile(strSaveFile, GENERIC_READ, FILE_SHARE_READ, _
                                        0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    'EG20 V2.1.0.1 ADD END

                '�t�@�C���I�[�v��������ɍs��ꂽ���H
                If lngHandle = INVALID_HANDLE_VALUE Then GoTo ErrorHandler
        
                '�t�@�C���^�C����GET
                bRet = GetFileTime(lngHandle, lpCreatTime, lpAccessTime, lpLastwTime)
                If bRet = False Then GoTo APIError                          '�擾������ɍs��ꂽ���H
        
                '�t�@�C���^�C�������[�J���^�C���ɕϊ�
'                bRet = FileTimeToLocalFileTime(lpCreatTime, lpLocalTime)    'EG20 V2.1.0.1 DEL �yMainte_03_01�z
                bRet = FileTimeToLocalFileTime(lpLastwTime, lpLocalTime)    'EG20 V2.1.0.1 ADD �yMainte_03_01�z
                If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H
        
                '���[�J���^�C�����V�X�e���^�C���ɕϊ�
                bRet = FileTimeToSystemTime(lpLocalTime, lpSystemTime)
                If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H
        
                '�n���h���̃N���[�Y
                Call CloseHandle(lngHandle)
        
                '�쐬���t��\������ (YYYY�NMM��DD��hh��mm��ss�b)
                lblSetteDate(intIndex).Caption = lpSystemTime.wYear & "�N " & _
                                                Right("  " & lpSystemTime.wMonth, 2) & "�� " & _
                                                Right("  " & lpSystemTime.wDay, 2) & "�� " & _
                                                Right("  " & lpSystemTime.wHour, 2) & "�� " & _
                                                Right("  " & lpSystemTime.wMinute, 2) & "�� " & _
                                                Right("  " & lpSystemTime.wSecond, 2) & "�b"
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
                blnExistFile = True
            End If
            
            lblSetteDate(intIndex).Visible = True
            intIndex = intIndex + 1
        Else
            lblSetteDate(intIndex).Visible = False
        End If
    'EG20 V2.1.0.1 ADD END
    Next

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    '�ۑ��t�@�C�����P���Ȃ��ꍇ�͕����{�^���A�}�̏o�͉����s��
    If blnExistFile = False Then
        cmdRenew.Enabled = False
        cmdOutput.Enabled = False
    Else
        cmdOutput.Enabled = True
    'EG20 V2.1.0.1 ADD END
        cmdRenew.Enabled = True     '�ݒ�l���f�{�^��������
    End If          'EG20 V2.1.0.1 ADD �yMainte_03_01�z
        
    Exit Sub

APIError:

    Call CloseHandle(lngHandle)             '�n���h���̃N���[�Y

ErrorHandler:

    '���݂��Ȃ��ꍇ�܂��̓G���[�����������A�u�����N��\��
    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'    lblSetteDate(K_SETTEI).Caption = "    �N   ��   ��   ��   ��   �b"
'    lblSetteDate(G_SETTEI).Caption = "    �N   ��   ��   ��   ��   �b"
    'EG20 V2.1.0.1 DEL END

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    For intCnt = intCnt To UBound(gudtSettiCorner)
        If lblSetteDate(intCnt).Visible = True Then
            lblSetteDate(intCnt).Caption = "    �N   ��   ��   ��   ��   �b"
        End If
    Next intCnt
    'EG20 V2.1.0.1 ADD END
    
    cmdRenew.Enabled = False    '�ݒ�l���f�{�^�������s��
    cmdOutput.Enabled = False   '�}�̏o�̓{�^�������s��     'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fCopySetteiFile
'//  �@�\����  : �ݒ�t�@�C���R�s�[����
'//  �@�\�T�v  : �ݒ�t�@�C�����w��t�H���_�փR�s�[����
'//�@�@�@�@�@�@�@�t�@�C���̍쐬�����A�A�N�Z�X�����A�X�V������ݒ肷��
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : String�@�@strFromFile�@�@[IN]�R�s�[���t�@�C����
'//  �@�@�@�@�@�@String�@�@strToFile�@�@�@[IN]�R�s�[��t�@�C����
'//  �@�@�@�@�@�@String�@�@strMutexName �@[IN]�~���[�e�b�N�X��
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : Boolean�@ True           ����I��
'//                     �@ False          �ُ�I��
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fCopySetteiFile(strFromFile As String, strToFile As String, strMutexName As String) As Boolean
    Dim lngHandle               As Long            '�n���h��
    Dim lngMuHandle             As Long            '�r�������p�n���h��

    Dim lpCreatTime             As FILETIME        '�쐬����
    Dim lpAccessTime            As FILETIME        '�ŏI�A�N�Z�X����
    Dim lpLastwTime             As FILETIME        '�X�V����
    Dim lpLocalTime             As FILETIME        '���[�J������
    Dim lpSystemTime            As SYSTEMTIME      '�V�X�e������
    Dim bRet                    As Boolean         '�߂�l

    On Error Resume Next

    '�����ݒ�
    fCopySetteiFile = False
    bRet = False

    '�ݒ�f�[�^�t�@�C���̗L���`�F�b�N
    If Dir(strFromFile) = "" Then Exit Function     '���݂��Ȃ��ꍇ�A�������I��

    lngMuHandle = dllOpenMutex(strMutexName)        '�r������(OPEN)

    If lngMuHandle <> 0 Then
        dllWaitForSingleObject (lngMuHandle)        '�r������(GET)
    End If

    bRet = CopyFile(strFromFile, strToFile, False)  '�ݒ�t�@�C���R�s�[����
    If bRet = False Then GoTo ErrorHandler          '�R�s�[�����͐���ɍs��ꂽ���H

    bRet = False    '�Đݒ�

    '�t�@�C���̍쐬�����A�A�N�Z�X�����A�X�V������ݒ�
    '�ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(strToFile, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                            0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then GoTo ErrorHandler

    '���[�J���������擾
    Call GetLocalTime(lpSystemTime)

    '�V�X�e���^�C�������[�J���^�C���ɕϊ�
    bRet = SystemTimeToFileTime(lpSystemTime, lpLocalTime)
    If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H

    '���[�J���^�C�����t�@�C���^�C���ɕϊ��i�쐬�����j
    bRet = LocalFileTimeToFileTime(lpLocalTime, lpCreatTime)
    If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H

    '���[�J���^�C�����t�@�C���^�C���ɕϊ��i�A�N�Z�X�����j
    bRet = LocalFileTimeToFileTime(lpLocalTime, lpAccessTime)
    If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H

    '���[�J���^�C�����t�@�C���^�C���ɕϊ��i�X�V�����j
    bRet = LocalFileTimeToFileTime(lpLocalTime, lpLastwTime)
    If bRet = False Then GoTo APIError                          '�ϊ�������ɍs��ꂽ���H

    '�t�@�C���̓��t��ݒ�
    bRet = SetFileTime(lngHandle, lpCreatTime, lpAccessTime, lpLastwTime)

APIError:

    Call CloseHandle(lngHandle)                     '�n���h���̃N���[�Y

ErrorHandler:

    If lngMuHandle <> 0 Then
        dllReleaseMutex (lngMuHandle)                   '�r������(FREE)
        dllCloseHandle (lngMuHandle)                    '�r������(CLOSE)
    End If

    '�R�s�[�����͐���ɍs��ꂽ���H
    If bRet = False Then Exit Function              '�ُ�̏ꍇ�A�������I��

    '����I��
    fCopySetteiFile = True
End Function

'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fCopySetteiFile
'//  �@�\����  : �@���񎩓����D�@�G���A�擾
'//  �@�\�T�v  : �@���񎩓����D�@�G���A�����擾���A
'//�@�@�@�@�@�@�@�@���񎩓����D�@�G���A�ۑ��t�@�C���𐶐�����
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : Boolean�@ True           ����I��
'//                     �@ False          �ُ�I��
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Function fKeepGateIniInf() As Boolean
'
'    Dim bRet            As Boolean          '�����݌���
'
'    On Error Resume Next
'
'    '�����ݒ�
'    fKeepGateIniInf = False
'
'    '�@���񎩓����D�@�G���A�ǂݍ��ݏ���
'    sGetGateIniInf
'
'    '�@���񎩓����D�@�G���A�ۑ��t�@�C���������ݏ���
'    bRet = fWriteGateIniInf
'
'    '���ʐݒ�
'    fKeepGateIniInf = bRet
'
'End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fCompareGateIniInf
'//  �@�\����  : �@���񎩓����D�@�G���A��r����
'//  �@�\�T�v  : �@���񎩓����D�@�G���A��
'//�@�@�@�@�@�@�@�@���񎩓����D�@�G���A�ۑ��t�@�C�����r����
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : Long�@    -1             �ُ�
'//                     �@ 0              �ύX����
'//                     �@ 1              �ύX�L��
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Function fCompareGateIniInf() As Long
'    Dim bRet            As Boolean      '�ǂݍ��ݏ�������
'    Dim bChgFlg         As Boolean      '�ύX�L��
'    Dim intCnt          As Integer      '�J�E���^
'    Dim lngHandle       As Long         '�n���h��
'
'    On Error GoTo ErrorHandler
'
'    '�����ݒ�
'    fCompareGateIniInf = RET_ERROR  '�ُ�
'
'    '�@���񎩓����D�@�G���A�ۑ��t�@�C���ǂݍ��ݏ���
'    bRet = fReadGateIniInf
'    '�ۑ��t�@�C���ǂݍ��݂�����ɍs��ꂽ���H
'    If bRet = False Then Exit Function
'
'    '�@���񎩓����D�@�G���A�ǂݍ��ݏ���
'    Call sGetGateIniInf
'
'    '�����ݒ�
'    bChgFlg = True      '�ύX����
'
'    '���@�����A�@���񎩓����D�@�G���A�Ƌ@���񎩓����D�@�G���A�ۑ��t�@�C���Ƃ̔�r
'    For intCnt = 0 To MAX_GATE_NO - 1
'        'NEG�^/C�^�E�W�D/���D/���p
'        If udtIniGate.Gate_Set(intCnt).nGate <> udtIniGateFile.Gate_Set(intCnt).nGate Or _
'            udtIniGate.Gate_Set(intCnt).nTuuro <> udtIniGateFile.Gate_Set(intCnt).nTuuro Then
'            '�����\���ύX�L��
'            bChgFlg = False
'            Exit For
'        End If
'    Next
'
'    '�����\���ɕύX�������������H
'    If bChgFlg = True Then
'        '�ύX���Ȃ������ꍇ�A�ύX������Ԃ�
'        fCompareGateIniInf = RET_NASI
'    Else
'        '�ύX���������ꍇ�A�ύX�L���Ԃ�
'        fCompareGateIniInf = RET_ARI
'    End If
'
'    Exit Function
'ErrorHandler:
'    '�ُ폈��
'End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sGetGateIniInf
'//  �@�\����  : �@���񎩓����D�@�G���A�ǂݍ��ݏ���
'//  �@�\�T�v  : �@���񎩓����D�@�G���A�����擾����
'//
'//              �^         ����     �@�@�@ �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^         �l        �@�@  �Ӗ�
'//  �߂�l    : Boolean    TRUE            ����I��
'//                         FALSE           �ُ�I��
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sGetGateIniInf()
'
'    Dim udtMapInf       As MAP_MEM          '�������}�b�s���O�I�u�W�F�N�g
'
'    Dim bRet            As Boolean          '�����݌���
'    Dim strIniData      As String * 1024    'INI�ݒ�l
'    Dim strKeyName      As String           '�L�[��
'    Dim strMutexName    As String           '�~���[�e�b�N�X��
'    Dim lSts            As Long             '�֐��߂�l
'    Dim lErrCode        As Long             '�G���[�R�[�h
'    Dim lngMuHandle     As Long             '�r�������p�n���h��
'    Dim iLoopCnt        As Integer          '���[�v�J�E���^
'
'    On Error Resume Next
'
'    strMutexName = "Mu_" & GIniGate
'    lngMuHandle = dllOpenMutex(strMutexName)         '�r������(OPEN)
'    If lngMuHandle <> 0 Then
'
'        dllCloseHandle (lngMuHandle)                 '�r������(CLOSE)
'
'        '�@�����`�G���A�̓��e���擾����B
'        '�G���A�̏�����
'        Call dllMemMappingInit(GIniGate, 0, MUTEXMODE_ON, udtMapInf)
'
'        '�G���A�̓��e���擾����B
'        Call dllMemMappingRead(udtMapInf.lngpAdr, LenB(udtIniGate), MUTEXMODE_ON, udtMapInf.lnghMutex, udtIniGate)
'
'        '�G���A���������
'        Call dllMemMappingEnd(udtMapInf, udtMapInf.lnghMutex)
'
'    Else
'
'        '�G���A�����݂��Ȃ��ꍇINI�t�@�C������擾����
'        For iLoopCnt = 0 To MAX_GATE_NO - 1
'            strIniData = ""
'            strKeyName = INI_GATE_KEY & Format(iLoopCnt + 1, "00")
'            lSts = GetPrivateProfileString(INI_GATE_SECTION, _
'                                           strKeyName, _
'                                           "Defo", _
'                                           strIniData, _
'                                           Len(strIniData), _
'                                           PATH_GATE_FILE)
'            If lSts > 0 Then
'                bRet = dllMemIniGate(udtIniGate.Gate_Set(iLoopCnt), strIniData, lErrCode)
'                If (bRet = False) Then
'                    '�ُ탍�O�o��
'                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, KAKARISET_GAMEN_GET_GATE_AREA_ERROR, lErrCode)
'                End If
'            Else
'                Exit Sub
'            End If
'        Next
'
'    End If
'
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fWriteGateIniInf
'//  �@�\����  : �@���񎩓����D�@�G���A�ۑ��t�@�C���������ݏ���
'//  �@�\�T�v  : �@���񎩓����D�@�G���A�ۑ��t�@�C���𐶐�����
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : Boolean�@ True           ����I��
'//                     �@ False          �ُ�I��
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Function fWriteGateIniInf() As Boolean
'    Dim lngHandle       As Long         '�n���h��
'    Dim lngRet          As Long         '�������܂ꂽ�o�C�g���̃A�h���X
'    Dim bRet            As Boolean      '�������݌���
'
'    On Error GoTo ErrorHandler
'
'    '�����ݒ�
'    fWriteGateIniInf = False
'
'    '�t�@�C�����쐬
'    lngHandle = CreateFile(glbSaveFoldePath & H_G_INFO_FILE, GENERIC_WRITE, FILE_SHARE_WRITE Or FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
'
'    '�t�@�C���쐬������ɍs��ꂽ���H
'    If lngHandle = INVALID_HANDLE_VALUE Then Exit Function
'
'    '�t�@�C���̏�������
'    bRet = WriteFile(lngHandle, udtIniGate, LenB(udtIniGate), lngRet, 0)
'
'    '�n���h���̃N���[�Y
'    Call CloseHandle(lngHandle)
'
'    '�������݌��ʐݒ�
'    fWriteGateIniInf = bRet
'
'    Exit Function
'
'ErrorHandler:
'    '�ُ폈��
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : fReadGateIniInf
''//  �@�\����  : �@���񎩓����D�@�G���A�ۑ��t�@�C���ǂݍ��ݏ���
''//  �@�\�T�v  : �@���񎩓����D�@�G���A�ۑ��t�@�C����ǂݍ���
''//
''//              �^        ����     �@�@�@�Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �@�@ �Ӗ�
''//  �߂�l    : Boolean�@ True           ����I��
''//                     �@ False          �ُ�I��
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Function fReadGateIniInf() As Boolean
'    Dim lngHandle       As Long         '�n���h��
'    Dim lngRet          As Long         '�ǂݍ��܂ꂽ�o�C�g���̃A�h���X
'    Dim bRet            As Boolean      '�ǂݍ��݌���
'
'    On Error GoTo ErrorHandler
'
'    '�����ݒ�
'    fReadGateIniInf = False
'
'    '�t�@�C�����I�[�v��
'    lngHandle = CreateFile(glbSaveFoldePath & H_G_INFO_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
'
'    '�t�@�C���I�[�v��������ɍs��ꂽ���H
'    If lngHandle = INVALID_HANDLE_VALUE Then Exit Function
'
'    '�@���񎩓����D�@�G���A�ۑ��t�@�C���ǂݍ��ݏ���
'    bRet = ReadFile(lngHandle, udtIniGateFile, LenB(udtIniGateFile), lngRet, 0)
'
'    '�n���h���̃N���[�Y
'    Call CloseHandle(lngHandle)
'
'    '�ǂݍ��݌��ʐݒ�
'    fReadGateIniInf = bRet
'
'    Exit Function
'
'ErrorHandler:
'    '�ُ폈��
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : fDeleteKeepFile
''//  �@�\����  : �e�ۑ��t�@�C���̍폜����
''//  �@�\�T�v  : �Ď��ݒ�f�[�^�ۑ��t�@�C���A�����ݒ�f�[�^�ۑ��t�@�C���A
''//�@�@�@�@�@�@�@�@���񎩓����D�@�G���A�ۑ��t�@�C�������݂��Ă����ꍇ�A�폜����
''//
''//              �^        ����     �@�@�@�Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �@�@ �Ӗ�
''//  �߂�l    : Boolean�@ True           ����I��
''//                     �@ False          �ُ�I��
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Function fDeleteKeepFile() As Boolean
'    Dim bRet    As Boolean      '��������
'    Dim lngRet  As Long
'
'    On Error GoTo ErrorHandler
'
'    '��������
'    fDeleteKeepFile = False
'
'    If Dir(glbSaveFoldePath & H_K_SETTEI_FILE) <> "" Then          '�Ď��ݒ�f�[�^�ۑ��t�@�C��
'        bRet = DeleteFile(glbSaveFoldePath & H_K_SETTEI_FILE)      '�t�@�C���폜����
'    End If
'
'    If Dir(glbSaveFoldePath & H_G_SETTEI_FILE) <> "" Then          '�����ݒ�f�[�^�ۑ��t�@�C��
'        bRet = DeleteFile(glbSaveFoldePath & H_G_SETTEI_FILE)      '�t�@�C���폜����
'    End If
'
'    If Dir(glbSaveFoldePath & H_G_INFO_FILE) <> "" Then            '�@���񎩓����D�@�G���A�ۑ��t�@�C��
'        bRet = DeleteFile(glbSaveFoldePath & H_G_INFO_FILE)        '�t�@�C���폜����
'    End If
'
'    '����I��
'    fDeleteKeepFile = True
'
'    Exit Function
'
'ErrorHandler:
'    '�ُ폈��
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : fKansiSetteiMailSend
''//  �@�\����  : ������w�����[�����M�����i�Ď��ݒ�j
''//  �@�\�T�v  : �ă}�v���Z�X�ցu�����ݒ�w���v���[���𑗐M����
''//
''//              �^        ����     �@�@�@�Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �@�@ �Ӗ�
''//  �߂�l    : Long�@ �@ �T�C�Y         ���[�����M�T�C�Y
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Function fKansiSetteiMailSend() As Long
'    Dim lngMSlot_KM As Long                 '�ă}�̃��[���X���b�g�n���h��
'    Dim udtMail     As MAIL_GATE_SET_ORD    '�����ݒ�w�����[�����M�G���A
'    Dim lngRet      As Long                 '�֐��߂�l
'    Dim intCnt      As Integer              '�J�E���^
'
'    On Error Resume Next
'
'    '�����ݒ�
'    fKansiSetteiMailSend = INVALID_HANDLE_VALUE
'
'    '���ʃw�b�_�ҏW
'    udtMail.mlHeader.dwId = ML_ID_GATE_SET_ORD
'    udtMail.mlHeader.dwSize = Len(udtMail)
'    udtMail.mlHeader.dwProid = RHOSHU_ID
'    udtMail.mlHeader.dwSubArea = 0
'
'    '�G���A��ʂ�ݒ�
'    udtMail.dwCmnFile = K_SETTEI_FILE_NO
'
'    '�ݒ���
'    udtMail.dwGateSet(0) = 1
'    For intCnt = 1 To MAX_GATE_NO - 1
'        udtMail.dwGateSet(intCnt) = 0
'    Next intCnt
'
'    '���[�����M
'    lngRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.GATE_SET_ORD, udtMail.mlHeader)
'
'    '���b�Z�[�W���M����
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SENDMAIL, 0)
'
'    '�������ʂ�Ԃ�
'    fKansiSetteiMailSend = 1
'
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : fGateSetteiMailSend
''//  �@�\����  : ������w�����[�����M�����i�����ݒ�j
''//  �@�\�T�v  : �ă}�v���Z�X�ցu�����ݒ�w���v���[���𑗐M����
''//
''//              �^        ����     �@�@�@�Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �@�@ �Ӗ�
''//  �߂�l    : Long�@ �@ �T�C�Y         ���[�����M�T�C�Y
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Function fGateSetteiMailSend() As Long
'    Dim lngMSlot_KM As Long                 '�ă}�̃��[���X���b�g�n���h��
'    Dim udtMail     As MAIL_GATE_SET_ORD    '�����ݒ�w�����[�����M�G���A
'    Dim lngRet      As Long                 '�֐��߂�l
'    Dim intCnt      As Integer              '�J�E���^
'
'    On Error Resume Next
'
'    '�����ݒ�
'    fGateSetteiMailSend = INVALID_HANDLE_VALUE
'
'    '���ʃw�b�_�ҏW
'    udtMail.mlHeader.dwId = ML_ID_GATE_SET_ORD
'    udtMail.mlHeader.dwSize = Len(udtMail)
'    udtMail.mlHeader.dwProid = RHOSHU_ID
'    udtMail.mlHeader.dwSubArea = 0
'
'    '�G���A��ʂ�ݒ�
'    udtMail.dwCmnFile = G_SETTEI_FILE_NO
'
'    '�ݒ���
'    For intCnt = 0 To MAX_GATE_NO - 1
'        If Not udtIniGate.Gate_Set(intCnt).intGate = GATE_NASI Then
'            udtMail.dwGateSet(intCnt) = 1
'        Else
'            udtMail.dwGateSet(intCnt) = 0
'        End If
'    Next intCnt
'
'    '���[�����M
'    lngRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.GATE_SET_ORD, udtMail.mlHeader)
'
'    '���b�Z�[�W���M����
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SENDMAIL, 0)
'
'    '�������ʂ�Ԃ�
'    fGateSetteiMailSend = 1
'End Function
'EG20 V2.1.0.1 DEL END

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
'//  �߂�l    : Long�@ �@ �T�C�Y         ���[�����M�T�C�Y
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmRenewData.Caption, False
        pfFormActive (frmRenewData.hwnd)
    End If
    
End Sub

'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sRcv_Renew
'//  �@�\����  : �W���ݒ蕜���v��RES��M����
'//  �@�\�T�v  : �W���ݒ蕜���v��RES��M���̏������s��
'//
'//              �^           ����     �@�@�@�Ӗ�
'//  ����      : ML_KYOTU_INF udtReadMail    ��M�f�[�^
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : Long�@ �@ �T�C�Y         ���[�����M�T�C�Y
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-13   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sRcv_Renew(ByRef udtReadMail As ML_KYOTU_INF, ByVal strMsgTitle)

    Dim intCounta As Integer
    Dim blnIsErr As Boolean
    Dim intCount As Integer
    Dim iResponse As Integer
    
    On Error Resume Next
    
    blnIsErr = False
    '�������ʔ���
    For intCount = 0 To lblSetteDate.UBound
        If udtReadMail.lngData(intCount) > 0 Then
            blnIsErr = True
            Exit For
        End If
    Next intCount
    
    '�t�@�C���쐬�������X�V����
    Call sFromInitialize
    
    '�������ʕ\��
    If blnIsErr = True Then
        iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly, strMsgTitle)
    Else
        iResponse = MsgBox("����I�����܂����B", vbOKOnly, strMsgTitle)
    End If
        
End Sub
'EG20 V2.1.0.1 ADD END



