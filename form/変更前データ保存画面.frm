VERSION 5.00
Begin VB.Form frmSetteiBefore 
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
      TabIndex        =   3
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�f�[�^���W�E�o��  ��ʂ֖߂�"
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
      Left            =   9500
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8040
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
      Caption         =   "�ύX�O�f�[�^�ۑ�"
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
      TabIndex        =   1
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmSetteiBefore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 ALL Rights Reserved
'//
'//  �t�@�C����  �F�ύX�O�f�[�^�ۑ�.frm
'//  �p�b�P�[�W���F�ύX�O�f�[�^�ۑ��̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-10   CODED   BY [TCC] T.Nakajima
'//                 �E�W���ݒ�ۑ�����(frmRenewData.frm)�𗬗p
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SET_BEF_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    glbSaveFoldePath = ""
    strPath = ""
    
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
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-22  REVISED BY [TCC] T.Nakajima
'//                 2016�N�x�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdKeep_Click()
    
    Dim iResponse       As Integer      'MsgBox�߂�l
    Dim strMessage      As String       'MsgBox����
    Dim bRet            As Boolean      '�߂�l
    Dim intCount        As Integer
    Dim udtSendData     As MAIL_KAKARIIN_SETTEI
    Dim lngRet          As Long
    Dim blnIsSelected   As Boolean
    Dim intGokiCount    As Integer      '���̃R�[�i�̍��@��
    Dim intComSts       As Integer      '���̎����̒ʐM���
    Dim bResult         As Boolean
    
    Dim fso             As FileSystemObject
    
    intComSts = False
    
    On Error Resume Next

    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SET_BEF_GAMEN_SAVE, 0)
    
    '�R�[�i�ʂ̍��@�����擾����
    Call gsGetGateInfo

    Erase glngTergetCorner
    blnIsSelected = False
    For intCount = 0 To RenewChk.UBound
        If RenewChk(intCount).Visible = True And RenewChk(intCount).Value = CMN_ONOFF.CMN_ON Then
            blnIsSelected = True
            If RenewChk(intCount).Visible = True Then
                glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON
                '���̃R�[�i�ɑ�������D�@�̒ʐM��Ԃ��`�F�b�N����
                '�Ď��ՋN���L���`�F�b�N
                If CheckAppStart(PROC_KANRI) <> 0 Then
                    For intGokiCount = 0 To gudtSettiCorner(intCount).intGokiNum - 1
                        gpfGetjikaiConectSts intComSts, gudtSettiCorner(intCount).intGateNo(intGokiCount)
                        If intComSts <> CONECTSTS_NORMAL Then
                            Exit For
                        End If
                    Next
                    '1��ł��ʐM�ُ�̉��D�@������΁A�x����\������̂ŁA�R�[�i�P�ʂ̃��[�v�𔲂���
                    If intComSts <> CONECTSTS_NORMAL Then
                        Exit For
                    End If
                Else
                    '�ێ�P�ƋN���̏ꍇ�͉��D�@�ێ�ݒ�f�[�^���ŐV�ł͖������Ƃ�ʒm����
                    Exit For
                End If
            End If
        End If
    Next intCount
    '�I���R�[�i�Ȃ��̏ꍇ�A���b�Z�[�W��\�����ďI��
    If blnIsSelected = False Then
        iResponse = MsgBox("�ΏۃR�[�i���I������Ă��܂���B" & vbCrLf & _
                            "�I�����Ă��������B", vbOKOnly + vbCritical, "�R�[�i���I��")
        Exit Sub
    End If
    
    'EG30 V32.1.0.1 ADD START
    If intComSts <> CONECTSTS_NORMAL Then
        iResponse = MsgBox("�I�������R�[�i�ɒʐM�ُ�̉��D�@������܂��B" & vbCrLf & _
                            "�ʐM�ُ퍆�@�̉��D�@�ێ�ݒ�f�[�^�͍ŐV�Ŗ����\��������܂��B", _
                            vbOKOnly + vbExclamation, "�ʐM�ُ���D�@�L��")
    End If
    'EG30 V32.1.0.1 ADD END
    
    '��ʂ����b�N����
    cmdKeep.Enabled = False
    cmdReturn.Enabled = False
    
    '�ύX�O�f�[�^�ۑ�
    pfCopySetteiFiles bResult
    
    If bResult = False Then
        iResponse = MsgBox("�ۑ��Ɏ��s�������ڂ�����܂��B", vbOKOnly + vbExclamation, "�ۑ����s")
    Else
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "����I��")
    End If
    '��ʃ��b�N����
    cmdKeep.Enabled = True
    cmdReturn.Enabled = True
    
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SET_BEF_GAMEN_END, 0)
    
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
    
    Dim blnExistFile            As Boolean          '�ۑ��t�@�C���L��
    Dim strSaveFile             As String
    Dim intIndex                As Integer
    
    On Error Resume Next

    blnExistFile = False
    intIndex = 0
    For intCnt = 0 To UBound(gudtSettiCorner)
        If gblnCornerSet(intCnt) = True Then
            '�ۑ��t�@�C���̓��t���擾
            strSaveFile = PATH_OPERATE & "CORNER" & CStr(intCnt + 1) & "\\SETTEI_BEF\\" & SET_BEF_DATE_FILE
            If Dir(strSaveFile) = "" Then
                lblSetteDate(intIndex).Caption = "    �N   ��   ��   ��   ��   �b"
            Else
                '�t�@�C�����I�[�v��
                lngHandle = CreateFile(strSaveFile, GENERIC_READ, FILE_SHARE_READ, _
                                        0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)

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
                blnExistFile = True
            End If
            
            lblSetteDate(intIndex).Visible = True
            intIndex = intIndex + 1
        Else
            lblSetteDate(intIndex).Visible = False
        End If
    Next
    
    Exit Sub

APIError:

    Call CloseHandle(lngHandle)             '�n���h���̃N���[�Y

ErrorHandler:

    '���݂��Ȃ��ꍇ�܂��̓G���[�����������A�u�����N��\��
    For intCnt = intCnt To UBound(gudtSettiCorner)
        If lblSetteDate(intCnt).Visible = True Then
            lblSetteDate(intCnt).Caption = "    �N   ��   ��   ��   ��   �b"
        End If
    Next intCnt
    
End Sub

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
        AppActivate frmSetteiBefore.Caption, False
        pfFormActive (frmSetteiBefore.hwnd)
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  �֐�����  : pfCopySetteiFiles
'//  �@�\����  : �����ݒ���A���݉w�ݒ�f�[�^�A�����ێ�ݒ�f�[�^��ύX�O�ۑ��p�Ƃ���B
'//  �@�\�T�v  :
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : boolean   TRUE/FALSE ����FTrue �ُ�Ffalse
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-10   CODED   BY [TCC] T.Nakajima
'//                 2016�N�{���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub pfCopySetteiFiles(bResult As Boolean)
    Dim fso, f              As Object
    Dim i, j                As Integer
    Dim bRet                As Boolean
    Dim strSetteiBefFolder  As String       '�ύX�O�ۑ��p�t�H���_�p�X
    Dim strSetteiBefFolderZero  As String   '�ύX�O�ۑ��p�t�H���_�p�X(�R�[�i0�j
    Dim strOperateSetteiFolder  As String   '�����ݒ�t�H���_
    Dim strJpCfgPath            As String   '���@�ʐݒ�R���t�B�O�t�@�C���p�X
    Dim bIsFileExists           As Boolean
    Dim textFile                As TextStream
    Dim lngMuHandle             As Long     '�r�������p�n���h��
    Dim strMutexName            As String   '�~���[�e�b�N�X��

    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    bRet = False
    
    bResult = True
    
    '�w�s�x�f�[�^���R�[�i�O�ɃR�s�[����
    strSetteiBefFolderZero = PATH_OPERATE & "CORNER" & CStr(0) & "\\SETTEI_BEF\\"
    '�ύX�O�ۑ��p�t�H���_�̃t�@�C�������ׂč폜����
    fso.DeleteFile strSetteiBefFolderZero & "*.*", True
    
    If (CopyFile(EKI_SETTI_FILE, strSetteiBefFolderZero & fso.GetFileName(EKI_SETTI_FILE), True) = False) Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & fso.GetFileName(EKI_SETTI_FILE), 0)
        bResult = False
    End If
    '�ύX�O�ۑ��f�[�^���t���쐬����(��ɂ��̃t�@�C�����R�[�i�ʂɃR�s�[����)
    Set textFile = fso.CreateTextFile(strSetteiBefFolderZero & SET_BEF_DATE_FILE, True)
    textFile.Close
        
    For i = 0 To RenewChk.UBound
        '�I�����ꂽ�R�[�i
        If RenewChk(i).Visible = True And RenewChk(i).Value = CMN_ONOFF.CMN_ON Then
            strSetteiBefFolder = PATH_OPERATE & "CORNER" & CStr(i + 1) & "\\SETTEI_BEF\\"
            '�ύX�O�ۑ��p�t�H���_�̃t�@�C�������ׂč폜����
            fso.DeleteFile strSetteiBefFolder & "*.*", True
            
            '�����ݒ�����R�s�[����
            strOperateSetteiFolder = PATH_OPERATE & "CORNER" & CStr(i + 1) & "\\SETTEI\\"
            'SETTEI�t�H���_�ɉ����t�@�C����������΁A�R�s�[�����͂��Ȃ��B
            If fso.GetFolder(strOperateSetteiFolder).files.Count > 0 Then
                For Each f In fso.GetFolder(strOperateSetteiFolder).files
                    
                    If (CopyFile(strOperateSetteiFolder & f.Name, strSetteiBefFolder & f.Name, True) = False) Then
                        '�R�s�[�������݂��Ȃ��̂Ŏ擾�ł��Ȃ�����
                        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & f.Name, 0)
                        bResult = False
                        '�����ݒ���̂����ꂩ�����s�����ꍇ�͑����ݒ���̃R�s�[�𒆎~���A�t�@�C�����폜
                        fso.DeleteFile strSetteiBefFolder & "*.*", True
                        Exit For
                    End If
                Next
            Else
                '�R�s�[���t�H���_�Ƀt�@�C�����Ȃ����߁A�G���[
                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & strOperateSetteiFolder, 0)
                bResult = False
            End If
            
            '�����ێ�ݒ�f�[�^���R�s�[����
            If gudtSettiCorner(i).intGokiNum > 0 Then
                '���̃R�[�i�ɑ�������D�@�����R�s�[����
                For j = 0 To gudtSettiCorner(i).intGokiNum - 1
                    strJpCfgPath = PATH_DATA & Replace(JP_CFG, "##", Format(gudtSettiCorner(i).intGateNo(j), "0#"))
                    
                    '�r��JP_CFG�t�@�C�����쐬���̏ꍇ�͑҂�
                    strMutexName = Replace(MU_N_CFG, "##", Format(gudtSettiCorner(i).intGateNo(j), "0#"))
                    lngMuHandle = dllOpenMutex(strMutexName)
                    
                    If lngMuHandle <> 0 Then
                        dllWaitForSingleObject (lngMuHandle)
                    End If
                    
                    If (CopyFile(strJpCfgPath, strSetteiBefFolder & fso.GetFileName(strJpCfgPath), True) = False) Then
                        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_NOTFOUND & ":" & fso.GetFileName(strJpCfgPath), 0)
                        bResult = False
                    End If
                    
                    If lngMuHandle <> 0 Then
                        dllReleaseMutex (lngMuHandle)                   '�r������(FREE)
                        dllCloseHandle (lngMuHandle)                    '�r������(CLOSE)
                    End If
                    
                Next j
            Else
                ' ���̃R�[�i��1������D�@�������ꍇ
                ' ���D�@�������������݂��Ȃ��̂ŁA���D�@�ێ�ݒ�f�[�^�����݂��Ȃ����߃G���[�Ƃ��Ȃ�
            End If
            
            '�ύX�O�ۑ��f�[�^���t���쐬����i�R�[�i�O����R�s�[����j
            fso.CopyFile strSetteiBefFolderZero & SET_BEF_DATE_FILE, strSetteiBefFolder, True
        End If
    Next i
    
    Set fso = Nothing
    
End Sub

