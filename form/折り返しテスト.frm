VERSION 5.00
Begin VB.Form frmOriTest 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�܂�Ԃ��e�X�g���"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   14.25
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
   Begin VB.Timer tmrMail 
      Left            =   8760
      Top             =   7800
   End
   Begin VB.Frame fraResource 
      Caption         =   "�e�X�g��w��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9600
      TabIndex        =   3
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton optSyubetu 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optSyubetu 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdTestStart 
      Caption         =   "�e�X�g�J�n"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "     ���j���[        ��ʂ֖߂�"
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
      Left            =   9600
      TabIndex        =   2
      Top             =   7320
      Width           =   2175
   End
   Begin VB.ListBox lstKan 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�܂�Ԃ��e�X�g"
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
      Height          =   403
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�X�e�[�^�X"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�T�[�o�[��"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "����"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmOriTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'*
'*   Ӽޭ�يT�v  : �ܕԂ��e�X�g��ʂ̃t�H�[�����W���[��
'*               :�i�W�v����ь𒲂ɑ΂���󋵊m�F��ʁj
'*
'*     ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'*              �EKK�Ή�
'*     REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Option Explicit
'���\�[�X�萔

'�I�𒆃��\�[�X��� =0�F�W�v�A=1:��
Dim iSelResource As Integer


Private Const MN_MAIL_INTERVAL = 1000 '���C���^�C�}�̃C���^�[�o���l

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  �T�v     : �ܕԂ��e�X�g���ʕ\��
'  ����     : �ܕԂ��e�X�g�̌��ʂ�\�����镶�����쐬����B
'  ���Ұ�   : strMsg, I ,string, �F�e�X�g���ʕ\������
'           :  �߂�l,O ,string, �F���X�g�{�b�N�X�\������
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              �EKK�Ή�
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Function fMakeListbox(strMsg As String) As String
    Dim strRet As String
    Dim strServer As String
    
    On Error Resume Next
    
    strRet = vbNullString
    
    '�V�X�e�������擾
    strRet = Format(Now, "YYYY/MM/DD HH:MM:SS")
    
    '�T�[�o�[�^�C�v�擾
    If (iSelResource = 0) Then
        strServer = "����"
    Else
        strServer = "��"
    End If

    '���X�g�{�b�N�X��(���� �T�[�o�[�� �ُ�I��)���L�ڂ���
    strRet = strRet & Space(5) & strServer & Space(13) & strMsg

    fMakeListbox = strRet

End Function

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2001 All Right Reserved
'
'  �T�v     : �u�ێ��ʂɖ߂�v�{�^���������̃C�x���g�v���V�[�W��
'  ����     : �ܕԂ��e�X�g���ʕ\����ʂ����B
'  ���Ұ�   :
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              �EKK�Ή�
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdReturn_Click()

    On Error Resume Next

    '���O���L�ڂ��A����ʂ���������
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ORI_TEST_GAMEN_END, 0)
    frmOriTest.ZOrder
    Unload Me

End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2001 All Right Reserved
'
'   �T�v    : �e�X�g�J�n�{�^���������̃C�x���g�v���V�[�W��
'   ����    : �ܕԂ��e�X�g�J�n�v�����ă}�ɑ��M����B
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              �EKK�Ή�
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdTestStart_Click()
    
    Dim udtMail As MAIL_ORI_TEST        '���M���[��
    Dim strServer As String             '�T�[�o�[���i�[
    
    Dim bRet As Boolean
    Dim strRet As String
    
    On Error Resume Next
    
    '�e�X�g�J�n�t�������O�L��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ORI_TEST_TEST_START_BUTTOM, 0)
    
    '���W�I�t����T�[�o�^�C�v�擾
    If optSyubetu(0) = True Then
        '����
        iSelResource = 0
    Else
        '��
        iSelResource = 1
    End If
    
    '�e�X�g�J�n�����X�g�ɕ\������B
    lstKan.AddItem fMakeListbox("�J�n")
    
    '�ă}�ɑ΂��ĐܕԂ��e�X�g�J�n�v���𑗐M����B
    udtMail.mlHeader.dwId = ML_ID_ORI_TEST_REQ
    udtMail.mlHeader.dwSize = MlSize.ORI_TEST_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwSvrType = iSelResource
    bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMail), udtMail.mlHeader)
    
    '���[�����M�����`�F�b�N
    If bRet = False Then
        '���M���s��
        
        '���[�����M���s���O�L��
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, ORI_TEST_TEST_MAIL_SEND_ERR, 0)

        '�\�����͍쐬
        strRet = fMakeListbox("�ُ�I��")
        
        '���͕\��
        lstKan.AddItem strRet
    Else
        '���M������
        
        '�t������s���B
        cmdTestStart.Enabled = False
        optSyubetu(0).Enabled = False
        optSyubetu(1).Enabled = False
        cmdReturn.Enabled = False

    End If
       
    Exit Sub
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  �T�v      : �ܕԂ��e�X�g���ʕ\����ʂ��A�N�e�B�u�ɂȂ������̃C�x���g�v���V�[�W��
'  ����      : ���C����M�p�̃^�C�}���N������B
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              �EKK�Ή�
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Activate()
    '���[����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  �T�v     : �ܕԂ��e�X�g���ʕ\����ʂ��ި��è�ނɂȂ������̲������ۼ��ެ
'  ����     : ���[����M�p�̃^�C�}���~�߂�B
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              �EKK�Ή�
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Deactivate()
    '���[����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2009 All Right Reserved
'
'  �T�v     : �ܕԂ��e�X�g���ʕ\����ʖʂ����[�h���ꂽ���̲������ۼ��ެ
'  ����     : �����������s��
'  ���Ұ�   :
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              �EKK�Ή�
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Load()
    
    Dim iKansiAplChk As Integer
    
    On Error Resume Next

    '�܂�Ԃ��e�X�g��ʕ\�����O�L��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ORI_TEST_GAMEN_START, 0)

    '�e�X�g���ʕ\���p�̃��X�g�{�b�N�X���N���A����B
    lstKan.Clear
        
    '���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

    '��ʃT�C�Y
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '�Ď��ՋN��/���N���`�F�b�N���s���B
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '�Ď��ՋN����
        '�����̖t��������悤�ɂ���
        cmdTestStart.Enabled = True
        optSyubetu(0).Enabled = True
        optSyubetu(1).Enabled = True
    Else
        '�Ď����N����
        '�����̖t�������Ȃ�����
        cmdTestStart.Enabled = False
        optSyubetu(0).Enabled = False
        optSyubetu(1).Enabled = False
    End If

End Sub


'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2001 All Right Reserved
'
'  �T�v     : ���[����M�p�^�C�}���^�C���A�b�v�������̃C�x���g�v���V�[�W��
'  ����     : ��M���[���̓��e�Ɋ�Â�����������B
'  ���Ұ�   :
'
'   ORIGINAL  :(V1.10.0.1) 2009-09-25   CODED   BY [TCC] T.Furuya
'              �EKK�Ή�
'   REVISIONS :(V0.0.0.0)  0000-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub tmrMail_Timer()
    Dim lLen As Long                    '���C���T�C�Y
    Dim bRet As Boolean                 '�߂�l
    Dim udtReadMail As ML_KYOTU_INF

    On Error Resume Next

    '���[�����͂��Ă��邩�m�F����B
    lLen = DssMailRead(plMSlot_MN, udtReadMail)
    
    '��M�������[�����T�C�Y0����Ȃ���Ή�͂���
    If lLen <> 0 Then
        
        Select Case udtReadMail.udtlHeader.dwId   '���[���h�c
        
        '�u�v���Z�X�I���w���v����M�����ꍇ
        Case ML_ID_PROEND_ORD
                    
            '�����I���������s��
            pfAbortProc

        '�u�ܕԂ��e�X�g�����ʒm�v����M�����ꍇ
        Case ML_ID_ORI_TEST_INF
            
            '���ʓ��e�Ɋ�Â��������s���B
            Select Case udtReadMail.lngData(1)
                Case 0
                    '�e�X�g����I����\������B
                    lstKan.AddItem fMakeListbox("����I��")
                Case 1
                    '�e�X�g�ُ�I����\������B
                    lstKan.AddItem fMakeListbox("�ُ�I��")
                Case Else
                    '�e�X�g���s�s�\��\������B
                    lstKan.AddItem fMakeListbox("���s�s�\")
            End Select
            
            ' �{�^���������\�ɂ���B
            cmdTestStart.Enabled = True
            ' ���W�I�{�^���������\�ɂ���B
            optSyubetu(0).Enabled = True
            optSyubetu(1).Enabled = True
            ' �ێ��ʂ֖߂�t�������s�ɂ���B
            cmdReturn.Enabled = True
            

        '�ێ��ʃA�N�e�B�u�\���̏ꍇ
        Case ML_ID_HOSHU_ACTIVE_REQ

            '�ܕԂ��e�X�g��ʂ��A�N�e�B�u�ɂ���B
            AppActivate frmOriTest.Caption, False

        '(���[���h�c�s���j
        Case Else
        End Select
    End If
    
End Sub

