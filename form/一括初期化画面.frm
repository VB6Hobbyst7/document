VERSION 5.00
Begin VB.Form frmALLSysformat 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�V�X�e���������@�\(�ꊇ������)"
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
      Left            =   9840
      Top             =   3360
   End
   Begin VB.Timer tmrAplTimer 
      Left            =   9840
      Top             =   2880
   End
   Begin VB.Timer tmrMail 
      Left            =   9840
      Top             =   2280
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
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.ListBox LstStatus 
      Height          =   6360
      Left            =   240
      TabIndex        =   2
      Top             =   600
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
   Begin VB.Label lblHelp 
      Caption         =   "�E�k �c �t �F�A�v���P�[�V�������O�A�ێ�v���O�������O�A���D�@���O�A���̑��f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Left            =   120
      TabIndex        =   10
      Top             =   8300
      Width           =   8895
   End
   Begin VB.Label lblHelp 
      Caption         =   $"�ꊇ���������.frx":0000
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   7890
      Width           =   8895
   End
   Begin VB.Label lblHelp 
      Caption         =   $"�ꊇ���������.frx":0097
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   7500
      Width           =   9015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�ꊇ�V�X�e���o�׎�������"
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
      TabIndex        =   7
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblHelp 
      Caption         =   $"�ꊇ���������.frx":012C
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7050
      Width           =   8895
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
      TabIndex        =   5
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Caption         =   "����������"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "frmALLSysformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmALLSysformat.frm
'//  �p�b�P�[�W���F�ꊇ�V�X�e�����������
'/
'//  �T�v�F�ꊇ�V�X�e�����������
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//         �t�F�[�Y�Q�Ή��@�ݒ�t�@�C��(�ۑ��p�j��ǉ�
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
'//
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private bChk() As Boolean

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
'V1.5.0.1 ADD START
Private Const APL_INTERVAL = 390000     '�A�v���N���^�C�}�f�t�H���g�l
Dim lngMAX_Time As Long                    'INI�擾�ݒ�l
Dim lngtime     As Long                    '���݃^�C�}�l
'V1.5.0.1 ADD END
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        '���O�N���^�C�}�f�t�H���g�l(30�b)
Dim lngLogMAX_Time As Long                'INI�擾�ݒ�l(���O�j
'V1.20.0.1 ADD END
'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �ꊇ�V�X�e�����������(�A�N�e�B�u��)
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
'//  �@�\����  : �ꊇ�V�X�e�����������(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �ꊇ�V�X�e�����������(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.4.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   
    On Error Resume Next

    '�u�ꊇ�V�X�e���������F�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ALL_SYSFORMAT_GAMEN_START, 0)

    '�z�u�ݒ�
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

    '������
    LstStatus.Clear
    lblKekka.Caption = ""
    
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
  
   'V1.4.0.1 ADD START
   'IDU�k�ރ`�F�b�N
    psIDUCheck
   'V1.4.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdZikko_Click
'//  �@�\����  : �u���������s�v�t����������
'//  �@�\�T�v  : �V�X�e�����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-19   REVISED BY [TCC] S.Terao
'//         �t�F�[�Y�Q�Ή��@�ݒ�t�@�C��(�ۑ��p�j��ǉ�
'//     REVISIONS  :(1.5.0.1) 2009-03-19   CODED   BY [TCC] S.Terao
'//         �t�F�[�Y�P�s��Ή� �A�v���N���`�F�b�N�������C��
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                         �ێ瑍�_�����ʏC��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Terao
'//                 EG-R�Ď��Ձ@�Q���Ή��@���O�^�C�}�ǉ��A�m�F�|�b�v�A�b�v�C��
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-23  REVISED BY [TCC] T.Koyama
'//                 EG20�t�F�[�Y�Q�Ή��y�c����54�z
'//                 �E�ێ烍�O�t�@�C���b�k�n�r�d�����ǉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZikko_Click()

    Dim iRet As Integer
    Dim sDBFormat As String
    Dim sLine As String
    Dim sExecName As String
    Dim sDbInitCmd As String
    Dim bRtn1 As Boolean
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim uMail As ML_KYOTU_INF           '���[��
    Dim bRet As Boolean
    Dim iKansiApp As Integer            '�Ď��ՃA�v���N���t���O
    Dim iRetIDULog As Integer           'IDU���O�N���t���O
    Dim iRetLDULog As Integer           'IDU���O�N���t���O
    Dim bKansiDB_Code As Boolean
    Dim bIDUDB_Code As Boolean
    Dim lExitCode As Long
    Dim iTargetDB As Integer            '�Ώ�DB�l
    ReDim bChk(9)
    Dim i As Integer                    '�J�E���^�[
    Dim bRtn As Boolean
    'V1.5.0.1  ADD START
    Dim bKansiRet As Boolean            '�Ď��ՃA�v����������
    Dim bIDURet   As Boolean            'IDU�A�v����������
    Dim bLDURet   As Boolean            'LDU�A�v����������
   
    bKansiRet = False
    bIDURet = False
    bLDURet = False
    'V1.5.0.1  ADD END
    On Error GoTo ERR_SPACE

    '�u�ꊇ�V�X�e����������ʁF���������s�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_START_BUTTOM, 0)

    '�\���̏�����
    LstStatus.Clear
    lblKekka.Caption = ""
    
    iKansiApp = 1
    iRetIDULog = 1
    iRetLDULog = 1
    
    '�u�������m�F�v�|�b�v�A�b�v��\��
    iRet = MsgBox("�������������s���܂��B��낵���ł����H", vbExclamation + vbOKCancel, "�������m�F")
    If iRet = vbOK Then
         cmdZikko.Enabled = False  '�u���������s�v�t�����s��
         cmdCancel.Enabled = False '�u���j���[��ʂ֖߂�v�t�����s��

         On Error GoTo ERR_SPACE2
         
         '�Ď���(�Ǘ��v���Z�X)���N�����Ă��邩�ǂ����`�F�b�N����B
         If CheckAppStart(PROC_KANRI) <> 0 Then
           'V1.20.0.1 DEL START
           'iRet = MsgBox("�Ď��ՁA�h�c���p���j�b�g�A�k�c���[�e�B���e�B�A�v���P�[�V�������I�����܂��B" & Chr(vbKeyReturn) & _
           '               "��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F")
           'If iRet = vbOK Then
           'V1.20.0.1 DEL END
              '�A�v���I���v�����Ǘ��ɑ��M����
               uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
               uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
               uMail.udtlHeader.dwProid = RHOSHU_ID
               uMail.udtlHeader.dwSubArea = 0
               'V1.5.0.1 DEL START
               'bRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               'If bRet = 0 Then
               'V1.5.0.1 DEL END
               'V1.5.0.1 ADD START
               bKansiRet = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
               If bKansiRet = 0 Then
                  'V1.5.0.1 ADD END
                  '�u�ꊇ�V�X�e����������ʁF���[�����M�ُ�v���O�o��
                  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                  Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lngErrCode)
                  GoTo ERR_SPACE2:
               Else
                  '�u�ꊇ�V�X�e����������ʁF���[�����M����v���O�o��
                  Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                  '�Ď��ՃA�v���I���m�F
                  'iKansiApp = CheckAppEndComplete(PROC_KANRI, lExitCode)            'V1.5.0.1 DEL
               End If

           'V1.20.0.1 DEL START
'
'               '���O�v���Z�X�N���`�F�b�N
'               If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
'
'                  'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
'                  iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
'                  If iRet = vbOK Then
'                     'IDU���O�I���v��CMD���M
'                     'V1.5.0.1 DEL START
'                     'bRtn = EndIDULog
'                     'If bRtn = False Then
'                     'V1.5.0.1 DEL END
'                     'V1.5.0.1 ADD START
'                     bIDURet = EndIDULog
'                     If bIDURet = False Then
'                     'V1.5.0.1 ADD END
'                        '���M�ُ폈��
'                        lblKekka.ForeColor = SYSFORMAT_ERROR
'                        lblKekka.Caption = "�������Ɏ��s���܂���"
'                        cmdZikko.Enabled = True
'                        cmdCancel.Enabled = True
'                        '�����𔲂���
'                        Exit Sub
'                     End If
'
'                     'IDU���O�v���Z�X�I���m�F
'                     'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) 'V1.5.0.1 DEL
'                  'V1.7.0.1 ADD START
'                  Else
'                    '���O�v���Z�X�I�����b�Z�[�W�u�L�����Z���v�t����������
'                    GoTo ERR_SPACE3
'                  'V1.7.0.1 ADD END
'                  End If
'               'V1.5.0.1 ADD START
'               Else
'                 bIDURet = True
'               'V1.5.0.1 ADD END
'               End If
'
'               'LDU���O�v���Z�X�N���`�F�b�N
'               If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
'
'                  'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
'                  iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
'
'                  If iRet = vbOK Then
'                     'LDU���O�I���v��CMD���M
'                     'V1.5.0.1 DEL START
'                     'bRtn = EndLDULog
'                     'If bRtn = False Then
'                     'V1.5.0.1 DEL END
'                     'V1.5.0.1 ADD START
'                     bLDURet = EndLDULog
'                     If bLDURet = False Then
'                     'V1.5.0.1 ADD END
'                        '���M�ُ폈��
'                        lblKekka.ForeColor = SYSFORMAT_ERROR
'                        lblKekka.Caption = "�������Ɏ��s���܂���"
'                        cmdZikko.Enabled = True
'                        cmdCancel.Enabled = True
'                        '�����𔲂���
'                        Exit Sub
'                     End If
'
'                     'LDU���O�v���Z�X�I���m�F
'                     'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode) 'V1.5.0.1 DEL
'                  'V1.7.0.1 ADD START
'                  Else
'                    '���O�v���Z�X�I�����b�Z�[�W�u�L�����Z���v�t����������
'                    GoTo ERR_SPACE3
'                  'V1.7.0.1 ADD END
'                  End If
'               'V1.5.0.1 ADD START
'                   Else
'                bLDURet = True
'               'V1.5.0.1 ADD END
'               End If
'           Else
'             '�u�L�����Z���t�����v
'             '�u�ꊇ�V�X�e����������ʁF���������������s�v���O�o��
'             Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
'             cmdZikko.Enabled = True  '�u���������s�v�t������
'             cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
'             '�����𔲂���
'             Exit Sub
'           End If
'        'V1.5.0.1 ADD START
        'V1.20.0.1 DEL END
         Else
            bKansiRet = True
        'V1.5.0.1 ADD END
        ' End If  'V1.5.0.1 DEL
          If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
             
             'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
             'V1.20.0.1 DEL START
             'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
             'If iRet = vbOK Then
             'V1.20.0.1 DEL END
             
                'IDU/LDU���O�I���v��CMD���M
                'V1.5.0.1 DEL START
                'bRet = EndIDULog
                'If bRet = False Then
                'V1.5.0.1 DEL END
                'V1.5.0.1 ADD START
                bIDURet = EndIDULog
                If bIDURet = False Then
                'V1.5.0.1 ADD END
                 '���M�ُ�
                 lblKekka.ForeColor = SYSFORMAT_ERROR
                 lblKekka.Caption = "�������Ɏ��s���܂���"
                 cmdZikko.Enabled = True
                 cmdCancel.Enabled = True
                 '�����𔲂���
                 Exit Sub
               End If
               
               'IDU���O�v���Z�X�I���m�F
               'iRetIDULog = CheckAppEndComplete(PROCESS_IDU_LOG, lExitCode) 'V1.5.0.1 DEL
             'V1.7.0.1 ADD START
          'V1.20.0.1 DEL START
'             Else
'              '���O�v���Z�X�I�����b�Z�[�W�u�L�����Z���v�t����������
'              GoTo ERR_SPACE3
'             'V1.7.0.1 ADD END
'             End If
'         'V1.5.0.1 ADD START
          'V1.20.0.1 DEL END
          Else
            bIDURet = True
          'V1.5.0.1 ADD END
          End If
         
          If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
             
             'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "�I���m�F") 'V1.8.0.1 DEL
             'V1.20.0.1 DEL START
             'iRet = MsgBox("���O�v���Z�X���I�����܂��B��낵���ł����H", vbQuestion + vbOKCancel, "���O�I���m�F")  'V1.8.0.1 ADD
             '
             'If iRet = vbOK Then
             'V1.20.0.1 DEL END
             'IDU/LDU���O�I���v��CMD���M
             'V1.5.0.1 DEL START
             'bRet = EndLDULog
             'If bRet = False Then
             'V1.5.0.1 DEL END
             'V1.5.0.1 ADD START
             bLDURet = EndLDULog
             If bLDURet = False Then
             'V1.5.0.1 ADD END
                '���M�ُ�
                lblKekka.ForeColor = SYSFORMAT_ERROR
                lblKekka.Caption = "�������Ɏ��s���܂���"
                cmdZikko.Enabled = True
                cmdCancel.Enabled = True
                '�����𔲂���
                Exit Sub
              End If
              
              'LDU���O�v���Z�X�I���m�F
              'iRetLDULog = CheckAppEndComplete(PROCESS_LDU_LOG, lExitCode)  'V1.5.0.1 DEL
            'V1.7.0.1 ADD START
          'V1.20.0.1 DEL START
'             Else
'              '���O�v���Z�X�I�����b�Z�[�W�u�L�����Z���v�t����������
'              GoTo ERR_SPACE3
'            'V1.7.0.1 ADD END
'             End If
'         'V1.5.0.1 ADD START
          'V1.20.0.1 DEL END
          Else
            bLDURet = True
          'V1.5.0.1 ADD END
         End If
       End If        'V1.5.0.1 ADD
'V1.5.0.1 ADD START
       '�Ď��ՁAIDU�ALDU�A�v���̃��[�����M�������S�Đ��킾�����ꍇ�̂݁A�A�v���N���^�C�}���N�������A
       '�A�v���N���`�F�b�N�ɂ��A�v���̋N��/���N���𔻒f����B
       'If (bKansiRet = True) And (bIDURet = True) And (bLDURet = True) Then  'V1.20.0.1 DEL
       If (bKansiRet = True) Then  'V1.20.0.1 ADD
           lngtime = 0
           lngtime = MN_MAIL_INTERVAL
           tmrAplTimer.Enabled = True
       Else
          '�Ď��ՁAIDU�ALDU�A�v���̃��[�����M�ɂĂЂƂł��ُ킪�������ꍇ�A�������������ُ�I���Ƃ���B
          '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "�������Ɏ��s���܂���"
          cmdZikko.Enabled = True
          cmdCancel.Enabled = True
          '�����𔲂���
          Exit Sub
       End If
   End If
'V1.5.0.1 ADD END
'V1.5.0.1 DEL START
'         '�A�v���܂��̓��O�v���Z�X�ŏI�������Ɏ��s�����ꍇ
'         If (iKansiApp <> 1) Or (iRetIDULog <> 1) Or (iRetLDULog <> 1) Then
'            '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "�������Ɏ��s���܂���"
'           cmdZikko.Enabled = True
'           cmdCancel.Enabled = True
'           '�����𔲂���
'           Exit Sub
'         End If
'
'        'V1.4.0.1 ADD START
'        If sCreateShokiFile = False Then
'           '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'           lblKekka.ForeColor = SYSFORMAT_ERROR
'           lblKekka.Caption = "�������Ɏ��s���܂���"
'           cmdZikko.Enabled = True
'           cmdCancel.Enabled = True
'           '�����𔲂���
'           Exit Sub
'        End If
'        'V1.4.0.1 ADD END
'
'        '�V�X�e���t�@�C���̍폜����
'        bRtn1 = sSysFileDelete()
'
'        '�V�X�e���t�@�C���폜�������������ꍇ�A
'        '�t�H���_�A�t�@�C���̍폜�������s��
'        If bRtn1 = True Then
'
'            '�Ď��ՃV�X�e��������
'            For i = 1 To 6
'               bChk(i) = True
'            Next
'            bChk(5) = False
'
'            If sFileDelete(stsKansi, KANSI_SYSTEMFILE) = False Then
'              '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "�������Ɏ��s���܂���"
'              cmdZikko.Enabled = True  '�u���������s�v�t������
'              cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
'              Exit Sub
'           End If
'
'           'IDU�V�X�e��������
'           For i = 2 To 8
'               bChk(i) = True
'           Next
'           bChk(1) = False
'           If sFileDelete(stsIDU, PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE) = False Then
'              '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "�������Ɏ��s���܂���"
'              cmdZikko.Enabled = True  '�u���������s�v�t������
'              cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
'              Exit Sub
'           End If
'
'           'LDU�V�X�e��������
'           For i = 2 To 9
'               bChk(i) = True
'           Next
'           bChk(1) = False
'           If sFileDelete(stsLDU, PATH_LDU_APP & PATH_LDU_DATA & PATH_LDU_SYSTEMFILE) = False Then
'              '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'              Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'              lblKekka.ForeColor = SYSFORMAT_ERROR
'              lblKekka.Caption = "�������Ɏ��s���܂���"
'              cmdZikko.Enabled = True  '�u���������s�v�t������
'              cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
'              Exit Sub
'           End If
'
'           '�Ď��ՁF�ꌏ����
'           Me.LstStatus.AddItem "DB������:" & "�W�v�֘A�f�[�^"
'           DoEvents
'           iTargetDB = stsKansiMeisai
'           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'           DoEvents
'           Me.Refresh
'           If bKansiDB_Code = True Then
'                '�Ď��ՁF�ʏW�D
'                iTargetDB = stsKansiBetu
'                '�Ď���DB����������
'                bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'
'           'IDUDB����������
'            Me.LstStatus.AddItem "DB������:" & "DB�f�[�^"
'            DoEvents
'            Me.Refresh
'
'           'IDUDB����������
'           'IDU:DB�f�[�^
'           iTargetDB = stsIDUMeisai
'           bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'           DoEvents
'           Me.Refresh
'           If bIDUDB_Code = True Then
'                Me.LstStatus.AddItem "DB������:" & "�A�v���P�[�V�������O"
'                DoEvents
'                Me.Refresh
'                'IDU�F�A�v���P�[�V�������O
'                iTargetDB = stsIDUAPLlog
'                'IDU�F�A�v��DB����������
'               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'               DoEvents
'               Me.Refresh
'           End If
'           If bIDUDB_Code = True Then
'                Me.LstStatus.AddItem "DB������:" & "�ێ�v���O����"
'                DoEvents
'                Me.Refresh
'                'IDU�F�ێ烍�O
'                iTargetDB = stsIDUMentelog
'                'IDU�F�ێ�DB����������
'                bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'           If bIDUDB_Code = True Then
'                Me.LstStatus.AddItem "DB������:" & "����IC���W���[�����O"
'                DoEvents
'                Me.Refresh
'                'IDU�F����IC-M���W���[�����O
'                iTargetDB = stsIDUICM
'                'IDU�F����IC-MDB����������
'                bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'           If bIDUDB_Code = True Then
'                'IDU�F�l�K���X�g
'                iTargetDB = stsIDUNega
'                'IDU�F�l�K���X�gDB����������
'                bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
'                DoEvents
'                Me.Refresh
'           End If
'
'           If bKansiDB_Code = True And bIDUDB_Code = True Then
'                '�u�ꊇ�V�X�e����������ʁF�V�X�e����������������v���O�o��
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
'                lblKekka.ForeColor = SYSFORMAT_OK
'                lblKekka.Caption = "�������͐������܂���"
'           Else
'                '�u�ꊇ�V�X�e����������ʁFDB�����������ُ�v���O�o��
'                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
'                lblKekka.ForeColor = SYSFORMAT_ERROR
'                lblKekka.Caption = "�������Ɏ��s���܂���"
'           End If
'       Else
'         '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
'         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
'         lblKekka.ForeColor = SYSFORMAT_ERROR
'         lblKekka.Caption = "�������Ɏ��s���܂���"
'       End If
'  End If
'
'  '�����������I��
'  cmdZikko.Enabled = True  '�u���������s�v�t������
'  cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
'V1.5.0.1 DEL END
Exit Sub

'V1.7.0.1 ADD START
ERR_SPACE3:
 '�u�L�����Z���t�����v
 '�u�ꊇ�V�X�e����������ʁF���������������s�v���O�o��
 Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_NOT_START, 0)
 cmdZikko.Enabled = True  '�u���������s�v�t������
 cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
 '�����𔲂���
 Exit Sub
'V1.7.0.1 ADD END

ERR_SPACE2:
  '�G���[�������̏���
  cmdZikko.Enabled = True  '�u���������s�v�t������
  cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
  '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
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

    '�u�ꊇ�V�X�e����������ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, ALL_SYSFORMAT_GAMEN_END, 0)
    frmALLSysformat.ZOrder
    Unload Me
End Sub

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
'//         �t�F�[�Y�Q�Ή��@�ݒ�t�@�C��(�ۑ��p�j��ǉ�
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
    
    '///////////////////////////////////////////////////////////
    'IDU�k�ރ`�F�b�N��IDU�t�@�C���֘A�̕ۑ��p�t�@�C�����쐬����B
    '///////////////////////////////////////////////////////////
    '�t�@�C���L���`�F�b�N
    If pbIDUSts = 1 Then
       sCreateShokiFile = True
       '�u�ꊇ�V�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬����v���O�o��
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
    '�u�ꊇ�V�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, SHOKI_CREATE_OK, 0)
    
    Exit Function

ERR_SPACE:
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�u�ꊇ�V�X�e����������ʁF�ۑ��p�ݒ�t�@�C���쐬�ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, SHOKI_CREATE_ERROR, lngErrCode)
    sCreateShokiFile = False
End Function
'V1.4.0.1�@ADD END

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
   Dim iRet As Integer           '�폜�����߂�l
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
    '�u�ꊇ�V�X�e����������ʁF�V�X�e���t�@�C���폜�ُ�v���O�o��
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
Private Function sFileDelete(iKikitype As Integer, Sys_FilePath As String)
    Dim iFileNo As Integer
    Dim sFileData As String
    Dim iMozi, iKbn As Integer
    Dim sShubetu As String
    Dim sRoot As String
    Dim sPass As String
    Dim sType As String
    Dim sKomoku As String
    Dim bSyori As Boolean
    Dim fs As Object
    Dim MyName As String
    Dim i As Integer
    Dim sChkPass As String
    Dim iRet As Integer
    Dim sDeletePath As String    '�폜�Ώۃt���p�X
    Dim lngErrCode As Long       '�G���[�R�[�h
    Dim lBool As Boolean                      ' EG20 V3.3.0.1�y����TR-240�z�ǉ�

    sFileDelete = False

    On Error GoTo ERR_SPACE
    
    '�t�@�C���L���`�F�b�N
    MyName = Dir(Sys_FilePath, vbNormal)
    If MyName = "" Then
        GoTo ERR_SPACE
    End If

' EG20 V3.3.0.1�y����TR-240�z�ǉ��J�n�i�ʒu�ړ��j
    ' �ێ烍�O�t�@�C��CLOSE
    lBool = dllCloseHoshuLogFile()
' EG20 V3.3.0.1�y����TR-240�z�ǉ��I���i�ʒu�ړ��j

    iFileNo = FreeFile                                              '���g�p�̃t�@�C���ԍ����擾����B
    Open Sys_FilePath For Input As #iFileNo                         '�V�X�e���������ݒ�t�@�C�����J���B
    Line Input #iFileNo, sFileData                                  ' �P�s�ڂ͑S�̃o�[�W�����Ȃ̂œǔ�΂��B
    Do While Not EOF(iFileNo)
        Line Input #iFileNo, sFileData                              ' �P�s���Ǎ��ށB
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
            Select Case iKikitype
                Case stsKansi
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
                
                Case stsIDU
                     Select Case sRoot
                        Case 1
                          '�A�v�����[�g
                          sPass = PATH_IDU_APP & "\\" & sPass
                        Case 2
                          '�o�b�N�A�b�v���[�g
                          sPass = PATH_BUC & "\\" & sPass
                        Case 4
                          '���O���[�g
                          sPass = PATH_IDU_LOG & "\\" & sPass
                     End Select
                     
                Case stsLDU
                   Select Case sRoot
                      Case 1
                        '�A�v�����[�g
                        sPass = PATH_LDU_APP & "\\" & sPass
                      Case 4
                        '���O���[�g
                         sPass = PATH_LDU_LOG & "\\" & sPass
                    End Select
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
                          DoEvents      'V1.5.0.1�@ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
                      '�t�H���_�̍폜�^�쐬
                      Case 3:
                          fs.DeleteFolder (sPass), True
                          fs.CreateFolder (sPass)
                          LstStatus.AddItem "�폜�^�쐬�����t�H���_ - " & sPass
                          DoEvents      'V1.5.0.1�@ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                  End Select
                '�I�u�W�F�N�g���
                Set fs = Nothing
            Else
                '�w��o�`�r�r�i�V
                Select Case sShubetu
                   Case 2:
                       LstStatus.AddItem "�w��t�@�C���Ȃ� - " & sPass
                       DoEvents      'V1.5.0.1�@ADD
                       LstStatus.Selected(LstStatus.ListCount - 1) = True           'V1.12.0.1 ADD
                   Case 3:
                       Set fs = CreateObject("Scripting.FileSystemObject")
                       '�t�@�C���L���`�F�b�N
'                       For i = 0 To Len(sPass)          'EG20 V2.1.0.1 DEL �y��-313�Ή��z
                       For i = 0 To Len(sPass) - 1       'EG20 V2.1.0.1 ADD �y��-313�Ή��z
                           If Mid(sPass, Len(sPass) - i, 1) = "\" Then
                               sChkPass = Left(sPass, Len(sPass) - i - 1)
                               Exit For
                           End If
                       Next
                       MyName = Dir(sChkPass, vbDirectory)
                       If MyName = "" Then
                           LstStatus.AddItem "�t�H���_�쐬���s - " & sPass
                          DoEvents      'V1.5.0.1�@ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
                       Else
                           fs.CreateFolder (sPass)
                           LstStatus.AddItem "�쐬�����t�H���_ - " & sPass
                          DoEvents      'V1.5.0.1�@ADD
                          LstStatus.Selected(LstStatus.ListCount - 1) = True        'V1.12.0.1 ADD
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
    '�u�ꊇ�V�X�e����������ʁF�t�@�C���E�t�H���_�������ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TARGET_FILE_FOLDER_DELETE_ERROR, lngErrCode)
    Set fs = Nothing
End Function

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
        AppActivate frmALLSysformat.Caption, False
        pfFormActive (frmALLSysformat.hwnd)
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
    'If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then�@'V1.20.0.1 DEL
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
         cmdZikko.Enabled = True
         cmdCancel.Enabled = True
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
    '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
    lblKekka.ForeColor = SYSFORMAT_ERROR
    lblKekka.Caption = "�������Ɏ��s���܂���"
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
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
'//    REVISIONS :(EG20 V2.0.1.1) 2011-11-23  REVISED BY [TCC] T.Koyama
'//               �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�z
'//               �E�ێ烍�O�t�@�C���b�k�n�r�d�����ǉ�
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-240�z
'//    REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()

  Dim lBool As Boolean                      ' EG20 V2.0.1.1�y�c����54�zADD

  On Error Resume Next

  '�҂����Ԃ�INI��`�𒴂������ǂ����`�F�b�N
  If lngtime <= lngLogMAX_Time Then
    '���O�N���`�F�b�N���s���B�S�ďI�������Ƃ��̂݁A�������������s���B
    If CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
      '���O�N���`�F�b�N�^�C�}���~����B
      tmrLogTimer.Enabled = False
      
' EG20 V3.3.0.1�y����TR-240�z�폜�J�n�i�ʒu�ړ��j
'      ' EG20 V2.0.1.1�y�c����54�zADD START
'      ' �ێ烍�O�t�@�C��CLOSE
'       lBool = dllCloseHoshuLogFile()
'      ' EG20 V2.0.1.1�y�c����54�zADD START
' EG20 V3.3.0.1�y����TR-240�z�폜�I���i�ʒu�ړ��j
      
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
    cmdZikko.Enabled = True
    cmdCancel.Enabled = True
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
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-240�z
'//     REVISIONS :(EG20 V7.5.0.1) 2013-12-07  CODED BY  [TCC] H.Kondoh
'//                 �ő�ڑ������e���͈͊m�F�s��Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub DeleteFile_Folder()
    
    Dim iRet As Integer
    Dim sDBFormat As String
    Dim sLine As String
    Dim sExecName As String
    Dim sDbInitCmd As String
    Dim bRtn1 As Boolean
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim bRet As Boolean
    Dim bKansiDB_Code As Boolean
    Dim bIDUDB_Code As Boolean
    Dim lExitCode As Long
    Dim iTargetDB As Integer            '�Ώ�DB�l
    ReDim bChk(9)
    Dim i As Integer                    '�J�E���^�[
    Dim bRtn As Boolean
    'EG20 V2.1.0.1 ADD START �y��-313�Ή��z
    Dim intLoop As Integer
    Dim lSts As Long
    'EG20 V2.1.0.1 ADD END
         
    
    On Error GoTo ERR_SPACE
   
    '�Ď��ՁAIDU�A�v���A�e�ݒ�t�@�C��(�ۑ��p)�쐬����
    If sCreateShokiFile = False Then
       '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
        lblKekka.ForeColor = SYSFORMAT_ERROR
        lblKekka.Caption = "�������Ɏ��s���܂���"
        cmdZikko.Enabled = True
        cmdCancel.Enabled = True
        '�����𔲂���
         Exit Sub
    End If
   
    '�V�X�e���t�@�C���̍폜����
    bRtn1 = sSysFileDelete()
     
    '�V�X�e���t�@�C���폜�������������ꍇ�A
    '�t�H���_�A�t�@�C���̍폜�������s��
    If bRtn1 = True Then

      '�Ď��ՃV�X�e��������
      For i = 1 To 6
          bChk(i) = True
      Next

'      bChk(5) = False                  ' EG20 V3.3.0.1�y����TR-240�z�폜
      bChk(6) = False                   ' EG20 V3.3.0.1�y����TR-240�z�ǉ�

      If sFileDelete(stsKansi, KANSI_SYSTEMFILE) = False Then
         '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
         Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
         lblKekka.ForeColor = SYSFORMAT_ERROR
         lblKekka.Caption = "�������Ɏ��s���܂���"
         cmdZikko.Enabled = True  '�u���������s�v�t������
         cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
         Exit Sub
       End If
           
       'IDU�V�X�e��������
       For i = 2 To 8
           bChk(i) = True
       Next

       bChk(1) = False

       If sFileDelete(stsIDU, PATH_IDU_APP & PATH_IDU_DATA & PATH_IDU_SYSTEMFILE) = False Then
          '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "�������Ɏ��s���܂���"
          cmdZikko.Enabled = True  '�u���������s�v�t������
          cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
          Exit Sub
       End If
           
       'LDU�V�X�e��������
       For i = 2 To 9
           bChk(i) = True
       Next

       bChk(1) = False

       If sFileDelete(stsLDU, PATH_LDU_APP & PATH_LDU_DATA & PATH_LDU_SYSTEMFILE) = False Then
          '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
          lblKekka.ForeColor = SYSFORMAT_ERROR
          lblKekka.Caption = "�������Ɏ��s���܂���"
          cmdZikko.Enabled = True  '�u���������s�v�t������
          cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
          Exit Sub
       End If
 
       '�Ď��ՁF�ꌏ����
        Me.LstStatus.AddItem "DB������:" & "�W�v�֘A�f�[�^"
        DoEvents
        LstStatus.Selected(LstStatus.ListCount - 1) = True              'V1.12.0.1 ADD
        
        '�Ď��ՁF�ꌏ����
        Me.LstStatus.AddItem "�ꌏ���׃R�[�i�P�@DB�������J�n"
        DoEvents
        iTargetDB = stsKansiMeisai
        bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
        Me.LstStatus.AddItem "�ꌏ���׃R�[�i�P�@DB�������I��"
        DoEvents
        
        If bKansiDB_Code = True Then
           '�Ď��ՁF�ꌏ���ׁi�R�[�i�Q�j
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�Q�@DB�������J�n"
           DoEvents
           iTargetDB = stsKansiMeisai2
           'DB����������
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�Q�@DB�������I��"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '�Ď��ՁF�ꌏ���ׁi�R�[�i�R�j
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�R�@DB�������J�n"
           DoEvents
           iTargetDB = stsKansiMeisai3
           'DB����������
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�R�@DB�������I��"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '�Ď��ՁF�ꌏ���ׁi�R�[�i�S�j
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�S�@DB�������J�n"
           DoEvents
'           bKansiDB_Code = stsKansiMeisai4     'EG20 V7.5.0.1 DEL
           iTargetDB = stsKansiMeisai4          'EG20 V7.5.0.1 ADD
           'DB����������
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�S�@DB�������I��"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '�Ď��ՁF�ꌏ���ׁi�R�[�i�T�j
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�T�@DB�������J�n"
           DoEvents
           iTargetDB = stsKansiMeisai5
           'DB����������
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�T�@DB�������I��"
           DoEvents
        End If

        If bKansiDB_Code = True Then
           '�Ď��ՁF�ꌏ���ׁi�R�[�i�U�j
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�U�@DB�������J�n"
           DoEvents
           iTargetDB = stsKansiMeisai6
           'DB����������
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
           Me.LstStatus.AddItem "�ꌏ���׃R�[�i�U�@DB�������I��"
           DoEvents
        End If
            
        If bKansiDB_Code = True Then
           '�Ď��ՁF�ʏW�D
           Me.LstStatus.AddItem "�ʏW�D�@DB�������J�n"
           DoEvents
           iTargetDB = stsKansiBetu
           '�Ď���DB����������
           bKansiDB_Code = DB_format(iTargetDB, stsKansi, Me.LstStatus)
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
            
        bIDUDB_Code = False
        
        If bKansiDB_Code = True Then
            'IDUDB����������
            Me.LstStatus.AddItem "DB������:" & "DB�f�[�^"
            DoEvents
            LstStatus.Selected(LstStatus.ListCount - 1) = True          'V1.12.0.1 ADD
           
            'IDUDB����������
            'IDU:DB�f�[�^
            iTargetDB = stsIDUMeisai
            bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
            DoEvents
            If bIDUDB_Code = True Then
               Me.LstStatus.AddItem "DB������:" & "�A�v���P�[�V�������O"
               DoEvents
               LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
               'IDU�F�A�v���P�[�V�������O
               iTargetDB = stsIDUAPLlog
               'IDU�F�A�v��DB����������
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
            If bIDUDB_Code = True Then
               Me.LstStatus.AddItem "DB������:" & "�ێ�v���O����"
               DoEvents
               LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
               'IDU�F�ێ烍�O
               iTargetDB = stsIDUMentelog
               'IDU�F�ێ�DB����������
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
            If bIDUDB_Code = True Then
               Me.LstStatus.AddItem "DB������:" & "����IC���W���[�����O"
               DoEvents
               LstStatus.Selected(LstStatus.ListCount - 1) = True       'V1.12.0.1 ADD
               'IDU�F����IC-M���W���[�����O
               iTargetDB = stsIDUICM
               'IDU�F����IC-MDB����������
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
            If bIDUDB_Code = True Then
               'IDU�F�l�K���X�g
               iTargetDB = stsIDUNega
               'IDU�F�l�K���X�gDB����������
               bIDUDB_Code = DB_format(iTargetDB, stsIDU, Me.LstStatus)
               DoEvents
            End If
        End If
        If bKansiDB_Code = True And bIDUDB_Code = True Then
           '�u�ꊇ�V�X�e����������ʁF�V�X�e����������������v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSFORMAT_END_OK, 0)
           lblKekka.ForeColor = SYSFORMAT_OK
           lblKekka.Caption = "�������͐������܂���"
        Else
           '�u�ꊇ�V�X�e����������ʁFDB�����������ُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, DBFORMAT_ERROR, 0)
           lblKekka.ForeColor = SYSFORMAT_ERROR
           lblKekka.Caption = "�������Ɏ��s���܂���"
        End If
    Else
     '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
     lblKekka.ForeColor = SYSFORMAT_ERROR
     lblKekka.Caption = "�������Ɏ��s���܂���"
    End If
 
  '�����������I��
  cmdZikko.Enabled = True  '�u���������s�v�t������
  cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
  
Exit Sub

ERR_SPACE2:
  '�G���[�������̏���
  cmdZikko.Enabled = True  '�u���������s�v�t������
  cmdCancel.Enabled = True '�u���j���[��ʂ֖߂�v�t������
  '�u�ꊇ�V�X�e����������ʁF�V�X�e�������������ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, SYSFORMAT_END_ERROR, 0)
  lblKekka.ForeColor = SYSFORMAT_ERROR
  lblKekka.Caption = "�������Ɏ��s���܂���"
ERR_SPACE:

End Sub
'V1.5.0.1 ADD�@END
