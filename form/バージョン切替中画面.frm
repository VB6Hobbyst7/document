VERSION 5.00
Begin VB.Form frmChangeVer 
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Timer tmrAplCheck 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�n �j"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmChangeVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmChangeVer.frm
'//  �p�b�P�[�W���F�o�[�W�����ؑ֒����
'//
'//  �T�v�F�o�[�W�����ؑ֒����
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//         �t�F�[�Y�Q�Ή��@�ؑ֒���ʒǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

Private iChangeSts As Integer           '�����ԍ�
Private iChangeVerFlag As Integer       '�����t���O
Private iChengeVerApl As Integer        '�Ď���=0�AIDU=1
Private bChangeVerSts As Boolean        '�ؑ֏����S�̖̂߂�l

'�t�H���_�\����������X�e�[�^�X
Private Const DLLFILE_AtoC = 1          '�o�[�W�����P���ꎞ�t�H���_
Private Const DLLFILE_BtoA = 2          '�o�[�W�����Q���o�[�W�����P
Private Const DLLFILE_CtoB = 3          '�ꎞ�t�H���_���o�[�W�����Q
Private Const PARA_BtoA = 4             '�o�[�W�����Q���o�[�W�����P(�p�����t�H���_)
Private Const BACK_CtoA = 5             '�ꎞ�t�H���_���o�[�W�����P
Private Const BACK_BtoC = 6             '�o�[�W�����Q���ꎞ�t�H���_
Private Const BACK_AtoB = 7             '�o�[�W�����P���o�[�W�����Q

'�t�@�C���p�X
Private DllFolderName As String         '�o�[�W�����P(�{��)��
Private DllFolderName2 As String        '�o�[�W�����Q(�ۑ��p)��
Private DllFolderName3 As String        '�ꎞ�t�H���_��
Private ParaFolderName1 As String       '�o�[�W�����P�p����
Private ParaFolderName2 As String       '�o�[�W�����Q�p����

'�t���O�l
Private Const CHANGE_END = 0            '�����I��
Private Const CHANGE_CONTINU = 1        '�������s
Private Const CHANGE_RENAME_ERROR = 2   '���l�[���ُ�
Private Const CHANGE_OK = 3             '��������
'V1.6.0.1 ADD START
Public lngMAX_Time As Long                    'INI�擾�ݒ�l
Public lngtime     As Long                    '���݃^�C�}�l
'V1.6.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����ؑ֒����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
  Dim bRet As Boolean
  
  On Error Resume Next
  
  bRet = True
  
  bChangeVerSts = True
 
  tmrMail.Enabled = True
  
  '�A�v���N���m�F���s���A�N�����Ă���ꍇ�̓A�v���I���������s���B
  bRet = pfAplEnd

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����ؑ֒����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

  On Error Resume Next
  
  cmdOK.Visible = False
  
  '�u�o�[�W�����ؑ֒���ʁF�\���v���O�o��
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_SHORIGAMEN_START, 0)
  
  '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
  tmrMail.Interval = MN_MAIL_INTERVAL
  tmrMail.Enabled = False
'V1.6.0.1 DEL START
'  '�A�v���N���p�C���^�o���^�C�}�l�ݒ�B
'  tmrAplCheck.Interval = MN_MAIL_INTERVAL
'  tmrAplCheck.Enabled = False
'V1.6.0.1 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����ؑ֒����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdOK_Click
'//  �@�\����  : �uOK�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    On Error Resume Next
    '�u�o�[�W�����ؑ֒���ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_SHORIGAMEN_END, 0)
       
    '����ʂ������B
    Unload Me
   
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrAplCheck_Timer
'//  �@�\����  : �A�v���N���`�F�b�N�p�^�C�}�A�^�C���A�b�v������
'//  �@�\�T�v  : �^�C���A�b�v���ɃA�v���N���`�F�b�N���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplCheck_Timer()

   Dim bRet As Boolean

'V1.6.0.1 DEL START
'
'   On Error Resume Next
'
'    If CheckAppStart(PROC_KANRI) = 0 And _
'         CheckAppStart(PROCESS_IDU_LOG) = 0 And _
'         CheckAppStart(PROCESS_LDU_LOG) = 0 Then
'         '�Ǘ��AIDU���O�ALDU���O���N�����Ă��Ȃ�=�A�v���I��
'         tmrAplCheck.Enabled = False
'         '�o�[�W�����ؑ֏������s���B
'         bRet = psVersionChange
'    End If
'V1.6.0.1 DEL END
   On Error Resume Next
'V1.6.0.1 ADD START
  '�҂����Ԃ�INI��`�𒴂������ǂ����`�F�b�N
  If lngtime <= lngMAX_Time Then
    '�A�v���N���`�F�b�N���s���B�S�A�v�����I�������Ƃ��̂݁A�������������s���B
    If CheckAppStart(PROC_KANRI) = 0 And CheckAppStart(PROCESS_IDU_LOG) = 0 And CheckAppStart(PROCESS_LDU_LOG) = 0 Then
      '�A�v���N���`�F�b�N�^�C�}���~����B
      tmrAplCheck.Enabled = False
      '�o�[�W�����ؑ֏������s���B
      bRet = psVersionChange
    Else
      '�N���A�v���L��̏ꍇ�A�^�C�}�𒣂蒼��
       tmrAplCheck.Interval = MN_MAIL_INTERVAL
      '���v�o�ߑ҂����Ԃ��A�b�v
       lngtime = lngtime + MN_MAIL_INTERVAL
    End If
  Else
    'INI��`�l�𒴂����ꍇ�A�����������ُ�Ƃ���B
    '�A�v���N���`�F�b�N�^�C�}���~����B
    tmrAplCheck.Enabled = False
    '�����ُ��\��
    psChangeVerEnd (1)
  End If
'V1.6.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v������
'//  �@�\�T�v  : ���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    On Error Resume Next
        
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmChangeVer.Caption, False
        pfFormActive (frmChangeVer.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psChangeVerEnd
'//  �@�\����  : �ؑ֌��ʕ\������
'//  �@�\�T�v  : �o�[�W�����ؑ֌��ʂ̌��ʕ�����\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iEnd�@�@�@[IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psChangeVerEnd(iEnd As Integer)
    Dim i As Integer       '�J�E���^
    Dim lngErrCode As Long '�G���[�R�[�h

    On Error Resume Next
      
    cmdOK.Visible = True
  
    If iEnd = 0 Then
       '�ؑ֐���I�����̕�����\������B
       lblMessage(0) = "����I�����܂����B"
       lblMessage(1) = ""
       '�u�o�[�W�����ؑ֏�������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_CHANGE_OK, 0)
    Else
       '�ؑֈُ펞�̕�����\������B
       lblMessage(0) = "�ُ�I�����܂����B"
       lblMessage(1) = ""
       '�u�o�[�W�����ؑ֏����ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_CHANGE_ERROR, lngErrCode)
    End If
    
  cmdOK.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfAplEnd
'//  �@�\����  : �A�v���I������
'//  �@�\�T�v  : EG-R�Ď��ՃA�v���I���������s��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfAplEnd() As Boolean
   Dim uMail As ML_KYOTU_INF           '���[��
   Dim bRtn As Boolean                 '���[���̖߂�l
   Dim lExitCode As Long
   
   On Error Resume Next
   
   If CheckAppStart(PROC_KANRI) <> 0 Then
      '�A�v���I���v�����Ǘ��ɑ��M����
      uMail.udtlHeader.dwId = ML_ID_APLEND_REQ
      uMail.udtlHeader.dwSize = MlSize.APLEND_REQ
      uMail.udtlHeader.dwProid = RHOSHU_ID
      uMail.udtlHeader.dwSubArea = 0
      bRtn = DssSendMail(MAIL_SLOT_KANRI, MlSize.APLEND_REQ, uMail.udtlHeader)
      If bRtn <> 0 Then
         '�u�A�v���N���E�I����ʁF���[�����M���팋�ʁv���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, APL_END_CMD, 0)
                        
         'IDU���O�m�F
         If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
            'IDU���O�I���v��CMD���M
            bRtn = EndIDULog
            If bRtn = False Then
               pfAplEnd = False
               psChangeVerEnd (1)
               Exit Function
            End If
         End If
            
         'LDU���O�m�F
         If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
            'LDU���O�I���v��CMD���M
            bRtn = EndLDULog
            If bRtn = False Then
               pfAplEnd = False
               psChangeVerEnd (1)
               Exit Function
            End If
         End If
         'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
         'V1.6.0.1 ADD END
       
 '        tmrAplCheck.Enabled = True  'V1.6.0.1 DEL
      Else
        '�u�A�v���N���E�I����ʁF���[�����M�ُ팋�ʁv���O�o��
        lExitCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, APL_END_CMD, lExitCode)
        '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
        pfAplEnd = False
        psChangeVerEnd (1)
      End If
   'IDU���O�m�F
   ElseIf CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
          'IDU���O�I���v��CMD���M
          bRtn = EndIDULog
          If bRtn = False Then
             pfAplEnd = False
             psChangeVerEnd (1)
          End If
          
          'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
         'V1.6.0.1 ADD END

'          tmrAplCheck.Enabled = True 'V1.6.0.1 DEL
    'LDU���O�m�F
    ElseIf CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
           'LDU���O�I���v��CMD���M
           bRtn = EndLDULog
           If bRtn = False Then
              pfAplEnd = False
              psChangeVerEnd (1)
           End If
       'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
        'V1.6.0.1 ADD END
        'tmrAplCheck.Enabled = True 'V1.6.0.1 DEL
    Else
        'V1.6.0.1 ADD START
          lngtime = 0
          lngtime = MN_MAIL_INTERVAL
          tmrAplCheck.Enabled = True
        'V1.6.0.1 ADD END
        'tmrAplCheck.Enabled = True 'V1.6.0.1 DEL
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psVersionChange
'//  �@�\����  : �o�[�W�����ؑ֏���
'//  �@�\�T�v  : �Ď���/IDU�o�[�W�����ؑ֏������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function psVersionChange() As Boolean

    Dim bRet As Boolean
    
    On Error Resume Next
     
    iChangeVerFlag = 0
        
    bRet = False
    
    '�����t�H���_�p�X�ݒ�
    psFolderPathSettei
        
    '�t�H���_�\���`�F�b�N�F����A���̓��J�o�������t�H���_�\���ȊO�͏����I��
    If iChengeVerApl = stsKansi Then
      '�Ď��ՃA�v��
      bRet = pfKansiChkFolder
    Else
      'IDU�A�v��
      bRet = pfIDUChkFolder
    End If
    
    If bRet = False Then
       '�����ΏۊO�t�H���_�\����Ԃ̂��ߏ������s�킸�ɏI���B
       psChangeVerEnd (1)
       Exit Function
    End If
    
    '�t�H���_�\�����򏈗�
    Select Case iChangeSts
       Case DLLFILE_AtoC
            '�t�H���_�\������F���폈���V�[�P���X���s���B
            bRet = pfDLLFILE_AtoC
       
       Case BACK_BtoC
            '�t�H���_�\���ُ�P�F�o�[�W�����Q���ꎞ�t�H���_�ւ̃��J�o���������s���B
            bRet = pfBACK_BtoC
    
       Case BACK_AtoB
            '�t�H���_�\���ُ�Q�F�o�[�W�����P���o�[�W�����Q�ւ̃��J�o���������s���B
            bRet = pfBACK_AtoB
    
       Case BACK_CtoA
            '�t�H���_�\���ُ�R�F�ꎞ�t�H���_���o�[�W�����P�ւ̃��J�o���������s���B
            bRet = pfBACK_CtoA
    End Select
    
    If bChangeVerSts = False Then
       psChangeVerEnd (1)
       Exit Function
    End If
    
    psVersionChange = bRet

    psChangeVerEnd (0)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psFolderPathSettei
'//  �@�\����  : �t�H���_�p�X�ݒ菈��
'//  �@�\�T�v  : �Ď���/IDU�o�[�W�����֑ؑΏۃt�H���_�p�X�ݒ���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function psFolderPathSettei() As Boolean

   On Error Resume Next
 
    Select Case Change_Version
         Case EGR_CHANGE_VER
            DllFolderName = Mid(PATH_GATE_E, 1, Len(PATH_GATE_E) - 2)               '�o�[�W�����P(�{��)��
            DllFolderName2 = PATH_GATE_ESAVE        '�o�[�W�����Q(�ۑ��p)��
            DllFolderName3 = PATH_GATE_ETEMP        '�ꎞ�t�H���_��
            ParaFolderName1 = PATH_GATE_EPARA       '�o�[�W�����P�p����
            ParaFolderName2 = PATH_GATE_ESAVE_PARA  '�o�[�W�����Q�p����
            iChengeVerApl = stsKansi
         Case NEG_CHANGE_VER
            DllFolderName = Mid(PATH_GATE_N, 1, Len(PATH_GATE_N) - 2)             '�o�[�W�����P(�{��)��
            DllFolderName2 = PATH_GATE_NSAVE        '�o�[�W�����Q(�ۑ��p)��
            DllFolderName3 = PATH_GATE_NTEMP        '�ꎞ�t�H���_��
            ParaFolderName1 = PATH_GATE_NPARA       '�o�[�W�����P�p����
            ParaFolderName2 = PATH_GATE_NSAVE_PARA  '�o�[�W�����Q�p����
            iChengeVerApl = stsKansi
         Case ICM_CHANGE_VER
            DllFolderName = PATH_IDU_APP & PATH_PROHAN_FOLDER  '�o�[�W�����P(�{��)��
            DllFolderName2 = PATH_IDU_APP & PATH_PROHAN_SAVE   '�o�[�W�����Q(�ۑ��p)��
            DllFolderName3 = PATH_IDU_APP & PATH_PROHAN_TEMP   '�ꎞ�t�H���_��
            iChengeVerApl = stsIDU
         Case PASMO_CHANGE_VER
            DllFolderName = PATH_IDU_APP & PATH_CMN_UNT_FOLDER     '�o�[�W�����P(�{��)��
            DllFolderName2 = PATH_IDU_APP & PATH_CMN_UNT_SAVE     '�o�[�W�����Q(�ۑ��p)��
            DllFolderName3 = PATH_IDU_APP & PATH_CMN_UNT_TEMP     '�ꎞ�t�H���_��
            iChengeVerApl = stsIDU
         Case JIKIUNCHIN_CHANGE_VER
            DllFolderName = PATH_UNAKI_SUB             '�o�[�W�����P(�{��)��
            DllFolderName2 = PATH_UNKAI_SUB_SAVE       '�o�[�W�����Q(�ۑ��p)��
            DllFolderName3 = PATH_UNKAI_SUB_TEMP       '�ꎞ�t�H���_��
            ParaFolderName1 = PATH_UNKAI_SUB_BACK      '�o�[�W�����P�o�b�N�A�b�v
            ParaFolderName2 = PATH_UNKAI_SUB_SAVE_BACK '�o�[�W�����Q�o�b�N�A�b�v
            iChengeVerApl = stsKansi
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfKansiChkFolder
'//  �@�\����  : �o�[�W�����ؑ֏���(�t�H���_�`�F�b�N)
'//  �@�\�T�v  : �t�H���_�\����ԃ`�F�b�N���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FTRUE=�t�H���_�L��@False=�t�H���_����
'///////////////////////////////////////////////////////////////////
Private Function pfKansiChkFolder() As Boolean
                                      
   Dim fso As New FileSystemObject
   Dim bRet1 As Boolean '�����P
   Dim bRet2 As Boolean '�����Q
   Dim bRet3 As Boolean '�����R
   Dim bRet4 As Boolean '�����S
   Dim bRet5 As Boolean '�����T
   
   On Error Resume Next
       
    pfKansiChkFolder = True
   
   '�u�o�[�W�����ؑ֒���ʁF�t�H���_��ԃ`�F�b�N�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_FOLDER_CHK, 0)
    
   '�o�[�W�����P(�{���t�H���_)�L���`�F�b�N
    bRet1 = fso.FolderExists(DllFolderName)
   '�o�[�W�����Q(�ۑ��p�t�H���_)�L���`�F�b�N
    bRet2 = fso.FolderExists(DllFolderName2)
   '�ꎞ�t�H���_�L���`�F�b�N
    bRet3 = fso.FolderExists(DllFolderName3)
   '�p�����[�^�t�H���_�L���`�F�b�N
    bRet4 = fso.FolderExists(ParaFolderName1)
   '�p�����[�^�t�H���_�Q�L���`�F�b�N
    bRet5 = fso.FolderExists(ParaFolderName2)
    Set fso = Nothing
 
    If bRet1 = True And bRet4 = True And bRet2 = True And bRet5 = False And bRet3 = False Then
      '�o�[�W�����P�F�L�A�o�[�W�����Q�F�L�A�ꎞ�t�H���_�F��
      '���ʁF������
      iChangeSts = DLLFILE_AtoC
      
    ElseIf bRet1 = True And bRet4 = False And bRet2 = True And bRet5 = True And bRet3 = False Then
      '�o�[�W�����P�F�L�A�p�����F���A�o�[�W�����Q�F�L�A�ꎞ�t�H���_�F��
      '���ʁF�p�����[�^���l�[�������ُ�
      iChangeSts = BACK_BtoC
    
    ElseIf bRet1 = True And bRet4 = False And bRet2 = False And bRet5 = False And bRet3 = True Then
      '�o�[�W�����P�F�L�A�o�[�W�����Q�F���A�ꎞ�t�H���_�F�L
      '���ʁF���l�[�������ُ�P
      iChangeSts = BACK_AtoB
    ElseIf bRet1 = False And bRet4 = False And bRet2 = True And bRet5 = False And bRet3 = True Then
      '�o�[�W�����P�F���A�o�[�W�����Q�F�L�A�ꎞ�t�H���_�F�L
      '���ʁF���l�[�������ُ�Q
      iChangeSts = BACK_CtoA
      
    Else
       pfKansiChkFolder = False
    End If
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfIDUChkFolder
'//  �@�\����  : �o�[�W�����ؑ֏���(�t�H���_�`�F�b�N)
'//  �@�\�T�v  : �t�H���_�\����ԃ`�F�b�N���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FTRUE=�t�H���_�L��@False=�t�H���_����
'///////////////////////////////////////////////////////////////////
Private Function pfIDUChkFolder() As Boolean
                                
  Dim fso As New FileSystemObject
  Dim bRet1 As Boolean '�����P
  Dim bRet2 As Boolean '�����Q
  Dim bRet3 As Boolean '�����R
   
  On Error Resume Next
  
  pfIDUChkFolder = True
   
  '�u�o�[�W�����ؑ֒���ʁF�t�H���_��ԃ`�F�b�N�v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_FOLDER_CHK, 0)
    
   '�o�[�W�����P(�{���t�H���_)�L���`�F�b�N
   bRet1 = fso.FolderExists(DllFolderName)
   '�o�[�W�����Q(�ۑ��p�t�H���_)�L���`�F�b�N
   bRet2 = fso.FolderExists(DllFolderName2)
   '�ꎞ�t�H���_�L���`�F�b�N
   bRet3 = fso.FolderExists(DllFolderName3)
   Set fso = Nothing
   
   If bRet1 = True And bRet2 = True And bRet3 = False Then
      '�o�[�W�����P�F�L�A�o�[�W�����Q�F�L�A�ꎞ�t�H���_�F��
      '���ʁF������
      iChangeSts = DLLFILE_AtoC
   ElseIf bRet1 = True And bRet2 = False And bRet3 = True Then
      '�o�[�W�����P�F�L�A�p�����F���A�o�[�W�����Q�F�L�A�ꎞ�t�H���_�F��
      '���ʁF���J�o�������Q
      iChangeSts = BACK_AtoB
      
   ElseIf bRet1 = False And bRet2 = True And bRet3 = True Then
      '�o�[�W�����P�F�L�A�o�[�W�����Q�F���A�ꎞ�t�H���_�F�L
      '���ʁF���l�[�������P
      iChangeSts = BACK_CtoA
      
   Else
      pfIDUChkFolder = False
   End If
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfDLLFILE_AtoC
'//  �@�\����  : �ؑ֏����P
'//  �@�\�T�v  : ���ʐؑ֏���(�o�[�W�����P���ꎞ�t�H���_)���l�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfDLLFILE_AtoC() As Boolean
                               
  On Error Resume Next
                             
  Dim fso As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
  Dim bRet As Boolean
  
  bRet = True
  
  '�o�[�W�����P���ꎞ�t�H���_�Ƀ��l�[��
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName, DllFolderName3
  Set fso = Nothing
 
  bRet = pfDLLFILE_BtoA
  If iChangeVerFlag = CHANGE_OK Then
     bRet = pfBACK_CtoA
  End If
  
  pfDLLFILE_AtoC = bRet
  Exit Function
  
FileCopyError:
  '���l�[���ُ펞�͏����I��
  '�u�o�[�W�����ؑ֒���ʁF�ꎞ�t�H���_���l�[���ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DELETEFOLDER_RENAME_ERROR, 0)
  pfDLLFILE_AtoC = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfDLLFILE_BtoA
'//  �@�\����  : �ؑ֏����Q
'//  �@�\�T�v  : ���ʐؑ֏���(�o�[�W�����Q���o�[�W�����P)���l�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfDLLFILE_BtoA() As Boolean
                                
  On Error Resume Next
                               
  Dim fso As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
  Dim bRet As Boolean
  
  bRet = True
  
  '�o�[�W�����Q���o�[�W�����P�Ƀ��l�[��
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName2, DllFolderName
  Set fso = Nothing
  
  bRet = pfDLLFILE_CtoB
  If iChangeVerFlag = CHANGE_RENAME_ERROR Then
     '�ꎞ�t�H���_���o�[�W�����Q���l�[���ُ펞�����F�o�[�W�����P���o�[�W�����Q���l�[��
     bRet = True
     If bRet = True Then
        bRet = pfBACK_AtoB
     End If
     If bRet = True Then
        '�o�[�W�����P���o�[�W�����Q���l�[�����펞�����F�ꎞ�t�H���_���o�[�W�����P���l�[��
        bRet = pfBACK_CtoA
     End If
  End If
  
  pfDLLFILE_BtoA = bRet
  
  Exit Function
  
FileCopyError:
  pfDLLFILE_BtoA = False
  iChangeVerFlag = CHANGE_RENAME_ERROR
  '�u�o�[�W�����ؑ֒���ʁFDLL�t�H���_���l�[���ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DLLFOLDER_RENAME_ERROR, 0)
  '�ꎞ�t�H���_���o�[�W�����P���l�[������
  bRet = pfBACK_CtoA()
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfDLLFILE_CtoB
'//  �@�\����  : �ؑ֏����R
'//  �@�\�T�v  : ���ʐؑ֏���(�ꎞ�t�H���_���o�[�W�����Q)���l�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfDLLFILE_CtoB() As Boolean
                                
   On Error Resume Next
                               
  Dim fso As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
  Dim bRet As Boolean
    
  bRet = True
  
  '�ꎞ�t�H���_���o�[�W�����Q�Ƀ��l�[��
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName3, DllFolderName2
  Set fso = Nothing
  
  If iChengeVerApl = 0 Then
     '�Ď��ՃA�v���̂݃p�����[�^���l�[�����s���B
     bRet = pfPARA_BtoA
     If iChangeVerFlag = CHANGE_RENAME_ERROR Then
        '�o�[�W�����Q�p�������o�[�W�����P�p�����ُ펞�����F�o�[�W�����Q���ꎞ�t�H���_���l�[��
        bRet = pfBACK_BtoC
     End If
  End If
  
  pfDLLFILE_CtoB = bRet

  Exit Function
  
FileCopyError:
  '�u�o�[�W�����ؑ֒���ʁF�ۑ��p�t�H���_���l�[���ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_BACKUPFOLDER_RENAME_ERROR, 0)
  pfDLLFILE_CtoB = False
  iChangeVerFlag = CHANGE_RENAME_ERROR
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfPARA_BtoA
'//  �@�\����  : �ؑ֏����S
'//  �@�\�T�v  : ���ʐؑ֏���(�o�[�W�����Q���o�[�W�����P)�p�������l�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfPARA_BtoA() As Boolean
                                
   On Error Resume Next
                               
  Dim fso As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
  Dim bRet As Boolean
  
  bRet = True
  
  '�o�[�W�����Q���o�[�W�����P�Ƀp�������l�[��
  On Error GoTo FileCopyError
  fso.MoveFolder ParaFolderName2, ParaFolderName1
  Set fso = Nothing
  
  pfPARA_BtoA = bRet
  
  Exit Function
  
FileCopyError:
  '�u�o�[�W�����ؑ֒���ʁF�T�u�t�H���_���l�[���ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_SUBFOLDER_RENAME_ERROR, 0)
  pfPARA_BtoA = False
  iChangeVerFlag = CHANGE_RENAME_ERROR
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfBACK_CtoA
'//  �@�\����  : �ؑ֏����T
'//  �@�\�T�v  : ���ʐؑ֏���(�ꎞ�t�H���_���o�[�W�����P)���l�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfBACK_CtoA() As Boolean
                                
  On Error Resume Next
                               
  Dim fso As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
  Dim bRet As Boolean
  
  bRet = True
  
  '�ꎞ�t�H���_���o�[�W�����P�Ƀ��l�[��
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName3, DllFolderName
  Set fso = Nothing
  
  '�t�H���_�\�����ȉ��̎��AC��A��ɐ���V�[�P���X���s���B
  If iChangeSts = BACK_CtoA Or iChangeSts = BACK_BtoC Then
     '�t�H���_����F���폈�����s���B
     bRet = pfDLLFILE_AtoC
  End If
  
  pfBACK_CtoA = bRet
  
  Exit Function
  
FileCopyError:
  '�u�o�[�W�����ؑ֒���ʁFDLL�t�H���_���l�[���ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DLLFOLDER_RENAME_ERROR, 0)
  pfBACK_CtoA = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfBACK_BtoC
'//  �@�\����  : �ؑ֏����U
'//  �@�\�T�v  : ���ʐؑ֏���(�o�[�W�����Q���ꎞ�t�H���_)���l�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfBACK_BtoC() As Boolean
                                  
  Dim fso As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
  Dim bRet As Boolean
  
  On Error Resume Next
  
  bRet = True
    
  '�o�[�W�����Q���ꎞ�t�H���_�Ƀ��l�[��
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName2, DllFolderName3
  Set fso = Nothing
  
  If iChangeSts = BACK_BtoC Then
     '�o�[�W�����P���o�[�W�����Q�ւ̃��J�o������
      bRet = pfBACK_AtoB

     If bRet = True Then
        '�ꎞ�t�H���_���o�[�W�����P�ւ̃��J�o������
        bRet = pfBACK_CtoA
     End If
  End If
  
  pfBACK_BtoC = bRet
  
  Exit Function
   
FileCopyError:
  '�u�o�[�W�����ؑ֒���ʁF�ꎞ�t�H���_���l�[���ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_DELETEFOLDER_RENAME_ERROR, 0)
  pfBACK_BtoC = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfBACK_AtoB
'//  �@�\����  : �ؑ֏����V
'//  �@�\�T�v  : ���ʐؑ֏���(�o�[�W�����P���o�[�W�����Q)���l�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfBACK_AtoB() As Boolean
                                                              
  Dim fso As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
  Dim bRet As Boolean
  
  On Error Resume Next
  
  bRet = True
  
  '�o�[�W�����P���o�[�W�����Q�Ƀp�������l�[��
  On Error GoTo FileCopyError
  fso.MoveFolder DllFolderName, DllFolderName2
  Set fso = Nothing
  
  If iChangeSts = BACK_AtoB Then
     '�ꎞ�t�H���_���o�[�W�����P�ւ̃��J�o���������s���B
      bRet = pfBACK_CtoA
      
      If bRet = True Then
         '�t�H���_����F���폈�����s���B
         bRet = pfDLLFILE_AtoC
      End If
  End If
  
  pfBACK_AtoB = bRet
  Exit Function
  
FileCopyError:
  '�u�o�[�W�����ؑ֒���ʁF�ۑ��p�t�H���_���l�[���ُ�v���O�o��
  Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_CHANGE_BACKUPFOLDER_RENAME_ERROR, 0)
  pfBACK_AtoB = False
  Set fso = Nothing
  bChangeVerSts = False
End Function

