VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEkimeiFD 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�w���f�[�^�}�̓���"
   ClientHeight    =   9000
   ClientLeft      =   2160
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9
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
   ScaleHeight     =   9424.084
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   12121.21
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "�}�̎�O"
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
      Left            =   9360
      TabIndex        =   3
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   600
      Top             =   7440
   End
   Begin VB.CommandButton cmdFDInput 
      Caption         =   "���s"
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
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "   ���j���[     ��ʂ֖߂�"
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
      Left            =   9360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�w���f�[�^�}�̓���"
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
Attribute VB_Name = "frmEkimeiFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmEkimeiFD.frm
'//  �p�b�P�[�W���F�w���f�[�^�}�̓������
'//
'//  �T�v�F�w���f�[�^�}�̓������
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

'�_�C�A���O�\���p�h�c
Private Enum FDEkimei
    REBOOT = 1                          '�P�F����I���_�C�A���O
    FD_INPUT_ERR = 2                    '�Q�F�ُ�I���_�C�A���O
End Enum

'���O�o�͈˗��pID
Private Enum LogID
    LOG_NORMAL = 0                      '�O�F�w���f�[�^�e�c�t�@�C���쐬����
    FILEDELETE_ERROR = 1                '�P�F�w���f�[�^�e�c�t�@�C���폜�ُ�
    FILECOPY_ERROR = 2                  '�Q�F�w���f�[�^�e�c�t�@�C���쐬�ُ�
End Enum




'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : cmdInstall_Click
'//  �T�v      : �u�}�̎�O�v�t��������
'//  ����      : �}�̂����O���B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
 On Error Resume Next
   
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)

End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �T�v      : �w���f�[�^�e�c������ʂ����[�h���ꂽ���̃C�x���g�v���V�[�W��
'//  ����      : ���[����M�p�̃^�C�}�l��ݒ肷��B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMEI_DATA_INPUT_GAMEN_START, 0)
    
    '----------------------------------------------------
    '��ʏ����l�ݒ�
    '----------------------------------------------------
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �T�v      : �w���f�[�^�e�c������ʂ��\�����ꂽ���̃C�x���g�v���V�[�W��
'//  ����      : �u���[����M�p�^�C�}�v���N������B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    On Error Resume Next

    '���[����M�p�^�C�}���N������
    tmrMail.Enabled = True

End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �T�v      : �w���f�[�^�e�c������ʂ��������ꂽ���̃C�x���g�v���V�[�W��
'//  ����      : �u���[����M�p�̃^�C�}�v��j������B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next

    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False

End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : cmdFDInput_Click
'//  �T�v      : �u���s�v�{�^���������̃C�x���g�v���V�[�W��
'//  ����      : �w���f�[�^�e�c�����������s���B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFDInput_Click()

    Dim iResponse   As Integer          '�_�C�A���O�{�^���R�[�h
    Dim bRet        As Boolean          '���[�����M����
    Dim strFilePath As String           '�t�@�C���I���_�C�A���O�ɂđI�����ꂽ�t�@�C��
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g

    On Error Resume Next

    '������
    bRet = False        '�����߂�l
    strFilePath = ""    '�t�@�C������NULL�����ŏ�����   ' V6.1.0.3 ADD

    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMEI_DATA_INPUT_BUTTOM, 0)

   '��ʂ̃{�^���������s�ɂ���
    Call sButtonEnabled(False)

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
    CommonDialog1.Filter = "���ׂẴt�@�C��(*.*)|*.*|"
    '�t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    '�I�������t�@�C�������擾
    strFilePath = CommonDialog1.FileName

    Call ChDrive("D")

    If strFilePath <> "" Then        '�t�@�C���I��L
        
        '�t�@�C���R�s�[���s���A�w���f�[�^�e�c�t�@�C�����쐬����
        bRet = fFileCopy(strFilePath)

        '���[�����M���s��
        If bRet = True Then '�t�@�C���R�s�[������

            '����I���_�C�A���O��\������
            iResponse = fMessageBox(FDEkimei.REBOOT)

            If iResponse = vbOK Then    '�n�j����

                '�A�v���N���`�F�b�N
                If CheckAppStart(PROC_KANRI) <> 0 Then '�Ď��ՃA�v���N����
                    '���[�����M����
                    bRet = fSendPowerOffReqMail()
                                    
                    '���O�o�͈˗�
                    '���[�����M������
                    If bRet = True Then
                        sLogRequest LOG_NORMAL
                    End If

                Else '�Ď��ՃA�v���I����
                    '�ێ�v���Z�X�I������
                    psEndHoshuProc
                    '���u�[�g����
                    dllAPLEndReboot
                End If

            Else                        '�L�����Z���F��ōċN��
                '��ʂ̃{�^���������\�ɂ���
                Call sButtonEnabled(True)
            End If

        End If
         
    End If

    '�����L�����Z���܂��́A���[�����M���s
    If bRet = False Then
        '��ʂ̃{�^���������\�ɂ���
        Call sButtonEnabled(True)
    End If

End Sub



'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �T�v      : �u���j���[��ʂɖ߂�v�{�^���������̃C�x���g�v���V�[�W��
'//  ����      : �w���f�[�^�}�̓�����ʂ����B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMEI_DATA_INPUT_GAMEN_END, 0)
    
    '��ʏ���
    Unload Me

End Sub


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : fFileCopy
'//  �T�v      : �w���f�[�^�e�c�t�@�C�����쐬����B
'//  ����      : �e�c���͂��ꂽ�t�@�C�����R�s�[���A�w���f�[�^�e�c�t�@�C�����쐬����B
'//  ���Ұ�    :�e�c�w���f�[�^�t�@�C���p�X
'//            :�߂�l  ,�R�s�[��������FTrue�@�@�ُ�FFalse
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fFileCopy(strFilePath As String) As Boolean
    
    Dim iRet    As Integer                 '�w���f�[�^�L���m�F
    Dim iErrorFlag As Integer              '�G���[�t���O

    On Error Resume Next

    '������
    fFileCopy = False      '�߂�l
    iRet = -1
    iErrorFlag = 0

    '�w���f�[�^�e�c�t�@�C���̗L���`�F�b�N
    iRet = GetAttr(FD_EKIMEI_FILE)
    
    '�G���[���[�`����錾
    On Error GoTo Err_LOG

    '�G���[�̕��ނ��t�@�C���폜�G���[�ɐݒ�
    iErrorFlag = 1

    '�t�@�C�������݂���ꍇ�A�w���f�[�^�e�c�t�@�C�����폜
    If iRet <> -1 Then
        Kill FD_EKIMEI_FILE
    End If
    
    '�G���[�̕��ނ��t�@�C���쐬�G���[�ɐݒ�
    iErrorFlag = 2
    '�e�c�w���f�[�^�t�@�C������w���f�[�^�e�c�t�@�C���ɃR�s�[����
    FileCopy strFilePath, FD_EKIMEI_FILE
    '�t�@�C���R�s�[����I��
    fFileCopy = True

    Exit Function

Err_LOG:

    '�G���[���[�`����錾
    On Error Resume Next
    
    '���O�o�͈˗�
    If iErrorFlag = 1 Then
        sLogRequest FILEDELETE_ERROR
    ElseIf iErrorFlag = 2 Then
        sLogRequest FILECOPY_ERROR
    End If
    '�ُ�I���_�C�A���O��\��
    fMessageBox (FDEkimei.FD_INPUT_ERR)
    '�w���f�[�^�폜
    Kill FD_EKIMEI_FILE

End Function
'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : fSendPowerOffReqMail
'//  �T�v      : �Ď��Փd��OFF�v�����[���𑗐M����B
'//  ����      : �Ď��Փd��OFF�v�����[�����쐬�A���M�B
'//  ���Ұ�    :�߂�l  ,���[�����M����FTrue�@�@�ُ�FFalse
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSendPowerOffReqMail() As Boolean
    
    Dim lngRet      As Long                 '�߂�l
    Dim udtMail     As MAIL_JIKI_UNKAI     '���[�����M�G���A

    On Error Resume Next

    '������
    fSendPowerOffReqMail = False

    '���[���f�[�^�쐬
    udtMail.mlHeader.dwId = ML_ID_KAN_PW_OFF_REQ        '���[���h�c�F�Ď��d��OFF�v��
    udtMail.mlHeader.dwSize = Len(udtMail)              '���[���T�C�Y
    udtMail.mlHeader.dwProid = RHOSHU_ID                '���M���v���Z�X�h�c�F�ێ�
    udtMail.mlHeader.dwSubArea = 0                      '�⏕���F�O�i�Œ�j
    udtMail.dwData = Ml_SyoriType.ML_DT_REBOOT          '������ʁF���u�[�g

    '���[�����M
    lngRet = DssSendMail(MAIL_SLOT_KANRI, Len(udtMail), udtMail.mlHeader)
    
    
    If lngRet = False Then '���[�����M�ُ�
        '�u�Ď��d��OFF�v���F���[�����M�ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
    Else
        '�u�Ď��d��OFF�v���F���[�����M����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_DENGENOFF_CMD_SEND, 0)
        '���[�����M����I��
        fSendPowerOffReqMail = True
   End If

End Function


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : sButtonEnabled
'//  �T�v      : ��ʕ\�����̃{�^���̉����\�^�s�ݒ���s���B
'//  ����      : ��ʂ̃{�^����Enabled��ݒ肷��B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sButtonEnabled(bSet As Boolean)

    On Error Resume Next

    cmdFDInput.Enabled = bSet       '���s�{�^��
    cmdReturn.Enabled = bSet        '�ێ��ʂ֖߂�{�^��
    cmdRemove.Enabled = bSet        '�ێ��ʂ֖߂�{�^��

End Sub
'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �T�v      : �u���[����M�p�^�C�}�v���^�C���A�b�v�������̃C�x���g�v���V�[�W��
'//  ����      : ���[����M�������s���B
'//  ���Ұ�    :
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    On Error Resume Next
    
    '�ėp���[����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmEkimeiFD.Caption, False
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : fMessageBox
'//  �T�v      : �_�C�A���O�\��
'//  ����      : �_�C�A���OID�ɂ��A�_�C�A���O���쐬���\������
'//  ���Ұ�    : �_�C�A���OID
'//            : �߂�l  ,�����t���
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMessageBox(iMsgID As Integer) As Integer

    Dim strMessage  As String           '�_�C�A���O�̕���
    Dim strTitle    As String           '�_�C�A���O�̃^�C�g��
    Dim lngOption   As Long             '�_�C�A���O�̕\���{�^���ƃA�C�R��

    fMessageBox = 0

    Select Case iMsgID
        Case FDEkimei.REBOOT         '����I��
            strMessage = "�w���f�[�^�}�̓���������ɍs���܂����B" & Chr(vbKeyReturn) & _
                         "�Ď��Ղ��ċN�����܂����H"
            lngOption = vbOKCancel + vbInformation      '�u�n�j�v�u�L�����Z���v�{�^���A�u���v�A�C�R��
            strTitle = "�w���f�[�^�}�̓���"

        Case FDEkimei.FD_INPUT_ERR   '�ُ�I��
            strMessage = "�w���f�[�^�}�̓����Ɏ��s���܂����B" & Chr(vbKeyReturn) & _
                         "�}�̂��ُ�łȂ������m�F���A�������}�̂�}�����A�ēx���s���Ă��������B"
            lngOption = vbOKOnly + vbCritical        '�u�n�j�v�{�^���A�u�x���v�A�C�R��
            strTitle = "�w���f�[�^�}�̓���"

        Case Else
    End Select

    If lngOption <> 0 Then
        '���b�Z�[�W�{�b�N�X��\�����A�߂�l��Function�̖߂�l�Ƃ���B
        fMessageBox = MsgBox(strMessage, lngOption, strTitle)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : sLogRequest
'//  �T�v      : ���O�o�͂��˗�����B
'//  ����      : �����Ɋւ��郍�O�o�͂��˗�����B
'//  ���Ұ�    :  iLog       ,I ,Integer        :���O�o�͈˗��pID
'//            :
'//
'//  ORIGINAL  �F(2.7.0.1) 2010-12-24  CODED BY  [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sLogRequest(iLogID As Integer)

    Dim udtLogParam As LOGPARAM  '���O�g���[�X�˗����Ұ�
    Dim lRet As Long   '���O�g���[�X�˗��֐��̖߂�l

'    '���O�˗��p�����[�^�̋��ʕ��ɒl���Z�b�g����B
    If iLogID = LOG_NORMAL Then
    '�w���f�[�^�e�c�t�@�C���쐬�����̏ꍇ�A
        '�u�w���ް�FḐ�ٍ쐬 ����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, "�w���ް�FḐ�ٍ쐬 ����", 0)
    ElseIf iLogID = FILEDELETE_ERROR Then
    '�w���f�[�^�e�c�t�@�C���폜�G���[�̏ꍇ�A
        '�u�w���ް�FḐ�ٍ폜 �ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, "�w���ް�FḐ�ٍ폜 �ُ�", 0)
    ElseIf iLogID = FILECOPY_ERROR Then
    '�w���f�[�^�e�c�t�@�C���R�s�[�G���[�̏ꍇ�A
        '�u�w���ް�FḐ�ٍ폜 �ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "�w���ް�FḐ�ٍ쐬 �ُ�", 0)
    Else
    End If
End Sub

