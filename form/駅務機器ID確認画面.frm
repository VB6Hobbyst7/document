VERSION 5.00
Begin VB.Form frmEkimKikiId 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�w���@��ID�m�F"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
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
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9120
      Top             =   3480
   End
   Begin VB.CommandButton cmdInstall 
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
      Height          =   735
      Left            =   9720
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "�e�L�X�g�\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   9720
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "�}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   9720
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ListBox ListEkimId 
      Height          =   7710
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   8775
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10425
      Width           =   3135
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
      Left            =   9360
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   8
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�w���@��ID�m�F"
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
Attribute VB_Name = "frmEkimKikiId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmEkimKikiId.frm
'//  �p�b�P�[�W���F�w���@��ID�m�F���
'//
'//  �T�v�F�w���@��ID�m�F���
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 �w���@��ID�����ݐ�f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 �y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
Private iSendType As Integer            '�v����ʒl
Private Const EKIMU_DEFU = "APL\APL_WORK"

Private Const APL = "APL"
Private Const LOG = "LOG"
Private Const Data = "DATA"
Private Const BACKUP = "BACKUP"

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �w���@��ID�m�F���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 �y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    On Error Resume Next
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
    
'V1.7.0.1 ADD START
    Dim bRet As Boolean                 '�߂�l
    Dim bFlag As Boolean                '�t���O
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim udtMail As MAIL_INFO_CMD          '��ʕ\���v��
    Dim uMail As ML_KYOTU_INF           '���[��
    Dim lLen  As Long
  
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
  
   '�o�b�t�@�t���b�V���v�������O�v���Z�X�ɑ��M����
   '���v��CMD(�w���@��ID=0)��ID����ɑ��M����
   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
   iSendType = MailCmdType.ML_DT_EKIMU_ID
   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '�u�w���@��ID�m�F�F���v��CMD���M�ُ�v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
      '�v���O���X�o�[����������
      Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
      
      Exit Sub
   Else
      '�u�w���@��ID�m�F�F���v��CMD���M����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
      '��ʃ��b�N
      cmdVer(1).Enabled = False
      cmdVer(2).Enabled = False
      cmdInstall.Enabled = False
      cmdCancel.Enabled = False
   End If
   
    '�o�b�t�@�t���b�V���I���ʒm��M
    bFlag = False
    Do Until bFlag = True
       '���[����M�������s��
       lLen = DssMailRead(plMSlot_MN, uMail)
       If lLen > 0 Then                            '��M����̎�
         If ML_ID_INFO_RES = uMail.udtlHeader.dwId Then '���[���h�c
            '���v��RES����M������A��ʕ\���p�t�@�C���쐬���s���B
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
            '�v����ʁA�������ʂ��擾
            Call psDispID(uMail.lngData(1))
           '��ʃ��b�N����
' EG20 V6.3.0.1�y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z�폜�J�n
'           cmdVer(1).Enabled = True
'           cmdVer(2).Enabled = True
' EG20 V6.3.0.1�y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z�폜�I��
' EG20 V6.3.0.1�y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z�ǉ��J�n
            If ListEkimId.ListCount > 0 Then
                cmdVer(1).Enabled = True
                cmdVer(2).Enabled = True
            End If
' EG20 V6.3.0.1�y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z�ǉ��I��
           cmdInstall.Enabled = True
           cmdCancel.Enabled = True
           Exit Do
         End If
        End If
        Sleep (MN_MAIL_INTERVAL)
    Loop
'V1.7.0.1 ADD END
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �w���@��ID�m�F���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    On Error Resume Next
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �w���@��ID�m�F���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 �y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
 'V1.7.0.1 DEL START
'   Dim udtMail As MAIL_INFO_CMD          '��ʕ\���v��
'   Dim iResponse As Integer            '���b�Z�[�W�{�b�N�X�߂�l
'   Dim bRet As Boolean                 '���[�����M�����߂�l
'   Dim lngErrCode As Long              '�G���[�R�[�h
'   Dim bFlag As Boolean
'   Dim lId As Long
 'V1.7.0.1 DEL END
 
   On Error Resume Next
   
   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
    
   '�u�w���@��ID�m�F��ʁF�\���v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_GAMEN_START, 0)
    
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
 
' EG20 V6.3.0.1�y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z�ǉ��J�n
    cmdVer(1).Enabled = False
    cmdVer(2).Enabled = False
' EG20 V6.3.0.1�y�e�L�X�g�o�́A�}�̏o�̓{�^���̗}�~�Ή��z�ǉ��I��
 'V1.7.0.1 DEL START
'   '���v��CMD(�w���@��ID=0)��ID����ɑ��M����
'   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
'   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
'   udtMail.mlHeader.dwProid = RHOSHU_ID
'   udtMail.mlHeader.dwSubArea = 0
'   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
'   iSendType = MailCmdType.ML_DT_EKIMU_ID
'   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
'   If bRet = False Then
'      '�u�w���@��ID�m�F�F���v��CMD���M�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
'   Else
'      '�u�w���@��ID�m�F�F���v��CMD���M����v���O�o��
'      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
'      '��ʃ��b�N
'      cmdVer(1).Enabled = False
'      cmdVer(2).Enabled = False
'      cmdInstall.Enabled = False
'      cmdCancel.Enabled = False
'   End If
 'V1.7.0.1 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  �֐�����  : cmdCancel_Click
'//  �T�v     : �u���j���[��ʂ֖߂�v�t��������
'//  ����     : ����ʂ���������B
'//  ���Ұ�   :
'//           :
'//
'//  ORIGINAL  �F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
   
    On Error Resume Next
    
    '�u�w���@��ID�m�F��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  �֐�����  : cmdInstall_Click
'//  �T�v     : �u�}�̎�O�v�t��������
'//  ����     : �}�̂����O���B
'//  ���Ұ�   :
'//           :
'//
'//  ORIGINAL  �F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdInstall_Click()
 On Error Resume Next
   
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  �֐�����  : cmdVer_Click
'//  �T�v     : �u�e�L�X�g�\���v�u�}�̏o�́v�t��������
'//  ����     : �e�t�����������s���B
'//  ���Ұ�   :
'//           :
'//
'//  ORIGINAL  �F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdVer_Click(Index As Integer)
    Dim lRetVal As Long             '�߂�l
    Dim sCommand As String          '�R�}���h������
    Dim lngErrCode As Long
    Dim bRet As Boolean
    
    On Error Resume Next
  
    Select Case Index

      Case 1
           '�u�e�L�X�g�\���t�F�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_TEXT_BUTTOM, 0)
           '���������s�R�}���h���쐬
           sCommand = MN_EXE_MEMO & MN_VERSI_FILE
           '���������N������
           lRetVal = Shell(sCommand, vbMaximizedFocus)
           '���������A�N�e�B�u�i�O�ʕ\���j�ɂ���
           AppActivate lRetVal, True
           SendKeys "{LEFT}", True
      Case 2
           '�u�}�̏o�͖t�F�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_OUTPUT_BUTTOM, 0)
           bRet = Text_OutPut
           If bRet = False Then
              '�u�}�̏o�ُ͈�v���O�o��
              Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, EKIMUKIKI_ID_OUTPUT_ERROR, 0)
           Else
              '�u�}�̏o�͐���v���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, EKIMUKIKI_ID_OUTPUT_OK, 0)
           End If
           
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  �֐�����  : Text_Output
'//  �T�v     : �u�}�̏o�́v����
'//  ����     : �}�̏o�͏������s���B
'//  ���Ұ�   :
'//           :
'//
'//  ORIGINAL  �F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS �F(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 �w���@��ID�����ݐ�f�B���N�g���ʒu�ύX
'//  REVISIONS �F(1.20.0.1) 2010-03-10   REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//  REVISIONS �F(EG20 V2.0.1.1) 2011-11-21  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�z
'//                 �E�o�̓t�@�C�����ύX
'//  REVISIONS �F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function Text_OutPut() As Boolean
    Dim sCopyfile As String         '�R�s�[��
    Dim sCopyTargetFile As String   '�R�s�[��
    Dim sLzhDirName As String
    Dim iResponse           As Integer          'MsgBox�߂�l
    
' EG20 V2.0.1.1 ADD START
    Dim strStationName As String                ' �w���擾�G���A
' EG20 V2.0.1.1 ADD END
' EG20 V3.0.0.2�ǉ��J�n
    Dim fso         As New FileSystemObject     ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim textWrite   As TextStream               ' �e�L�X�g�i���C�g�j
    Dim textRead    As TextStream               ' �e�L�X�g�i���[�h�j
    Dim bWOpen      As Boolean
    Dim bROpen      As Boolean
    Dim strRecord   As String                   ' ���[�N
' EG20 V3.0.0.2�ǉ��I��
    
On Error GoTo FileCopyError
  
    Text_OutPut = False

' EG20 V3.0.0.2�ǉ��J�n
    bWOpen = False
    bROpen = False
' EG20 V3.0.0.2�ǉ��I��
   
    '�t�H���_�I����ʂ�\�������A�t�@�C���i�[�f�B���N�g�����𓾂�B
'    sLzhDirName = pfDirSelection("a:", "�w���@��ID�����ݐ�̃f�B���N�g���I��")     'V1.12.0.1 DEL
    'sLzhDirName = pfDirSelection("H:", "�w���@��ID�����ݐ�̃f�B���N�g���I��")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
    sLzhDirName = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
    If sLzhDirName = "" Then
       Text_OutPut = True
       Exit Function
    End If
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
' EG20 V2.0.1.1 DEL START
'    sCopyfile = sLzhDirName & EKIMU_ID_TXT
' EG20 V2.0.1.1 DEL END
' EG20 V2.0.1.1 ADD START
    '�w���擾
    strStationName = gsGetStationEkiName
    ' �o�̓t�@�C�����쐬
    sCopyfile = sLzhDirName & strStationName & "_" & EKIMU_ID_TXT
' EG20 V2.0.1.1 ADD END
    
    sCopyTargetFile = MN_VERSI_FILE
    
' EG20 V3.0.0.2�폜�J�n
'    FileCopy sCopyTargetFile, sCopyfile
' EG20 V3.0.0.2�폜�I��
    
' EG20 V3.0.0.2�ǉ��J�n
    Set textWrite = fso.CreateTextFile(sCopyfile, True)
    bWOpen = True
    textWrite.WriteLine ("�ݒu�w�@�F" & strStationName)
    textWrite.WriteBlankLines (1)
    Set textRead = fso.OpenTextFile(sCopyTargetFile, ForReading, False)
    bROpen = True
    Do Until textRead.AtEndOfStream = True
        strRecord = textRead.ReadLine
        textWrite.WriteLine strRecord
    Loop
    textWrite.Close
    bWOpen = False
    textRead.Close
    bROpen = False
    Set textWrite = Nothing
    Set textRead = Nothing
    Set fso = Nothing
' EG20 V3.0.0.2�ǉ��I��
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    iResponse = MsgBox("����I�����܂����B", _
                       vbOKOnly, _
                       "�}�̏o�͌���")
    
    
    '�f�B�X�N�����擾
    Text_OutPut = True
    
    Exit Function

FileCopyError:
' EG20 V3.1.0.2�ǉ��J�n
    If bWOpen = True Then
        textWrite.Close
        bWOpen = False
    End If
    If bROpen = True Then
        textRead.Close
        bROpen = False
    End If
    Set textWrite = Nothing
    Set textRead = Nothing
    Set fso = Nothing
' EG20 V3.1.0.2�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    iResponse = MsgBox("�ُ�I�����܂����B", _
                       vbOKOnly, _
                       "�}�̏o�͌���")
End Function

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
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
 'V1.7.0.1 DEL START
'  Dim lLen  As Long
'  Dim uMail As ML_KYOTU_INF           '���[��
'
'  On Error Resume Next
'
'  '���[����M
'  lLen = DssMailRead(plMSlot_MN, uMail)
'  If lLen > 0 Then                            '��M����̎�
'
'      Select Case uMail.udtlHeader.dwId  '���[���h�c
'        Case ML_ID_HOSHU_ACTIVE_REQ
'            '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
'            AppActivate frmEkimKikiId.Caption, False
'            pfFormActive (frmEkimKikiId.hwnd)
'        Case ML_ID_INFO_RES
'            '���v��RES����M������A��ʕ\���p�t�@�C���쐬���s���B
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
'
'            '�v����ʁA�������ʂ��擾
'            Call psDispID(uMail.lngData(1))
'        Case Else
'     End Select
'  End If
'  '��ʃ��b�N����
'  cmdVer(1).Enabled = True
'  cmdVer(2).Enabled = True
'  cmdInstall.Enabled = True
'  cmdCancel.Enabled = True
'V1.7.0.1 DEL END
'V1.7.0.1 ADD START
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmEkimKikiId.Caption, False
        pfFormActive (frmEkimKikiId.hwnd)
    End If
'V1.7.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psDispID
'//  �@�\����  : ��ʕ\������
'//  �@�\�T�v  : �w���@��ID����ʕ\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Long     lngSts    [IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Function psDispID(lngSts As Long)
    Dim sEkimuIDFile    As String   '�w���@��ID�t�@�C���p�X
    Dim iRet            As Integer  'INI�擾�߂�l
    Dim sFolder         As String * MAX_PATH_SIZE  '�t�H���_��
    Dim sFile           As String   '�t�@�C����
    Dim MyName          As String   '�t�@�C����������
    Dim bRet            As Boolean  '�߂�l
    Dim lngErrCode      As Long     '�G���[�R�[�h
    Dim intFileNo       As Integer  '�t�@�C���ԍ�
    Dim strWork         As String   '��ƃG���A
    Dim dwErrsts        As Long
    Dim sFolderName     As String
        
    '�������ʃX�e�[�^�X������̏ꍇ�AiF���������s���B
    If lngSts = 0 Then
      sFolder = ""
      
      '�������ʁF���펞�͉�ʕ\������
      iRet = GetPrivateProfileString(IDU_SECTION_NAME, _
                                     IDU_EKIMUID_KEY, _
                                     EKIMU_DEFU, sFolder, Len(sFolder), _
                                     PATH_IDU_INI_FILE)
      If iRet = 0 Then
        sFolder = EKIMU_DEFU
      End If
      sEkimuIDFile = ""
      '�v����ʒl���t�@�C�����쐬
      sFile = Replace(EKIMU_ID_FILE, "##", Format(iSendType, "0#"))
      If iRet = 0 Then
         sFolderName = RTrim(sFolder)
      Else
         sFolderName = Mid(sFolder, 1, iRet)
      End If
      '�p�X�ϊ�����
      sFolderName = pfChangeFolderName(sFolderName)
      '�w���@��ID�t�@�C���p�X�쐬
      sEkimuIDFile = sFolderName & "\" & sFile
      '�t�@�C���L���`�F�b�N
      If Dir(sEkimuIDFile, vbNormal) = "" Then
         Exit Function
      End If
      
      '/////////////////////////////////////////////////////////////////////
      '//�ێ��p�֐��F�w���@��ID��ʕ\���p�t�@�C���쐬����
      '////////////////////////////////////////////////////////////////////
      'bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE) 'V1.8.0.1 DEL
      bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE, PATH_IDU_APP) 'V1.8.0.1 ADD

      If dwErrsts = 1 Then
         '�G���[�R�[�h�F����
         '���X�g������
         ListEkimId.Clear

        'VB�G���[����
        On Error GoTo Error_psVersionDisp
    
        '�w���@��ID��ʕ\���p�t�@�C���̃t�@�C���ԍ����擾����B
        intFileNo = FreeFile
      
        '�w���@��ID��ʕ\���p�t�@�C���I�[�v��
        Open MN_VERSI_FILE For Input As #intFileNo
    
        '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1�폜
        Do While Not EOF(intFileNo)                         ' EG20 V3.3.0.1�ǉ�
           '��ƃG���A��������
           strWork = ""

           Line Input #intFileNo, strWork
           
           '���s�R�[�h�݂͓̂ǂ݂Ƃ΂�
           If Trim(strWork) <> "" Then
              '���X�g�ɏo��
              ListEkimId.AddItem (strWork)
           End If
        Loop
         
        '�t�@�C���N���[�Y
        Close #intFileNo
      Else
        '�G���[�R�[�h�F�ُ�
        Exit Function
     End If
   Else
     '�������ʁF�ُ펞�͉������Ȃ�
   End If
Exit Function

'VB�G���[����
Error_psVersionDisp:
    'V1.21.0.1 ADD  START
    If intFileNo > 0 Then
        Close #intFileNo
    End If
    'V1.21.0.1 ADD  END
    '�u�w���@��ID�m�F��ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfChangeFolderName
'//  �@�\����  : �t�H���_�p�X�ϊ�����
'//  �@�\�T�v  : INI�t�@�C�����擾�����t�H���_��`�̕ϊ����s���B
'//
'//              �^        ����         �Ӗ�
'//  ����      : String sFolderName    [IN]INI��`
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Function pfChangeFolderName(sFolderName As String) As String
   Dim iPath As Integer
   Dim sRootPath As String
   Dim sFolder As String
      
   '�u���v�ʒu���擾
   iPath = InStr(sFolderName, "\")
   If iPath = 0 Then
     sRootPath = Mid(sFolderName, 1)
   Else
     '�u���v�O��������擾
     sRootPath = Mid(sFolderName, 1, iPath - 1)
     '�u���v�㕶������擾
     sFolder = Mid(sFolderName, iPath + 1)
   End If
   Select Case sRootPath
      Case APL
        '�A�v�����[�g
        sRootPath = PATH_IDU_APP
      Case LOG
        '���O���[�g
        sRootPath = PATH_IDU_LOG
      Case Data
        'DB���[�g
        sRootPath = PATH_IDU_DB
      Case BACKUP
        '�o�b�N�A�b�v���[�g
        sRootPath = PATH_BUC
   End Select
    '�p�X�A��
    pfChangeFolderName = sRootPath + "\" + sFolder
End Function
