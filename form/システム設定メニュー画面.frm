VERSION 5.00
Begin VB.Form frmSystemSetteiMenu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�����[�g�����e�i���X"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Left            =   6600
      Top             =   6240
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�����ԏ��ݒ�"
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
      Index           =   5
      Left            =   6360
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "ID�T�[�o�n�I�Ɛݒ�"
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
      Index           =   4
      Left            =   2040
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "��ԊĎ��@�\�ݒ�"
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
      Index           =   3
      Left            =   6360
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "���ԑѕʃf�[�^�ݒ�"
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
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�V�X�e�����t�ݒ�"
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
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "���u�[�g�ݒ�"
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
      Index           =   1
      Left            =   6360
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " �����e�i���X   ��ʂ֖߂�"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�V�X�e���ݒ�"
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
      TabIndex        =   3
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmSystemSetteiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmSystemSetteiMenu.frm
'//  �p�b�P�[�W���F�V�X�e���ݒ胁�j���[���
'//  �T�v        �F���O�Ǘ����j���[���
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Option Explicit



Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FForm_Activate
'//  �@�\����    �F�V�X�e���ݒ胁�j���[���(�A�N�e�B�u��)
'//  �@�\�T�v    �F��ʍĕ\���������s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FForm_Deactivate
'//  �@�\����    �F�V�X�e���ݒ胁�j���[���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v    �F���[����M�p�̃^�C�}��~
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FForm_Load
'//  �@�\����    �FForm_Load������
'//  �@�\�T�v    �FForm_Load���������s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    '�u�V�X�e���ݒ胁�j���[��� �\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    ' /////////////////////////////////////////////////////
    ' IDU�k�ރ`�F�b�N
    psIDUCheck

    If pbIDUSts = 1 Then
       cmdFixedExe(3).Visible = False       ' ��ԊĎ��@�\�ݒ�
       cmdFixedExe(4).Visible = False       ' ID�T�[�o�n�I�Ɛݒ�
       cmdFixedExe(5).Visible = False       ' �����ԏ��ݒ�t
    End If
   
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdFixedExe_Click
'//  �@�\����    �F�e�t��������
'//  �@�\�T�v    �F�t�����ɂ���ʑJ�ڂ���B
'//
'//                 �^          ����            �Ӗ�
'//  ����        �F Integer     Index           �����t�C���f�b�N�X�l
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
  
   Dim udtMail As ML_DISP_INF          '��ʕ\���v��
   Dim iResponse As Integer            '���b�Z�[�W�{�b�N�X�߂�l

   On Error Resume Next
  
  
    Select Case Index
        Case 0                                 ' �V�X�e�����t�ݒ�
            '�u�V�X�e���ݒ胁�j���[��ʁF�V�X�e�����t�ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_SYSDATE_BUTTOM, 0)
            Load frmSystemDateSettei
            frmSystemDateSettei.Show 1
        Case 1                                 ' ���u�[�g�����ݒ�
            '�u�V�X�e���ݒ胁�j���[��ʁF���u�[�g�����ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_REBOOT_BUTTOM, 0)
            Load frmRebootTimeSettei
            frmRebootTimeSettei.Show 1
        Case 2                                 ' ���ԑѕʃf�[�^�ݒ�
            '�u�V�X�e���ݒ胁�j���[��ʁF���ԑѕʃf�[�^�ݒ�ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_TIMEDATA_BUTTOM, 0)
            Load frmTimeDataSettei
            frmTimeDataSettei.Show 1
        Case 3                                 ' ��ԊĎ��@�\�ݒ�
            '�u�V�X�e���ݒ胁�j���[��ʁF��ԊĎ��@�\�ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_IDCONDITION_BUTTOM, 0)
        
            '��ʕ\���v���i��ԊĎ��@�\�ݒ�j��ID����ɑ��M����
            If (SendMessageDispInfo(ML_DT_IDU_SYSTEM_CONDITION) = False) Then
         
                iResponse = MsgBox("��ԊĎ��@�\�ݒ�t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "��ԊĎ��@�\��ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
            End If
        
        Case 4                                 ' ID�T�[�o�n�I�Ɛݒ�
            '�u�V�X�e���ݒ胁�j���[��ʁFID�T�[�o�n�I�Ɛݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_IDJOBTIME_BUTTOM, 0)
        
            '��ʕ\���v���iID�T�[�o�n�I�Ɛݒ�j��ID����ɑ��M����
            If (SendMessageDispInfo(ML_DT_IDU_SYSTEM_JOBTIME) = False) Then
         
                iResponse = MsgBox("ID�T�[�o�n�I�Ɛݒ�t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "ID�T�[�o�n�I�Ɛݒ��ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
            End If
        Case 5                                 ' �����ԏ��z�M�ݒ�
            '�u�V�X�e���ݒ胁�j���[��ʁF�����ԏ��z�M�ݒ�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_IDPERIOD_BUTTOM, 0)
            
            '��ʕ\���v���i�����ԏ��z�M�ݒ�j��ID����ɑ��M����
            If (SendMessageDispInfo(ML_DT_IDU_SYSTEM_PERIOD) = False) Then
         
                iResponse = MsgBox("�����ԏ��z�M�ݒ�t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "�����ԏ��z�M�ݒ��ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
            End If
    End Select

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdReturn_Click
'//  �@�\����    �F�߂�t��������
'//  �@�\�T�v    �F�߂�t�����������������s���B
'//
'//                 �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '�u�V�X�e���ݒ胁�j���[��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_SETTEI_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FSendMessageDispInfo
'//  �@�\����    �F��ʕ\����Ԓʒm
'//  �@�\�T�v    �F��ʕ\����Ԓʒm���s���B
'//
'//                 �^      ����                �Ӗ�
'//  ����         : Long    lDispInfo           ��ʗv�����
'//
'//  �߂�l       : TRUE    �i����j
'//                 FALSE   �i�ُ�j
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l         :
'///////////////////////////////////////////////////////////////////////////////////////////////
Private Function SendMessageDispInfo(ByVal lDispInfo As Long) As Boolean

    Dim udtMail As ML_DISP_INF          '��ʕ\���v��
    Dim bRet As Boolean                 '���[�����M�����߂�l
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    '��ʕ\���v����ID����ɑ��M����
    udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
    udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
    udtMail.udtlHeader.dwProid = RHOSHU_ID
    udtMail.udtlHeader.dwSubArea = 0
    udtMail.dwDisp_Type = lDispInfo
    bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
    If bRet = False Then
        '�u��ʕ\���v�����[�����M�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
    Else
   
        '�u��ʕ\���v�����[�����M����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
    End If
    
    SendMessageDispInfo = bRet

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmSystemSetteiMenu.Caption, False
        pfFormActive (frmSystemSetteiMenu.hwnd)
    End If
End Sub
