VERSION 5.00
Begin VB.Form frmLogMenu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "���O�Ǘ�"
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�����"
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
   Begin VB.Timer tmrMail 
      Left            =   6840
      Top             =   6360
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�k�c�t"
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
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�h�b�l"
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
      TabIndex        =   4
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�h�c�t"
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
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "���D�@"
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
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�����Ď���"
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
      Caption         =   "���O�Ǘ�"
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
      TabIndex        =   6
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmLogMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmLogMenu.frm
'//  �p�b�P�[�W���F���O�Ǘ����j���[���
'//
'//  �T�v�F���O�Ǘ����j���[���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�q�x�s���O�Ǘ���ʒǉ�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���O�Ǘ����j���[���(�A�N�e�B�u��)
'//  �@�\�T�v  : ��ʍĕ\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
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
'//  �@�\����  : ���O�Ǘ����j���[���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : ���O�Ǘ����j���[���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    '�u���O�Ǘ����j���[��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'IDU�k�ރ`�F�b�N
    psIDUCheck

    If pbIDUSts = 1 Then
      'IDU�Ɩ���\��
       cmdFixedExe(1).Visible = False
       cmdFixedExe(4).Visible = False
    End If
   'V1.3.0.1 ADD START
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �t�����ɂ���ʑJ�ڂ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�����t�C���f�b�N�X�l
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-23  CODED BY  [TCC] M.Matsumoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
   Dim udtMail As ML_DISP_INF          '��ʕ\���v��
   Dim iResponse As Integer            '���b�Z�[�W�{�b�N�X�߂�l
   Dim bRet As Boolean                 '���[�����M�����߂�l
   Dim lngErrCode As Long              '�G���[�R�[�h

   On Error Resume Next

 Select Case Index
        Case 0                                 '���O�Ǘ�(�Ď���)
           '�u���O�Ǘ����j���[��ʁF�Ď��Ֆt�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_KANSI_BUTTOM, 0)
           Load frmKansiLogKanri
           frmKansiLogKanri.Show 1
        Case 1                                 '���O�Ǘ�(IDU)
           '�u���O�Ǘ����j���[��ʁF�h�c�t�t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_IDU_BUTTOM, 0)
           Load frmIDULogkanri
           frmIDULogkanri.Show 1
        Case 2                                 '���O�Ǘ�(LDU)
           '�u���O�Ǘ����j���[��ʁF�k�c�t�t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_LDU_BUTTOM, 0)
           Load frmLDULogkanri
           frmLDULogkanri.Show 1
        Case 3                                 '���O�Ǘ�(���D�@)
          '�u���O�Ǘ����j���[��ʁF���D�@�t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_KAISATUKI_BUTTOM, 0)

           '��ʕ\���v��(���D�@)��LD����ɑ��M����
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_KAISATUKI_LOG
            bRet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
              '�u���O�Ǘ����j���[��ʁF��ʕ\���v�����[�����M�ُ�v���O�o��
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
                iResponse = MsgBox("���D�@�t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "���D�@���O�Ǘ���ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
                Exit Sub
            End If
              '�u���O�Ǘ����j���[��ʁF��ʕ\���v�����[�����M����v���O�o��
               Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        Case 4                                 '���O�Ǘ�(����IC-M)
          '�u���O�Ǘ����j���[��ʁF����h�b�|�l�t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_ICM_BUTTOM, 0)
            
            '��ʕ\���v��(����IC-M)��ID����ɑ��M����
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_HANTEI_LOG
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
                '�u���O�Ǘ����j���[��ʁF��ʕ\���v�����[�����M�ُ�v���O�o��
                lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                 Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
         
                iResponse = MsgBox("����IC-M�t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "����IC-M���O�Ǘ���ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
                Exit Sub
            End If
            '�u���O�Ǘ����j���[��ʁF��ʕ\���v�����[�����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
' EG20 V2.1.0.1[Mainte_03_01] �폜�J�n
'    'V1.6.0.1 ADD START
'      Case 5                                 '���O�Ǘ�(RYT)
'           '�u���O�Ǘ����j���[��ʁF�q�x�s�t�����v���O�o��
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_RYT_BUTTOM, 0)
'           Load frmRYTSyusyu
'           frmRYTSyusyu.Show 1
'    'V1.6.0.1 ADD END
' EG20 V2.1.0.1[Mainte_03_01] �폜�I��
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
      Case 5                                 '���O�Ǘ��i�����j
           '�u���O�Ǘ����j���[��ʁF�����t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_TAKU_BUTTOM, 0)
           'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'           Load frmRYTSyusyu
'           frmRYTSyusyu.Show 1
           'EG20 V2.1.0.1 DEL END
           'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
           Load frmTakuLogKanri
           frmTakuLogKanri.Show 1
           'EG20 V2.1.0.1 ADD START
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�����e�i���X��ʂ֖߂�v�t��������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      :�Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
   On Error Resume Next
   
   '�u���O�Ǘ����j���[��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_MENU_GAMEN_END, 0)
    Unload Me
End Sub

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
        AppActivate frmLogMenu.Caption, False
        pfFormActive (frmLogMenu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
