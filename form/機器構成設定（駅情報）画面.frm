VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmKikiData 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�@��\���ݒ�"
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
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "   �ݺ��޺��    ���@����`��ʂ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   7
      Left            =   7200
      TabIndex        =   14
      Top             =   7800
      Width           =   2175
   End
   Begin VB.ComboBox CmbCornerName 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8280
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   13
      Top             =   480
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "���D�@��ʂ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   6
      Left            =   7200
      TabIndex        =   11
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   120
      Top             =   480
   End
   Begin VB.ComboBox cmbEkiInfo 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   " �@����ݒ�   ��ʂ֖߂�"
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
      TabIndex        =   8
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "�}�̎�O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   5
      Left            =   4850
      TabIndex        =   6
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "�ꎞ�ۑ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "�ꎞ�ۑ��f�[�^ �捞"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   3
      Left            =   2450
      TabIndex        =   4
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "�ݒ蔽�f"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "�}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   1
      Left            =   2450
      TabIndex        =   2
      Top             =   7800
      Width           =   2175
   End
   Begin VB.TextBox txtDummy 
      Height          =   330
      IMEMode         =   3  '�̌Œ�
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "�}�̓���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   4
      Left            =   4850
      TabIndex        =   1
      Top             =   7800
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   6195
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10927
      _Version        =   393216
      Rows            =   12
      Cols            =   4
      FixedCols       =   2
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblEkiName 
      Caption         =   "�w���F����������������������������������������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�@��\���ݒ�i�w���j"
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
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �F�@����ݒ�i�w���j���.frm
'//  �p�b�P�[�W���F�@����ݒ�i�w���j��ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�u������ʂցv�t�����ǉ�
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//                 �R���s���[�^���A�l�b�g���[�N�ύX�����ǉ�
'//                 �f�B�X�N���擾�ʒu�ύX
'//                 �t�@�C�����������폜
'//                 �}�̃t�@�C�������Œ薼�̂ɕύX
'//                 ��ʃ��b�N�����^��ʃ��b�N���������ǉ�
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                 �u�ꎞ�ۑ��f�[�^�捞�v�t�������C��
'//                  �{�^�����̕ύX�ɂ��|�b�v�A�b�v�ύX
'//     REVISIONS :(1.17.0.1) 2009-12-24  REVISED BY [TCC] E.Watanabe
'//                 �s��C��
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ��ʍđO�ʕ\���C��(�s��C��)
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//                 �J�[�\���ړ��̏������폜
'//                 �ݒ蔽�f�{�^���������ꂸ�ɉ�ʑJ�ڂ���Ƃ��̌x���\����ǉ�
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �|�b�v�A�b�v�m�F��ʂ̃^�C�g���C��
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(2.7.0.1) 2011-01-18  REVISED BY [TCC] S.Terao
'//                 �d�f�q_�i�d�Ή�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.12.0.1) 2012-05-18  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   '���C���^�C�}�̃C���^�[�o���l
Private bScroll As Boolean
Private GamenDataTbl() As EKI_DATA_TBL                  '��ʕ\���p�f�[�^�e�[�u��(�z��̗v�f[0]�͖��g�p)
Private KikiDataUpDateFlg As Boolean                    '�@����f�[�^�X�V�t���O

Private SetteiHaneiFlg As Boolean                       '�ݒ蔽�f�t���O     'V1.20.0.1 ADD

Private Const TITOL_EKI_NAME = "�w���F"                 '�w���^�C�g��       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �@����ݒ�i�w���j���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �őO�O�\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.17.0.1) 2009-12-24  REVISED BY [TCC] E.Watanabe
'//                 �s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    '�G���[���[�`����錾
    On Error Resume Next
    
    '����ʍőO�ʕ\���������s���B
    pfFormActive (hwnd)
    
'V1.17.0.1 ADD START
    '�t�H�[�J�X�ʒu��ݒ�
    cmdCancel.SetFocus
'V1.17.0.1 ADD END
    
    '�^�C�}���N������
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �@����ݒ�i�w���j���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �@����ݒ�i�w���j���(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                �ݒ蔽�f�t���O�ǉ�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim bRet                 As Boolean         '�֐��߂�l
    Dim lErrCode             As Long            '�G���[�R�[�h
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_START, 0)
    
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
    
    '�@����f�[�^�X�V�t���O�ݒ�i�X�V�ݒ�j
    KikiDataUpDateFlg = True
    
    '�@����ݒ�i�w���j�C���[�W�t�@�C���쐬
    bRet = dllGetKikiIniData(0, 0, KIKI_DATA_SET_EKI_INFO_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
    If bRet = False Then
        '�@����ݒ�i�w���j�C���[�W�t�@�C���폜
        Kill KIKI_DATA_SET_EKI_INFO_FILE
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
    End If
    
    '�w���R���{�{�b�N�X�����l�ݒ�
    cmbEkiInfo.Clear
    cmbEkiInfo.AddItem "�Ď�"
    cmbEkiInfo.AddItem "�l�b�g���[�N"
    cmbEkiInfo.ListIndex = 0
    
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    '�R�[�i�ݒ�R���{�{�b�N�X�̏���������
    Call InitCornerComboBox
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
    
    '�@����f�[�^�X�V�t���O�ݒ�i�ʏ�ݒ�j
    KikiDataUpDateFlg = False
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '�ݒ蔽�f�t���O�i�ύX�Ȃ��j
    SetteiHaneiFlg = False     'V1.20.0.1 ADD
    
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
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
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ��ʍđO�ʕ\���C��(�s��C��)
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
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
'                AppActivate frmInputMstData.Caption, False     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
                AppActivate frmKikiData.Caption, False          ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
                pfFormActive (frmKikiData.hwnd)                 ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '�u�ێ瑀���v���O�������M�v���v����M�����ꍇ
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
                    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f�����ݒ蔽�f����")
                Else
                    iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "���f�����ݒ蔽�f����")
                End If
                Call SetEnableTrue
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
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                �ݒ蔽�f�t�̖��������b�Z�[�W�ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    Dim iResponse           As Integer          'MsgBox�߂�l    'V1.20.0.1 ADD
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    'V1.20.0.1 ADD START
    If SetteiHaneiFlg = True Then
        iResponse = MsgBox("��ʕ\�����ɐݒ肳�ꂽ�f�[�^�������܂��B" & Chr(vbKeyReturn) & _
                            "��낵���ł����H", _
                            vbYesNo + vbQuestion, _
                            "�ݒ蔽�f�t������")
        
        If iResponse = vbNo Then Exit Sub
    End If
    'V1.20.0.1 ADD END
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_END, 0)
    
    Unload Me

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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim strFileName          As String          '�t�@�C����
    Dim nCornerIndex         As Integer         ' �R�[�i�I�����

    '�G���[���[�`����錾
    On Error Resume Next

' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    If CmbCornerName.ListIndex < 0 Then
        Exit Sub
    End If
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��

    '�����l�ݒ�
    strFileName = ""                            '�t�@�C����
    cmbEkiInfo.Enabled = False                  '�w���R���{�{�b�N�X�I��s�ݒ�
    CmbCornerName.Enabled = False               ' �R�[�i�I�𕔑I��s��      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    LblEkiName.Caption = TITOL_EKI_NAME         '�w�����x��������           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    '----------------------------------------------------
    '�O���b�h�^�C�g���ݒ�
    '----------------------------------------------------
    Call sDispGridTitol
    Erase GamenDataTbl
    ReDim GamenDataTbl(0)
    
    '�@����f�[�^�X�V�t���O�`�F�b�N
    If KikiDataUpDateFlg = True Then
        Erase KikiDataTbl
        ReDim KikiDataTbl(0)
        Call pfKikiDataSet
    End If
    
    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '�O���b�h�f�[�^���N���A����
        Call sDispDataClear(1, GridIni.Rows)

        '�����t�����s�\�ݒ�
        CmdKikiSetMenu(0).Enabled = False           '�@��\�����ڐݒ蔽�f
        CmdKikiSetMenu(1).Enabled = False           '�@��\�����ڔ}�̏o��
        CmdKikiSetMenu(2).Enabled = False           '�@��\�����ړ����ۑ�

        Exit Sub
        
    End If
    
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    '----------------------------------------------------
    '�w�����x���X�V
    '----------------------------------------------------
    LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo(NotEkiVer)
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
    
    '�@��\�����i�w���j�C���[�W�t�@�C������
    strFileName = Dir(KIKI_DATA_SET_EKI_INFO_FILE)
    
    '�t�@�C�������݂����ꍇ
    If strFileName <> "" Then
    
        '�O���b�h�f�[�^���ݒ�
'        Call sDispDataSet(pfGetCodeDaiBunrui(cmbEkiInfo))                              ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        nCornerIndex = CmbCornerName.ListIndex
        Call sDispDataSet(pfGetCodeDaiBunrui(cmbEkiInfo), nCornerIndex + 1)
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
    
        cmbEkiInfo.Enabled = True                   '�w���R���{�{�b�N�X�I���ݒ�
        CmbCornerName.Enabled = True                ' �R�[�i�I�𕔑I��s��      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
        '�����t�����\�ݒ�
        CmdKikiSetMenu(0).Enabled = True            '�@��\�����ڐݒ蔽�f
        CmdKikiSetMenu(1).Enabled = True            '�@��\�����ڔ}�̏o��
        CmdKikiSetMenu(2).Enabled = True            '�@��\�����ړ����ۑ�

    Else
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKIINFO_IMAGE, 0)
        
        '�O���b�h�f�[�^���N���A����
        Call sDispDataClear(1, GridIni.Rows)
    
        '�����t�����s�\�ݒ�
        CmdKikiSetMenu(0).Enabled = False           '�@��\�����ڐݒ蔽�f
        CmdKikiSetMenu(1).Enabled = False           '�@��\�����ڔ}�̏o��
        CmdKikiSetMenu(2).Enabled = False           '�@��\�����ړ����ۑ�

    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispGridTitol
'//  �@�\����  : �O���b�h�^�C�g�����ݒ菈��
'//  �@�\�T�v  : �O���b�h�̏����l�A�^�C�g����ݒ肷��
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
Private Sub sDispGridTitol()
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�O���b�h�^�C�g���ݒ�
    With GridIni
    
        '----------------------------------
        '�O���b�h�̏�����
        '----------------------------------
        .Clear
        .Width = 11550
        
        '----------------------------------
        '�O���b�h�Z�����ݒ�
        '----------------------------------
        .Rows = 9
        .Cols = 4
        
        '----------------------------------
        '�O���b�h���ݒ�
        '----------------------------------
        .ColWidth(0) = 400
        .ColWidth(1) = 3700
        .ColWidth(2) = 2500
'        .ColWidth(3) = 4825        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        .ColWidth(3) = 5000         ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        
        '----------------------------------
        '�^�C�g���ݒ�
        '----------------------------------
        '���ڐݒ�
        .Col = 1
        .Row = 0: .Text = "����"
        .CellAlignment = flexAlignCenterCenter

        '�ݒ�l�ݒ�
        .Col = 2
        .Text = "�ݒ�l"
        .CellAlignment = flexAlignCenterCenter

        '�ڍאݒ�
        .Col = 3
        .Text = "�ݒ�l�ڍ�"
        .CellAlignment = flexAlignCenterCenter
        
'        .RowHeight(0) = 700        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        .RowHeight(0) = 450         ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    End With
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispDataClear
'//  �@�\����  : �O���b�h�f�[�^���N���A����
'//  �@�\�T�v  : �O���b�h�f�[�^�����N���A����
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   intStartRow  �J�n�s�ʒu
'//              Integer   intEndRow    �I���s�ʒu
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear(intStartRow As Integer, intEndRow As Integer)
    
    Dim iLoopCnt             As Integer         '���[�v�J�E���^
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�O���b�h������
    With GridIni
            .Rows = intEndRow

        For iLoopCnt = intStartRow To intEndRow - 1

            '�ʔԐݒ�
            .Col = 0
            .Row = iLoopCnt: .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '���ڐݒ�
            .Col = 1
            .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '�ݒ�l�ݒ�
            .Col = 2
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
            .CellAlignment = flexAlignLeftCenter

            '�ڍאݒ�
            .Col = 3
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
            .CellAlignment = flexAlignLeftCenter

            .RowHeight(iLoopCnt) = 700
        Next

    End With
        
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispDataSet
'//  �@�\����  : �O���b�h�f�[�^���ݒ菈��
'//  �@�\�T�v  : �O���b�h�f�[�^����ݒ肷��
'//
'//              �^        ����         �Ӗ�
'//  ����      : Integer   iBunrui_Dai  �啪��
'//            : Integer   iCorner      �R�[�i  ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(2.7.0.1) 2011-01-18  REVISED BY [TCC] S.Terao
'//                 �d�f�q_�i�d�Ή�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sDispDataSet(iBunrui_Dai As Integer)                       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
Private Sub sDispDataSet(iBunrui_Dai As Integer, iCorner As Integer)    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    Dim intFileNumber       As Integer                      ' �t�@�C���|�C���^
    Dim iLoopCnt            As Integer                      ' ���[�v�J�E���^
    Dim iKikiDataCnt        As Integer                      ' �@����f�[�^�J�E���^
    Dim iRowCnt             As Integer                      ' �s�J�E���^
    
    Dim strBunrui_Dai       As String                       ' �啪��
    Dim strBunrui_Tyu       As String                       ' ������
    Dim strBunrui_Sho       As String                       ' ������
    Dim strNo               As String                       ' �ʔ�
    Dim strKomoku           As String                       ' ����
    Dim strKubun            As String                       ' �敪
    Dim strData             As String                       ' �ݒ�l
    Dim strSetShosai        As String                       ' �ݒ�l�ڍ�
                
    Dim byBuff()            As Byte                         ' �o�C�g�o�b�t�@
    Dim iLoopCnt2           As Integer
    Dim iBuffCnt            As Integer                      'V2.7.0.1 ADD�@�G���A�m�ې�
    
    Dim strCorner           As String                       ' �R�[�i    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Dim iCmpCorner          As Integer                      ' �R�[�i    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�����l�ݒ�
    iKikiDataCnt = 0
    iLoopCnt = 1
    GridIni.Rows = 1
    
    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '�@��\�����i�w���j�C���[�W�t�@�C�����I�[�v������B
    Open KIKI_DATA_SET_EKI_INFO_FILE For Input As #intFileNumber
    
    GridIni.Visible = False
    Do While Not EOF(intFileNumber)
        
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'        '�P �s�ǂݍ���
'        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strNo, _
'                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        '�P �s�ǂݍ���
        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strCorner, strNo, _
                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
        
        '�@����f�[�^�X�V�t���O�`�F�b�N
        If KikiDataUpDateFlg = False Then
            strData = KikiDataTbl(iKikiDataCnt).strData
            strData = StrConv(strData, vbUnicode)
        End If
        
        '�@����f�[�^�J�E���^�C���N�������g
        iKikiDataCnt = iKikiDataCnt + 1
        
        If iBunrui_Dai = strBunrui_Dai Then
        
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        ' �R�[�i����ǉ�
        ' �I�������R�[�i�A�������̓R�[�i���֌W�̃��R�[�h�͍̗p����
        iCmpCorner = CInt(strCorner)
        If ((iCorner = iCmpCorner) Or (iCmpCorner = 0)) Then
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
        
            With GridIni
                
                '�ő�s���C���N�������g
                .Rows = iLoopCnt + 1
                
                '�ʔԐݒ�
                .Col = 0
                .Row = iLoopCnt: If .Text = "" Then .Text = CStr(iLoopCnt)
                .CellAlignment = flexAlignLeftCenter
    
                '���ڐݒ�
                .Col = 1
                If .Text = "" Then .Text = strKomoku
                .CellAlignment = flexAlignLeftCenter
    
                '�ݒ�l�ݒ�
                .Col = 2
                .Text = strData
                .CellAlignment = flexAlignLeftCenter
    
                '��ʕ\���p�f�[�^�ۑ�
                ReDim Preserve GamenDataTbl(.Rows - 1)
                GamenDataTbl(.Rows - 1).iBunrui_Dai = CInt(strBunrui_Dai)     '�啪��
                GamenDataTbl(.Rows - 1).iBunrui_Tyu = CInt(strBunrui_Tyu)     '������
                GamenDataTbl(.Rows - 1).iBunrui_Sho = CInt(strBunrui_Sho)     '������
                GamenDataTbl(.Rows - 1).iBunrui_Corner = iCmpCorner           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
                
                '�ݒ�l
                If strData <> "" Then 'V2.7.0.1 ADD
                    byBuff = StrConv(strData, vbFromUnicode)
                'V2.7.0.1 ADD START
                Else
                   '���p�X�y�[�X����
                   '���p�X�y�[�X�ŕϊ��������Ă��A
                   '�u�ݒ蔽�f�v���͉�ʓ��e����������邽��
                   '���p�X�y�[�X�ϊ�����(32)��(0�j�ƂȂ�B
                    strData = "  "
                    byBuff = StrConv(strData, vbFromUnicode)
                End If
                'V2.7.0.1 ADD END
                '���I�z��̓��e�����O�p�����[�^�\���̂̐ÓI�z��Ɋi�[����B
                For iLoopCnt2 = 0 To UBound(GamenDataTbl(.Rows - 1).strData)
                    'Null�l�ɂȂ����珈���𔲂���B
                    If byBuff(iLoopCnt2) = vbVEmpty Then Exit For
                    
                    GamenDataTbl(.Rows - 1).strData(iLoopCnt2) = byBuff(iLoopCnt2)
                    
                    '���I�z��̍ő�v�f�ɂȂ����珈���𔲂���
                    If iLoopCnt2 = UBound(byBuff) Then Exit For
                Next
                 
                '�ڍאݒ�
                .Col = 3
                If .Text = "" Then .Text = strSetShosai
                .CellAlignment = flexAlignLeftCenter
    
                .RowHeight(iLoopCnt) = 700
        
                '�C���N�������g
                iLoopCnt = iLoopCnt + 1             '�\���s��
            
                '�\���f�[�^���P��ʂɕ\��������Ȃ��ꍇ
                If .Rows > 8 Then
                    '�X�N���[���o�[���A�O���b�h���L����
                    .Width = 11775
                End If
                
            End With
        
        End If          ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        End If
    
    Loop

    GridIni.Visible = True
    '�t�@�C�����N���[�Y����B
    Close #intFileNumber
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�\���s���ɖ����Ȃ��ꍇ�f�[�^���N���A����
    If GridIni.Rows < 9 Then
        '�O���b�h�f�[�^���N���A����
'        Call sDispDataClear(GridIni.Rows - 1, 9) 'V1.8.0.1 DEL
         Call sDispDataClear(GridIni.Rows, 9) 'V1.8.0.1 ADD
    End If

    Exit Sub

'�G���[����
Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
    
    '�O���b�h�^�C�g���ݒ�
    Call sDispGridTitol
    
    '�O���b�h�f�[�^���N���A����
'    Call sDispDataClear(1, GridIni.Rows - 1)   'V1.8.0.1 DEL
     Call sDispDataClear(1, GridIni.Rows)       'V1.8.0.1 ADD
     
    GridIni.Visible = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmbEkiInfo_Change
'//  �@�\����  : �w���I������
'//  �@�\�T�v  : �O���b�h�f�[�^���Đݒ肷��
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmbEkiInfo_Click()
    
    Dim iIndex          As Integer                  '�C���f�b�N�X
    
    '�G���[���[�`����錾
    On Error Resume Next

'V1.12.0.1 ADD START
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_EKIINFO_SELECT, 0)
    
    '��ʕ\������
    Call sDisp

'V1.12.0.1 ADD START
    '�S�{�^���������Ƃ���B
    Call SetEnableTrue
'V1.12.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : GridIni_Click
'//  �@�\����  : �O���b�h��I�����ꂽ���̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �_�~�[�e�L�X�g�̃Z�b�g
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Click()
    
    Dim iColSave        As Integer          '��ۑ��G���A
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�N���b�N���ꂽ�ʒu�Ƀ_�~�[�e�L�X�g���ړ����A�t�H�[�J�X�����킹��
    With GridIni
    
        '�ڍאݒ�l�͏������Ȃ�
        If .Col = 3 Then Exit Sub
        
        iColSave = .Col
        .Col = 0
        If .Text = "" Then
            .Col = iColSave
            Exit Sub
        End If
        .Col = iColSave
        
        txtDummy.Left = .Left + .CellLeft
        txtDummy.Top = .Top + .CellTop
        txtDummy.Width = .CellWidth
        txtDummy.Height = .CellHeight
        txtDummy.Text = .Text
        txtDummy.Visible = True
        txtDummy.SetFocus
        
        '�_�~�[�e�L�X�g�̍ŏI�Ƀt�H�[�J�X�ړ�
        SendKeys "{END}"
    
    End With

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : GridIni_KeyPress
'//  �@�\����  : �O���b�h���������ꂽ���̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �_�~�[�e�L�X�g�̃Z�b�g
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_KeyPress(KeyAscii As Integer)
    
    '�G���[���[�`����錾
    On Error Resume Next

' EG20 V3.5.0.1�ǉ��J�n
    If GridIni.Col = 3 Then
        Exit Sub
    End If
' EG20 V3.5.0.1�ǉ��I��

    '�N���b�N���ꂽ�ʒu�Ƀ_�~�[�e�L�X�g���ړ����A�t�H�[�J�X�����킹��
    With GridIni
        
        txtDummy.Left = .Left + .CellLeft
        txtDummy.Top = .Top + .CellTop
        txtDummy.Width = .CellWidth
        txtDummy.Height = .CellHeight
        If KeyAscii <> 13 Then
            txtDummy.Text = .Text & Chr(KeyAscii)
        Else
            txtDummy.Text = .Text
        End If
        txtDummy.Visible = True
        txtDummy.SetFocus
        
        '�_�~�[�e�L�X�g�̍ŏI�Ƀt�H�[�J�X�ړ�
        SendKeys "{END}"
    
    End With

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : GridIni_Scroll
'//  �@�\����  : �O���b�h���X�N���[���������̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �_�~�[�e�L�X�g�̔�\��
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Scroll()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�O���b�h���X�N���[�����ꂽ���A�_�~�[�e�L�X�g���\���ɂ���
    If bScroll = False Then
        txtDummy.Visible = False
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtDummy_Change
'//  �@�\����  : �_�~�[�e�L�X�g���ύX���ꂽ���̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �O���b�h�ւ̔��f
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                �ݒ蔽�f�t���O�ǉ�
'//                �󔒃Z�����͎��̉��������ǉ�
'//     REVISIONS :(EG20 V5.12.0.1) 2012-05-18  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.4.0.1) 2012-06-17 REVISED BY [TCC] H.Sugimoto
'//                �y���_���C���Ή��F���p�X�y�[�X�̓��͂�}�~����Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtDummy_Change()
    
    Dim iLoopCnt            As Integer                      ' ���[�v�J�E���^
    Dim iLoopCnt2           As Integer                      ' ���[�v�J�E���^
    Dim byBuff()            As Byte                         ' �o�C�g�o�b�t�@
    Dim szWork              As String                       ' ���[�N    ' EG20 V6.4.0.1�ǉ�
    
    '�G���[���[�`����錾
    On Error Resume Next
    
' EG20 V6.4.0.1�ǉ��J�n
    If InStr(txtDummy.Text, " ") > 0 Then
        szWork = Replace(txtDummy.Text, " ", "")
        txtDummy.Text = szWork
        MsgBox "�X�y�[�X�̓��͂ł��܂���B" & vbCrLf & _
                "���͓��e���m�F���Ă��������B", vbOKOnly + vbCritical, "�ݒ�l���ُ͈�"
        Exit Sub
    End If
' EG20 V6.4.0.1�ǉ��I��
    
    'V1.20.0.1 ADD START
    If GridIni.Text <> txtDummy.Text Then
        '�ݒ蔽�f�t���O�i�ύX����j
        SetteiHaneiFlg = True
    End If
    'V1.20.0.1 ADD END
    
    '�O���b�h�ɓ��͍��ڂ𔽉f������
    GridIni.Text = txtDummy.Text
    
    'V1.20.0.1 ADD START
    '�f�[�^���Ȃ��s���N���b�N�����ꍇ
    If UBound(GamenDataTbl) < GridIni.Row Then
        Exit Sub
    End If
    'V1.20.0.1 ADD END
    
    '��ʕ\���f�[�^�ۑ�
    byBuff = StrConv(GridIni.Text, vbFromUnicode)
    '���I�z��̓��e�����O�p�����[�^�\���̂̐ÓI�z��Ɋi�[����B
    Erase GamenDataTbl(GridIni.Row).strData
    For iLoopCnt2 = 0 To UBound(GamenDataTbl(GridIni.Row).strData)
        'Null�l�ɂȂ����珈���𔲂���B
        If byBuff(iLoopCnt2) = vbVEmpty Then Exit For
        
        GamenDataTbl(GridIni.Row).strData(iLoopCnt2) = byBuff(iLoopCnt2)
        
        '���I�z��̍ő�v�f�ɂȂ����珈���𔲂���
        If iLoopCnt2 = UBound(byBuff) Then Exit For
    Next
    
    For iLoopCnt = 0 To UBound(KikiDataTbl) - 1
    
        '��ʏ��Ƌ@����̑啪�ށA�����ށA�����ނ���v�����ꍇ
' EG20 V5.12.0.1�폜�J�n
'        If (GamenDataTbl(GridIni.Row).iBunrui_Dai = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
'           (GamenDataTbl(GridIni.Row).iBunrui_Tyu = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
'           (GamenDataTbl(GridIni.Row).iBunrui_Sho = KikiDataTbl(iLoopCnt).iBunrui_Sho) Then
' EG20 V5.12.0.1�폜�I��
' EG20 V5.12.0.1�ǉ��J�n
        If (GamenDataTbl(GridIni.Row).iBunrui_Dai = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
           (GamenDataTbl(GridIni.Row).iBunrui_Tyu = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
           (GamenDataTbl(GridIni.Row).iBunrui_Sho = KikiDataTbl(iLoopCnt).iBunrui_Sho) And _
           (GamenDataTbl(GridIni.Row).iBunrui_Corner = KikiDataTbl(iLoopCnt).iBunrui_Corner) Then
' EG20 V5.12.0.1�ǉ��I��
            '�@��\�����f�[�^�ۑ�
            '���I�z��̓��e�����O�p�����[�^�\���̂̐ÓI�z��Ɋi�[����B
            Erase KikiDataTbl(iLoopCnt).strData
            For iLoopCnt2 = 0 To UBound(KikiDataTbl(iLoopCnt).strData)
                'Null�l�ɂȂ����珈���𔲂���B
                If byBuff(iLoopCnt2) = vbVEmpty Then Exit For
                
                KikiDataTbl(iLoopCnt).strData(iLoopCnt2) = byBuff(iLoopCnt2)
                
                '���I�z��̍ő�v�f�ɂȂ����珈���𔲂���
                If iLoopCnt2 = UBound(byBuff) Then Exit For
            Next
            Exit For
        End If

    Next

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtDummy_KeyDown
'//  �@�\����  : �L�[�{�[�h�������̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �_�~�[�e�L�X�g�̃Z�b�g
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                 �J�[�\���ړ��̏������폜
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iColSave        As Integer          '��ۑ��G���A
    Dim iRowSave        As Integer          '�s�ۑ��G���A
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '����L�[���������ꂽ���A���L�̏������s��
    bScroll = True
    On Err GoTo ShoriErr
    
    With GridIni
        'V1.20.0.1 DEL START
'        '�����������ꂽ��
'        If KeyCode = 38 Then
'            If .Row <> 0 And .Row <> 1 Then
'                '�Z������Ɉ�ړ�
'                .Row = .Row - 1
'            End If
'        '���A�܂���enter���������ꂽ��
'        ElseIf KeyCode = 40 Or KeyCode = 13 Then
        'V1.20.0.1 DEL END
        If KeyCode = 13 Then    'V1.20.0.1 ADD

            iColSave = .Col
            iRowSave = .Row
            .Col = 0
            .Row = .Row + 1
            If .Text = "" Then
                .Col = iColSave
                .Row = iRowSave
                Exit Sub
            End If
            .Col = iColSave
            .Row = iRowSave
            
            If .Row <> .Rows - 1 Then
                '�Z�������Ɉ�ړ�
                .Row = .Row + 1
            End If
        'V1.20.0.1 DEL START
'        '���A�܂��́����������ꂽ��
'        ElseIf KeyCode = 37 Or KeyCode = 39 Then
'            Exit Sub
        'V1.20.0.1 DEL END
        End If

        '�_�~�[�e�L�X�g�̃Z�b�g
        txtDummy.Left = .Left + .CellLeft
        txtDummy.Top = .Top + .CellTop
        txtDummy.Width = .CellWidth
        txtDummy.Height = .CellHeight
        txtDummy.Text = .Text
        txtDummy.Visible = True
        txtDummy.SetFocus
    End With
    bScroll = False

ShoriErr:

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtDummy_Change
'//  �@�\����  : �_�~�[�e�L�X�g����t�H�[�J�X���ړ��������̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �_�~�[�e�L�X�g���\���ɂ���
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtDummy_LostFocus()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�_�~�[�e�L�X�g���\���ɂ���
    txtDummy.Visible = False

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : CmdKikiSetMenu_Click
'//  �@�\����  : �e�t����������
'//  �@�\�T�v  : �e�t���������ɏ]��
'//
'//              �^        ����     �@�@�@�Ӗ�
'//  ����      : Integer�@ Index          �I��t�̃C���f�b�N�X
'//
'//              �^        �l        �@�@ �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�u������ʂցv�t�����ǉ�
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                �ݒ蔽�f�t�̖��������b�Z�[�W�ǉ�
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmdKikiSetMenu_Click(Index As Integer)
    Dim iResponse           As Integer          'MsgBox�߂�l   'V1.20.0.1 ADD
    Dim bUnlock             As Boolean          ' ���b�N�����t���O      ' EG20 V3.0.0.2 �ǉ�

    '�G���[���[�`����錾
    On Error Resume Next
    
'V1.12.0.1 ADD START
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
'V1.12.0.1 ADD END
 ' EG20 V3.0.0.2 �ǉ��J�n
' ���������t�ɉ����ă��b�N�����𐧌�����
' �����[����M��҂���
    bUnlock = True
' EG20 V3.0.0.2 �ǉ��I��
   
    Select Case Index
        
        Case 0                                 ' �@��\�����ڐݒ蔽�f
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_INSTOL, 0)
            
            '�@��\�����ڐݒ蔽�f����
            Call sInstolKikiData
            bUnlock = False                     ' EG20 V3.0.0.2 �ǉ�

        Case 1                                 ' �@��\�����ڔ}�̏o��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_OUTPUT, 0)
            
            '�@��\�����ڔ}�̏o�͏���
            Call sKikiDataOutPut
    
        Case 2                                 ' �@��\�����ړ����ۑ�
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_SAVE, 0)
            
            '�@��\�����ړ����ۑ�����
            Call sKikiDataSave
        
        Case 3                                 ' �@��\���ݒ�f�[�^�I��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_KIKIDATA_SELECT, 0)
            
            '�@��\���ݒ�f�[�^�I������
            Call sKikiDataSelect
    
        Case 4                                 ' �}�̓���
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_EKIINFO_GAMEN_MEDIUM_INPUT, 0)
            
            '�}�̓��͏���
            Call sInputMedium
    
        Case 5                                 ' �}�̎�O
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
            
            '�}�̎�O����
            Call pfRemove(Me)
'V1.4.0.1 ADD START
        Case 6                                 ' ������ʂ�
            'V1.20.0.1 ADD START
            If SetteiHaneiFlg = True Then
                iResponse = MsgBox("��ʕ\�����ɐݒ肳�ꂽ�f�[�^�������܂��B" & Chr(vbKeyReturn) & _
                                    "��낵���ł����H", _
                                    vbYesNo + vbQuestion, _
                                    "�ݒ蔽�f�t������")
                If iResponse = vbNo Then
                    '�S�{�^���������Ƃ���B
                    Call SetEnableTrue
                    Exit Sub
                End If
            End If
            '�ݒ蔽�f�t���O�i�ύX�Ȃ��j
            SetteiHaneiFlg = False
            'V1.20.0.1 ADD END
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)
            Unload Me
            Load frmKikiDataGate
            frmKikiDataGate.Show 1
            Exit Sub     'V1.20.0.1 ADD
'V1.4.0.1 ADD END
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        Case 7                                 ' �G���R�[�h�R�[�i���@��ʂ�
            If SetteiHaneiFlg = True Then
                iResponse = MsgBox("��ʕ\�����ɐݒ肳�ꂽ�f�[�^�������܂��B" & Chr(vbKeyReturn) & _
                                    "��낵���ł����H", _
                                    vbYesNo + vbQuestion, _
                                    "�ݒ蔽�f�t������")
                If iResponse = vbNo Then
                    '�S�{�^���������Ƃ���B
                    Call SetEnableTrue
                    Exit Sub
                End If
            End If
            '�ݒ蔽�f�t���O�i�ύX�Ȃ��j
            SetteiHaneiFlg = False
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_SUBGATE, 0)
            
            '�\������ʃA�����[�h
            Unload Me
            
            '�G���R�[�h�R�[�i���@��ʕ\��
            Load frmKikiDataSubGate
            frmKikiDataSubGate.Show 1
            Exit Sub
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��

        Case Else
            '�����Ȃ�
            
    End Select
'V1.12.0.1 ADD START
    '�S�{�^���������Ƃ���B
' EG20 V3.0.0.2 �ǉ��J�n
    If bUnlock = True Then
        Call SetEnableTrue
    End If
' EG20 V3.0.0.2 �ǉ��I��
'    Call SetEnableTrue                 ' EG20 V3.0.0.2 �폜
'V1.12.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sInstolKikiData
'//  �@�\����  : �u�@��\�����ڐݒ蔽�f�v�t����������
'//  �@�\�T�v  : ��ʕ\���f�[�^��INI�t�@�C���֔��f����
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
'//     REVISIONS :(1.10.0.1) 2009-10-23   REVISED BY [TCC] D.Yamashita
'//                 �t�F�[�Y�R�c�����ڑΉ��@�L�����Z���s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-16   REVISED BY [TCC] C.Terui
'//                 �R���s���[�^���A�l�b�g���[�N�ύX�����ǉ�
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                �ݒ蔽�f�t���O�ǉ�
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �|�b�v�A�b�v�m�F��ʂ̃^�C�g���C��
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.76�C���Ή��z
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-08  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sInstolKikiData()

    Dim iResponse           As Integer          'MsgBox�߂�l
    Dim bRet                As Boolean          '�֐��߂�l
    Dim lErrCode            As Long             '�G���[�R�[�h
    Dim strFileName         As String           '�}�̃t�@�C����
    
    Dim bData()             As Byte             '�o�C�i���f�[�^
    Dim iLoopCnt            As Integer          '���[�v�J�E���^
    Dim bSysChange          As Boolean          '�R���s���[�^���A�l�b�g���[�N�ύX��������   'V1.12.0.1 ADD
    
    Dim lResult             As Long             ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�
' EG20 V30.0.3.1 �ǉ��J�n�i�v�Z�ɗ��p����ϐ���LONG�^�ɕύX�j
    Dim lLoop               As Long             ' ���[�v�J�E���^
    Dim lRecord             As Long             ' ���R�[�h
    Dim lIndex              As Long             ' �C���f�b�N�X
    Dim lSize               As Long             ' �T�C�Y
' EG20 V30.0.3.1 �ǉ��I���i�v�Z�ɗ��p����ϐ���LONG�^�ɕύX�j
    
    '�G���[���[�`����錾
    On Error Resume Next
'V1.8.0.1 DEL START
'    iResponse = MsgBox("�@��\���f�[�^���C���X�g�[�����܂��B��낵���ł����H" & Chr(vbKeyReturn) & _
'                        "���f�͍ċN����ɂȂ�܂��B", _
'                        vbYesNo + vbExclamation, _
'                        "�}�̓��͊m�F")
'V1.8.0.1 DEL END
'V1.8.0.1 ADD START
'V1.21.0.1 DEL START
'    iResponse = MsgBox("�@��\���f�[�^���C���X�g�[�����܂��B��낵���ł����H" & Chr(vbKeyReturn) & _
'                        "���f�͍ċN����ɂȂ�܂��B", _
'                        vbOKCancel + vbExclamation, _
'                        "�}�̓��͊m�F")
'V1.21.0.1 DEL END
'V1.21.0.1 ADD START
    iResponse = MsgBox("�@��\���f�[�^���C���X�g�[�����܂��B��낵���ł����H" & Chr(vbKeyReturn) & _
                        "���f�͍ċN����ɂȂ�܂��B", _
                        vbOKCancel + vbExclamation, _
                        "�ݒ蔽�f�m�F")
'V1.21.0.1 ADD END
'V1.8.0.1 ADD END
'    If iResponse = vbNo Then Exit Sub      'V1.10.0.1 DEL
    If iResponse = vbCancel Then
        Call SetEnableTrue                      ' EG20 V5.0.2.1�y����TR-No.76�C���Ή��z�ǉ�
        Exit Sub   'V1.10.0.1 ADD
    End If
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
' EG20 V30.3.0.1 �폜�J�n�i�v�Z�ɗ��p����ϐ���LONG�^�ɕύX�j
'    '�\���̔z����o�C�i���z��ɕϊ�
'    ReDim bData((UBound(KikiDataTbl) + 1) * Len(KikiDataTbl(0))) As Byte
'    For iLoopCnt = 0 To UBound(KikiDataTbl)
'          MoveMemory bData(iLoopCnt * Len(KikiDataTbl(0))), KikiDataTbl(iLoopCnt), Len(KikiDataTbl(iLoopCnt))
'    Next
' EG20 V30.3.0.1 �폜�I���i�v�Z�ɗ��p����ϐ���LONG�^�ɕύX�j
' EG20 V30.3.0.1 �ǉ��J�n�i�v�Z�ɗ��p����ϐ���LONG�^�ɕύX�j
    lSize = Len(KikiDataTbl(0))
    lRecord = UBound(KikiDataTbl)
    ReDim bData((lRecord + 1) * lSize) As Byte
    For lLoop = 0 To lRecord
        lIndex = lLoop * lSize
        MoveMemory bData(lIndex), KikiDataTbl(lLoop), lSize
    Next
' EG20 V30.3.0.1 �ǉ��I���i�v�Z�ɗ��p����ϐ���LONG�^�ɕύX�j
    
    '�@��\���f�[�^�C���X�g�[������
    bRet = dllInstolKikiData(KIKI_DATA_FILE, EKI_SETTI_FILE, bData(0), UBound(KikiDataTbl) + 1, lErrCode)

    If bRet = False Then
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
        
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
        
        '�ُ�I��
        'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̓��͌���") '1.21.0.1 DEL
        iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f�����ݒ蔽�f����") 'V1.21.0.1 ADD
        Call SetEnableTrue                      ' EG20 V5.0.2.1�y����TR-No.76�C���Ή��z�ǉ�
    Else
'        '���O�o��
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'
'        '����I��
'        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�̓��͌���")
'    End If
'V1.12.0.1 ADD START
            '�R���s���[�^���A�l�b�g���[�N�ύX����
            bSysChange = pfNetWorkChng(Me)
            If bSysChange = False Then
                '�ُ탍�O�o��
                Call pfOutPutErrLog(lErrCode)

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
                Call SetEnableTrue                      ' EG20 V5.0.2.1�y����TR-No.76�C���Ή��z�ǉ�
            Else
                
' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
                ' //////////////////////////////////////////////
                ' // �����v���O��������
                ' //////////////////////////////////////////////
                lResult = pubfuncTakuProgramData(2, EKI_SETTI_FILE)
                If lResult = 0 Then
                   '�v���O���X�o�[����������
                   Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                   ' �ُ�I��
                   iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f�����ݒ蔽�f����")
                   Call SetEnableTrue
                   Exit Sub
                ElseIf lResult = 1 Then
                   ' ���[�����M��
                   ' ���O�o��
                   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
                   ' �ݒ蔽�f�t���O�i�ύX�Ȃ��j
                   SetteiHaneiFlg = False
                    
                   Exit Sub
                End If
' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
                
                '���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
                
                '����I��
                'iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�̓��͌���")  'V1.21.0.1 DEL
                 iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "���f�����ݒ蔽�f����")  'V1.21.0.1 ADD
                '�ݒ蔽�f�t���O�i�ύX�Ȃ��j
                SetteiHaneiFlg = False      'V1.20.0.1 ADD
                Call SetEnableTrue                      ' EG20 V5.0.2.1�y����TR-No.76�C���Ή��z�ǉ�
            End If
        End If
'V1.12.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sKikiDataOutPut
'//  �@�\����  : �u�@��\�����ڔ}�̏o�́v�t����������
'//  �@�\�T�v  : �@��\���f�[�^�t�@�C�����O���}�̂ɏo�͂���
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
'//                 �f�B�X�N���擾�ʒu�ύX
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                  �{�^�����̕ύX�ɂ��|�b�v�A�b�v�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataOutPut()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
    Dim iResponse            As Integer         'MsgBox�߂�l

    Dim iRet                 As Integer         '���b�Z�[�W�{�b�N�X�߂�l
    Dim lSekuta              As Long            '�Z�N�^�i�N���X�^����j
    Dim lByte                As Long            '�o�C�g���i�Z�N�^����j
    Dim lKurasuta            As Long            '�t���[�N���X�^��
    Dim lDrive               As Long               '�h���C�u�̃N���X�^���i���v�j
    Dim strDrive             As String       '�h���C�u
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '�@��\���f�[�^�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(KIKI_DATA_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_KIKI_DATA, 0)
        
        '�ُ�I��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", _
                vbOKOnly + vbExclamation, _
                 "�f�[�^���x��"
        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '�}�̏o�͏���
    '----------------------------------------------------
    ''V1.20.0.1 DEL START
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
    'V1.20.0.1 DEL END
    
    'sWriteDir = pfDirSelection(strDrive, "�@��\���t�@�C�������ݐ�̃f�B���N�g���I��")    'V1.20.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
    If sWriteDir <> "" Then
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
'        FileCopy KIKI_DATA_FILE, sWriteDir & Dir(KIKI_DATA_FILE)                                       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        FileCopy KIKI_DATA_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(KIKI_DATA_FILE)    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '����I��
        'iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�@��\�����ڔ}�̏o�͌���") 'V1.13.0.1 DEL
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�̏o�͌���")              'V1.13.0.1 ADD
    
    End If
  
  Exit Sub
 
COPY_ERROR:

    '�ُ탍�O�o��
    Select Case Err.Number
        Case 61 ' �}�̏o�͋󂫗e�ʕs��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case 71 ' �}�̂Ȃ�
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_NOT_DISK, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    '�ُ�I��
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbInformation, "�@��\�����ڔ}�̏o�͌���") 'V1.8.0.1 DEL
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�@��\�����ڔ}�̏o�͌���") 'V1.8.0.1 ADD 'V1.13.0.1 DEL
     iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̏o�͌���") 'V1.13.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sKikiDataSave
'//  �@�\����  : �u�@��\�����ړ����ۑ��v�t����������
'//  �@�\�T�v  : �@��\���f�[�^�t�@�C�����w��t�H���_�ɏo�͂���
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
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 �t�@�C�����������폜
'//     REVISIONS :(1.13.0.1) 2009-11-19  REVISED BY [TCC] S.Terao
'//                 �t���ύX�ɂ��A�|�b�v�A�b�v�^�C�g���ύX
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSave()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
'    Dim sMyPath(1 To 3)      As String          '�t�@�C���p�X      'V1.12.0.1 DEL
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim iLoopCount           As Integer         '���[�v�J�E���^
    Dim intFileNo            As Integer         '�t�@�C���ԍ�

    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    strFileName = ""

'V1.12.0.1 DEL START
'    '----------------------------------------------------
'    '�@��\���f�[�^�t�@�C������
'    '----------------------------------------------------
'    strFileName = Dir(KIKI_DATA_FILE)
'
'    '�t�@�C�������݂��Ȃ��ꍇ
'    If strFileName = "" Then
'
'        '�ُ탍�O�o��
'        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_KIKI_DATA, 0)
'
'        '�ُ�I��
'        MsgBox "�}�̏o�͂���f�[�^������܂���B", _
'                vbOKOnly + vbExclamation, _
'                 "�f�[�^���x��"
'        Exit Sub
'
'    End If
'
'    '----------------------------------------------------
'    '�����ۑ��t�@�C������
'    '----------------------------------------------------
'    For iLoopCount = 1 To 3
'
'        '�t�@�C���p�X�擾
'        sMyPath(iLoopCount) = Replace(KIKI_DATA_S_FILE, "##", Format(iLoopCount, "0#"))
'
'        '�t�@�C������
'        strFileName = Dir(sMyPath(iLoopCount))
'
'        '�t�@�C�������݂��Ȃ��ꍇ
'        If strFileName = "" Then
'
'            intFileNo = FreeFile                                        '���g�p�̃t�@�C���ԍ����擾����
'            Open sMyPath(iLoopCount) For Output Access Write As #intFileNo
'            Close #intFileNo
'
'        End If
'
'    Next
'V1.12.0.1 DEL END

    '----------------------------------------------------
    '�����ۑ�����
    '----------------------------------------------------
'V1.12.0.1 ADD START
    iResponse = MsgBox("�@��\���ݒ���ꎞ�ۑ����܂��B" & vbCrLf & "��낵���ł����H", _
    vbOKCancel + vbQuestion, "�ꎞ�ۑ��m�F")
    
    If iResponse = vbCancel Then Exit Sub
    
    '�t�@�C������
    strFileName = Dir(KIKI_DATA_S_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then

        intFileNo = FreeFile                                        '���g�p�̃t�@�C���ԍ����擾����
        Open KIKI_DATA_S_FILE For Output Access Write As #intFileNo
        Close #intFileNo
    End If
    
   '�ꎞ�ۑ��t�@�C�����쐬����
    Name KIKI_DATA_S_FILE As KIKI_DATA_S_TEMP_FILE
'V1.12.0.1 ADD END
    
    '�t�@�C�����擾
'    sWriteDir = pfDispFileSelect("d:", FOLDER_KIKI_DATA, FILE_NAME_KIKI_DATA_S, "�����ۑ�̧�ّI��")    'V1.12.0.1 DEL
    sWriteDir = KIKI_DATA_S_FILE  'V1.12.0.1 ADD
    
    If sWriteDir <> "" Then
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
        FileCopy KIKI_DATA_FILE, sWriteDir
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
'V1.12.0.1 ADD START
        '�ꎞ�ۑ��t�@�C���폜
        Kill KIKI_DATA_S_TEMP_FILE
'V1.12.0.1 ADD END
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '����I��
        'iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�@��\�����ړ����ۑ�����")�@�@'V1.13.0.1 DEL
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�ꎞ�ۑ�����")     'V1.13.0.1 ADD
    End If
  
  Exit Sub
 
COPY_ERROR:

    '�ُ탍�O�o��
    Select Case Err.Number
        Case 61 ' �󂫗e�ʕs��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

'V1.12.0.1 ADD START
        '�t�@�C������
        strFileName = Dir(KIKI_DATA_S_FILE)
        If strFileName <> "" Then
            '�t�@�C���폜
            Kill KIKI_DATA_S_FILE
        End If
        '�t�@�C�����̂����ɖ߂�
        Name KIKI_DATA_S_TEMP_FILE As KIKI_DATA_S_FILE
'V1.12.0.1 ADD END

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    '�ُ�I��
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbInformation, "�@��\�����ړ����ۑ�����") 'V1.8.0.1 DEL
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�@��\�����ړ����ۑ�����")    'V1.8.0.1 ADD  'V1.13.0.1 DEL
     iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�ꎞ�ۑ�����")    'V1.13.0.1 ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sKikiDataSelect
'//  �@�\����  : �u�@��\���ݒ�f�[�^�I���v�t����������
'//  �@�\�T�v  : �@��\���f�[�^�����ۑ��t�@�C�����@��\���f�[�^�t�@�C���ɃR�s�[����
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
'//     REVISIONS :(1.12.0.1) 2009-11-16  REVISED BY [TCC] C.Terui
'//                 �t�@�C�����������폜
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                 �R�s�[�t�@�C���p�X�w����C��
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                �ݒ蔽�f�t���O�ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSelect()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
'    Dim sMyPath(1 To 3)      As String          '�t�@�C���p�X      'V1.12.0.1 DEL
    Dim sMyPath              As String          '�t�@�C���p�X       'V1.12.0.1 ADD
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim iLoopCount           As Integer         '���[�v�J�E���^
    Dim intFileNo            As Integer         '�t�@�C���ԍ�
    Dim bRet                 As Boolean         '�֐��߂�l
    Dim lErrCode             As Long            '�G���[�R�[�h

    '�G���[���[�`����錾
    On Error Resume Next
    
'V1.12.0.1 DEL START
'    '----------------------------------------------------
'    '�����ۑ��t�@�C������
'    '----------------------------------------------------
'    For iLoopCount = 1 To 3
'
'        '�t�@�C���p�X�擾
'        sMyPath(iLoopCount) = Replace(KIKI_DATA_S_FILE, "##", Format(iLoopCount, "0#"))
'
'        '�����l�ݒ�
'        strFileName = ""
'
'        '�t�@�C������
'        strFileName = Dir(sMyPath(iLoopCount))
'
'        '�t�@�C�������݂��Ȃ��ꍇ
'        If strFileName = "" Then
'
'            intFileNo = FreeFile                                        '���g�p�̃t�@�C���ԍ����擾����
'            Open sMyPath(iLoopCount) For Output Access Write As #intFileNo
'            Close #intFileNo
'        End If
'
'    Next
'V1.12.0.1 DEL END

    '----------------------------------------------------
    '�@��\���f�[�^�t�@�C���X�V����
    '----------------------------------------------------
'V1.12.0.1 ADD START
    iResponse = MsgBox("�@��\���ݒ�̈ꎞ�ۑ��f�[�^���捞�݂܂��B" & vbCrLf & "��낵���ł����H", _
    vbOKCancel + vbQuestion, "�ꎞ�ۑ��f�[�^�捞�m�F")
    
    If iResponse = vbCancel Then Exit Sub
    
    '�t�@�C������
    strFileName = Dir(KIKI_DATA_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then

        intFileNo = FreeFile                                        '���g�p�̃t�@�C���ԍ����擾����
        Open KIKI_DATA_FILE For Output Access Write As #intFileNo
        Close #intFileNo
    End If
    
   '�ꎞ�ۑ��t�@�C�����쐬����
    Name KIKI_DATA_FILE As KIKI_DATA_BACKUP_FILE
'V1.12.0.1 ADD END
    
    '�t�@�C�����擾
'    sWriteDir = pfDispFileSelect("d:", FOLDER_KIKI_DATA, FILE_NAME_KIKI_DATA_S, "�@��\��̧�ّI��")    'V1.12.0.1 DEL
'V1.12.0.1 ADD START
    strFileName = Dir(KIKI_DATA_S_FILE)
    sWriteDir = strFileName
'V1.12.0.1 ADD START
    If sWriteDir <> "" Then

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
        'FileCopy sWriteDir, KIKI_DATA_FILE�@'V1.13.0.1 DEL
         FileCopy KIKI_DATA_S_FILE, KIKI_DATA_FILE  'V1.13.0.1 ADD
        
        '�@����ݒ�i�w���j�C���[�W�t�@�C���쐬
        bRet = dllGetKikiIniData(0, 1, KIKI_DATA_SET_EKI_INFO_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)
'V1.12.0.1 ADD START
            '�t�@�C���폜
            Kill KIKI_DATA_FILE
            '�t�@�C�����̂����ɖ߂�
            Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
'V1.12.0.1 ADD END

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
            
            '�ُ�I��
            'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbInformation, "�}�̓��͌���") 'V1.8.0.1 DEL
            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̓��͌���")  'V1.8.0.1 ADD
            Exit Sub
        End If
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'V1.12.0.1 ADD START
        '�ꎞ�ۑ��t�@�C�����폜����
        Kill KIKI_DATA_BACKUP_FILE
'V1.12.0.1 ADD END
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
        
        '����I��
'        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�@��\���ݒ�f�[�^�I������") 'V1.13.0.1 DEL
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�ꎞ�ۑ��f�[�^�捞����")      'V1.13.0.1 ADD
    
        '�@����f�[�^�X�V�t���O�ݒ�i�X�V�ݒ�j
        KikiDataUpDateFlg = True
        '��ʕ\������
        Call sDisp
        '�@����f�[�^�X�V�t���O�ݒ�i�ʏ�ݒ�j
        KikiDataUpDateFlg = False
        
        '�ݒ蔽�f�t���O�i�ύX����j
        SetteiHaneiFlg = True       'V1.20.0.1 ADD
    End If
  
  Exit Sub
 
COPY_ERROR:

    '�ُ탍�O�o��
    Select Case Err.Number
        Case 61 ' �󂫗e�ʕs��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

'V1.12.0.1 ADD START
        '�t�@�C������
        strFileName = Dir(KIKI_DATA_FILE)
        If strFileName <> "" Then
            '�t�@�C���폜
             Kill KIKI_DATA_FILE
        End If
        '�t�@�C�����̂����ɖ߂�
        Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
'V1.12.0.1 ADD END
    
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
    '�ُ�I��
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbInformation, "�@��\���ݒ�f�[�^�I������") 'V1.8.0.1 DEL
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�@��\���ݒ�f�[�^�I������") 'V1.8.0.1 ADD  'V1.13.0.1 DEL
    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�ꎞ�ۑ��f�[�^�捞����")      'V1.13.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sInputMedium
'//  �@�\����  : �u�}�̓��́v�t����������
'//  �@�\�T�v  : �O���}�̂��@��\���f�[�^�t�@�C���ɃR�s�[����
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
'//                 �}�̃t�@�C�������Œ薼�̂ɕύX
'//                 �f�B�X�N���擾�ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                �ݒ蔽�f�t���O�ǉ�
'//                �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.77�C���Ή��z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sInputMedium()

    Dim iResponse               As Integer      'MsgBox�߂�l
    Dim bRet                    As Boolean      '�֐��߂�l
    Dim lErrCode                As Long         '�G���[�R�[�h
    Dim strFileName             As String       '�}�̃t�@�C����
    
    Dim iRet                    As Integer      '���b�Z�[�W�{�b�N�X�߂�l
    Dim lSekuta                 As Long         '�Z�N�^�i�N���X�^����j
    Dim lByte                   As Long         '�o�C�g���i�Z�N�^����j
    Dim lKurasuta               As Long         '�t���[�N���X�^��
    Dim lDrive                  As Long         '�h���C�u�̃N���X�^���i���v�j
    Dim strDrive                As String       '�h���C�u
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g  'V1.20.0.1 ADD
    
    '�G���[���[�`����錾
    On Error Resume Next
'V1.12.0.1 ADD START
    iResponse = MsgBox("�@��\���ݒ�̔}�̓��͂��s���܂��B" & vbCrLf & "��낵���ł����H", _
    vbOKCancel + vbQuestion, "�}�̓��͊m�F")
    
    'V1.20.0.1 DEL START
    'If iResponse = vbCancel Then Exit Sub
''V1.12.0.1 ADD END
    '
    ''�f�B�X�N�����擾
''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)      'V1.12.0.1 DEL
    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)       'V1.12.0.1 ADD
    '
    'If lDrive = 0 Then
    '    strDrive = "d:"
    'Else
''        strDrive = "a:"        'V1.12.0.1 DEL
    '    strDrive = "H:"         'V1.12.0.1 ADD
    'End If
    '
    ''�}�̃t�@�C�����擾
''    strFileName = pfFileSelection(strDrive, "*.csv", "�}�̓���̧�ّI��")          'V1.12.0.1 DEL
    'strFileName = pfFileSelection(strDrive, "KIKI_DATA.CSV", "�}�̓���̧�ّI��")   'V1.12.0.1 ADD
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    If iResponse = vbCancel Then
        Set objFso = Nothing
        Exit Sub
    End If
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
'    CommonDialog1.Filter = "�@��\���f�[�^�t�@�C���iKIKI_DATA.CSV�j|KIKI_DATA.CSV|"    ' EG20 V5.0.2.1�폜
    CommonDialog1.Filter = "�@��\���f�[�^�t�@�C���iKIKI_DATA.CSV�j|*KIKI_DATA.CSV|"    ' EG20 V5.0.2.1�ǉ�
    '�t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    '�I�������t�@�C�������擾
    strFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
    
    Call ChDrive("D")  'V2.5.0.1 ADD
    
    '�t�@�C�����݃`�F�b�N
    If strFileName <> "" Then

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��

        On Error GoTo COPY_ERROR

' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL Start
        '�t�@�C���R�s�[
'        FileCopy strFileName, KIKI_DATA_FILE
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL End
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
        '�ꎞ�ۑ��t�H���_�Ƀf�[�^���R�s�[���ǎ��p����������
        If pfChangeAttrNormal(strFileName, PATH_HOSHUTMP_KIKI_DATA, KIKI_DATA_FILE) = False Then
            GoTo COPY_ERROR
       End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
        
        '�@����ݒ�i�w���j�C���[�W�t�@�C���쐬
        bRet = dllGetKikiIniData(0, 1, KIKI_DATA_SET_EKI_INFO_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)
            
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
            
            '�ُ�I��
            'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbInformation, "�}�̓��͌���") 'V1.8.0.1 DEL
            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̓��͌���") 'V1.8.0.1 ADD
            
            Exit Sub
        End If
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
        
        '����I��
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�̓��͌���")
    
        '�@����f�[�^�X�V�t���O�ݒ�i�X�V�ݒ�j
        KikiDataUpDateFlg = True
        '��ʕ\������
        Call sDisp
        '�@����f�[�^�X�V�t���O�ݒ�i�ʏ�ݒ�j
        KikiDataUpDateFlg = False
        
        '�ݒ蔽�f�t���O�i�ύX����j
        SetteiHaneiFlg = True       'V1.20.0.1 ADD
    End If

  Exit Sub
  
COPY_ERROR:

    '�ُ탍�O�o��
    Select Case Err.Number
        Case 61 ' �󂫗e�ʕs��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case Else
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
    '�ꎞ�ۑ��t�H���_���폜����
    psDeleteFolder PATH_HOSHUTMP
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��

    '�ُ�I��
    'iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbInformation, "�}�̓��͌���") 'V1.8.0.1 DEL
    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̓��͌���")  'V1.8.0.1 ADD

End Sub
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������s�Ƃ���B
    CmdKikiSetMenu(3).Enabled = False
    CmdKikiSetMenu(4).Enabled = False
    CmdKikiSetMenu(5).Enabled = False
    CmdKikiSetMenu(6).Enabled = False
    CmdKikiSetMenu(7).Enabled = False               ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    cmdCancel.Enabled = False
    
    '�R���{�{�b�N�X�ACmdKikiSetMenu(0)�`(2)�͏����ɂ���Ă͌��X�����s�̂��ߔ�����s��
    If cmbEkiInfo.Enabled = True Then
        cmbEkiInfo.Enabled = False
        DoEvents
    End If
    
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    If CmbCornerName.Enabled = True Then
        CmbCornerName.Enabled = False
        DoEvents
    End If
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
    
    If CmdKikiSetMenu(0).Enabled = True Then
        CmdKikiSetMenu(0).Enabled = False
    End If
    
    If CmdKikiSetMenu(1).Enabled = True Then
        CmdKikiSetMenu(1).Enabled = False
    End If
    
    If CmdKikiSetMenu(2).Enabled = True Then
        CmdKikiSetMenu(2).Enabled = False
    End If
    
    DoEvents
    
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    Dim strFileName          As String          '�t�@�C����
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������Ƃ���B
    CmdKikiSetMenu(3).Enabled = True
    CmdKikiSetMenu(4).Enabled = True
    CmdKikiSetMenu(5).Enabled = True
    CmdKikiSetMenu(6).Enabled = True
    CmdKikiSetMenu(7).Enabled = True                ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    cmdCancel.Enabled = True
    
    '�R���{�{�b�N�X��CmdKikiSetMenu(0)�`(2)�͏����ɂ���Ă͌��X�����s�̂��߁A��ʕ\���̗L���Ŕ�����s��
    strFileName = Dir(KIKI_DATA_SET_EKI_INFO_FILE)
    '�t�@�C�������݂���ꍇ
    If strFileName <> "" Then
        cmbEkiInfo.Enabled = True
        CmbCornerName.Enabled = True                ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        CmdKikiSetMenu(0).Enabled = True
        CmdKikiSetMenu(1).Enabled = True
        CmdKikiSetMenu(2).Enabled = True
    End If
    
    DoEvents
    
End Sub
'V1.12.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : CmbCornerName_Click
'//  �@�\����  : �R�[�i�I�𕔑I������
'//  �@�\�T�v  : �O���b�h�f�[�^���Đݒ肷��
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//                 cmbEkiInfo_Click()���p
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmbCornerName_Click()

    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_GAMEN_CORNER_SELECT, 0)
    
    '��ʕ\������
    Call sDisp

    '�S�{�^���������Ƃ���B
    Call SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : InitCornerComboBox
'//  �@�\����  : �R�[�i�ݒ�R���{�{�b�N�X�̏���������
'//  �@�\�T�v  : �����������s���B
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
Private Sub InitCornerComboBox()

    Dim intLoop   As Integer            ' ���[�v�J�E���^
    Dim strCorner As String             ' ������i�[�G���A
    
    On Error Resume Next
    
    ' /////////////////////////////////////////////////////
    ' // ����������
    ' /////////////////////////////////////////////////////
    ' �R�[�i���̐ݒ菈��
    Call gsGetCornerName
    
    CmbCornerName.Clear
    For intLoop = 0 To 5
    
        '�ݒ肠��̃R�[�i�������ɂ���
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
            strCorner = gstrCornerName(intLoop)
            CmbCornerName.AddItem strCorner
        End If
    Next intLoop
    CmbCornerName.ListIndex = 0

End Sub

