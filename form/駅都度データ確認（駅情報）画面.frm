VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEkiData 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�w�s�x�f�[�^�m�F"
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
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   12
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton CmdMoveSubGateGamen 
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
      Left            =   7080
      TabIndex        =   11
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMoveGateGamen 
      Caption         =   "���D�@��ʂ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7080
      TabIndex        =   10
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "�e�L�X�g�}�̏o�́i�w���j"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   7800
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "�w�ݒ�o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "�w�ݒ����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
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
      Height          =   555
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   8880
      Top             =   960
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
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   10425
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   6195
      Left            =   120
      TabIndex        =   5
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
      FocusRect       =   0
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
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�w�s�x�f�[�^�m�F�i�w���j"
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
      Height          =   388
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   8295
   End
End
Attribute VB_Name = "frmEkiData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �F�w�s�x�f�[�^�m�F�i�w���j���.frm
'//  �p�b�P�[�W���F�w�s�x�f�[�^�m�F�i�w���j��ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-04-06   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//                   �u�w�ݒ�o�́v�u�w�ݒ���́v
'//�@�@�@�@�@�@�@�@�@ �u�w�ݒ�e�L�X�g�o�́v�u�}�̎�O�v�t�����ǉ�
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ�
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//                 �w�ݒ�t�@�C�������ݐ�f�B���N�g���ύX
'//                 �f�B�X�N���擾�ʒu�ύX
'//                 �e�L�X�g�o�͓��e�ύX
'//                 ��ʃ��b�N�����^��ʃ��b�N���������ǉ�
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                �t�H���_�I����ʂł́u����v�t���������ǉ�
'//                �u�e�L�X�g�}�̏o��(�w���)�v�t���������C��
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ��ʍđO�ʕ\���C��(�s��C��)
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 �P���b�`�����g�p�w�Ή�
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.76�C���Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   '���C���^�C�}�̃C���^�[�o���l
' Private Const TITOL_EKI_NAME = "�w���@�@�@�F"           '�w���^�C�g��     ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
Private Const TITOL_EKI_NAME = "�w���F"                 '�w���^�C�g��       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�

'V1.12.0.1 ADD START
'�@��\���f�[�^�i�w���j�C���[�W�t�@�C���ǎ�p�̍\����
Private Type EKIINFO_IMAGE_FILE
    sType       As String                '���
    sGoki       As String                '���@
    sNo         As String                '��ʖ��ʔ�
    sCorner     As String                '�R�[�i        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    sTuuban     As String                '�ʔ�
    sKoumoku    As String                '����
    sKubun      As String                '�敪
    sSettei     As String                '�ݒ�l
    sSyosai     As String                '�ݒ�l�ڍ�
End Type
'V1.12.0.1 ADD END
Private gstrFileName        As String                       ' �o�̓t�@�C����    ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �w�s�x�f�[�^�m�F�i�w���j���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �őO�O�\���������s���B
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

    '�G���[���[�`����錾
    On Error Resume Next
    
    '����ʍőO�ʕ\���������s���B
    pfFormActive (hwnd)
    
    '�^�C�}���N������
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �w�s�x�f�[�^�m�F�i�w���j���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �w�s�x�f�[�^�m�F�i�w���j���(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �����������s���B
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
Private Sub Form_Load()

    Dim bRet                 As Boolean         '�֐��߂�l
    Dim lErrCode             As Long            '�G���[�R�[�h
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_EKIINFO_GAMEN_START, 0)
    
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
    
    '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���쐬
    bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���폜
        Kill EKI_TUDO_CHK_EKI_INFO_FILE
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
    End If
    
    '�w���R���{�{�b�N�X�����l�ݒ�
    cmbEkiInfo.Clear
    cmbEkiInfo.AddItem "�w���"
    cmbEkiInfo.AddItem "�Ď�"
    cmbEkiInfo.AddItem "�l�b�g���[�N"
    cmbEkiInfo.AddItem "���"                           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    cmbEkiInfo.ListIndex = 0
    
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    '�R�[�i�ݒ�R���{�{�b�N�X�̏���������
    Call InitCornerComboBox
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
    
    '���C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
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
                AppActivate frmEkiData.Caption, False           ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
                pfFormActive (frmEkiData.hwnd)                  ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '�u�ێ瑀���v���O�������M�v���v����M�����ꍇ
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
                    '�v���O���X�o�[����������
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
                    Kill gstrFileName
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ���͌���")
                    Call SetEnableTrue
                Else
                    Call pfuncInstallEkiSettei
                End If
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_EKIINFO_GAMEN_END, 0)
    
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
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ�
'//     REVISIONS :(2.1.0.1)  2010-05-28  REVISED BY [TCC] S.Yoshimori
'//                 �P���b�`�����g�p�w�Ή�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim strFileName          As String          '�t�@�C����
    Dim iLoopCnt             As Integer         '���[�v�J�E���^
    Dim bRet                 As Boolean         '�֐��߂�l
    Dim strKubun             As String          '�敪
    Dim strIniData           As String          'INI�t�@�C���ݒ�l
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
    LblEkiName.Caption = TITOL_EKI_NAME         '�w�����x��������
    
    '----------------------------------------------------
    '�O���b�h�^�C�g���ݒ�
    '----------------------------------------------------
    Call sDispGridTitol
    
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

        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '�w�����x���X�V
    '----------------------------------------------------
'   LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo              'V2.1.0.1 DEL
    LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo(NotEkiVer)   'V2.1.0.1 ADD
    
    '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C������
    strFileName = Dir(EKI_TUDO_CHK_EKI_INFO_FILE)
    
    '�t�@�C�������݂����ꍇ
    If strFileName <> "" Then
    
        '�O���b�h�f�[�^���ݒ�
'        Call sDispDataSet(cmbEkiInfo.ListIndex + 1)                    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        nCornerIndex = CmbCornerName.ListIndex
        Call sDispDataSet(pfGetCodeDaiBunrui(cmbEkiInfo), nCornerIndex + 1)
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
        
        'INI�f�[�^�`�F�b�N
        With GridIni
            For iLoopCnt = 1 To .Rows - 1
            
                .Row = iLoopCnt
'V1.11.0.1 DEL START
'                .Col = 3
'
'                '�f�[�^�`�F�b�N
'                bRet = pfDispDataChk(.Text)
'                If bRet = False Then
'                    '�f�[�^�s��v�̏ꍇ�A�Z���̔w�i�F��ԐF�ɂ���
'                    .CellBackColor = QBColor(12)
'                End If
'V1.11.0.1 DEL END
                'V1.11.0.1 ADD START
                .Col = 0
                If .Text <> "" Then
                    .Col = 3
                    
                    '�f�[�^�`�F�b�N
                    bRet = pfDispDataChk(.Text)
                    If bRet = False Then
                        '�f�[�^�s��v�̏ꍇ�A�Z���̔w�i�F��ԐF�ɂ���
                        .CellBackColor = QBColor(12)
                    End If
                End If
                'V1.11.0.1 ADD END
            Next
        End With
    
        cmbEkiInfo.Enabled = True                  '�w���R���{�{�b�N�X�I���ݒ�
        CmbCornerName.Enabled = True               ' �R�[�i�I�𕔑I��s��      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Else
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKIINFO_IMAGE, 0)
        
        '�O���b�h�f�[�^���N���A����
        Call sDispDataClear(1, GridIni.Rows)
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
        .Cols = 5
        
        '----------------------------------
        '�O���b�h���ݒ�
        '----------------------------------
        .ColWidth(0) = 500
        .ColWidth(1) = 3800
        .ColWidth(2) = 730
        .ColWidth(3) = 2700
        .ColWidth(4) = 3700
        
        '----------------------------------
        '�^�C�g���ݒ�
        '----------------------------------
        '���ڐݒ�
        .Col = 1
        .Row = 0: .Text = "����"
        .CellAlignment = flexAlignCenterCenter

        '�敪�ݒ�
        .Col = 2
        .Text = "�敪"
        .CellAlignment = flexAlignCenterCenter

        '�ݒ�l�ݒ�
        .Col = 3
        .Text = "�ݒ�l"
        .CellAlignment = flexAlignCenterCenter

        '�ڍאݒ�
        .Col = 4
        .Text = "�ݒ�l�ڍ�"
        .CellAlignment = flexAlignCenterCenter
        
'        .RowHeight(0) = 700        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        .RowHeight(0) = 440         ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    End With
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispDataClear
'//  �@�\����  : �O���b�h�f�[�^���N���A����
'//  �@�\�T�v  : �O���b�h�f�[�^�����N���A����
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   intStartRow  �J�n�s�ʒu
'//              Integer   intEndRow    �I���s�ʒu
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear(intStartRow As Integer, intEndRow As Integer)
    
    Dim iLoopCnt             As Integer         '���[�v�J�E���^
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�O���b�h������
    With GridIni

        .Rows = intEndRow   'V1.11.0.1 ADD
        For iLoopCnt = intStartRow To intEndRow - 1

            '�ʔԐݒ�
            .Col = 0
            .Row = iLoopCnt: .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '���ڐݒ�
            .Col = 1
            .Text = ""
            .CellAlignment = flexAlignLeftCenter

            '�敪�ݒ�
            .Col = 2
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'            .Text = "" & vbCrLf & _
'                    "" & vbCrLf & _
'                    ""
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
            
            .CellAlignment = flexAlignCenterCenter

            '�ݒ�l�ݒ�
            .Col = 3
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'            .Text = "" & vbCrLf & _
'                    "" & vbCrLf & _
'                    ""
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
            .CellAlignment = flexAlignLeftCenter

            '�ڍאݒ�
            .Col = 4
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'            .Text = "" & vbCrLf & _
'                    "" & vbCrLf & _
'                    ""
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
            .Text = "" & vbCrLf & _
                    "" & vbCrLf & _
                    "" & vbCrLf & _
                    ""
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
            .CellAlignment = flexAlignLeftCenter

            .RowHeight(iLoopCnt) = 938
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
'//     REVISIONS :(1.11.0.1) 2009-10-26  REVISED BY [TCC] D.Yamashita
'//                 EG-R�t�F�[�Y3�c�����ڑΉ�
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
    Dim iRowCnt             As Integer                      ' �s�J�E���^
    Dim iBunrui_Sho()      As Integer                       ' �����ރe�[�u��
    
    Dim strBunrui_Dai       As String                       ' �啪��
    Dim strBunrui_Tyu       As String                       ' ������
    Dim strBunrui_Sho       As String                       ' ������
    Dim strNo               As String                       ' �ʔ�
    Dim strKomoku           As String                       ' ����
    Dim strKubun            As String                       ' �敪
    Dim strData             As String                       ' �ݒ�l
    Dim strSetShosai        As String                       ' �ݒ�l�ڍ�
    Dim strCorner           As String                       ' �R�[�i    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Dim iCmpCorner          As Integer                      ' �R�[�i    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '������
    ReDim iBunrui_Sho(0)

    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C�����I�[�v������B
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber
    
    GridIni.Visible = False
    iLoopCnt = 1
    GridIni.Rows = 1
    Do While Not EOF(intFileNumber)
    
        '�P �s�ǂݍ���
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strNo, _
'                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strCorner, strNo, _
                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
        
        If iBunrui_Dai = strBunrui_Dai Then

' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        ' �R�[�i����ǉ�
        ' �I�������R�[�i�A�������̓R�[�i���֌W�̃��R�[�h�͍̗p����
        iCmpCorner = CInt(strCorner)
        If ((iCorner = iCmpCorner) Or (iCmpCorner = 0)) Then
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
        
            '�O���b�h������
            With GridIni
            
                '���ڌ���
                For iRowCnt = 0 To iLoopCnt - 2
                    If CStr(iBunrui_Sho(iRowCnt)) = strBunrui_Sho Then Exit For
                Next
                
                '���ڂ�������Ȃ������ꍇ
                If iRowCnt = .Rows - 1 Then
                    
                    '�����ރe�[�u���o�^
                    ReDim Preserve iBunrui_Sho(.Rows - 1)
                    iBunrui_Sho(iRowCnt) = CInt(strBunrui_Sho)
                    
                    '�\���s�����C���N�������g
                    iLoopCnt = iLoopCnt + 1
                    .Rows = iLoopCnt
                    
                End If
            
                '�\���f�[�^���P��ʂɕ\��������Ȃ��ꍇ
                If .Rows > 6 Then
                    '�X�N���[���o�[���A�O���b�h���L����
                    .Width = 11775
                End If

                '�ʔԐݒ�
                .Col = 0
                .Row = iLoopCnt - 1: If .Text = "" Then .Text = CStr(iLoopCnt - 1)
                .CellAlignment = flexAlignLeftCenter
    
                '���ڐݒ�
                .Col = 1
                If .Text = "" Then .Text = strKomoku
                .CellAlignment = flexAlignLeftCenter
    
                '�敪�ݒ�
                .Col = 2
                .Text = pfDispAplBunrui(.Text, strKubun)
                .CellAlignment = flexAlignCenterCenter
    
                '�ݒ�l�ݒ�
                .Col = 3
                .Text = pfDispIniData(.Text, strData, strKubun)
                .CellAlignment = flexAlignLeftCenter
    
                '�ڍאݒ�
                .Col = 4
                If .Text = "" Then .Text = strSetShosai
                .CellAlignment = flexAlignLeftCenter
    
                .RowHeight(iLoopCnt - 1) = 938
        
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
    If GridIni.Rows < 7 Then
        '�O���b�h�f�[�^���N���A����
'        Call sDispDataClear(GridIni.Rows - 1, 9)   'V1.11.0.1 DEL
        Call sDispDataClear(GridIni.Rows, 7)        'V1.11.0.1 ADD
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
    Call sDispDataClear(1, GridIni.Rows - 1)

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
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_GAMEN_EKIINFO_SELECT, 0)
    
    '��ʕ\������
    Call sDisp

'V1.12.0.1 ADD START
    '�S�{�^���������Ƃ���B
    Call SetEnableTrue
'V1.12.0.1 ADD END

End Sub

'V1.4.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : CmdMenu_Click
'//  �@�\����  : �u�w�ݒ�o�́v�u�w�ݒ���́v�u�w�ݒ�e�L�X�g�o�́v
'//              �u�}�̎�O�v�t��������
'//  �@�\�T�v  : �e�t���̏������s���B
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmdMenu_Click(Index As Integer)
  
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
       Case 0                                  '�w�ݒ�o��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_OUTPUT, 0)
            
            '�w�ݒ�o�͏���
            Call sEkiSetteiOutPut
        
        Case 1                                  '�w�ݒ����
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_EKISET_INPUT, 0)
            
            '�w�ݒ���͏���
            Call sInstolEkiSettei
            bUnlock = False                     ' EG20 V3.0.0.2 �ǉ�
        
        Case 2                                  '�w�ݒ�e�L�X�g�o��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKISETKAKUNINMENU_GAMEN_DISP_TEXT, 0)
            
            '�w�ݒ�e�L�X�g�o�͏���
            Call sDispTextEkiDataNow
        
        Case 3                                  '�}�̎�O
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
            
            '�}�̎�O����
            Call pfRemove(Me)
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
'//  �֐�����  : sEkiSetteiOutPut
'//  �@�\����  : �u�w�ݒ�o�́v�t����������
'//  �@�\�T�v  : ���݉w�ݒ�t�@�C�����O���}�̂ɏo�͂���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �w�ݒ�t�@�C�������ݐ�f�B���N�g���ύX
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
Private Sub sEkiSetteiOutPut()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
    Dim iResponse            As Integer         'MsgBox�߂�l

    '�G���[���[�`����錾
    On Error Resume Next
'V1.8.0.1 DEL START
'    iResponse = MsgBox("�I������Ă���w�̌��݂̉w�s�x�f�[�^�P�w�����o�͂��܂��B" & Chr(vbKeyReturn) & _
'                        "��낵���ł����H", _
'                        vbYesNo + vbQuestion, _
'                        "�w�ݒ�o�͊m�F")
'V1.8.0.1 DEL END
'V1.8.0.1 ADD START
    iResponse = MsgBox("�I������Ă���w�̌��݂̉w�s�x�f�[�^�P�w�����o�͂��܂��B" & Chr(vbKeyReturn) & _
                        "��낵���ł����H", _
                        vbOKCancel + vbQuestion, _
                        "�w�ݒ�o�͊m�F")
'V1.8.0.1 ADD END
'    If iResponse = vbNo Then Exit Sub              'V1.12.0.1 DEL
    If iResponse = vbCancel Then Exit Sub           'V1.12.0.1 DEL
    
'�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '�ُ�I��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", _
                vbOKOnly + vbExclamation, _
                 "�f�[�^���x��"
        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '�}�̏o�͏���
    '----------------------------------------------------
'    sWriteDir = pfDirSelection("a:", "�w�ݒ�t�@�C�������ݐ�̃f�B���N�g���I��")   'V1.12.0.1 DEL
    'sWriteDir = pfDirSelection("H:", "�w�ݒ�t�@�C�������ݐ�̃f�B���N�g���I��")    'V1.12.0.1 ADD     'V1.20.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
    If sWriteDir <> "" Then
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
'        FileCopy EKI_SETTI_FILE, sWriteDir & Dir(EKI_SETTI_FILE)       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        FileCopy EKI_SETTI_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(EKI_SETTI_FILE)        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
       '����I��
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�w�ݒ�o�͌���")
    
    End If
    
  Exit Sub
 
COPY_ERROR:

    Select Case Err.Number
        Case 61 ' �}�̏o�͋󂫗e�ʕs��
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_SHORT_VOLUME, 0)
        Case 71 ' �}�̂Ȃ�
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_NOT_DISK, 0)
        Case Else
            '�ُ탍�O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, ERROR_MEDIUM_OTHER_ERR, 0)
    End Select

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    iResponse = MsgBox("�ُ�I�����܂���", vbOKOnly + vbCritical, "�w�ݒ�o�͌���")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sInstolEkiSettei
'//  �@�\����  : �u�w�ݒ���́v�t����������
'//  �@�\�T�v  : �O���}�̂��猻�݉w�ݒ�t�@�C���C���X�g�[������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �f�B�X�N���擾�ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.76�C���Ή��z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sInstolEkiSettei()

    Dim iResponse           As Integer          'MsgBox�߂�l
    Dim bRet                As Boolean          '�֐��߂�l
    Dim strFileName         As String           '�}�̃t�@�C����

    Dim objFso As New FileSystemObject          '�t�@�C���V�X�e���I�u�W�F�N�g

    Dim lResult             As Long             ' ��������

    '�G���[���[�`����錾
    On Error Resume Next
    iResponse = MsgBox("�w�s�x�f�[�^�P�w�����C���X�g�[�����܂��B" & Chr(vbKeyReturn) & _
                        "��낵���ł����H", _
                        vbOKCancel + vbQuestion, _
                        "�w�ݒ���͊m�F")
    If iResponse = vbCancel Then
        Set objFso = Nothing
        Call SetEnableTrue                      ' EG20 V5.0.2.1�y����TR-No.76�C���Ή��z�ǉ�
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
    ' �g���q��ݒ�
    CommonDialog1.Filter = "�b�r�u�i�J���}��؂�j(*.csv)|*.csv|"
    ' �t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    ' �I�������t�@�C�������擾
    strFileName = CommonDialog1.FileName
    
    Call ChDrive("D")  'V2.5.0.1 ADD

    '�t�@�C�����݃`�F�b�N
    If strFileName <> "" Then
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL Start
'        ' �o�͐�t�@�C������ۑ�
'        gstrFileName = strFileName
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL End
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
        ' �o�͐�t�@�C������ۑ�
        gstrFileName = PATH_HOSHUWRK_EKI_INFO
        '�ꎞ�ۑ��t�H���_�Ƀf�[�^���R�s�[���ǎ��p����������
        If pfChangeAttrNormal(strFileName, PATH_HOSHUTMP_EKI_INFO, PATH_HOSHUWRK_EKI_INFO) = False Then
            Kill gstrFileName
            '�ꎞ�ۑ��t�H���_���폜����
            psDeleteFolder PATH_HOSHUTMP
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' �ُ�I��
            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ���͌���")
            Call SetEnableTrue
            Exit Sub
        End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End

        ' //////////////////////////////////////////////
        ' // �����v���O��������
        ' //////////////////////////////////////////////
        lResult = pubfuncTakuProgramData(2, gstrFileName)
        If lResult = 0 Then
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
            Kill gstrFileName
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' �ُ�I��
            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ���͌���")
            Call SetEnableTrue
            Exit Sub
        ElseIf lResult = 1 Then
            ' ���[�����M��
            ' ���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
            Exit Sub
        End If


        ' //////////////////////////////////////////////
        ' // �����Ď��Ք񓮍쒆�̂��߃��[��������҂�����
        ' // �����X�V
        ' //////////////////////////////////////////////
        bRet = pfuncInstallEkiSettei

    End If
    Call SetEnableTrue                      ' EG20 V5.0.2.1�y����TR-No.76�C���Ή��z�ǉ�
End Sub

' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�폜�J�n�i�S�̌������j
'Private Sub sInstolEkiSettei()
'
'    Dim iResponse           As Integer          'MsgBox�߂�l
'    Dim bRet                As Boolean          '�֐��߂�l
'    Dim lErrCode            As Long             '�G���[�R�[�h
'    Dim strFileName         As String           '�}�̃t�@�C����
'
'    Dim iRet                    As Integer      '���b�Z�[�W�{�b�N�X�߂�l
'    Dim lSekuta                 As Long         '�Z�N�^�i�N���X�^����j
'    Dim lByte                   As Long         '�o�C�g���i�Z�N�^����j
'    Dim lKurasuta               As Long         '�t���[�N���X�^��
'    Dim lDrive                  As Long         '�h���C�u�̃N���X�^���i���v�j
'    Dim strDrive                As String       '�h���C�u
'    Dim bSysChange              As Boolean      '�V�X�e���ݒ菈���߂�l�@'V1.8.0.1�@ADD
'    Dim bUpData                 As Boolean      '��ʍX�V�����߂�l�@�@�@'V1.8.0.1�@ADD
'
'    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g  'V1.20.0.1 ADD
'
'    Dim lResult             As Long             ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'
'    '�G���[���[�`����錾
'    On Error Resume Next
''V1.8.0.1 DEL START
''    iResponse = MsgBox("�w�s�x�f�[�^�P�w�����C���X�g�[�����܂��B" & Chr(vbKeyReturn) & _
''                        "��낵���ł����H", _
''                        vbYesNo + vbQuestion, _
''                        "�w�ݒ���͊m�F")
''V1.8.0.1 DEL END
''V1.8.0.1 ADD START
'    iResponse = MsgBox("�w�s�x�f�[�^�P�w�����C���X�g�[�����܂��B" & Chr(vbKeyReturn) & _
'                        "��낵���ł����H", _
'                        vbOKCancel + vbQuestion, _
'                        "�w�ݒ���͊m�F")
''V1.8.0.1 ADD END
'    'V1.20.0.1 DEL START
'''    If iResponse = vbNo Then Exit Sub              'V1.12.0.1 DEL
'    'If iResponse = vbCancel Then Exit Sub           'V1.12.0.1 ADD
'    '
'    ''�f�B�X�N�����擾
'''    iRet = GetDiskFreeSpace("A:\", lSekuta, lByte, lKurasuta, lDrive)  'V1.12.0.1 DEL
'    'iRet = GetDiskFreeSpace("H:\", lSekuta, lByte, lKurasuta, lDrive)   'V1.12.0.1 ADD
'    '
'    'If lDrive = 0 Then
'    '    strDrive = "d:"
'    'Else
'''        strDrive = "a:"    'V1.12.0.1 DEL
'    '    strDrive = "H:"     'V1.12.0.1 ADD
'    'End If
'    '
'    ''�}�̃t�@�C�����擾
'    'strFileName = pfFileSelection(strDrive, "*.csv", "�w�ݒ�̧�ّI��")
'    'V1.20.0.1 DEL END
'    'V1.20.0.1 ADD START
'    If iResponse = vbCancel Then
'        Set objFso = Nothing
'        Exit Sub
'    End If
'    '�擾�t�@�C������������
'    CommonDialog1.FileName = ""
'    '�����f�B���N�g����ݒ�
'    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    '�t�H���_�I����ʃf�t�H���g�p�X�P�����݂��邩
'        '���݂��邽�߁A�f�t�H���g�p�X�P�iH:�j��ݒ�
'        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
'    Else
'        '���݂��Ȃ����߁A�f�t�H���g�p�X�Q�iC:�j��ݒ�
'        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
'    End If
'    Set objFso = Nothing
'    '�g���q��ݒ�
'    CommonDialog1.Filter = "�b�r�u�i�J���}��؂�j(*.csv)|*.csv|"
'    '�t�@�C���I����ʂ��J��
'    CommonDialog1.ShowOpen
'    '�I�������t�@�C�������擾
'    strFileName = CommonDialog1.FileName
'    'V1.20.0.1 ADD END
'
'    Call ChDrive("D")  'V2.5.0.1 ADD
'
'    '�t�@�C�����݃`�F�b�N
'    If strFileName <> "" Then
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'        '�v���O���X�o�[��\������
'        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'        '���݉w�ݒ�f�[�^�C���X�g�[������
'        bRet = dllInstolEkiDataNow(strFileName, EKI_SETTI_FILE, lErrCode)
'
'        If bRet = False Then
'
'            '�ُ탍�O�o��
'            Call pfOutPutErrLog(lErrCode)
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'            '�v���O���X�o�[����������
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'            '�ُ�I��
'            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ���͌���")
'
'        Else
''V1.8.0.1 ADD START
'            '----------------------------------------------------
'            '�R���s���[�^���A�l�b�g���[�N�ύX����
'            '----------------------------------------------------
'            'Call pfNetWorkChng(Me)
'            bSysChange = True
'            bUpData = True
'            bSysChange = pfNetWorkChng(Me)
''V1.8.0.1 ADD END
'             '���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
''V1.8.0.1 ADD START
'           '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���쐬
'            bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
'            If bRet = False Then
'                '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���폜
'                Kill EKI_TUDO_CHK_EKI_INFO_FILE
'                '�ُ탍�O�o��
'                Call pfOutPutErrLog(lErrCode)
'                bUpData = False
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'                '�v���O���X�o�[����������
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'            End If
'
'' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'            ' //////////////////////////////////////////////
'            ' // �����v���O��������
'            ' //////////////////////////////////////////////
'             lResult = pubfuncTakuProgramData(2)
'             If lResult = 0 Then
'                '�v���O���X�o�[����������
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'                ' �ُ�I��
'                iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�ݒ蔽�f����")
'                Call SetEnableTrue
'                Exit Sub
'             ElseIf lResult = 1 Then
'                ' ���[�����M��
'                ' ���O�o��
'                Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'
'                Exit Sub
'             End If
'' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'
'            '�w���R���{�{�b�N�X�����l�ݒ�
'            cmbEkiInfo.Clear
'            cmbEkiInfo.AddItem "�w���"
'            cmbEkiInfo.AddItem "�Ď�"
'            cmbEkiInfo.AddItem "�l�b�g���[�N"
'            cmbEkiInfo.AddItem "���"                           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'            cmbEkiInfo.ListIndex = 0
'
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'            '�R�[�i�ݒ�R���{�{�b�N�X�̏���������
'            Call InitCornerComboBox
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'            '�v���O���X�o�[����������
'            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'            If bSysChange = True And bUpData = True Then
''V1.8.0.1 ADD END
'            '����I��
'            iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�w�ݒ���͌���")
'            End If         'V1.8.0.1 ADD
'        End If
'    End If
'
'End Sub
' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�폜�I���i�S�̌������j

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sDispTextEkiDataNow
'//  �@�\����  : �u�w�ݒ�e�L�X�g�o�́v�t����������
'//  �@�\�T�v  : ���݉w�ݒ�t�@�C�����e�L�X�g�\������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �e�L�X�g�o�͓��e�ύX
'//     REVISIONS :(1.13.0.1) 2009-11-19   REVISED BY [TCC] S.Terao
'//                �t�H���_�I����ʂł́u����v�t���������ǉ�
'//                �u�e�L�X�g�}�̏o��(�w���)�v�t���������C��
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(EG20 V6.6.0.1)  2012-06-20  CODED BY  [TCC] H.Sugimoto
'//                 �y�I���R�[�i�ʂ̏o�͑Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispTextEkiDataNow()

    Dim strFileName          As String          '�t�@�C����
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim lRetVal              As Long            '�߂�l
    Dim sCommand             As String          '�R�}���h������
'V1.12.0.1 ADD START
    Dim sWriteDir            As String              '�������ݐ�t�H���_��
    Dim intFileNumber        As Integer             '�t�@�C���|�C���^
    Dim strLineCount         As String              '�s���J�E���^
    Dim i                    As Integer             '���[�v�J�E���^�P
    Dim j                    As Integer             '���[�v�J�E���^�Q
    Dim k                    As Integer             '���[�v�J�E���^�R
    Dim ReadFileSettei()     As EKIINFO_IMAGE_FILE  '�t�@�C���Ǎ��p�\����
    Dim fso         As New FileSystemObject         '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim FsoTS As TextStream

    Set fso = CreateObject("Scripting.FileSystemObject")
'V1.12.0.1 ADD END
'V1.13.0.1 ADD START
    Dim skansi               As String  '�O��敪����i�Ď��p�j
    Dim sidu                 As String  '�O��敪����iIDU�p�j
    Dim sldu                 As String  '�O��敪����iLDU�p�j
    Dim sTaku                As String  '�O��敪����i�����p�j   ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Dim sSyoKoumoku          As String  '�O�񏬍��ڔ���
    Dim nProcMode            As Integer ' ���ݏ�������              ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    Dim szCornerName         As String          ' �R�[�i����        ' EG20 V6.6.0.1�ǉ�
    Dim nNullIndex           As Integer         ' ���������[�N      ' EG20 V6.6.0.1�ǉ�
    Dim nCornerIndex         As Integer         ' �R�[�i�I�����    ' EG20 V6.6.0.1�ǉ�
    Dim strSaveFileName      As String          ' �ۑ��t�@�C����    ' EG20 V6.6.0.1�ǉ�
    
    '������
    skansi = ""
    sidu = ""
    sldu = ""
'V1.13.0.1 ADD END

    '�G���[���[�`����錾
    On Error Resume Next
'V1.8.0.1 DEL START
'    iResponse = MsgBox("�I������Ă���w�̌��݂̉w�s�x�f�[�^�P�w�����e�L�X�g�\�����܂��B" & Chr(vbKeyReturn) & _
'                        "��낵���ł����H", _
'                        vbYesNo + vbQuestion, _
'                        "�w�ݒ�e�L�X�g�o�͊m�F")
'V1.8.0.1 DEL END
'V1.12.0.1 DEL START
''V1.8.0.1 ADD START
'    iResponse = MsgBox("�I������Ă���w�̌��݂̉w�s�x�f�[�^�P�w�����e�L�X�g�\�����܂��B" & Chr(vbKeyReturn) & _
'                        "��낵���ł����H", _
'                        vbOKCancel + vbQuestion, _
'                        "�w�ݒ�e�L�X�g�o�͊m�F")
''V1.8.0.1 ADD END
''    If iResponse = vbNo Then Exit Sub
'V1.12.0.1 DEL START
    
'V1.12.0.1 ADD START
    '�������ݐ�t�@�C���I��
    'sWriteDir = pfDirSelection("H:", "�@��\���t�@�C�������ݐ�̃f�B���N�g���I��")         'V1.20.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
'V1.12.0.1 ADD START
'V1.13.0.1 ADD START
    If sWriteDir = "" Then
       '�t�H���_�I����ʁu����v�t�������͏����I��
       Exit Sub
    End If
'V1.13.0.1 ADD END
    
   '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '���݉w�ݒ�t�@�C������
    '----------------------------------------------------
'    strFileName = Dir(EKI_SETTI_FILE)              'V1.12.0.1 DEL
    strFileName = Dir(EKI_TUDO_CHK_EKI_INFO_FILE)   'V1.12.0.1 ADD

    '�t�@�C�������݂��Ȃ��ꍇ
    If strFileName = "" Then
    
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '�ُ�I��
        MsgBox "�e�L�X�g�\������f�[�^������܂���B", _
                vbOKOnly + vbExclamation, _
                 "�f�[�^���x��"
        Exit Sub
        
    End If
'V1.12.0.1 ADD START
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
    On Error GoTo OUTPUT_ERROR
    
    '�t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    'CSV�t�@�C���I�[�v��
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber
    
    'CSV�t�@�C���s���J�E���g�i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1�폜
        Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1�ǉ�
            Line Input #intFileNumber, strLineCount
            j = j + 1
        Loop
    
    'CSV�t�@�C���N���[�Y
    Close #intFileNumber

    '�t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '�Đݒ�
    ReDim ReadFileSettei(j) As EKIINFO_IMAGE_FILE        '�t�@�C���Ǎ��p�G���A
        
    'CSV�t�@�C���I�[�v��
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber

    '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
        For i = 0 To j - 1
            Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, _
            ReadFileSettei(i).sCorner, ReadFileSettei(i).sTuuban, ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, _
            ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
        Next

    'CSV�t�@�C���N���[�Y
    Close #intFileNumber
    
    '�ꎞ�t�@�C�������
    Set FsoTS = fso.CreateTextFile(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, True)
       
'    FsoTS.Write ("�ݒu�w�F" & Trim(pfGetEkiNameInfo(NotEkiVer)) & vbCrLf & vbCrLf)     ' EG20 V6.6.0.1�폜
' EG20 V6.6.0.1�ǉ��J�n
    FsoTS.Write ("�ݒu�w�@�@�F" & Trim(pfGetEkiNameInfo(NotEkiVer)) & vbCrLf)
    ' �R�[�i���̂̕t��
    nNullIndex = InStr(gstrCornerName(CmbCornerName.ListIndex), Chr(0))
    If nNullIndex <> 0 Then
        szCornerName = Left(gstrCornerName(CmbCornerName.ListIndex), nNullIndex - 1)
    Else
        szCornerName = gstrCornerName(CmbCornerName.ListIndex)
    End If
    FsoTS.Write ("�ݒu�R�[�i�F" & szCornerName & vbCrLf & vbCrLf)
    nCornerIndex = CmbCornerName.ListIndex + 1
' EG20 V6.6.0.1�ǉ��I��
       
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�R�����g�ǉ��J�n
' �]�������ڒʔԂ͑區�ڒP�ʘA�Ԃɑ΂��ĂP����A�Ԃł��邱�Ƃ�
' �O��ł��������O��͂Ȃ��Ȃ����B
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�R�����g�ǉ��I��
    
    nProcMode = 0       ' ���ݏ�������              ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    For k = 0 To j - 1
        
' EG20 V6.6.0.1�����ǉ��J�n
      If ((ReadFileSettei(k).sCorner = nCornerIndex) Or (ReadFileSettei(k).sCorner = 0)) Then
' EG20 V6.6.0.1�����ǉ��I��
        
        '����
        If ReadFileSettei(k).sType = 1 Then
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sTuuban = 1 Then        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
            If nProcMode <> 1 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
                '�^�C�g���\�������i�w���j
                FsoTS.Write ("�y�w���z" & vbCrLf & "����,�敪,�ݒ�l" & vbCrLf)
                nProcMode = 1                                                           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
            End If
'            FsoTS.Write (ReadFileSettei(k).sKoumoku & ",")     'V1.13.0.1 DEL
        
        ElseIf ReadFileSettei(k).sType = 2 Then
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sNo <> ReadFileSettei(k - 1).sNo Then   ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
            If nProcMode <> 2 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
                '�^�C�g���\�������i�Ď��j
                FsoTS.Write (vbCrLf & "�y�Ď��z" & vbCrLf & "����,�敪,�ݒ�l" & vbCrLf)
                nProcMode = 2                                                           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
            End If
'            FsoTS.Write (ReadFileSettei(k).sKoumoku & ",")     'V1.13.0.1 DEL
        
'        Else                                       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        ElseIf ReadFileSettei(k).sType = 3 Then     ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sNo <> ReadFileSettei(k - 1).sNo Then   ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
            If nProcMode <> 3 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
                '�^�C�g���\������
                FsoTS.Write (vbCrLf & "�y�l�b�g���[�N�z" & vbCrLf & "����,�敪,�ݒ�l" & vbCrLf)
                nProcMode = 3                                                           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
            End If
'            FsoTS.Write (ReadFileSettei(k).sKoumoku & ",")     'V1.13.0.1 DEL
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        Else
'            If ReadFileSettei(k).sNo = 1 And ReadFileSettei(k).sNo <> ReadFileSettei(k - 1).sNo Then   ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
            If nProcMode <> 7 Then                                                      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
                '�^�C�g���\������
                FsoTS.Write (vbCrLf & "�y��ʁz" & vbCrLf & "����,�敪,�ݒ�l" & vbCrLf)
                nProcMode = 7                                                           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
            End If
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
        End If

'V1.13.0.1 ADD START
        '����ƑO��̏����ڂ��������ǂ������肷��
        If ReadFileSettei(k).sNo <> sSyoKoumoku Then
                '���݂̏����ڂ�ۑ�����i�敪�ʁj
'                If ReadFileSettei(k).sKubun = "�Ď�" Then      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
                If ReadFileSettei(k).sKubun = "����" Then       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
                    skansi = ReadFileSettei(k).sNo
                ElseIf ReadFileSettei(k).sKubun = "IDU" Then
                    sidu = ReadFileSettei(k).sNo
                ElseIf ReadFileSettei(k).sKubun = "LDU" Then
                    sldu = ReadFileSettei(k).sNo
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
                ElseIf ReadFileSettei(k).sKubun = "����" Then
                    sTaku = ReadFileSettei(k).sNo
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
                End If
                '���݂̏����ڂ�ۑ�����i�S�́j
                sSyoKoumoku = ReadFileSettei(k).sNo
                '�t�@�C���ɏo�͂���
                FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                ReadFileSettei(k).sSettei & vbCrLf)
        Else
            '�����ڂ������������ꍇ�A�敪���������ǂ����m�F����B�����ł���Ώo�͂��Ȃ��B
            Select Case ReadFileSettei(k).sKubun
'                Case "�Ď�"            ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
                Case "����"             ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
                    If ReadFileSettei(k).sNo = skansi Then
                        '�����Ȃ�
                    Else
                        '�t�@�C���ɏo�͂���
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
                Case "IDU"
                    If ReadFileSettei(k).sNo = sidu Then
                        '�����Ȃ�
                    Else
                        '�t�@�C���ɏo�͂���
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
                Case "LDU"
                    If ReadFileSettei(k).sNo = sldu Then
                        '�����Ȃ�
                    Else
                        '�t�@�C���ɏo�͂���
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
                Case "����"
                    If ReadFileSettei(k).sNo = sTaku Then
                        '�����Ȃ�
                    Else
                        '�t�@�C���ɏo�͂���
                        FsoTS.Write (ReadFileSettei(k).sKoumoku & "," & ReadFileSettei(k).sKubun & "," & _
                        ReadFileSettei(k).sSettei & vbCrLf)
                    End If
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
            End Select
        End If
      End If              ' EG20 V6.6.0.1�ǉ�
'V1.13.0.1 ADD END
'V1.13.0.1 DEL START
'            '�敪
'            FsoTS.Write (ReadFileSettei(k).sKubun & ",")
'
'            '�ݒ�l
'            FsoTS.Write (ReadFileSettei(k).sSettei & vbCrLf)
'V1.13.0.1 DEL END
    Next
    
    '�t�@�C�����N���[�Y����B
    FsoTS.Close
        
    '�ꎞ�t�@�C����}�̂ɃR�s�[����
'    Call FileCopy(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, sWriteDir & EKI_SETTI_EKI_INFO_FILE)        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
' EG20 V6.6.0.1�폜�J�n
'    Call FileCopy(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & EKI_SETTI_EKI_INFO_FILE)        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
' EG20 V6.6.0.1�폜�I��
' EG20 V6.6.0.1�ǉ��J�n
    strSaveFileName = sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Replace(szCornerName, " ", "") & "_" & EKI_SETTI_EKI_INFO_FILE
    Call FileCopy(PATH_WORK & EKI_SETTI_EKI_INFO_FILE, strSaveFileName)
' EG20 V6.6.0.1�ǉ��I��
    
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
    '����I��
    iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�w�ݒ�e�L�X�g�o�͌���")
'V1.12.0.1 ADD END

'V1.12.0.1 DEL START
'    sCommand = MN_EXE_MEMO & EKI_SETTI_FILE         '���������s�R�}���h���쐬����
'    lRetVal = Shell(sCommand, vbMaximizedFocus)     '�m�[�g�p�b�h���N������
'    AppActivate lRetVal, True                       '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
'
'    SendKeys "{LEFT}", True
'V1.12.0.1 DEL END

'V1.12.0.1 ADD START

    Exit Sub
    
OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    '�ُ탍�O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, CREATE_FILE_ERROR, 0)

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
    '�ُ�I��
    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ�e�L�X�g�o�͌���")
'V1.12.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfStartUpProc
'//  �@�\����  : �t�@�C���I����ʏ���
'//  �@�\�T�v  : �t�@�C���I����ʂ�\�����A�I�����ꂽ�t�@�C������Ԃ��B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sDrive�@�@[IN]�����\���h���C�u��
'//  �@�@      : String�@�@sPattern�@[IN]�I��Ώۃt�@�C���g���q
'//  �@�@      : String�@�@sTitle�@�@[IN]��ʕ\�����x��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :String�@�@�@�@�@�@�@ [OUT]�߂�l
'//                                      �I�����ꂽ�t�@�C���p�X:����@""�F�G���[
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Function pfFileSelection(sDrive As String, _
                                sPattern As String, _
                                sTitle As String) As String
                                
    Dim sWorkDrive As String                    '���[�N�p�����\���h���C�u��

    '�h���C�u�ُ폈�����`����B
    On Error GoTo Drive_Error
    
    sWorkDrive = sDrive                         '�����\���h���C�u�������[�N�p�ɃZ�b�g����B
    frmFil.filSelection.Pattern = sPattern      '�I��Ώۊg���q���Z�b�g����B
    frmFil.lblFileSelection = sTitle            '�T�u�^�C�g�����Z�b�g����B

Retry:
    frmFil.drvSelection.Drive = sWorkDrive      '�h���C�u���Z�b�g����B
    frmFil.dirSelection.Path = sWorkDrive & "\" '�f�B���N�g�����Z�b�g����B
    
    '�t�@�C���I����ʂ�\������B
    frmFil.Show 1
    
    '�I�����ꂽ�t�@�C������Ԃ��B
    pfFileSelection = gstrMyPath
    
    Exit Function

'**�h���C�u�w��ُ폈��**
Drive_Error:

'    If Left$(sWorkDrive, 1) = "a" Then     'V1.12.0.1 DEL
    If Left$(sWorkDrive, 1) = "H" Then      'V1.12.0.1 ADD
        'a:�h���C�u���ُ�Ȃ�A�J�����g�h���C�u��\��������B
        sWorkDrive = Left$(App.Path, 2)
        GoTo Retry
    End If
    
    '���̑��̃h���C�u�Ȃ�A�t�@�C���I���Ȃ��Ŗ߂�B
    pfFileSelection = ""

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : CmdMoveGateGamen_Click
'//  �@�\����  : �u������ʂցv�t��������
'//  �@�\�T�v  : �w�s�x�f�[�^�m�F(����)��ʂ�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :�Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�̉����^�s�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveGateGamen_Click()
   
'V1.12.0.1 ADD START
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    DoEvents
'V1.12.0.1 ADD END
   
   '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)
    Unload Me
    Load frmEkiDataGate
    frmEkiDataGate.Show 1
    
End Sub
'V1.4.0.1 ADD END
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
    cmbEkiInfo.Enabled = False
    CmdMenu(0).Enabled = False
    CmdMenu(1).Enabled = False
    CmdMenu(2).Enabled = False
    CmdMenu(3).Enabled = False
    CmdMoveGateGamen.Enabled = False
    cmdCancel.Enabled = False
    
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    CmdMoveSubGateGamen.Enabled = False
    CmbCornerName.Enabled = False
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��

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
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������Ƃ���B
    cmbEkiInfo.Enabled = True
    CmdMenu(0).Enabled = True
    CmdMenu(1).Enabled = True
    CmdMenu(2).Enabled = True
    CmdMenu(3).Enabled = True
    CmdMoveGateGamen.Enabled = True
    cmdCancel.Enabled = True
 
 ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    CmdMoveSubGateGamen.Enabled = True
    CmbCornerName.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��

    DoEvents
    
End Sub
'V1.12.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : CmdMoveGateGamen_Click
'//  �@�\����  : �G���R�[�h�R�[�i�ݒ��ʐؑ�
'//  �@�\�T�v  : �w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i�ݒ�j��ʂ�\������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EGR HK1.1.0.1) 2011-05-11  CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//                 EGR HK���p
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveSubGateGamen_Click()
    
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    DoEvents
   
   '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SUBGATE_GAMEN_GO_BUTTOM, 0)

    '�\������ʃA�����[�h
    Unload Me
                
    Load frmEkiDataSubGate
    frmEkiDataSubGate.Show 1

End Sub

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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : pfuncInstallEkiSettei
'//  �@�\����  : �w�ݒ�C���X�g�[������
'//  �@�\�T�v  :
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
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfuncInstallEkiSettei() As Boolean

    Dim iResponse           As Integer          'MsgBox�߂�l
    Dim bRet                As Boolean          '�֐��߂�l
    Dim lErrCode            As Long             '�G���[�R�[�h

    Dim bSysChange              As Boolean      '�V�X�e���ݒ菈���߂�l
    Dim bUpData                 As Boolean      '��ʍX�V�����߂�l

    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse

    pfuncInstallEkiSettei = True

    '���݉w�ݒ�f�[�^�C���X�g�[������
    bRet = dllInstolEkiDataNow(gstrFileName, EKI_SETTI_FILE, lErrCode)
    
    If bRet = False Then
            
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
            
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        pfuncInstallEkiSettei = False
        '�ُ�I��
        iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�w�ݒ���͌���")
            
    Else
        
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        '----------------------------------------------------
        '�R���s���[�^���A�l�b�g���[�N�ύX����
        '----------------------------------------------------
        bUpData = True
        bSysChange = True
        bSysChange = pfNetWorkChng(Me)
         '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
            
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
            
        '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���쐬
        bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
        If bRet = False Then
            '�w�s�x�f�[�^�m�F�i�w���j�C���[�W�t�@�C���폜
            Kill EKI_TUDO_CHK_EKI_INFO_FILE
               
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)
            bUpData = False
            pfuncInstallEkiSettei = False
        End If

        '�w���R���{�{�b�N�X�����l�ݒ�
        cmbEkiInfo.Clear
        cmbEkiInfo.AddItem "�w���"
        cmbEkiInfo.AddItem "�Ď�"
        cmbEkiInfo.AddItem "�l�b�g���[�N"
        cmbEkiInfo.AddItem "���"
        cmbEkiInfo.ListIndex = 0

        '�R�[�i�ݒ�R���{�{�b�N�X�̏���������
        Call InitCornerComboBox
            
        '��ʕ\������
        Call sDisp
            
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
        If bSysChange = True And bUpData = True Then
            
            '����I��
            iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�w�ݒ���͌���")
        End If
    End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
    Kill gstrFileName
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
    gstrFileName = ""
    Call SetEnableTrue
End Function

