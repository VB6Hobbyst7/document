VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmKikiDataSubGate 
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleMode       =   0  'հ�ް
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
      TabIndex        =   15
      Top             =   700
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox cmbGoki 
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
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
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
      Index           =   7
      Left            =   7250
      TabIndex        =   13
      Top             =   7800
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdKikiSetMenu 
      Caption         =   "�w����ʂ�"
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
      Left            =   7250
      TabIndex        =   12
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   9120
      Top             =   2160
   End
   Begin VB.ComboBox CmbDummy 
      Height          =   345
      Left            =   4080
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   11
      Top             =   9720
      Width           =   2655
   End
   Begin VB.ListBox ListDummy 
      Height          =   510
      Left            =   120
      TabIndex        =   10
      Top             =   9720
      Width           =   1935
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
      TabIndex        =   0
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   7800
      Width           =   2175
   End
   Begin VB.TextBox txtDummy 
      Height          =   495
      IMEMode         =   3  '�̌Œ�
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   10425
      Width           =   3135
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
      TabIndex        =   2
      Top             =   7800
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   5730
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10107
      _Version        =   393216
      Rows            =   18
      Cols            =   8
      FixedCols       =   2
      RowHeightMin    =   350
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
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   720
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�@��\���ݒ�i�G���R�[�h�R�[�i���@����`�j"
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
      TabIndex        =   8
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiDataSubGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �F�@��\���ݒ�i�G���R�[�h�j���.frm
'//  �p�b�P�[�W���F�@��\���ݒ�i�G���R�[�h�j��ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�@��\���ݒ�i�G���R�[�h�j���
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή�
'//                 �yHKRK_Kansi07_003_01�zSub_gate_kan.ini�t�H�[�}�b�g�������Ή�
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   '���C���^�C�}�̃C���^�[�o���l
Private Const TITOL_EKI_NAME = "�w���F"                 '�w���^�C�g��       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
Private bScroll As Boolean

'�ݒ蔽�f�t���O
Private SetteiHaneiFlg As Boolean

'�@����f�[�^�X�V�t���O
Private KikiDataUpDateFlg As Boolean

'�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���ǎ�p�̍\����
Private Type SUBGATE_IMAGE_FILE
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

'Private Const START_DATA_COL_INDEX = 1       '1�s�̃f�[�^�ݒ���J�n����J�����C���f�b�N�X  'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
Private Const START_DATA_COL_INDEX = 2       '1�s�̃f�[�^�ݒ���J�n����J�����C���f�b�N�X   'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
'Private Const MAX_DATA_COL_INDEX = 6         '1�s�̍ő�ݒ�J������        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'Private Const MAX_DATA_COL_INDEX = 3         '1�s�̍ő�ݒ�J������         ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�   'EG20 V30.1.0.1 DEL
'Private Const MAX_DATA_COL_INDEX = 6         '1�s�̍ő�ݒ�J������         ' EG20 V30.1.0.1 ADD   'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
Private Const MAX_DATA_COL_INDEX = 7         '1�s�̍ő�ݒ�J������         ' EG20 V30.1.0.1 ADD    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
Private Const FM_CORNER_COL = 2                 'FM���R�[�i�ԍ��̗�iCOL�v���p�e�B�j
Private Const FM_GOKI_COL = 3                   'FM�����@�ԍ��̗�iCOL�v���p�e�B�j
Private Const SINKANSENIC_CORNER_COL = 4        '�V����IC�R�[�i�ԍ��iCOL�v���p�e�B�j
Private Const SINKANSENIC_GOKI_COL = 5          '�V����IC���@�ԍ��iCOL�v���p�e�B�j
Private Const NRZ_CORNER_COL = 6                'NRZ���R�[�i�ԍ�(COL�v���p�e�B)
Private Const NRZ_GOKI_COL = 7                  'NRZ�����@�ԍ�(COL�v���p�e�B)
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zADD END


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �@����ݒ�i�G���R�[�h�j���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �őO�O�\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    '�G���[���[�`����錾
    On Error Resume Next
    
    '����ʍőO�ʕ\���������s���B
    pfFormActive (hwnd)
    
    '�t�H�[�J�X�ʒu��ݒ�
    cmdCancel.SetFocus
    
    '�^�C�}���N������
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �@����ݒ�i�G���R�[�h�j���(�f�B�A�N�e�B�u��)
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
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �@����ݒ�i�G���R�[�h�j���(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim bRet                 As Boolean         '�֐��߂�l
    Dim lErrCode             As Long            '�G���[�R�[�h
    Dim iLoopCnt             As Integer         '���[�v�J�E���^
    
    '�G���[���[�`����錾
    On Error Resume Next

    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_START, 0)
    
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
    
    '�@����ݒ�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���쐬
    bRet = dllGetKikiIniData(2, 0, KIKI_DATA_SET_SUBGATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
    If bRet = False Then
        '�@����ݒ�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���폜
        Kill KIKI_DATA_SET_SUBGATE_FILE
        '�ُ탍�O�o��
        Call pfOutPutErrLog(lErrCode)
    End If

    '���@�R���{�{�b�N�X�����l
    cmbGoki.Clear

    'For iLoopCnt = 0 To 15 'EG20 V30.1.0.1 DEL
    For iLoopCnt = 0 To 31  'EG20 V30.1.0.1 ADD
            cmbGoki.AddItem iLoopCnt + 1 & "���@"
    Next
    cmbGoki.ListIndex = 0

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
    SetteiHaneiFlg = False

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(V30.1.0.1) 2014-06-04 REVISED BY [TCC] T.Nakajima
'//                 �k���V�������s�Ή�
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
                'AppActivate frmInputMstData.Caption, False     'EG20 V30.1.0.1 DEL
                'EG20 V30.1.0.1 ADD START
                AppActivate frmKikiDataSubGate.Caption, False
                pfFormActive (frmKikiDataSubGate.hwnd)
                'EG20 V30.1.0.1 ADD END
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
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    Dim iResponse           As Integer          'MsgBox�߂�l
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    If SetteiHaneiFlg = True Then
        iResponse = MsgBox("��ʕ\�����ɐݒ肳�ꂽ�f�[�^�������܂��B" & Chr(vbKeyReturn) & _
                            "��낵���ł����H", _
                            vbYesNo + vbQuestion, _
                            "�ݒ蔽�f�t������")
        
        If iResponse = vbNo Then Exit Sub
    End If
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_END, 0)
    
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDisp()

    Dim strFileName          As String          '�t�@�C����
    Dim nCornerIndex         As Integer         ' �R�[�i�I�����
    
    ' EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
    Dim iLoopCnt            As Integer          '���[�v�J�E���^�i�Z���c�����j
    Dim iLoopCnt2           As Integer          '���[�v�J�E���^�i�Z���������j
    ' EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END

    '�G���[���[�`����錾
    On Error Resume Next

' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    If CmbCornerName.ListIndex < 0 Then
        Exit Sub
    End If
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��

    '�����l�ݒ�
    strFileName = ""                            '�t�@�C����
    cmbGoki.Enabled = False                     '���@�I���R���{�{�b�N�X�I��s�ݒ�
    CmbCornerName.Enabled = False               ' �R�[�i�I�𕔑I��s��      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    LblEkiName.Caption = TITOL_EKI_NAME         '�w�����x��������           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    '----------------------------------------------------
    '�O���b�h�^�C�g���ݒ�
    '----------------------------------------------------
    Call sDispGridTitol
    
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
        Call sDispDataClear
        GridIni.Enabled = False
        
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
    
    '�@��\�����i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C������
    strFileName = Dir(KIKI_DATA_SET_SUBGATE_FILE)
    
    '�t�@�C�������݂����ꍇ
    If strFileName <> "" Then
    
        '�O���b�h�f�[�^���ݒ�
'        Call sDispDataSet(cmbGoki.ListIndex + 1)                               ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'        nCornerIndex = CmbCornerName.ListIndex
'        Call sDispDataSet(cmbGoki.ListIndex + 1, nCornerIndex + 1)
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        '�啪��:5�̉w�s�x�f�[�^�̓R�[�i�w��ł͂Ȃ��Ȃ������߁A�R�[�i��0�Œ�Ƃ���B�i�w�s�x�̃R�[�i��0�Ō�������j
        '1�`32���@
        For iLoopCnt = 0 To 31
            '���ڇ@�`�E
            For iLoopCnt2 = 0 To 5
                Call sDispDataSet(iLoopCnt + 1, 0, iLoopCnt2 + 1)
            Next
        Next
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
    Else
        '�ُ탍�O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_SUBGATE_IMAGE, 0)
        
        '�O���b�h�f�[�^���N���A����
        Call sDispDataClear
        GridIni.Enabled = False
        
        '�����t�����s�\�ݒ�
        CmdKikiSetMenu(0).Enabled = False           '�@��\�����ڐݒ蔽�f
        CmdKikiSetMenu(1).Enabled = False           '�@��\�����ڔ}�̏o��
        CmdKikiSetMenu(2).Enabled = False           '�@��\�����ړ����ۑ�

    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispGridTitol()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    Dim ColCount                As Integer         ' �J�����J�E���^
    Dim RowCount                As Integer         '���[�v�J�E���^

    GridIni.Visible = False             '�ݒ蒆�͔�\���ɐݒ�
    
    '�O���b�h�^�C�g���ݒ�
    With GridIni
    
        '----------------------------------
        '�O���b�h�̏�����
        '----------------------------------
        .Clear
        
        '----------------------------------
        '�O���b�h�Z�����ݒ�
        '----------------------------------
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'        .Rows = 18
'        .Cols = 7
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'        .Rows = 5
'        .Cols = 4
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
' EG20 V30.1.0.1 DEL END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zDEL START
'' EG20 V30.1.0.1 ADD START
'        .Rows = 2
'        .Cols = 7
'' EG20 V30.1.0.1 ADD END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zDEL END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zADD START
        .Rows = 33
        .Cols = 8
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zADD END
        For ColCount = 2 To (.Cols - 1)
            .ColWidth(ColCount) = 1748
        Next

        '----------------------------------
        '�O���b�h���ݒ�
        '----------------------------------
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000         'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
        
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'        For ColCount = 1 To (.Cols - 1)
'            .ColWidth(ColCount) = 1748
'        Next
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        For ColCount = 2 To (.Cols - 1)
            .ColWidth(ColCount) = 1548
        Next
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END

        '----------------------------------
        '�^�C�g���ݒ�
        '----------------------------------
        For RowCount = 1 To (.Rows - 1)
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
            '���@�\��
            .Col = 0
            .Row = RowCount: .Text = RowCount & "���@"
            .CellAlignment = flexAlignLeftCenter
            
            '���ЁE���Е\���i�k���V�����ł͉w�s�x�͎��Ђ̂݁j
            .Col = 1
            .Row = RowCount: .Text = "����"
            .CellAlignment = flexAlignLeftCenter
            
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'            If RowCount = 1 Then
'                '���Аݒ�
'                .Col = 0
'                .Row = RowCount: .Text = "����"
'                .CellAlignment = flexAlignLeftCenter
'
'            Else
'                '���Аݒ�
'                .Col = 0
'                .Row = RowCount: .Text = "����" & RowCount - 1
'                .CellAlignment = flexAlignLeftCenter
'            End If
        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
        
        Next

        .RowHeight(0) = 500

    End With

    GridIni.Visible = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sDispDataClear
'//  �@�\����  : �O���b�h�f�[�^���N���A����
'//  �@�\�T�v  : �O���b�h�f�[�^�����N���A����
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear()
    
    Dim iRowCnt             As Integer         '���[�v�J�E���^
    Dim ColCount             As Integer         '���[�v�J�E���^
    
    '�G���[���[�`����錾
    On Error Resume Next

    GridIni.Visible = False             '�ݒ蒆�͔�\���ɐݒ�
    
    '�O���b�h������
    With GridIni

        For iRowCnt = 1 To (.Rows - 1)
        
            .Row = iRowCnt

            '���ڐݒ�
            For ColCount = 2 To (.Rows - 1)
                .Col = ColCount
                .Text = ""
                .CellAlignment = flexAlignLeftCenter
            Next

        Next

    End With

    GridIni.Visible = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sDispDataSet
'//  �@�\����  : �O���b�h�f�[�^���ݒ菈��
'//  �@�\�T�v  : �O���b�h�f�[�^����ݒ肷��
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sDispDataSet(iGoki As Integer)                       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer)    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ� 'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer, iKomoku As Integer)    ' EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
    
    Dim intFileNumber       As Integer                      ' �t�@�C���|�C���^
    Dim iKikiDataCnt        As Integer                      ' �@����f�[�^�J�E���^
    Dim ColCount            As Integer                      ' �J�����J�E���^
    Dim RowCount            As Integer                      ' �s�J�E���^
    
    Dim strBunrui_Dai       As String                       ' �啪��
    Dim strBunrui_Tyu       As String                       ' ������
    Dim strBunrui_Sho       As String                       ' ������
    Dim strKomoku           As String                       ' ����
    Dim strKubun            As String                       ' �敪
    Dim strData             As String                       ' �ݒ�l
    Dim strSetShosai        As String                       ' �ݒ�l�ڍ�
    
    Dim strDispData         As String                       ' �\���f�[�^
    Dim strCorner           As String                       ' �R�[�i    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Dim iCmpCorner          As Integer                      ' �R�[�i    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�����l�ݒ�
    iKikiDataCnt = 0
    
    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C�����I�[�v������B
    Open KIKI_DATA_SET_SUBGATE_FILE For Input As #intFileNumber
    
    GridIni.Visible = False             '�ݒ蒆�͔�\���ɐݒ�

    ColCount = START_DATA_COL_INDEX     '�f�[�^�ݒ�̃X�^�[�g�J�����C���f�b�N�X
    RowCount = 1                        '�f�[�^�ݒ�̃X�^�[�g�s�C���f�b�N�X
    Do While Not EOF(intFileNumber)
        '�P �s�ǂݍ���
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, _
'                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
        Input #intFileNumber, strBunrui_Dai, strBunrui_Tyu, strBunrui_Sho, strCorner, _
                              strKomoku, strKubun, strData, strSetShosai
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��

        '�@����f�[�^�X�V�t���O�`�F�b�N
        If KikiDataUpDateFlg = False Then
            For iKikiDataCnt = 0 To UBound(KikiDataTbl)
        
                '�Y���f�[�^����
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'                If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(iKikiDataCnt).iBunrui_Dai) And _
'                   (CInt(strBunrui_Tyu) = KikiDataTbl(iKikiDataCnt).iBunrui_Tyu) And _
'                   (CInt(strBunrui_Sho) = KikiDataTbl(iKikiDataCnt).iBunrui_Sho) Then
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
                If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(iKikiDataCnt).iBunrui_Dai) And _
                   (CInt(strBunrui_Tyu) = KikiDataTbl(iKikiDataCnt).iBunrui_Tyu) And _
                   (CInt(strBunrui_Sho) = KikiDataTbl(iKikiDataCnt).iBunrui_Sho) And _
                   (CInt(strCorner) = KikiDataTbl(iKikiDataCnt).iBunrui_Corner) Then
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
                  
                    strData = KikiDataTbl(iKikiDataCnt).strData
                    strData = StrConv(strData, vbUnicode)
                    
                End If
            Next
        End If
                
        '���@�ԍ��`�F�b�N
        If CStr(iGoki) = strBunrui_Tyu Then
            If iKomoku = CInt(strBunrui_Sho) Then       'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
                ' �R�[�i����ǉ�
                ' �I�������R�[�i�A�������̓R�[�i���֌W�̃��R�[�h�͍̗p����
                iCmpCorner = CInt(strCorner)
                'If (iCorner = iCmpCorner) Then 'EG20 V30.1.0.1 DEL
                If ((iCorner = iCmpCorner) Or (iCmpCorner = 0)) Then 'EG20 V30.1.0.1 ADD
        
        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
                
                    '�O���b�h�ݒ�
                    With GridIni
                
                        '�J�����C���f�b�N�X�ݒ�
                        '.Col = ColCount                    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
                        .Col = ColCount + (iKomoku - 1)     'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
                        
                       '�^�C�g���ݒ�
                        If (strKomoku <> "") Then
                            .Row = 0
                            .Text = strKomoku
                            .CellAlignment = flexAlignLeftCenter
                        End If
        
                        '���ڐݒ�
                        '.Row = RowCount        'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
                        .Row = iGoki            'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
                        .Text = Format(pfDispIniData(.Text, strData, strKubun), "000")
                        .CellAlignment = flexAlignLeftCenter
                        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD START
                        '�w�s�x�f�[�^1���R�[�h���̐ݒ�l���Z���ɃZ�b�g�����̂ŁA��U�I��炷�B
                        Exit Do
                        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD END
                        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL START ��LExit Do�Ƀ��W�b�N�ύX�������ߕs�v�ɂȂ����B
        '                ColCount = ColCount + 1
        '                If ColCount > MAX_DATA_COL_INDEX Then
        '                 ColCount = START_DATA_COL_INDEX
        '                 RowCount = RowCount + 1
        '                End If
                       'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL END
        
                    End With
                
                End If          ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
            End If          'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�zADD
        End If
    
    Loop

    GridIni.Visible = True
    
    '�t�@�C�����N���[�Y����B
    Close #intFileNumber
    
    '���@�I���R���{�{�b�N�X�I���ݒ�
    cmbGoki.Enabled = True
    CmbCornerName.Enabled = True                ' �R�[�i�I�𕔑I����      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�

    '�����t�����\�ݒ�
    CmdKikiSetMenu(0).Enabled = True            '�@��\�����ڐݒ蔽�f
    CmdKikiSetMenu(1).Enabled = True            '�@��\�����ڔ}�̏o��
    CmdKikiSetMenu(2).Enabled = True            '�@��\�����ړ����ۑ�

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
    Call sDispDataClear
    GridIni.Enabled = False

    GridIni.Visible = True

    '�����t�����s�\�ݒ�
    CmdKikiSetMenu(0).Enabled = False           '�@��\�����ڐݒ蔽�f
    CmdKikiSetMenu(1).Enabled = False           '�@��\�����ڔ}�̏o��
    CmdKikiSetMenu(2).Enabled = False           '�@��\�����ړ����ۑ�

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Click()
    
    Dim iLoopCnt As Integer
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�N���b�N���ꂽ�ʒu�Ƀ_�~�[�e�L�X�g���ړ����A�t�H�[�J�X�����킹��
    With GridIni
        
        CmbDummy.Left = .Left + .CellLeft
        CmbDummy.Top = .Top + .CellTop
        CmbDummy.Width = .CellWidth
        CmbDummy.Height = .CellHeight
        CmbDummy.Text = .Text
        CmbDummy.Visible = True
        CmbDummy.SetFocus

        '�_�~�[�R���{�{�b�N�X�����l
        CmbDummy.Clear
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        '�N���b�N���ꂽ��ɂ���ăR���{�{�b�N�X����I�ׂ�l��؂�ւ���
        Select Case .Col
            'FM���R�[�i�ԍ��A�V����IC�R�[�i�ԍ��ANRZ���R�[�i�ԍ�
            Case FM_CORNER_COL, SINKANSENIC_CORNER_COL, NRZ_CORNER_COL
                '���͒l��00�`99
                '�w�s�x�d�l�ł�00�`99�����A��O�I�Ȓl���l������255�܂œ������悤�A�Ƃ肠�������Ă����ė~�����Ɠ��ŗl����̎w���ɂ��
                '000�`255�ɕύX
                For iLoopCnt = 0 To 255
                    CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
                    
                    '�R���{�{�b�N�X�̃C���f�b�N�X��ݒ�
                    If .Text <> "" Then
                        If iLoopCnt = CInt(.Text) Then
                            
                            '�l����v������C���f�b�N�X�ݒ�
                            CmbDummy.ListIndex = iLoopCnt
                            
                        End If
                    End If
                Next
            'FM�����@�ԍ��A�V����IC���@�ԍ��ANRZ�����@�ԍ�
            Case FM_GOKI_COL, SINKANSENIC_GOKI_COL, NRZ_GOKI_COL
                '���͒l��01�`16
                '�w�s�x�d�l�ł�01�`16�����A��O�I�Ȓl���l������255�܂œ������悤�A�Ƃ肠�������Ă����ė~�����Ɠ��ŗl����̎w���ɂ��
                '000�`255�ɕύX
                For iLoopCnt = 0 To 255
                    'CmbDummy.AddItem Format(CStr(iLoopCnt + 1), "000")
                    CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
                    
                    '�R���{�{�b�N�X�̃C���f�b�N�X��ݒ�
                    If .Text <> "" Then
                        'If (iLoopCnt + 1) = CInt(.Text) Then
                        If (iLoopCnt) = CInt(.Text) Then
                            
                            '�l����v������C���f�b�N�X�ݒ�
                            CmbDummy.ListIndex = iLoopCnt
                            
                        End If
                    End If
                Next
            Case Else
                For iLoopCnt = 0 To 255
                    CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
                    
                    '�R���{�{�b�N�X�̃C���f�b�N�X��ݒ�
                    If iLoopCnt = CInt(.Text) Then
                        
                        '�l����v������C���f�b�N�X�ݒ�
                        CmbDummy.ListIndex = iLoopCnt
                        
                    End If
                Next
            End Select
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'        For iLoopCnt = 0 To 255
'            CmbDummy.AddItem Format(CStr(iLoopCnt), "000")
'
'            '�R���{�{�b�N�X�̃C���f�b�N�X��ݒ�
'            If iLoopCnt = CInt(.Text) Then
'
'                '�l����v������C���f�b�N�X�ݒ�
'                CmbDummy.ListIndex = iLoopCnt
'
'            End If
'        Next
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
    
    End With

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub GridIni_Scroll()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�O���b�h���X�N���[�����ꂽ���A�_�~�[�e�L�X�g���\���ɂ���
    If bScroll = False Then
        CmbDummy.Visible = False
        CmbDummy.Clear
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : CmbDummy_Click
'//  �@�\����  : �_�~�[�e�L�X�g���I�����ꂽ���̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �O���b�h�ւ̔��f
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�C���Ή�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_Click()

    Dim iLoopCnt            As Integer                      ' ���[�v�J�E���^
    Dim iLoopCnt2           As Integer                      ' ���[�v�J�E���^
    Dim byBuff()            As Byte                         '�o�C�g�o�b�t�@
    Dim iGoki               As Integer                      ' ���@�ԍ�
    Dim iBunrui_Sho         As Integer                      ' ������
    Dim iBunrui_Corner      As Integer                      ' �R�[�i����            ' EG20 V3.0.0.2 �i�w�s�x�C���Ή��j�ǉ�

    '�G���[���[�`����錾
    On Error Resume Next

    If GridIni.Text <> CmbDummy.Text Then
        '�ݒ蔽�f�t���O�i�ύX����j
        SetteiHaneiFlg = True
    End If

    GridIni.Text = CmbDummy.Text
    GridIni.CellAlignment = flexAlignLeftCenter
    
    'iGoki = cmbGoki.ListIndex + 1                                              'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
    iGoki = GridIni.Row                                                         'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
    'iBunrui_Sho = ((GridIni.Row - 1) * MAX_DATA_COL_INDEX) + GridIni.Col       'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
    iBunrui_Sho = GridIni.Col - 1                                               'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
    'iBunrui_Corner = CmbCornerName.ListIndex + 1                                    ' EG20 V3.0.0.2 �i�w�s�x�C���Ή��j�ǉ� EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zDEL
    iBunrui_Corner = 0                                                           'EG20 V30.3.0.1�yHKRK_Kansi07_003_01�z �R�[�i�ʂł͂Ȃ��̂�0�Œ� ADD

    For iLoopCnt = 0 To UBound(KikiDataTbl)

        '�Y���f�[�^����
        If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(iLoopCnt).iBunrui_Dai) And _
           (iGoki = KikiDataTbl(iLoopCnt).iBunrui_Tyu) And _
           (iBunrui_Sho = KikiDataTbl(iLoopCnt).iBunrui_Sho) And _
           (iBunrui_Corner = KikiDataTbl(iLoopCnt).iBunrui_Corner) Then             ' EG20 V3.0.0.2 �i�w�s�x�C���Ή��j�ǉ�

            '�@��\�����f�[�^�ۑ�
            byBuff = StrConv(GridIni.Text, vbFromUnicode)

            Erase KikiDataTbl(iLoopCnt).strData

            '���I�z��̓��e�����O�p�����[�^�\���̂̐ÓI�z��Ɋi�[����B
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
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : CmbDummy_KeyDown
'//  �@�\����  : �L�[�{�[�h�������̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �_�~�[�e�L�X�g�̃Z�b�g
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '����L�[���������ꂽ���A���L�̏������s��
    bScroll = True
    
    With GridIni
    
        '�_�~�[�e�L�X�g�̃Z�b�g
        CmbDummy.Left = .Left + .CellLeft
        CmbDummy.Top = .Top + .CellTop
        CmbDummy.Width = .CellWidth
        CmbDummy.Height = .CellHeight
        CmbDummy.Text = .Text
        CmbDummy.Visible = True
        CmbDummy.SetFocus

    End With
    bScroll = False

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : CmbDummy_LostFocus
'//  �@�\����  : �_�~�[�e�L�X�g����t�H�[�J�X���ړ��������̃C�x���g�v���V�[�W��
'//  �@�\�T�v  : �_�~�[�e�L�X�g���\���ɂ���
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmbDummy_LostFocus()
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '�_�~�[�e�L�X�g���\���ɂ���
    CmbDummy.Visible = False
    CmbDummy.Clear

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmdKikiSetMenu_Click(Index As Integer)
    
    Dim iResponse           As Integer          'MsgBox�߂�l
    Dim bUnlock             As Boolean          ' ���b�N�����t���O      ' EG20 V3.0.0.2 �ǉ�

    '�G���[���[�`����錾
    On Error Resume Next
    
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    
' EG20 V3.0.0.2 �ǉ��J�n
' ���������t�ɉ����ă��b�N�����𐧌�����
' �����[����M��҂���
    bUnlock = True
' EG20 V3.0.0.2 �ǉ��I��
    Select Case Index
        
        Case 0                                 ' �@��\�����ڐݒ蔽�f
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_INSTOL, 0)
            
            '�@��\�����ڐݒ蔽�f����
            Call sInstolKikiData
            bUnlock = False                     ' EG20 V3.0.0.2 �ǉ�
        Case 1                                 ' �@��\�����ڔ}�̏o��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_OUTPUT, 0)
            
            '�@��\�����ڔ}�̏o�͏���
            Call sKikiDataOutPut
    
        Case 2                                 ' �@��\�����ړ����ۑ�
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_SAVE, 0)
            
            '�@��\�����ړ����ۑ�����
            Call sKikiDataSave
        
        Case 3                                 ' �@��\���ݒ�f�[�^�I��
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_KIKIDATA_SELECT, 0)
            
            '�@��\���ݒ�f�[�^�I������
            Call sKikiDataSelect
    
        Case 4                                 ' �}�̓���
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_MEDIUM_INPUT, 0)
            
            '�}�̓��͏���
            Call sInputMedium
    
        Case 5                                 ' �}�̎�O
            '��ʑ��샍�O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
    
            '�}�̎�O����
            Call pfRemove(Me)

        Case 6                                 ' �w����ʂ�
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
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIINFO_GAMEN_GO_BUTTOM, 0)
            Unload Me
            Load frmKikiData
            frmKikiData.Show 1
            Exit Sub

        Case 7                                 ' ������ʂ�
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
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)

            '�\������ʃA�����[�h
            Unload Me

            '������ʕ\��
            Load frmKikiDataGate
            frmKikiDataGate.Show 1
            Exit Sub

        Case Else
            '�����Ȃ�
            
    End Select

    '�S�{�^���������Ƃ���B
' EG20 V3.0.0.2 �ǉ��J�n
    If bUnlock = True Then
        Call SetEnableTrue
    End If
' EG20 V3.0.0.2 �폜�I��
'    Call SetEnableTrue                 ' EG20 V3.0.0.2 �폜

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.76�C���Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sInstolKikiData()

    Dim iResponse           As Integer          'MsgBox�߂�l
    Dim bRet                As Boolean          '�֐��߂�l
    Dim lErrCode            As Long             '�G���[�R�[�h
    Dim strFileName         As String           '�}�̃t�@�C����
    
    Dim bData()             As Byte             '�o�C�i���f�[�^
    Dim lLoopCnt            As Long             '���[�v�J�E���^
    Dim lLoopCnt2           As Long             '���[�v�J�E���^
    Dim bSysChange          As Boolean          '�R���s���[�^���A�l�b�g���[�N�ύX��������
    Dim byBuff()            As Byte             '�o�C�g�o�b�t�@
    Dim strSetteiData       As String           ' �ݒ�l

    Dim lResult             As Long             ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�

    '�G���[���[�`����錾
    On Error Resume Next
    
    iResponse = MsgBox("�@��\���f�[�^���C���X�g�[�����܂��B" & Chr(vbKeyReturn) & _
                        "��낵���ł����H", _
                        vbOKCancel + vbExclamation, _
                        "�ݒ蔽�f�m�F")
    If iResponse = vbCancel Then
        Call SetEnableTrue                      ' EG20 V5.0.2.1�y����TR-No.76�C���Ή��z�ǉ�
        Exit Sub
    End If
    
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
    
    '�@��\���f�[�^�e�[�u���i�啪�ށF�G���R�[�h�R�[�i���@�j�̐ݒ�l���u999�v�̏����ɕϊ�
    For lLoopCnt = 0 To UBound(KikiDataTbl)
        '�Y���f�[�^����
        If (BUNRUI_DAI.DAI_SubGate = KikiDataTbl(lLoopCnt).iBunrui_Dai) Then
            
            strSetteiData = KikiDataTbl(lLoopCnt).strData
            strSetteiData = StrConv(strSetteiData, vbUnicode)
            strSetteiData = Format(strSetteiData, "000")
    
            '�@��\�����f�[�^�ۑ�
            byBuff = StrConv(strSetteiData, vbFromUnicode)

            Erase KikiDataTbl(lLoopCnt).strData

            '���I�z��̓��e�����O�p�����[�^�\���̂̐ÓI�z��Ɋi�[����B
            For lLoopCnt2 = 0 To UBound(KikiDataTbl(lLoopCnt).strData)
                'Null�l�ɂȂ����珈���𔲂���B
                If byBuff(lLoopCnt2) = vbVEmpty Then Exit For

                KikiDataTbl(lLoopCnt).strData(lLoopCnt2) = byBuff(lLoopCnt2)

                '���I�z��̍ő�v�f�ɂȂ����珈���𔲂���
                If lLoopCnt2 = UBound(byBuff) Then Exit For
            Next
        End If
    Next

    '�\���̔z����o�C�i���z��ɕϊ�
    ReDim bData((UBound(KikiDataTbl) + 1) * Len(KikiDataTbl(0))) As Byte
    For lLoopCnt = 0 To UBound(KikiDataTbl)
          MoveMemory bData(lLoopCnt * Len(KikiDataTbl(0))), KikiDataTbl(lLoopCnt), Len(KikiDataTbl(lLoopCnt))
    Next
    
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
        iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f�����ݒ蔽�f����")
        Call SetEnableTrue              ' EG20 V3.0.0.2�i���[�����M���Ȃ��ꍇ�̂݃��b�N�����Ή��j�ǉ�
    Else
        '�R���s���[�^���A�l�b�g���[�N�ύX����
        
        bSysChange = pfNetWorkChng(Me)
        If bSysChange = False Then

            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
            '�ُ�I��
             iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f�����ݒ蔽�f����")
            Call SetEnableTrue              ' EG20 V3.0.0.2�i���[�����M���Ȃ��ꍇ�̂݃��b�N�����Ή��j�ǉ�
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
            iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "���f�����ݒ蔽�f����")
            
            '�ݒ蔽�f�t���O�i�ύX�Ȃ��j
            SetteiHaneiFlg = False
            Call SetEnableTrue              ' EG20 V3.0.0.2�i���[�����M���Ȃ��ꍇ�̂݃��b�N�����Ή��j�ǉ�
        End If
    End If


End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
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
    Dim lDrive               As Long            '�h���C�u�̃N���X�^���i���v�j
    Dim strDrive             As String          '�h���C�u
    
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
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
'        FileCopy KIKI_DATA_FILE, sWriteDir & Dir(KIKI_DATA_FILE)                                       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        FileCopy KIKI_DATA_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(KIKI_DATA_FILE)    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
        '����I��
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�̏o�͌���")
    
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

    '�ُ�I��
    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̏o�͌���")

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSave()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
    Dim sMyPath(1 To 3)      As String          '�t�@�C���p�X
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim iLoopCount           As Integer         '���[�v�J�E���^
    Dim intFileNo            As Integer         '�t�@�C���ԍ�

    '�G���[���[�`����錾
    On Error Resume Next
    
    '�����l�ݒ�
    strFileName = ""

    '----------------------------------------------------
    '�����ۑ�����
    '----------------------------------------------------
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
    
    '�t�@�C�����擾
    sWriteDir = KIKI_DATA_S_FILE
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
        FileCopy KIKI_DATA_FILE, sWriteDir
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
        '�ꎞ�ۑ��t�@�C���폜
        Kill KIKI_DATA_S_TEMP_FILE
        
        '����I��
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�ꎞ�ۑ�����")
    
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

        '�t�@�C������
        strFileName = Dir(KIKI_DATA_S_FILE)
        If strFileName <> "" Then
            '�t�@�C���폜
            Kill KIKI_DATA_S_FILE
        End If
        '�t�@�C�����̂����ɖ߂�
        Name KIKI_DATA_S_TEMP_FILE As KIKI_DATA_S_FILE
    
    '�ُ�I��
     iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�ꎞ�ۑ�����")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sKikiDataSelect()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
    Dim sMyPath(1 To 3)      As String          '�t�@�C���p�X
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim iLoopCount           As Integer         '���[�v�J�E���^
    Dim intFileNo            As Integer         '�t�@�C���ԍ�
    Dim bRet                 As Boolean         '�֐��߂�l
    Dim lErrCode             As Long            '�G���[�R�[�h

    '�G���[���[�`����錾
    On Error Resume Next
    
    '----------------------------------------------------
    '�@��\���f�[�^�t�@�C���X�V����
    '----------------------------------------------------
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
    
    '�t�@�C�����擾
    strFileName = Dir(KIKI_DATA_S_FILE)
    sWriteDir = strFileName
    If sWriteDir <> "" Then
    
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
        
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
         FileCopy KIKI_DATA_S_FILE, KIKI_DATA_FILE
        
        '�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C��
        bRet = dllGetKikiIniData(2, 1, KIKI_DATA_SET_SUBGATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)
            '�t�@�C���폜
            Kill KIKI_DATA_FILE
            '�t�@�C�����̂����ɖ߂�
            Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
            '�ُ�I��
            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�ꎞ�ۑ��f�[�^�捞����")
            Exit Sub
        End If
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
        
        '�ꎞ�ۑ��t�@�C���폜
        Kill KIKI_DATA_BACKUP_FILE
        
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
        
        '����I��
        iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�ꎞ�ۑ��f�[�^�捞����")
    
        '�@����f�[�^�X�V�t���O�ݒ�i�X�V�ݒ�j
        KikiDataUpDateFlg = True
        '��ʕ\������
        Call sDisp
        '�@����f�[�^�X�V�t���O�ݒ�i�ʏ�ݒ�j
        KikiDataUpDateFlg = False
        
        '�ݒ蔽�f�t���O�i�ύX����j
        SetteiHaneiFlg = True
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

   '�t�@�C������
   strFileName = Dir(KIKI_DATA_FILE)
   If strFileName <> "" Then
    '�t�@�C���폜
    Kill KIKI_DATA_FILE
   End If
   '�t�@�C�����̂����ɖ߂�
   Name KIKI_DATA_BACKUP_FILE As KIKI_DATA_FILE
   
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
   '�ُ�I��
   iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�ꎞ�ۑ��f�[�^�捞����")

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-12  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.77�C���Ή��z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
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
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    iResponse = MsgBox("�@��\���ݒ�̔}�̓��͂��s���܂��B" & vbCrLf & "��낵���ł����H", _
    vbOKCancel + vbQuestion, "�}�̓��͊m�F")
    
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
    
    Call ChDrive("D")

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
        
        '�@����ݒ�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���쐬
        bRet = dllGetKikiIniData(2, 1, KIKI_DATA_SET_SUBGATE_FILE, EKI_SETTI_FILE, KIKI_DATA_FILE, lErrCode)
        If bRet = False Then
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
            
            '�ُ�I��
            iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̓��͌���")
            
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
        SetteiHaneiFlg = True
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
    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�̓��͌���")

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
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
    CmdKikiSetMenu(7).Enabled = False
    cmdCancel.Enabled = False
    
    'CmdKikiSetMenu(0)�`(2)�͏����ɂ���Ă͌��X�����s�̂��ߔ�����s��
    If cmbGoki.Enabled = True Then
        cmbGoki.Enabled = False     '���@�I���R���{�{�b�N�X�I��s�ݒ�
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
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
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
    CmdKikiSetMenu(7).Enabled = True
    cmdCancel.Enabled = True
    
    '�R���{�{�b�N�X��CmdKikiSetMenu(0)�`(2)�͏����ɂ���Ă͌��X�����s�̂��߁A��ʕ\���̗L���Ŕ�����s��
    strFileName = Dir(KIKI_DATA_SET_SUBGATE_FILE)
    '�t�@�C�������݂���ꍇ
    If strFileName <> "" Then
        cmbGoki.Enabled = True              '���@�I���R���{�{�b�N�X�I���ݒ�
        CmbCornerName.Enabled = True        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        CmdKikiSetMenu(0).Enabled = True
        CmdKikiSetMenu(1).Enabled = True
        CmdKikiSetMenu(2).Enabled = True
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmbGoki_Click
'//  �@�\����  : ���@�I������
'//  �@�\�T�v  : �O���b�h�f�[�^���Đݒ肷��
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmbGoki_Click()
    
    Dim iIndex          As Integer                  '�C���f�b�N�X
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKISET_SUBGATE_GAMEN_GOKI_SELECT, 0)
    
    '��ʕ\������
    Call sDisp

    '�S�{�^���������Ƃ���B
    Call SetEnableTrue

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

