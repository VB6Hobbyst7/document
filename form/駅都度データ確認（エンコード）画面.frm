VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEkiDataSubGate 
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
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton CmdMoveEkiInfoGamen 
      Caption         =   "�w����ʂ�"
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
      TabIndex        =   11
      Top             =   8400
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   8520
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "�e�L�X�g�}�̏o��(�ݺ��޺�ō��@�ݒ�)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   4680
      TabIndex        =   8
      Top             =   7800
      Width           =   2295
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
      TabIndex        =   7
      Top             =   8400
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
      TabIndex        =   6
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   8160
      Top             =   6000
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
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
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
      Height          =   5730
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10107
      _Version        =   393216
      Rows            =   33
      Cols            =   9
      FixedCols       =   3
      RowHeightMin    =   350
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
      Caption         =   "�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j"
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
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   7815
   End
End
Attribute VB_Name = "frmEkiDataSubGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �F�w�s�x�f�[�^�m�F�i�G���R�[�h�j���.frm
'//  �p�b�P�[�W���F�w�s�x�f�[�^�m�F�i�G���R�[�h�j��ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�w�s�x�f�[�^�m�F�i�G���R�[�h�j���.frm
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//                 EG-R��}�@���p
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή�
'//                 �yHKRK_Kansi07_003_01�z Sub_gate_kan.ini�t�H�[�}�b�g�������Ή�
'//                 �yHKRK_Kansi07_008_01�z �w�s�x�f�[�^�̏����ނ��ӎ����ĕ\���A�}�̏o�͂��s��
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 REVISED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000                   '���C���^�C�}�̃C���^�[�o���l
'Private Const TITOL_EKI_NAME = "�w���@�@�@�F"           '�w���^�C�g��      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
Private Const TITOL_EKI_NAME = "�w���F"                 '�w���^�C�g��       ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�

'�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���ǎ�p�̍\����
Private Type SUBGATE_IMAGE_FILE
    sType       As String                '���
    sGoki       As String                '���@
    sNo         As String                '��ʖ��ʔ�
    sCorner     As String                '�R�[�i        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    sKoumoku    As String                '����
    sKubun      As String                '�敪
    sSettei     As String                '�ݒ�l
    sSyosai     As String                '�ݒ�l�ڍ�
End Type

'�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�o�̓t�@�C���쐬�p�̒萔��`
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'Private Const ONE_PARAGRAPH_OUTPUT_KOUMOKU = 48  '1�i���o�͍��ڐ�
'Private Const ONE_PARAGRAPH_OUTPUT_ROW = 16      '1�i���o�͗p�z��̗v�f��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'Private Const ONE_PARAGRAPH_OUTPUT_KOUMOKU = 24  '1�i���o�͍��ڐ�
'Private Const ONE_PARAGRAPH_OUTPUT_ROW = 3       '1�i���o�͗p�z��̗v�f��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
Private Const ONE_PARAGRAPH_OUTPUT_KOUMOKU = 48  '1�i���o�͍��ڐ�
Private Const ONE_PARAGRAPH_OUTPUT_ROW = 0       '1�i���o�͗p�z��̗v�f��
'EG20 V30.1.0.1 ADD END
Private Const ONE_PARAGRAPH_OUTPUT_GOKI = 8      '1�i���ɏo�͂��鍆�@��
Private Const GOKI_JISHA_START = "1"             '�����ށu���Ёv�̍ŏ�����
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
'Private Const GOKI_JISHA_END = "6"               '�����ށu���Ёv�̍ő區��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'Private Const GOKI_JISHA_END = "3"               '�����ށu���Ёv�̍ő區��
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'EG20 V30.1.0.1 DEL END
Private Const GOKI_JISHA_END = "6"               '�����ށu���Ёv�̍ő區��

'�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�o�̓t�@�C���쐬�p�̃e�[�u����`
Private Type SUBGATE_OUT_DEF_TBL
    iRow            As Integer                   '�s�ԍ�
    iNoStart        As Integer                   '�����ނ̍ŏ�����
    iNoEnd          As Integer                   '�����ނ̍ő區��
End Type

'�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�o�̓t�@�C���쐬�p�̍\���́i1�s�j
Private Type SUBGATE_IMAGE_FILE_ONE_ROW
    sShakyoku       As String                    '�Ћ�
    sKubun          As String                    '�敪
    sSettei         As String                    '�ݒ�l
End Type

'Private Const START_DATA_COL_INDEX = 2           '1�s�̃f�[�^�ݒ���J�n����J�����C���f�b�N�X  'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
Private Const START_DATA_COL_INDEX = 3           '1�s�̃f�[�^�ݒ���J�n����J�����C���f�b�N�X   'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
'Private Const MAX_DATA_COL_INDEX = 7             '1�s�̍ő�ݒ�J������    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'Private Const MAX_DATA_COL_INDEX = 4             '1�s�̍ő�ݒ�J������     ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ� 'EG20 V30.1.0.1 DEL
'Private Const MAX_DATA_COL_INDEX = 7             '1�s�̍ő�ݒ�J������     ' EG20 V30.1.0.1 ADD 'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
Private Const MAX_DATA_COL_INDEX = 8             '1�s�̍ő�ݒ�J������     ' EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD

Private gstrFileName        As String                       ' �o�̓t�@�C����    ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : CmbCornerName_Click
'//  �@�\����  : �R�[�i�I������
'//  �@�\�T�v  : �O���b�h�f�[�^���Đݒ肷��
'//
'//              �^        ����         �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l           �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�C���Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmbCornerName_Click()
    
    Dim iIndex          As Integer                  '�C���f�b�N�X
    
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
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �w�s�x�f�[�^�m�F�i�G���R�[�h�j���(�A�N�e�B�u���F�C�x���g�v���V�[�W��)
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
    
    '�^�C�}���N������
    tmrMail.Enabled = True
    
End Sub

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �w�s�x�f�[�^�m�F�i�G���R�[�h�j���(�f�B�A�N�e�B�u��)
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

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �w�s�x�f�[�^�m�F�i�G���R�[�h�j���(���[�h���F�C�x���g�v���V�[�W��)
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
'//     REVISIONS :(V30.1.0.1) 2014-05-20 CODED BY [TCC] T.Nakajima
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_SUBGATE_GAMEN_START, 0)
    
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
    
    '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���쐬
    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���폜
        Kill EKI_TUDO_CHK_SUBGATE_FILE
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
    Call sDisp
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
'//                 �k���V�����J�ƑΉ�
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
    Dim iLoopCnt As Integer          ' ���[�v
    
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
                'AppActivate frmInputMstData.Caption, False      'EG20 V30.1.0.1 DEL
                'EG20 V30.1.0.1 ADD START
                AppActivate frmEkiDataSubGate.Caption, False
                pfFormActive (frmEkiDataSubGate.hwnd)
                'EG20 V30.1.0.1 ADD END
            Case ML_ID_HOSHU_OPERATE_PROG_SNDREQ_RES
                '�u�ێ瑀���v���O�������M�v���v����M�����ꍇ
                If pubfuncRespCheckTakuProgramData(udtReadMail) = False Then
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
                    Kill gstrFileName
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
                    '�v���O���X�o�[����������
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
    
    '�G���[���[�`����錾
    On Error Resume Next
    
    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_SUBGATE_GAMEN_END, 0)
    
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
    Dim iLoopCnt             As Integer         '���[�v�J�E���^
    Dim iLoopCnt2            As Integer         '���[�v�J�E���^ EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
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
    cmbGoki.Enabled = False                     '���@�R���{�{�b�N�X�I��s�ݒ�
    CmbCornerName.Enabled = False               ' �R�[�i�I�𕔑I��s��      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    LblEkiName.Caption = TITOL_EKI_NAME         '�w�����x��������
    
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
    '��ʃ��C�A�E�g�ύX�ɂ��A�R�[�i�ƍ��@�̃R���{�{�b�N�X�͕s�v�ɂȂ����B
    cmbGoki.Visible = False
    CmbCornerName.Visible = False
    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
    
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
        Call sDispDataClear

        Exit Sub
        
    End If
    
    '----------------------------------------------------
    '�w�����x���X�V
    '----------------------------------------------------
    LblEkiName.Caption = TITOL_EKI_NAME & pfGetEkiNameInfo(NotEkiVer)
    
    '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C��
    strFileName = Dir(EKI_TUDO_CHK_SUBGATE_FILE)
    
    '�t�@�C�������݂����ꍇ
    If strFileName <> "" Then
    
        '�O���b�h�f�[�^���ݒ�
'        Call sDispDataSet(cmbGoki.ListIndex + 1)                                   ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'        nCornerIndex = CmbCornerName.ListIndex
'        Call sDispDataSet(cmbGoki.ListIndex + 1, nCornerIndex + 1)
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
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
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28 CODED BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19 CODED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispGridTitol()
    
    Dim ColCount                As Integer         ' �J�����J�E���^
    Dim RowCount                As Integer         '���[�v�J�E���^
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�ݒ蒆�͔�\���ɐݒ�
    GridIni.Visible = False
    
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
'        .Cols = 8
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'EG20 V30.1.0.1 DEL START
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'        .Rows = 5
'        .Cols = 5
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'EG20 V30.1.0.1 DEL END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
''EG20 V30.1.0.1 ADD START
'        .Rows = 2
'        .Cols = 8
''EG20 V30.1.0.1 ADD END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        .Rows = 33
        .Cols = 9
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
        
        '----------------------------------
        '�O���b�h���ݒ�
        '----------------------------------
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zDEL START
'        .ColWidth(0) = 900
'        .ColWidth(1) = 700
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        .ColWidth(0) = 700
        .ColWidth(1) = 700
        .ColWidth(2) = 700
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL START
'        For ColCount = 2 To (.Cols - 1)
'            .ColWidth(ColCount) = 1675
'        Next
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        For ColCount = 3 To (.Cols - 1)
            .ColWidth(ColCount) = 1550
        Next
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END
        
        '----------------------------------
        '�^�C�g���ݒ�
        '----------------------------------
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD START
        '���@�ݒ�
        .Col = 0
        .Row = 0: .Text = "���@"
        .CellAlignment = flexAlignCenterCenter
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD END

        '�敪�ݒ�
        '.Col = 1    'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zDEL
        .Col = 2     'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zADD
        .Row = 0: .Text = "�敪"
        .CellAlignment = flexAlignCenterCenter
        For RowCount = 1 To (.Rows - 1)
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zADD START
            '�c�����̌Œ�\������ݒ肷��
            '���@�ݒ�i�P�`�R�Q�j
            .Col = 0
            .Row = RowCount: .Text = RowCount
            .CellAlignment = flexAlignCenterCenter
            
            '���ЁE���Аݒ�i�k���V�����̉w�s�x�͎��Ђ̂݁j
            .Col = 1
            .Row = RowCount: .Text = "����"
            .CellAlignment = flexAlignCenterCenter
            
            '�敪�i�k���V�����ł͋敪�͓����̂݁j
            .Col = 2
            .Row = RowCount: .Text = "����"
            .CellAlignment = flexAlignCenterCenter
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zADD END
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�zDEL START
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
'
'            .Col = 1
'            .Row = RowCount
''            .Text = "�Ď�"                 ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'            .Text = "����"                  ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'            .CellAlignment = flexAlignCenterCenter
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
'//              �^        ����      �Ӗ�
'//  ����      : Integer   intStartRow  �J�n�s�ʒu
'//              Integer   intEndRow    �I���s�ʒu
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  REVISED BY [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi07_003_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispDataClear()
    
    Dim iLoopCnt             As Integer         '���[�v�J�E���^
    Dim ColCount             As Integer         '���[�v�J�E���^
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�ݒ蒆�͔�\���ɐݒ�
    GridIni.Visible = False
    
    '�O���b�h������
    With GridIni

        For iLoopCnt = 1 To (.Rows - 1)

            '���ڐݒ�
            'For ColCount = 2 To (.Rows - 1) 'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z DEL
            For ColCount = 3 To (.Rows - 1) 'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z ADD
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
'//  ����      : Integer   iBunrui_Dai  �啪��
'//            : Integer   iCorner      �R�[�i  ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
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
'//                 �k���V�����t�F�[�Y�Q�Ή�
'//                 �yHKRK_Kansi07_003_01�z�A�yHKRK_Kansi07_008_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sDispDataSet(iGoki As Integer)                             ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer)          ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ� 'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL
Private Sub sDispDataSet(iGoki As Integer, iCorner As Integer, iKomoku As Integer)    'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
    
    Dim intFileNumber       As Integer                      ' �t�@�C���|�C���^
    Dim iLoopCnt            As Integer                      ' ���[�v�J�E���^
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

    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C�����I�[�v������B
    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
    
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

        '���@�ԍ��`�F�b�N
        If CStr(iGoki) = strBunrui_Tyu Then
            If iKomoku = CInt(strBunrui_Sho) Then       'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
                ' �R�[�i����ǉ�
                ' �I�������R�[�i�̃��R�[�h���̗p����
                iCmpCorner = CInt(strCorner)
                If (iCorner = iCmpCorner) Then
        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
                    '�O���b�h�ݒ�
                    With GridIni
                
                        '�J�����C���f�b�N�X�ݒ�
                        '.Col = ColCount                    'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�zDEL
                        .Col = ColCount + (iKomoku - 1)     'EG20 V30.3.0.1 �yHKRK_Kansi07_007_01�zADD
        
                       '�^�C�g���ݒ�
                        If (strKomoku <> "") Then
                            .Row = 0
                            .Text = strKomoku
                            .CellAlignment = flexAlignLeftCenter
                        End If
        
                        '���ڐݒ�
                        '.Row = RowCount        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL
                        .Row = iGoki            'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD
                        .Text = Format(pfDispIniData(.Text, strData, strKubun), "000")
                        .CellAlignment = flexAlignLeftCenter
                        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD START
                        '�w�s�x�f�[�^1���R�[�h���̐ݒ�l���Z���ɃZ�b�g�����̂ŁA��U�I��炷�B
                        Exit Do
                        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z ADD END
                        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL START ��LExit Do�Ƀ��W�b�N�ύX�������ߕs�v�ɂȂ����B
'                        ColCount = ColCount + 1
'                        If ColCount > MAX_DATA_COL_INDEX Then
'                         ColCount = START_DATA_COL_INDEX
'                         RowCount = RowCount + 1
'                        End If
                        'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�z DEL END
        
                    End With
                
                End If          ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
            End If          'EG20 V30.3.0.1 �yHKRK_Kansi07_008_01�zADD
        End If
    
    Loop

    GridIni.Visible = True
    
    '�t�@�C�����N���[�Y����B
    Close #intFileNumber

    '���@�R���{�{�b�N�X�I���ݒ�
    cmbGoki.Enabled = True
    CmbCornerName.Enabled = True               ' �R�[�i�I�𕔑I��s��      ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�

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

    GridIni.Visible = True
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
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKITUDOKAKUNIN_SUBGATE_GAMEN_GOKI_SELECT, 0)
    
    '��ʕ\������
    Call sDisp

    '�S�{�^���������Ƃ���B
    Call SetEnableTrue

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmdMenu_Click(Index As Integer)
  
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

  '�S�{�^���������Ƃ���B
' EG20 V3.0.0.2 �ǉ��J�n
    If bUnlock = True Then
        Call SetEnableTrue
    End If
' EG20 V3.0.0.2 �ǉ��I��
'    Call SetEnableTrue                 ' EG20 V3.0.0.2 �폜

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sEkiSetteiOutPut()

    Dim strFileName          As String          '�t�@�C����
    Dim sWriteDir            As String          '�t�H���_��
    Dim iResponse            As Integer         'MsgBox�߂�l

    '�G���[���[�`����錾
    On Error Resume Next
    iResponse = MsgBox("�I������Ă���w�̌��݂̉w�s�x�f�[�^�P�w�����o�͂��܂��B" & Chr(vbKeyReturn) & _
                        "��낵���ł����H", _
                        vbOKCancel + vbQuestion, _
                        "�w�ݒ�o�͊m�F")

    If iResponse = vbCancel Then Exit Sub

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
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    If sWriteDir <> "" Then
    
        On Error GoTo COPY_ERROR
        '�t�@�C���R�s�[
'        FileCopy EKI_SETTI_FILE, sWriteDir & Dir(EKI_SETTI_FILE)                   ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
        FileCopy EKI_SETTI_FILE, sWriteDir & Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & Dir(EKI_SETTI_FILE)    ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
        
        '���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
    
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

    iResponse = MsgBox("�ُ�I�����܂���", vbOKOnly + vbCritical, "�w�ݒ�o�͌���")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
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
'    Dim bSysChange              As Boolean      '�V�X�e���ݒ菈���߂�l
'    Dim bUpData                 As Boolean      '��ʍX�V�����߂�l
'    Dim iLoopCnt                As Integer      '���[�v�J�E���^
'
'    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
'
'    Dim lResult             As Long             ' ��������     ' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'
'    '�G���[���[�`����錾
'    On Error Resume Next
'    iResponse = MsgBox("�w�s�x�f�[�^�P�w�����C���X�g�[�����܂��B" & Chr(vbKeyReturn) & _
'                        "��낵���ł����H", _
'                        vbOKCancel + vbQuestion, _
'                        "�w�ݒ���͊m�F")
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
'
'    Call ChDrive("D")
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
'            '----------------------------------------------------
'            '�R���s���[�^���A�l�b�g���[�N�ύX����
'            '----------------------------------------------------
'            bSysChange = True
'            bUpData = True
'            bSysChange = pfNetWorkChng(Me)
'             '���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_SHORI_OK, 0)
'           '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���쐬
'            bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
'            If bRet = False Then
'                '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���폜
'                Kill EKI_TUDO_CHK_SUBGATE_FILE
'
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
'                '�v���O���X�o�[����������
'                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��
'
'                '�ُ탍�O�o��
'                Call pfOutPutErrLog(lErrCode)
'                bUpData = False
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
'            '���@�R���{�{�b�N�X�����l
'            cmbGoki.Clear
'            For iLoopCnt = 0 To 15
'                    cmbGoki.AddItem iLoopCnt + 1 & "���@"
'            Next
'            cmbGoki.ListIndex = 0
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
'            '����I��
'            iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�w�ݒ���͌���")
'            End If
'        End If
'    End If
'
'End Sub
' EG20 V3.0.0.2[Mainte_03_01 �w�s�x�Ή�]�폜�I���i�S�̌������j

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-28 CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19 CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��yHKRK_Kansi07_003_01�z�A�yHKRK_Kansi07_008_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDispTextEkiDataNow()

    Dim strFileName          As String          '�t�@�C����
    Dim iResponse            As Integer         'MsgBox�߂�l
    Dim lRetVal              As Long            '�߂�l
    Dim sCommand             As String          '�R�}���h������
    Dim sWriteDir            As String          '�������ݐ�t�H���_��
    Dim intFileNumber        As Integer         '�t�@�C���|�C���^
    Dim ColCount             As Integer         '�J�����J�E���^
    Dim RowCount             As Integer         '���[�v�J�E���^
    Dim TypeCount            As Integer         '���[�v�J�E���^
    Dim sData                As String          '���͗p������
    Dim strData_Kansi()      As String          '�Ď��Տ��ۑ��z��
    Dim strData_Ldu()        As String          'LDU���ۑ��z��
    Dim iLength              As Integer         '���s�R�[�h�����p�i�����j
    Dim iLeft                As Integer         '���s�R�[�h�����p�i�擪�j
    Dim iRight               As Integer         '���s�R�[�h�����p�i�I�[�j
    
    Dim ReadFileSettei()     As SUBGATE_IMAGE_FILE          '�t�@�C���Ǎ��p�\����
    Dim OutFileData1()       As SUBGATE_IMAGE_FILE_ONE_ROW  '�t�@�C���o�͗p�\����
    Dim OutFileData2()       As SUBGATE_IMAGE_FILE_ONE_ROW  '�t�@�C���o�͗p�\����
    Dim strOutDefTbl()       As SUBGATE_OUT_DEF_TBL         '�o�͏��e�[�u��
    Dim i                    As Integer             '���[�v�J�E���^�P
    Dim j                    As Integer             '���[�v�J�E���^�Q
    Dim k                    As Integer             '���[�v�J�E���^�R
    Dim strLineCount         As String              '�s���J�E���^
    Dim fso                  As New FileSystemObject        '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim FsoTS                As TextStream
    Dim strSaveFileName      As String          ' �ۑ��t�@�C����        ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Dim szCornerName         As String          ' �R�[�i����            ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
    Dim nNullIndex           As Integer         ' ���������[�N          ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�������ݐ�t�@�C���I��
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    If sWriteDir = "" Then
       '�t�H���_�I����ʁu����v�t�������͏����I��
       Exit Sub
    End If

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
        MsgBox "�e�L�X�g�\������f�[�^������܂���B", _
                vbOKOnly + vbExclamation, _
                 "�f�[�^���x��"
        Exit Sub

    End If

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_EKITSUDO)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��

    On Error GoTo OUTPUT_ERROR
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z�A�yHKRK_Kansi07_008_01�zADD START
    '���g�p�̃t�@�C���ԍ��擾
    intFileNumber = FreeFile
    
    '���݉w�ݒ�t�@�C�����I�[�v������
    Open PATH_WORK & EKI_SETTI_SUBGATE_FILE For Output As #intFileNumber
    
    '�^�C�g���\��
    Print #intFileNumber, "�ݒu�w�@�@�F" & Trim(pfGetEkiNameInfo(NotEkiVer))
    Print #intFileNumber, "�y�G���R�[�h�R�[�i���@����`�z"
    
    '�O���b�h�^�C�g���ݒ�
    With GridIni
    
        '�s�������[�v������
        For RowCount = 0 To .Rows - 1
            'sData������
            sData = ""
            '�e���ڕ\��
            If RowCount = 0 Then
                For ColCount = 0 To .Cols - 1
                    .Col = ColCount
                    .Row = RowCount
                    
                    If ColCount <> .Cols - 1 Then
                        sData = sData & Replace(.Text, " ", "") & ","
                    Else
                        sData = sData & Replace(.Text, " ", "")
                    End If
                Next
                Print #intFileNumber, sData
            Else
                .Row = RowCount
                .Col = 0
                
                '�Ē�`
                
                '���ڕ����[�v����
                For ColCount = 0 To .Cols - 1
                    .Col = ColCount
                    .Row = RowCount
                   
                    '�ݒ�l�擾
                    If ColCount <> .Cols - 1 Then
                        sData = sData & .Text & ","
                    Else
                        sData = sData & .Text
                    End If
                Next
                Print #intFileNumber, sData
            End If
        Next
    End With
    
    '�t�@�C�����N���[�Y����
    Close #intFileNumber
    
    ' �R�[�i���̂̕t��
    'strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & szCornerName & "_" & EKI_SETTI_SUBGATE_FILE
    strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & EKI_SETTI_SUBGATE_FILE
    '�ꎞ�t�@�C����}�̂ɃR�s�[����
    Call FileCopy(PATH_WORK & EKI_SETTI_SUBGATE_FILE, sWriteDir & strSaveFileName)
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z�A�yHKRK_Kansi07_008_01�zDEL END
'EG20 V30.3.0.1 �}�̏o�̓t�H�[�}�b�g�啝�������ɕt���폜�yHKRK_Kansi07_003_01�z�A�yHKRK_Kansi07_008_01�zDEL START
'    '///////////////////////////////////////////////////////////////////////////
'    '�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C�����ǂݍ���
'    '///////////////////////////////////////////////////////////////////////////
'    '�t�@�C���ԍ��擾
'    intFileNumber = FreeFile
'
'    'CSV�t�@�C���I�[�v��
'    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
'
'    'CSV�t�@�C���s���J�E���g�i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
''    Do While Not EOF(1)                                    ' EG20 V3.3.0.1�폜
'    Do While Not EOF(intFileNumber)                         ' EG20 V3.3.0.1�ǉ�
'        Line Input #intFileNumber, strLineCount
'        j = j + 1
'    Loop
'
'    'CSV�t�@�C���N���[�Y
'    Close #intFileNumber
'
'    '�t�@�C���ԍ��擾
'    intFileNumber = FreeFile
'
'    '�Đݒ�
'    ReDim ReadFileSettei(j) As SUBGATE_IMAGE_FILE        '�t�@�C���Ǎ��p�G���A
'
'    'CSV�t�@�C���I�[�v��
'    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
'
'    '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'    For i = 0 To j - 1
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
''        Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, _
''         ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'        Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, ReadFileSettei(i).sCorner, _
'         ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'    Next
'
'    'CSV�t�@�C���N���[�Y
'    Close #intFileNumber
'
'    '�Đݒ�
'    ReDim OutFileData1(ONE_PARAGRAPH_OUTPUT_ROW) As SUBGATE_IMAGE_FILE_ONE_ROW         '�t�@�C���o�͗p�\����
'    ReDim OutFileData2(ONE_PARAGRAPH_OUTPUT_ROW) As SUBGATE_IMAGE_FILE_ONE_ROW         '�t�@�C���o�͗p�\����
'
'    '�e���ڐݒ�l���o�͗p�\���̂ɕϊ�
'
'    ReDim strOutDefTbl(ONE_PARAGRAPH_OUTPUT_ROW) As SUBGATE_OUT_DEF_TBL                '�o�͏��e�[�u��
'    strOutDefTbl(0).iRow = 0
'    strOutDefTbl(0).iNoStart = GOKI_JISHA_START
'    strOutDefTbl(0).iNoEnd = GOKI_JISHA_END
'
'    For RowCount = 1 To ONE_PARAGRAPH_OUTPUT_ROW
'        strOutDefTbl(RowCount).iRow = RowCount
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
''        strOutDefTbl(RowCount).iNoStart = strOutDefTbl(RowCount - 1).iNoStart + 6
''        strOutDefTbl(RowCount).iNoEnd = strOutDefTbl(RowCount - 1).iNoEnd + 6
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'' EG20 V30.1.0.1 DEL START
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
''        strOutDefTbl(RowCount).iNoStart = strOutDefTbl(RowCount - 1).iNoStart + 3
''        strOutDefTbl(RowCount).iNoEnd = strOutDefTbl(RowCount - 1).iNoEnd + 3
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
''EG20 V30.1.0.1 DEL END
''EG20 V30.1.0.1 ADD START
'        strOutDefTbl(RowCount).iNoStart = strOutDefTbl(RowCount - 1).iNoStart + 6
'        strOutDefTbl(RowCount).iNoEnd = strOutDefTbl(RowCount - 1).iNoEnd + 6
''EG20 V30.1.0.1 ADD END
'
'    Next
'
'    '1���@�`8���@
'    For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'        '�Ћǂ�ݒ�
'        If RowCount = 0 Then
'            OutFileData1(RowCount).sShakyoku = "����"
'        Else
'            OutFileData1(RowCount).sShakyoku = "����" & StrConv(CStr(RowCount), vbWide)
'        End If
'
'        '�敪��ݒ�
''        OutFileData1(RowCount).sKubun = "�Ď�"         ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'        OutFileData1(RowCount).sKubun = "����"          ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'
'        '�ݒ�l��ݒ�
'        If RowCount = strOutDefTbl(RowCount).iRow Then
'            'i = 0  EG20 V30.0.1.1 DEL
'            i = CmbCornerName.ListIndex * MAX_GATE_NO * GOKI_JISHA_END  ' �R�[�i���ɏo�͂���  EG20V30.0.1.1 ADD
'
'            Do While (i < j)
'             If (CInt(ReadFileSettei(i).sGoki) <= ONE_PARAGRAPH_OUTPUT_GOKI) And _
'                (CInt(ReadFileSettei(i).sNo) >= strOutDefTbl(RowCount).iNoStart) And _
'                (CInt(ReadFileSettei(i).sNo) <= strOutDefTbl(RowCount).iNoEnd) Then
'
'                 If (CInt(ReadFileSettei(i).sGoki) = ONE_PARAGRAPH_OUTPUT_GOKI) And _
'                    (CInt(ReadFileSettei(i).sNo) = strOutDefTbl(RowCount).iNoEnd) Then
'
'                     OutFileData1(RowCount).sSettei = OutFileData1(RowCount).sSettei + _
'                                                      Format(ReadFileSettei(i).sSettei, "000") & vbCrLf
'                    Exit Do
'
'                 Else
'
'                     OutFileData1(RowCount).sSettei = OutFileData1(RowCount).sSettei + _
'                                                       Format(ReadFileSettei(i).sSettei, "000") & ","
'                 End If
'
'             End If
'
'             i = i + 1
'
'            Loop
'
'        End If
'
'   Next
'
'    '9���@�`16���@
'   For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'       '�Ћǂ�ݒ�
'       If RowCount = 0 Then
'           OutFileData2(RowCount).sShakyoku = "����"
'       Else
'           OutFileData2(RowCount).sShakyoku = "����" & StrConv(CStr(RowCount), vbWide)
'       End If
'
'       '�敪��ݒ�
''       OutFileData2(RowCount).sKubun = "�Ď�"          ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜
'       OutFileData2(RowCount).sKubun = "����"           ' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ�
'
'       '�ݒ�l��ݒ�
'       If RowCount = strOutDefTbl(RowCount).iRow Then
'           'i = 0       'EG20 V30.0.1.1 DEL
'           i = (CmbCornerName.ListIndex * MAX_GATE_NO * GOKI_JISHA_END) + 8  ' �R�[�i���ɏo�͂���  EG20V30.0.1.1 ADD
'
'           Do While (i < j)
'            If (CInt(ReadFileSettei(i).sGoki) > ONE_PARAGRAPH_OUTPUT_GOKI) And _
'               (CInt(ReadFileSettei(i).sGoki) <= (ONE_PARAGRAPH_OUTPUT_GOKI * 2)) And _
'               (CInt(ReadFileSettei(i).sNo) >= strOutDefTbl(RowCount).iNoStart) And _
'               (CInt(ReadFileSettei(i).sNo) <= strOutDefTbl(RowCount).iNoEnd) Then
'
'                If (CInt(ReadFileSettei(i).sGoki) = (ONE_PARAGRAPH_OUTPUT_GOKI * 2)) And _
'                   (CInt(ReadFileSettei(i).sNo) = strOutDefTbl(RowCount).iNoEnd) Then
'
'                    OutFileData2(RowCount).sSettei = OutFileData2(RowCount).sSettei + _
'                                                     Format(ReadFileSettei(i).sSettei, "000") & vbCrLf
'                   Exit Do
'
'                Else
'
'                    OutFileData2(RowCount).sSettei = OutFileData2(RowCount).sSettei + _
'                                                     Format(ReadFileSettei(i).sSettei, "000") & ","
'                End If
'
'            End If
'
'                i = i + 1
'
'           Loop
'
'       End If
'
'    Next
'
'
'    '///////////////////////////////////////////////////////////////////////////
'    '�@��\���f�[�^�i�G���R�[�h�R�[�i���@����`�j�t�@�C���o�͏���
'    '///////////////////////////////////////////////////////////////////////////
'    '�ꎞ�t�@�C�������
'    Set FsoTS = fso.CreateTextFile(PATH_WORK & EKI_SETTI_SUBGATE_FILE, True)
'
'    '�^�C�g���o��
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'    ' �R�[�i���̂̕t��
'    nNullIndex = InStr(gstrCornerName(CmbCornerName.ListIndex), Chr(0))
'    If nNullIndex <> 0 Then
'        szCornerName = Left(gstrCornerName(CmbCornerName.ListIndex), nNullIndex - 1)
'    Else
''        szCornerName = ""                                          ' EG20 V3.3.0.1�폜
'        szCornerName = gstrCornerName(CmbCornerName.ListIndex)      ' EG20 V3.3.0.1�ǉ�
'    End If
'
'    FsoTS.Write ("�ݒu�w�@�@�F" & Trim(pfGetEkiNameInfo(NotEkiVer)) & vbCrLf)
'    FsoTS.Write ("�ݒu�R�[�i�F" & szCornerName & vbCrLf)
'    FsoTS.Write (vbCrLf)
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'    FsoTS.Write ("�y�G���R�[�h�R�[�i���@����`�z" & vbCrLf)
'
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
''    '���ڃ^�C�g���o��
''    FsoTS.Write ("����,�敪,1���@,,,,,,2���@,,,,,,3���@,,,,,,4���@,,,,,,5���@,,,,,,6���@,,,,,,7���@,,,,,,8���@" & vbCrLf)
''
''    '���ڃ^�C�g���o��
''    FsoTS.Write (",,")
''    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
''        FsoTS.Write ("�@,�A,�B,�C,�D,�E,")     '1�`7���@
''    Next
''    FsoTS.Write ("�@,�A,�B,�C,�D,�E" & vbCrLf) '8���@
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'    '���ڃ^�C�g���o��
'    'FsoTS.Write ("����,�敪,1���@,,,2���@,,,3���@,,,4���@,,,5���@,,,6���@,,,7���@,,,8���@" & vbCrLf)   'EG20 V30.1.0.1 DEL
'    FsoTS.Write ("����,�敪,1���@,,,,,,2���@,,,,,,3���@,,,,,,4���@,,,,,,5���@,,,,,,6���@,,,,,,7���@,,,,,,8���@" & vbCrLf)   'EG20 V30.1.0.1 ADD
'
'    '���ڃ^�C�g���o��
'    FsoTS.Write (",,")
'    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
'        'FsoTS.Write ("�@,�A,�B,")     '1�`7���@     'EG20 V30.1.0.1 DEL
'        FsoTS.Write ("�@,�A,�B,�C,�D,�E,")     '1�`7���@     'EG20 V30.1.0.1 ADD
'    Next
'    'FsoTS.Write ("�@,�A,�B" & vbCrLf) '8���@   'EG20 V30.1.0.1 DEL
'    FsoTS.Write ("�@,�A,�B,�C,�D,�E" & vbCrLf) '8���@   'EG20 V30.1.0.1 ADD
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'
'    '�e���ڐݒ�l�o��
'    '1���@�`8���@
'    For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'        FsoTS.Write (OutFileData1(RowCount).sShakyoku & ",")
'        FsoTS.Write (OutFileData1(RowCount).sKubun & ",")
'        FsoTS.Write (OutFileData1(RowCount).sSettei)
'    Next
'
'    '��s�o��
'    FsoTS.Write (vbCrLf)
'
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
''    '���ڃ^�C�g���o��
''    FsoTS.Write ("����,�敪,9���@,,,,,,10���@,,,,,,11���@,,,,,,12���@,,,,,,13���@,,,,,,14���@,,,,,,15���@,,,,,,16���@" & vbCrLf)
''
''    '���ڃ^�C�g���o��
''    FsoTS.Write (",,")
''    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
''        FsoTS.Write ("�@,�A,�B,�C,�D,�E,")     '1�`7���@
''    Next
''    FsoTS.Write ("�@,�A,�B,�C,�D,�E" & vbCrLf) '8���@
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'    '���ڃ^�C�g���o��
'    'FsoTS.Write ("����,�敪,9���@,,,10���@,,,11���@,,,12���@,,,13���@,,,14���@,,,15���@,,,16���@" & vbCrLf)    'EG20�@V30.1.0.1 DEL
'    FsoTS.Write ("����,�敪,9���@,,,,,,10���@,,,,,,11���@,,,,,,12���@,,,,,,13���@,,,,,,14���@,,,,,,15���@,,,,,,16���@" & vbCrLf)    'EG20 V30.1.0.1 ADD
'
'    '���ڃ^�C�g���o��
'    FsoTS.Write (",,")
'    For i = 0 To ONE_PARAGRAPH_OUTPUT_GOKI - 2
'        'FsoTS.Write ("�@,�A,�B,")     '1�`7���@    'EG20 V30.1.0.1 DEL
'        FsoTS.Write ("�@,�A,�B,�C,�D,�E,")     '1�`7���@    'EG20 V30.1.0.1 ADD
'    Next
'    'FsoTS.Write ("�@,�A,�B" & vbCrLf) '8���@   'EG20 V30.1.0.1 DEL
'    FsoTS.Write ("�@,�A,�B,�C,�D,�E" & vbCrLf) '8���@   'EG20 V30.1.0.1 ADD
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'
'    '9���@�`16���@
'    For RowCount = 0 To ONE_PARAGRAPH_OUTPUT_ROW
'        FsoTS.Write (OutFileData2(RowCount).sShakyoku & ",")
'        FsoTS.Write (OutFileData2(RowCount).sKubun & ",")
'        FsoTS.Write (OutFileData2(RowCount).sSettei)
'    Next
'
'    '�t�@�C�����N���[�Y����B
'    FsoTS.Close
'
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�J�n
''    '�ꎞ�t�@�C����}�̂ɃR�s�[����
''    Call FileCopy(PATH_WORK & EKI_SETTI_SUBGATE_FILE, sWriteDir & EKI_SETTI_SUBGATE_FILE)
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�폜�I��
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
'    ' �R�[�i���̂̕t��
'    strSaveFileName = Trim(pfGetEkiNameInfo(NotEkiVer)) & "_" & szCornerName & "_" & EKI_SETTI_SUBGATE_FILE
'    '�ꎞ�t�@�C����}�̂ɃR�s�[����
'    Call FileCopy(PATH_WORK & EKI_SETTI_SUBGATE_FILE, sWriteDir & strSaveFileName)
'' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��
'EG20 V30.3.0.1 �yHKRK_Kansi07_003_01�z�A�yHKRK_Kansi07_008_01�zDEL END

' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V3.0.0.2[Mainte_03_01 �v���O���X�o�[�Ή�]�ǉ��I��

    '����I��
    iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�w�ݒ�e�L�X�g�o�͌���")
    
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

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
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

    If Left$(sWorkDrive, 1) = "H" Then
        'a:�h���C�u���ُ�Ȃ�A�J�����g�h���C�u��\��������B
        sWorkDrive = Left$(App.Path, 2)
        GoTo Retry
    End If
    
    '���̑��̃h���C�u�Ȃ�A�t�@�C���I���Ȃ��Ŗ߂�B
    pfFileSelection = ""

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveGateGamen_Click()
   
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    DoEvents
   
   '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_GAMEN_GO_BUTTOM, 0)
    Unload Me
    Load frmEkiDataGate
    frmEkiDataGate.Show 1
    
End Sub

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
    cmbGoki.Enabled = False
    CmdMenu(0).Enabled = False
    CmdMenu(1).Enabled = False
    CmdMenu(2).Enabled = False
    CmdMenu(3).Enabled = False
    CmdMoveGateGamen.Enabled = False
    CmdMoveEkiInfoGamen.Enabled = False
    cmdCancel.Enabled = False
    
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
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
'//     ORIGINAL  :(1.1.0.1) 2011-05-11   CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-10-28  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z�w�s�x�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()

    Dim strFileName         As String           '�t�@�C����

    '�����l�ݒ�
    strFileName = ""                            '�t�@�C����
    
    '�G���[���[�`����錾
    On Error Resume Next

    '�S�{�^���������Ƃ���B
    CmdMenu(0).Enabled = True
    CmdMenu(1).Enabled = True
    CmdMenu(2).Enabled = True
    CmdMenu(3).Enabled = True
    CmdMoveGateGamen.Enabled = True
    CmdMoveEkiInfoGamen.Enabled = True
    cmdCancel.Enabled = True

    '�R���{�{�b�N�X�͏����ɂ���Ă͌��X�����s�̂��߁A��ʕ\���p�t�@�C���̗L���Ŕ�����s��
    strFileName = Dir(EKI_TUDO_CHK_SUBGATE_FILE)
    '�t�@�C�������݂���ꍇ
    If strFileName <> "" Then
        cmbGoki.Enabled = True
    End If

' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��J�n
    CmbCornerName.Enabled = True
' EG20 V2.1.0.1[Mainte_03_01 �w�s�x�Ή�]�ǉ��I��

    DoEvents
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : CmdMoveGateGamen_Click
'//  �@�\����  : �G���R�[�h�R�[�i�ݒ��ʐؑ�
'//  �@�\�T�v  : �w�s�x�f�[�^�m�F�i�w���j��ʂ�\������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EGR HK1.1.0.1) 2011-05-11  CODED   BY [TCC] M.Kuroki
'//                 EG-R��}�@�V�K�J��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub CmdMoveEkiInfoGamen_Click()
    
    '�S�{�^���������s�Ƃ���B
    Call SetEnableFalse
    DoEvents
   
   '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIINFO_GAMEN_GO_BUTTOM, 0)

    '�\������ʃA�����[�h
    Unload Me
                
    Load frmEkiData
    frmEkiData.Show 1

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

    Dim bSysChange              As Boolean      '�V�X�e���ݒ菈���߂�l�@�fV1.8.0.1�@ADD
    Dim bUpData                 As Boolean      '��ʍX�V�����߂�l�@�@�@'V1.8.0.1�@ADD
    Dim iLoopCnt                As Integer      '���[�v�J�E���^

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
            
        '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���쐬
         bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
        If bRet = False Then
            '�w�s�x�f�[�^�m�F�i�G���R�[�h�R�[�i���@����`�j�C���[�W�t�@�C���폜
            Kill EKI_TUDO_CHK_SUBGATE_FILE
               
            '�ُ탍�O�o��
            Call pfOutPutErrLog(lErrCode)
            bUpData = False
            pfuncInstallEkiSettei = False
        End If

        '���@�R���{�{�b�N�X�����l
        cmbGoki.Clear
        'For iLoopCnt = 0 To 15 'EG20 V30.1.0.1 DEL
        For iLoopCnt = 0 To 31  'EG20 V30.1.0.1 ADD
                cmbGoki.AddItem iLoopCnt + 1 & "���@"
        Next
        cmbGoki.ListIndex = 0

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


