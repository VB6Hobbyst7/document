VERSION 5.00
Begin VB.Form frmVersion 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�o�[�W�����Ǘ�"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   14.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
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
      Caption         =   "�@�V�����@���D�@"
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
      Index           =   9
      Left            =   3320
      TabIndex        =   21
      Top             =   8040
      Width           =   1600
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
      Left            =   120
      TabIndex        =   18
      Top             =   7185
      Width           =   1600
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
      Left            =   4920
      TabIndex        =   17
      Top             =   7185
      Width           =   1600
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
      Left            =   1720
      TabIndex        =   16
      Top             =   7185
      Width           =   1600
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
      Left            =   3320
      TabIndex        =   15
      Top             =   7185
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "IC���ʉ^��"
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
      Left            =   120
      TabIndex        =   14
      Top             =   8040
      Width           =   1600
   End
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
      Index           =   6
      Left            =   1720
      TabIndex        =   13
      Top             =   8040
      Width           =   1600
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�@Ver�ꗗ�@USB�o��"
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
      Index           =   8
      Left            =   8120
      TabIndex        =   12
      Top             =   7185
      Width           =   1600
   End
   Begin VB.ListBox lstKan 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   11530
   End
   Begin VB.CommandButton cmdFixedExe 
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
      Height          =   855
      Index           =   7
      Left            =   8120
      TabIndex        =   10
      Top             =   8040
      Width           =   1600
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
      Left            =   6520
      TabIndex        =   9
      Top             =   7185
      Width           =   1600
   End
   Begin VB.Timer tmrMail 
      Left            =   11400
      Top             =   7320
   End
   Begin VB.Frame fraAllKansiVersion 
      Caption         =   "�S�̃o�[�W�����FZ9.Z9.Z9.Z9"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11535
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   5
         Left            =   8520
         TabIndex        =   8
         Top             =   650
         Width           =   2895
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   3
         Left            =   4500
         TabIndex        =   7
         Top             =   650
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�h�c�t�A�v���P�[�V�����F"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   6
         Top             =   345
         Width           =   3255
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9.Z9.Z9.Z9"
         Height          =   375
         Index           =   1
         Left            =   450
         TabIndex        =   5
         Top             =   650
         Width           =   2295
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�����Ď��ՁF "
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   350
         Width           =   2535
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�k�c�t�A�v���P�[�V�����F"
         Height          =   375
         Index           =   4
         Left            =   8355
         TabIndex        =   2
         Top             =   345
         Width           =   3015
      End
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
   Begin VB.Label lbltitle 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�^�C�g��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�t�@�C����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   19
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�o�[�W�����Ǘ�"
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
      TabIndex        =   4
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmVersion.frm
'//  �p�b�P�[�W���F�o�[�W�����Ǘ����
'//
'//  �T�v�F�o�[�W�����Ǘ����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �E�t�F�[�Y�Q�Ή��@�Ď��t�@�[���ARAS�}�C�R���ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͒ǉ�
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK�Ή�
'//     REVISIONS :(1.20.0.1) 2010-03-11   REVISED BY [TCC] S.Yamazaki
'//                 ��ʃ��C�A�E�g�ύX
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.18�C���Ή��z
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 �y�h�b�l�o�[�W�����t�@�C�����X�g�Ή��z
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 �y�o�[�W�����\���s���Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'//                 �y�k���V�����J�ƑΉ��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
Dim uVersion() As MN_VERSION_JIKAI      '�o�[�W�������i�[�G���A

'V1.20.0.1 ADD START
Private Type DISP_FILE_INFO      '�\���o�[�W����csv�t�@�C�����e
    sTitle As String             '�^�C�g��
    sFilePath As String          '�t�@�C���p�X
    iType As Integer             '�\���^�C�v
    iIdu As Integer              '�h�c�t�k�ޑΏۃt�@�C���L��
    iMaker As Integer            ' ���[�J�ԍ��i�^�C�v�Q�j             ' EG20 V5.6.0.1�ǉ�
End Type

Private Const CSV_COMMENT_CHAR = ":"  'csv�t�@�C���ŃR�����g�Ƃ��镶����
'V1.20.0.1 ADD END

' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
Dim FileList() As String                     '�t�@�C�������X�g�ꗗ�i�[�G���A
Dim FileListType() As String                 '�t�@�C�����X�g�ꗗ�i�[�G���A�i�����㎩���^�C�v���܂ށj
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����Ǘ����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
Private Sub Form_Activate()
   On Error Resume Next
    
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����Ǘ����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
Private Sub Form_Deactivate()
    On Error Resume Next
   
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����Ǘ����(���[�h��)
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
'//     REVISIONS :(1.4.0.1) 2009-03-18   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   Dim strWork         As String   '��ƃG���A
 
   On Error Resume Next
 
   '�u�o�[�W�����Ǘ���ʁF�\���v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_GAMEN_START, 0)

   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
   
   'V1.4.0.1 DEL START
   '�o�[�W�����擾����
   'psGetVersion
   'V1.4.0.1 DEL END
   
   '���C�^���Ή��`�F�b�N
'   psJikiCheck                                ' EG20 V2.1.0.1[Mainte_03_01] �폜
    
   'IDU�k�ރ`�F�b�N
   psIDUCheck
    
   If pbIDUSts = 1 Then
     'IDU�Ɩ���\��
      cmdFixedExe(1).Visible = False
'      cmdFixedExe(5).Visible = False          ' EG20 V2.1.0.1[Mainte_03_01] �폜
'      cmdFixedExe(6).Visible = False          ' EG20 V2.1.0.1[Mainte_03_01] �폜
      cmdFixedExe(4).Visible = False           ' EG20 V2.1.0.1[Mainte_03_01] �ǉ�
      cmdFixedExe(5).Visible = False           ' EG20 V2.1.0.1[Mainte_03_01] �ǉ�
   End If
   
   'V1.4.0.1 ADD START
   '�o�[�W�����擾����
   psGetVersion
   'V1.4.0.1 ADD END

   'V1.3.0.1 ADD START
   '���[����M�p�̃^�C�}�l��ݒ肷��B
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
   '1.3.0.1 ADD END
   'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
   'EG20 V30.1.0.1 ADD START
'    If fIsExistCornerType(CORNER_TYPE_ZAIRAI) = False Then
'        '�ݗ����R�[�i�[��������݂��Ȃ��̂ŁA���D�@�t�͉����s�ɂ���B
'        cmdFixedExe(3).Enabled = False
'    End If
'
'    If fIsExistCornerType(CORNER_TYPE_KANSEN) = False Then
'        '�����R�[�i�[��������݂��Ȃ��̂ŁA�V�������D�@�t�͉����s�ɂ���B
'        cmdFixedExe(9).Enabled = False
'    End If
   'EG20 V30.1.0.1 ADD END
   'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �e��ʑJ�ږt����������
'//  �@�\�T�v  : �t���̉�ʂɑJ�ڂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-18   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�Ď��t�@�[���A�o�[�W�����ؑւ�ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-18   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή��@�o�[�W�����}�̏o�͂�ǉ�
'//     REVISIONS :(1.20.0.1) 2010-03-17  REVISED BY [TCC] S.Yamazaki
'//                 �o�[�W�����ؑւ�}�̎�O�ɕύX
'//                 �o�[�W�����}�̏o�͖t��Ver�ꗗUSB�o�͖t�ɕύX
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.18�C���Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
   Dim udtMail As ML_DISP_INF          '��ʕ\���v��
   Dim iResponse As Integer            '���b�Z�[�W�{�b�N�X�߂�l
   Dim bRet As Boolean                 '���[�����M�����߂�l
   Dim lngErrCode As Long              '�G���[�R�[�h
    
    On Error Resume Next
    
    Select Case Index
        Case 0                                 '�o�[�W�����Ǘ��i�Ď��Ձj
             '�u�o�[�W�����Ǘ���ʁF�Ď��Ֆt�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_KANSIBAN_BUTTOM, 0)
            Load frmKVer
            frmKVer.Show 1
        Case 1                                 '�o�[�W�����Ǘ��iIDU�j
            '�u�o�[�W�����Ǘ���ʁFIDU�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_IDU_BUTTOM, 0)
            Load frmIDUVer
            frmIDUVer.Show 1
        Case 2                                 '�o�[�W�����Ǘ��iLDU�j
            '�u�o�[�W�����Ǘ���ʁFLDU�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_LDU_BUTTOM, 0)
            Load frmLduVer
            frmLduVer.Show 1
' EG20 V2.1.0.1[Mainte_03_01] �폜�J�n
'        Case 3                                 '�o�[�W�����Ǘ��iEG-R�����j
'            '�u�o�[�W�����Ǘ���ʁFEG-R�����t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EGRJIKAI_BUTTOM, 0)
'            gStrCurrentForm = sFormName_EJVer
'            Load frmJVer
'            frmJVer.Show 1
'        Case 4                                 '�o�[�W�����Ǘ��iNEG�����j
'            '�u�o�[�W�����Ǘ���ʁFNEG�����t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_NEGJIKAI_BUTTOM, 0)
'            gStrCurrentForm = sFormName_NJVer
'            Load frmJVer
'            frmJVer.Show 1
'        Case 5                                 '�o�[�W�����Ǘ��i����IC-M�j
' EG20 V2.1.0.1[Mainte_03_01] �폜�I��
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
        Case 3                                 '�o�[�W�����Ǘ��i���D�@�j
            '�u�o�[�W�����Ǘ���ʁFEG20�����t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EG20JIKAI_BUTTOM, 0)
            gStrCurrentForm = sFormName_EG20JVer
            Load frmGateVerKanri
            frmGateVerKanri.Show 1
        Case 4                                 '�o�[�W�����Ǘ��i����IC-M�j
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��
            '�u�o�[�W�����Ǘ���ʁF����IC-M�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_ICM_BUTTOM, 0)
            '��ʕ\���v��(����IC�|M���)��ID����ɑ��M����
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_HANTEI_VER
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
               '�u�o�[�W�����Ǘ���ʁF�ێ��ʕ\���v�����[�����M�ُ�v���O�o��
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
               '�N�����s�|�b�v�A�b�v�\��
' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�폜�J�n
'               iResponse = MsgBox("����IC-M�t�A��`�G���[�B" & _
'                                   Chr(vbKeyReturn) & _
'                                   "����IC-M�o�[�W�����Ǘ���ʂ��N���ł��܂���B", _
'                                   vbOKOnly, _
'                                   "��ʋN���G���[")
' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�폜�I��
' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�ǉ��J�n
               iResponse = MsgBox("�h�b�l�t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "�h�b�l�o�[�W�����Ǘ���ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�ǉ��I��
               Exit Sub
            End If
            '�u�o�[�W�����Ǘ���ʁF�ێ��ʕ\���v�����[�����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
        Case 5                                 '�o�[�W�����Ǘ��iIC���ʉ^���j
            '�u�o�[�W�����Ǘ���ʁFIC���ʉ^���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_ICUNCHIN_BUTTOM, 0)
            '�ێ��ʕ\���v��(IC���ʉ^�����)��ID����ɑ��M����
            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
            udtMail.udtlHeader.dwProid = RHOSHU_ID
            udtMail.udtlHeader.dwSubArea = 0
            udtMail.dwDisp_Type = ML_DT_PASMO_VER
            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
            If bRet = False Then
               '�u�o�[�W�����Ǘ���ʁF�ێ��ʕ\���v�����[�����M�ُ�v���O�o��
               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
               '�N�����s�|�b�v�A�b�v�\��
               iResponse = MsgBox("IC���ʉ^���t�A��`�G���[�B" & _
                                  Chr(vbKeyReturn) & _
                                  "IC���ʉ^���f�[�^�o�[�W�����Ǘ���ʂ��N���ł��܂���B", _
                                  vbOKOnly, _
                                  "��ʋN���G���[")
                Exit Sub
            End If
            '�u�o�[�W�����Ǘ���ʁF�ێ��ʕ\���v�����[�����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        Case 6                                 '�o�[�W�����Ǘ��i�����j
' EG20 V2.1.0.1[Mainte_03_01 �����o�[�W�����Ǘ����]�ǉ��J�n
            '�u�o�[�W�����Ǘ���ʁFEG20�����t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_TAKU_BUTTOM, 0)
            gStrCurrentForm = sFormName_EJVer
            Load frmSousaTakuVerKanri
            frmSousaTakuVerKanri.Show 1
' EG20 V2.1.0.1[Mainte_03_01 �����o�[�W�����Ǘ����]�ǉ��I��
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��
' EG20 V2.1.0.1[Mainte_03_01] �폜�J�n
'        Case 6                                 '�o�[�W�����Ǘ��iPASMO�^���j
'            '�u�o�[�W�����Ǘ���ʁFPASMO�^���t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_PASMO_BUTTOM, 0)
'            '�ێ��ʕ\���v��(PASMO�^�����)��ID����ɑ��M����
'            udtMail.udtlHeader.dwId = ML_ID_DISP_STS_CMD
'            udtMail.udtlHeader.dwSize = MlSize.DISP_STS_CMD
'            udtMail.udtlHeader.dwProid = RHOSHU_ID
'            udtMail.udtlHeader.dwSubArea = 0
'            udtMail.dwDisp_Type = ML_DT_PASMO_VER
'            bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
'            If bRet = False Then
'               '�u�o�[�W�����Ǘ���ʁF�ێ��ʕ\���v�����[�����M�ُ�v���O�o��
'               lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'               Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
'               '�N�����s�|�b�v�A�b�v�\��
'               iResponse = MsgBox("PASMO���ʉ^���t�A��`�G���[�B" & _
'                                  Chr(vbKeyReturn) & _
'                                  "PASMO���ʉ^���f�[�^�o�[�W�����Ǘ���ʂ��N���ł��܂���B", _
'                                  vbOKOnly, _
'                                  "��ʋN���G���[")
'                Exit Sub
'            End If
'            '�u�o�[�W�����Ǘ���ʁF�ێ��ʕ\���v�����[�����M����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
'        Case 7                                 '�o�[�W�����Ǘ�(���C�^��)
'            '�u�o�[�W�����Ǘ���ʁF���C�^���t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_JIKIUNCHIN_BUTTOM, 0)
'            Load frmJikiUnkaiFD
'            frmJikiUnkaiFD.Show 1
'        'V1.4.0.1 ADD START
'        Case 8                                 '�o�[�W�����Ǘ�(�Ď��t�@�[��)
'            '�u�o�[�W�����Ǘ���ʁF�Ď��t�@�[���t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_FIRMWARE_BUTTOM, 0)
'            Load frmFirmWareVer
'            frmFirmWareVer.Show 1
' EG20 V2.1.0.1[Mainte_03_01] �폜�I��
'        Case 9                                 '�o�[�W�����Ǘ�(�o�[�W�����ؑ�)         ' EG20 V1.1.1.1 �폜
        Case 7                                 '�}�̎�O                                ' EG20 V1.1.1.1 �ǉ�
            'V1.20.0.1 DEL START
'            '�u�o�[�W�����Ǘ���ʁF�o�[�W�����֖ؑt�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_CHANGEVER_BUTTOM, 0)
'            Load frmVerChang
'            frmVerChang.Show 1
            'V1.20.0.1 DEL END
            'V1.20.0.1 ADD START
            '�u�}�̎�O�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
            '�}�̎�O����
            Call pfRemove(Me)
            'V1.20.0.1 ADD END
        'V1.4.0.1 ADD END
        'V1.6.0.1�@ADD START
'        Case 10                                 '�o�[�W�����Ǘ�(�o�[�W�����}�̏o��)    ' EG20 V2.1.0.1[Mainte_03_01] �폜
        Case 8                                  '�o�[�W�����Ǘ�(�o�[�W�����}�̏o��)     ' EG20 V2.1.0.1[Mainte_03_01] �ǉ�
        'V1.20.0.1 DEL START
'           '�u�o�[�W�����Ǘ���ʁF�o�[�W�����}�̏o�͖t�����v���O�o��
'           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_VERSION_OUTPUT_BUTTOM, 0)
'           Load frmVerOutput
'           frmVerOutput.Show 1
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
           '�u�o�[�W�����Ǘ���ʁFVer�ꗗUSB�o�͖t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_VER_USB_OUTPUT_BUTTOM, 0)
           Call cmdVer_Output
        'V1.20.0.1 ADD END
        'V1.6.0.1 ADD END
        'V30.1.0.1 ADD START
        Case 9                                 '�o�[�W�����Ǘ��i�V�������D�@�j
            '�u�o�[�W�����Ǘ���ʁF�V���������t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EG30JIKAI_BUTTOM, 0)
            gStrCurrentForm = sFormName_EG30JVer
            Load frmKansenGateVerKanri
            frmKansenGateVerKanri.Show 1
        'V30.1.0.1 ADD END
    
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�����e�i���X��ʂ֖߂�v�t����������
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
Private Sub cmdReturn_Click()
    
    On Error Resume Next
    
    '�u�o�[�W�����Ǘ���ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psGetVersion
'//  �@�\����  : �o�[�W�����擾����
'//  �@�\�T�v  : �o�[�W�����擾�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �EIDU�k�ގ���\�������A�Ď��t�@�[���ARAS�}�C�R���o�[�W���������ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�q�`�r�}�C�R���\���s�v�̂��ߍ폜
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK�Ή�
'//     REVISIONS :(1.20.0.1) 2010-03-17  REVISED BY [TCC] S.Yamazaki
'//                 ���x���ւ̕\�����A���X�g�ւ̕\���ɕύX
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub psGetVersion()
  Dim sVersion  As String 'V1.4.0.1 ADD
  Dim sGetJikiVer As String 'V1.10.0.1 ADD
 
 '�Ď��ՁAEG-R�S�̃o�[�W�����擾
  psKansiGetVersion
 'IDU�S�̃o�[�W�����擾
 'psIDUGetVersion       'V1.4.0.1�@DEL
 
 'V1.4.0.1�@ADD�@START
 If pbIDUSts = 1 Then
    'IDU�o�[�W������\��
    lblVerName(2).Enabled = False
    lblVerName(3).Caption = ""
 Else
    '��k�ގ��͕\������
    psIDUGetVersion
 End If
 'V1.4.0.1�@ADD�@END
 
 'LDU�S�̃o�[�W�����擾
  psLDUVersion

'V1.4.0.1�@DEL�@START
' 'EG-R�����o�[�W�����擾
'  psEGRJVersion
' 'NEG�����o�[�W�����擾
'  psNEGJVersion
'V1.4.0.1�@DEL�@END

'V1.20.0.1 ADD START
Call psListVersion
'V1.20.0.1 ADD END

'V1.20.0.1 DEL START
''V1.4.0.1�@ADD START
' 'EG-R�����o�[�W�����擾
'  '����CPU
'  sVersion = psEGRJVersion(HANTEI_CPU)
'  lblVerName(11).Caption = sVersion
'  '���C��CPU
'  sVersion = psEGRJVersion(MAIN_CPU)
'  lblVerName(12).Caption = sVersion
' '�T�uCPU
'  sVersion = psEGRJVersion(SUB_CPU)
'  lblVerName(13).Caption = sVersion
' '���C��OS
'  sVersion = psEGRJVersion(MAIN_OS)
'  lblVerName(14).Caption = sVersion
' '�\���P
'  sVersion = psEGRJVersion(YOBI1)
'  lblVerName(15).Caption = sVersion
' '�\���Q
'  sVersion = psEGRJVersion(YOBI2)
'  lblVerName(16).Caption = sVersion
' '�o�[�W�����`�F�b�N
'  sVersion = psEGRJVersion(VER_CHK)
'  lblVerName(17).Caption = sVersion
'
' 'NEG�����o�[�W�����擾
'  sVersion = psNEGJVersion
'  lblVerName(20).Caption = sVersion
''V1.4.0.1�@ADD END
'
' 'IC-M�o�[�W�����擾
' 'psICMGetVersion     'V1.4.0.1�@DEL
' 'V1.4.0.1�@ADD�@START
' If pbIDUSts = 1 Then
'    'IDU�o�[�W������\��
'    lblVerName(21).Enabled = False
'    lblVerName(31).Caption = ""
' Else
'    '��k�ގ��͕\������
'    sVersion = psICMGetVersion
'    lblVerName(31).Caption = sVersion
' End If
' 'V1.4.0.1�@ADD�@END
'
' '���ʉ^���o�[�W�����擾
' 'psICUnchinGetVersion  'V1.4.0.1�@DEL
' 'V1.4.0.1�@ADD�@START
' If pbIDUSts = 1 Then
'    'IDU�o�[�W������\��
'    lblVerName(22).Enabled = False
'    lblVerName(33).Caption = ""
' Else
'    '��k�ގ��͕\������
'    sVersion = psICUnchinGetVersion
'    lblVerName(33).Caption = sVersion
' End If
'
' '�Ď��t�@�[���o�[�W�����\������
' sVersion = psKansiFirmVersion
' lblVerName(25).Caption = sVersion
'
''V1.10.0.1 ADD START
' '���C�^���ǂݍ���
' sGetJikiVer = psJikiUnchinVersion
' lblVerName(27).Caption = CStr(sGetJikiVer)
''V1.10.0.1 ADD END
'
' 'V1.6.0.1 DEL START
' ''RAS�}�C�R���o�[�W�����\������
' 'sVersion = psRASMICOMVersion
' 'lblVerName(27).Caption = sVersion
' 'V1.6.0.1 DEL END
'
' 'V1.4.0.1�@ADD�@END
'V1.20.0.1 DEL END
 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psKansiGetVersion
'//  �@�\����  : �Ď����u�S�́A�Ď��Ճo�[�W�����擾����
'//  �@�\�T�v  : KansiVersion.ini���o�[�W�������擾����B
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
Public Function psKansiGetVersion()
    Dim lSts As Long                                       '�֐��߂�l
    Dim strKansiVersion As String * VERSION_GATE_SIZE      '�Ď��ՑS�̃o�[�W����
    Dim strKansiVersion2 As String * VERSION_GATE_SIZE     '�Ď����u�S�̃o�[�W����
    
    On Error Resume Next
    
    strKansiVersion = ""
    strKansiVersion2 = ""

    ' KansiVersion.ini����Ď����u�̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSISYSTEMVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion, _
                                   Len(strKansiVersion), _
                                   KANSI_VERSION_FILE)
    If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
        fraAllKansiVersion.Caption = "�S�̃o�[�W�����F " & Left$(strKansiVersion, lSts) & ""
    Else
        '�o�[�W�����ԍ��̎擾�ُ�̏ꍇ�A�u--,--,--,--�v��\��
        fraAllKansiVersion.Caption = "�S�̃o�[�W�����F--.--.--.-- "
    End If
 
    ' KansiVersion.ini����Ď��Ղ̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion2, _
                                   Len(strKansiVersion2), _
                                   KANSI_VERSION_FILE)
     If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
        lblVerName(1).Caption = Left$(strKansiVersion2, lSts)
    Else
        '�o�[�W�����ԍ��̎擾�ُ�̏ꍇ�A�u--,--,--,--�v��\��
        lblVerName(1).Caption = "--.--.--.-- "
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psIDUGetVersion
'//  �@�\����  : ID���p���j�b�g�o�[�W�����擾����
'//  �@�\�T�v  : ID���p���j�b�g�o�[�W�����Ǘ��t�@�C�����A
'//              ID���p���j�b�g�̃o�[�W�������擾����B
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
Public Function psIDUGetVersion()
    Dim strWork     As String       '��ƃG���A
    Dim iFileNumber As Integer      '���g�p�t�@�C���ԍ�
    
    On Error Resume Next
    
    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
        
   'ID���p���j�b�g�o�[�W�����Ǘ��t�@�C�����I�[�v���B
    Open PATH_IDU_APP & PATH_IDU_VERKANRI For Input As #iFileNumber

    '���s�o�[�W�������擾����B
    Input #iFileNumber, strWork
    If (Trim(strWork) = "") Then
        '�o�[�W�����ԍ��擾�ُ�̏ꍇ
        lblVerName(3).Caption = "--.--.--.--"
    Else
       '�S�̃o�[�W����������쐬
        lblVerName(3).Caption = Trim(strWork)
    End If
      
   'ID���p���j�b�g�o�[�W�����Ǘ��t�@�C�����N���[�Y�B
    Close #iFileNumber
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psLDUVersion
'//  �@�\����  : LD���[�e�B���e�B�o�[�W�����擾����
'//  �@�\�T�v  : LD���[�e�B���e�B�o�[�W�����Ǘ��t�@�C�����A
'//              LD���[�e�B���e�B�̃o�[�W�������擾����B
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
Public Function psLDUVersion()
    Dim strWork     As String       '��ƃG���A
    Dim iFileNumber As Integer      '���g�p�t�@�C���ԍ�
    
    On Error Resume Next
    
    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
    
   'LD���[�e�B���e�B�o�[�W�����Ǘ��t�@�C�����I�[�v���B
    Open PATH_LDU_APP & PATH_LDU_VERKANRI For Input As #iFileNumber

    '���s�o�[�W�������擾����B
    Input #iFileNumber, strWork
    If (Trim(strWork) = "") Then
        '�o�[�W�����ԍ��擾�ُ�̏ꍇ
        lblVerName(5).Caption = "--.--.--.--"
    Else
       '�S�̃o�[�W����������쐬
        lblVerName(5).Caption = Trim(strWork)
    End If
      
   'LD���[�e�B���e�B�o�[�W�����Ǘ��t�@�C�����N���[�Y�B
    Close #iFileNumber

End Function

'V1.4.0.1 DEL START
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : psEGRJVersion
''//  �@�\����  : EG-R�������D�@�o�[�W�����擾����
''//  �@�\�T�v  : GATEVER_FILE.INI�t�@�C�����A��\�t�@�C�������擾���A
''//              ��\�t�@�C�����o�[�W�������擾����B
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Public Function psEGRJVersion()
'    Dim strWork     As String                         '��ƃG���A
'    Dim iFileNumber As Integer                        '���g�p�t�@�C���ԍ�
'    Dim lSts As Long                                  '�֐��߂�l
'    Dim strVerFileName As String * VERSION_GATE_SIZE  '�擾�t�@�C����
'    Dim sGetVer     As String                         '��ƃG���A
'    Dim lngErrCode As Long
'
'    On Error Resume Next
'
'    ' GATEVER_FILE.INI���画��f�[�^CPU-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_HANTEI_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EHAN1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 11
'    Else
'       '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(11).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI���烁�C��CPU-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_MAIN_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EPRO1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 12
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(12).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI����T�uCPU-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_SUB_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_ESCPUNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 13
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(13).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI���烁�C��OS-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_MAIN_OS, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EOSNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 14
'    Else
'       '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(14).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI����\��1�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_YOBI1, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'     If lSts > 0 Then
'       strWork = E_EYOBI1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'       psJVerGet strWork, 15
'     Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(15).Caption = "--"
'     End If
'
'    ' GATEVER_FILE.INI����\��2�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_E, _
'                                   GATE_YOBI2, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = E_EYOBI2NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 16
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(16).Caption = "--"
'    End If
'End Function
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : psNEGJVersion
''//  �@�\����  : NEG�������D�@�o�[�W�����擾����
''//  �@�\�T�v  : GATEVER_FILE.INI�t�@�C�����A��\�t�@�C�������擾���A
''//              ��\�t�@�C�����o�[�W�������擾����B
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Public Function psNEGJVersion()
'    Dim strWork     As String                         '��ƃG���A
'    Dim iFileNumber As Integer                        '���g�p�t�@�C���ԍ�
'    Dim lSts As Long                                  '�֐��߂�l
'    Dim strVerFileName As String * VERSION_GATE_SIZE  '�擾�t�@�C����
'    Dim sGetVer     As String                         '��ƃG���A
'    Dim lngErrCode As Long
'
'     On Error Resume Next
'
'    ' GATEVER_FILE.INI���画��f�[�^CPU-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_HANTEI_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NHAN1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 17
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(17).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI���烁�C��CPU-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_MAIN_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NPRO1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 18
'    Else
'       '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(18).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI����T�uCPU-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_SUB_PRO, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NSCPUNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 19
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(19).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI���烁�C��OS-PRO�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_MAIN_OS, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NOSNOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 20
'    Else
'    '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'     lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'     Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'     lblVerName(20).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI����\��1�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_YOBI1, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NYOBI1NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 21
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(21).Caption = "--"
'    End If
'
'    ' GATEVER_FILE.INI����\��2�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_N, _
'                                   GATE_YOBI2, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = N_NYOBI2NOW & "\\" & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psJVerGet strWork, 22
'    Else
'       '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(22).Caption = "--"
'    End If
'End Function

''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : psJVerGet
''//  �@�\����  : ��\�t�@�C���̃o�[�W�������擾
''//  �@�\�T�v  : ��\�t�@�C���̃o�[�W�������擾���A��ʕ\������B
''//
''//              �^        ����      �Ӗ�
''//  ����      : String�@�@sPath�@�@[IN]��\�t�@�C����
''//  �@�@      : Integer�@ iIndex�@ [IN]�\���C���f�b�N�X�ԍ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Function psJVerGet(sPath As String, iCnt As Integer)
'
'    Dim i As Integer                    '�J�E���^
'    Dim j As Integer                    '�J�E���^
'    Dim iFileNumber As Integer          '�t�@�C���ԍ�
'    Dim lLen As Long                    '�t�@�C���T�C�Y
'    Dim uFooter As MN_FOOT              '�t�b�^���i�[�G���A
'    Dim lPos As Long                    '�o�[�W�������i�[�ʒu
'    Dim sDateTime As String
'    Dim lngErrCode As Long              '�G���[�R�[�h
'
'On Error GoTo FileGetError
'
'    If Dir(sPath) <> "" Then            '�t�@�C�������݂���?
'
'      lLen = FileLen(sPath)             '�t�@�C���T�C�Y�̎擾
'
'      iFileNumber = FreeFile            '���g�p�̃t�@�C���ԍ����擾����
'
'      '�t�@�C���̃I�[�v��
'      Open sPath For Binary Access Read As #iFileNumber
'
'      '�t�b�^���̎擾
'      Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
'
'      '�o�[�W�����l��\��
'      lblVerName(iCnt).Caption = CStr(uFooter.sVersion)
'      Close #iFileNumber                  '�t�@�C������܂�
'    Else
'      '�t�@�C�������݂��Ȃ��B
'      lblVerName(iCnt).Caption = "--"
'    End If
' Exit Function
'
'FileGetError:
'   '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'   lblVerName(iCnt).Caption = "--"
'   Close #iFileNumber                  '�t�@�C������܂�
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : psICMGetVersion
''//  �@�\����  : IC-M�o�[�W�����擾����
''//  �@�\�T�v  : GATEVER_FILE.INI�t�@�C�����A��\�t�@�C�������擾���A
''//              ��\�t�@�C�����o�[�W�������擾����B
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Public Function psICMGetVersion()
'    Dim strWork     As String                         '��ƃG���A
'    Dim iFileNumber As Integer                        '���g�p�t�@�C���ԍ�
'    Dim lSts As Long                                  '�֐��߂�l
'    Dim strVerFileName As String * VERSION_GATE_SIZE  '�擾�t�@�C����
'    Dim sGetVer     As String                         '��ƃG���A
'    Dim lngErrCode As Long
'
'    On Error Resume Next
'    strWork = ""
'
'    ' GATEVER_FILE.INI���画��IC-M�f�[�^�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(GATE_TYPE_ICM, _
'                                   GATE_ICM, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_GATEVER_FILE)
'    If lSts > 0 Then
'    strWork = PATH_IDU_APP & PATH_IDU_IC_M & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psIDUVerGet strWork, 31
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(31).Caption = "--------------------"
'    End If
'
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : psICUnchinGetVersion
''//  �@�\����  : IC���ʉ^���o�[�W�����擾����
''//  �@�\�T�v  : kansi.ini�t�@�C�����A��\�t�@�C�������擾���A
''//              ��\�t�@�C�����o�[�W�������擾����B
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Public Function psICUnchinGetVersion()
'    Dim strWork     As String                        '��ƃG���A
'    Dim iFileNumber As Integer                       '���g�p�t�@�C���ԍ�
'    Dim lSts As Long                                 '�֐��߂�l
'    Dim strVerFileName As String * VERSION_GATE_SIZE '�擾�t�@�C����
'    Dim sGetVer     As String                        '��ƃG���A
'    Dim lngErrCode As Long
'
'    strWork = ""
'
'    ' �Ď��Րݒu�\��INI�t�@�C��(kansi.ini)���IC���ʉ^���f�[�^�̑�\�t�@�C�������擾����B
'    strVerFileName = ""
'    lSts = GetPrivateProfileString(IDU_KANSI_SECTION_NAME, _
'                                   IDU_KANSI_KEY_NAME, _
'                                   DEFAILT, _
'                                   strVerFileName, _
'                                   Len(strVerFileName), _
'                                   PATH_IDU_APP & IDU_KANSI_INI)
'    If lSts > 0 Then
'    strWork = PATH_IDU_APP & PATH_IDU_ICUNCHIN & Left$(strVerFileName, (InStr(strVerFileName, vbNullChar) - 1))
'    psIDUVerGet strWork, 33
'    Else
'      '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'      lblVerName(33).Caption = "--------------------"
'    End If
'
'End Function
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : psIDUVerGet
''//  �@�\����  : ��\�t�@�C���̃o�[�W�������擾
''//  �@�\�T�v  : ��\�t�@�C���̃o�[�W�������擾���A��ʕ\������B
''//
''//              �^        ����      �Ӗ�
''//  ����      : String�@�@sPath�@�@[IN]��\�t�@�C����
''//  �@�@      : Integer�@ iIndex�@ [IN]�\���C���f�b�N�X�ԍ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Function psIDUVerGet(sPath As String, iCnt As Integer)
'
'    Dim i As Integer                    '�J�E���^
'    Dim j As Integer                    '�J�E���^
'    Dim sMyName As String               '�t�@�C����
'    Dim iFileNumber As Integer          '�t�@�C���ԍ�
'    Dim lLen As Long                    '�t�@�C���T�C�Y
'    Dim uFooter As MN_IDU_FOOT          '�t�b�^���i�[�G���A
'    Dim lPos As Long                    '�o�[�W�������i�[�ʒu
'    Dim sDateTime As String
'    Dim lngErrCode As Long              '�G���[�R�[�h
'
'On Error GoTo FileGetError
'
'    If Dir(sPath) <> "" Then            '�t�@�C�������݂���?
'
'      lLen = FileLen(sPath)             '�t�@�C���T�C�Y�̎擾
'
'      iFileNumber = FreeFile            '���g�p�̃t�@�C���ԍ����擾����
'
'        '�t�@�C���̃I�[�v��
'        Open sPath For Binary Access Read As #iFileNumber
'        '�t�b�^���̎擾
'        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
'
'        '�f�[�^���{�o�[�W������\��
'        lblVerName(iCnt).Caption = CStr(uFooter.sDataName) & CStr(uFooter.sVersion)
'        Close #iFileNumber                  '�t�@�C������܂�
'    Else
'      '�t�@�C�������݂��Ȃ��ꍇ
'      lblVerName(iCnt).Caption = "--------------------"
'    End If
'
'    Exit Function
'FileGetError:
'   '�u�o�[�W�����Ǘ���ʁF�o�[�W�����擾�ُ�v���O�o��
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_GETVER_ERROR, lngErrCode)
'   lblVerName(iCnt).Caption = "--------------------"
'   Close #iFileNumber                  '�t�@�C������܂�
'End Function
'V1.4.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psJikiCheck
'//  �@�\����  : ���C�^���Ή����[�U�`�F�b�N����
'//  �@�\�T�v  : HOSHU.INI���A���C�^���Ή����[�U�ł��邩�ǂ����`�F�b�N����B
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
Public Sub psJikiCheck()
    Dim iFlag As Integer '�擾���[�U�t���O
 
    On Error Resume Next
 
  ' HOSHU.INI��莥�C�^���Ή����[�U�t���O���擾����B
    iFlag = GetPrivateProfileInt(KANS_JIKI, _
                                 KANSI_JIKI_FLAG, _
                                 DEFAILT_Int, _
                                 HOSHU_FILE)
     If iFlag = 0 Then
      '�t���O��0�̏ꍇ�u���C�^���v�t�͔�\��
      cmdFixedExe(7).Visible = False
     End If
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
        AppActivate frmVersion.Caption, False
        pfFormActive (frmVersion.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'V1.20.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdVer_Output
'//  �@�\����  : �uVer�ꗗ�@USB�o�́v�t����������
'//  �@�\�T�v  : Ver�ꗗ��USB�o�́B�o�[�W�����}�̏o��frm�̔}�̏o�͂Ɠ��ꏈ��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21 REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�z
'//                  �E�o�[�W�����ꗗ�o�̓t�@�C�����ύX
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-21  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.18�C���Ή��z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function cmdVer_Output() As Boolean
    
    Dim sWriteDir As String                 '�}�̏o�͐�
    Dim bRet      As Boolean         '�߂�l
    Dim lRetVal   As Long            '�e�L�X�g�\�������߂�l
    Dim sCommand  As String          '�R�}���h������
    Dim iResponse As Integer   'MsgBox�߂�l
    Dim lngErrCode As Long     '�G���[�R�[�h
    Dim fso         As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim strWriteDir As String               '�o�͐�t�H���_
' EG20 V2.0.1.1 ADD START
    Dim strStationName As String    ' �w��
    Dim strSrcName     As String    ' �R�s�[���t�@�C���p�X
' EG20 V2.0.1.1 ADD END
    
   On Error GoTo COPY_ERROR

    cmdVer_Output = False

    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    
    If sWriteDir <> "" Then

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       
       '�f�B���N�g�����w�肳���΁A�o�[�W�����t�@�C������o��
        bRet = dllEGRCreateVersionFile(PATH_IDU_APP, PATH_LDU_APP)
        If bRet = False Then

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
           '�v���O���X�o�[����������
           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
           
           '�u�t�@�C���쐬�ُ�v�|�b�v�A�b�v��ʕ\��
'           MsgBox "�t�@�C���̍쐬�Ɏ��s���܂����B", vbOKOnly + vbCritical, "�t�@�C���쐬�ُ�"              ' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�폜
           MsgBox "�ُ�I�����܂����B", vbCritical, "Ver�ꗗUSB�o��"                                        ' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�ǉ�
           '�u�t�@�C���쐬�ُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERSION_OUTPUT_FILE_CREATE_ERROR, 0)

           '�uVer�ꗗUSB�o�͏����ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, VERASION_KANRI_MENU_VER_USB_OUTPUT_ERROR, 0)
            Exit Function
        Else
          '�t�@�C���R�s�[
' EG20 V2.0.1.1 DEL START
'          FileCopy EGR_KANSI_VERSION_FILE_PATH, sWriteDir & EGR_KANSI_VERSION_FILE
' EG20 V2.0.1.1 DEL END
' EG20 V2.0.1.1 ADD START
          '�w���擾
          strStationName = gsGetStationEkiName
          ' �R�s�[���t�@�C���p�X
          strSrcName = PATH_HOSHU_DATA & EGR_KANSI_VERSION_FILE
          
          FileCopy strSrcName, sWriteDir & strStationName & "_" & EGR_KANSI_VERSION_FILE
' EG20 V2.0.1.1 ADD START
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
          '�v���O���X�o�[����������
          Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

          '�u�}�̏o�͐���I���v�|�b�v�A�b�v��ʕ\��
'          MsgBox "�}�̏o�͂͐���I�����܂����B", vbOKOnly + vbInformation, "�}�̏o�͌���"          ' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�폜
          MsgBox "����I�����܂����B", vbOKOnly + vbInformation, "Ver�ꗗUSB�o��"                   ' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�ǉ�

          '�uVer�ꗗUSB�o�͏�������v���O�o��
          Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_KANRI_MENU_VER_USB_OUTPUT_OK, 0)
        End If
     Else
         '�uVer�ꗗUSB�o�͏��������s�v���O�o��
          Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_VER_USB_OUTPUT_MISHORI, 0)
     End If
  cmdVer_Output = True
  
  Exit Function
COPY_ERROR:
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   
   '�����ُ�̏ꍇ�A�o�͌��ʃ|�b�v�A�b�v(�ُ�)�\��
'    MsgBox "�}�̏o�ُ͈͂�I�����܂����B", vbCritical, "�}�̏o�͌���"                              ' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�폜
    MsgBox "�ُ�I�����܂����B", vbCritical, "Ver�ꗗUSB�o��"                                       ' EG20 V3.6.0.1�y03����TR-No.18�C���Ή��z�ǉ�
   
   '�uVer�ꗗUSB�o�͏����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_KANRI_MENU_VER_USB_OUTPUT_ERROR, lngErrCode)
   cmdVer_Output = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : psListVersion
'//  �@�\����  : ���X�g�\��
'//  �@�\�T�v  : �\���p�o�[�W�����t�@�C����ǂݍ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-16   CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 �y�h�b�l�o�[�W�����t�@�C�����X�g�Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-08  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion()
    
    Dim iCnt As Integer
    Dim iMax As Integer
    Dim structDispInfo() As DISP_FILE_INFO
    Dim fsoObj As New FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim iFileNumber As Integer
    
    Dim sLine As String             '1�s����csv�ǂݎ��f�[�^
    Dim sLineSplit() As String      '1�P�ꂸ��csv�ǂݎ��f�[�^
    
    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
    Dim sErrExe As String              '�G���[���O�pINI�g���q
    Dim lngErrCode As Long             '�G���[�R�[�h
    Dim sErrEventName As String        '�G���[���N�����C�x���g��
    
    '�G���[�g���b�v
    On Error GoTo Err_FILE
    
    '���X�g�{�b�N�X������
    lstKan.Clear
    
    '���݃`�F�b�N
    If fsoObj.FileExists(DISP_VERFILE_FILE) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        GoTo Err_FILE
    End If
    
    '���g�p�̃t�@�C���ԍ��擾
    iFileNumber = FreeFile
    
    '�t�@�C�����I�[�v������B
    sErrEventName = LOG_ERR_FILE_OPEN       '�t�@�C���I�[�v���ُ�
    Open DISP_VERFILE_FILE For Input As #iFileNumber
    
    iCnt = 0
    
    sErrEventName = LOG_ERR_FILE_READ       '�t�@�C���Ǎ��ُ�
    Do While Not EOF(iFileNumber)
        
        '�P �s�Âϐ��ǂݍ���
        Line Input #iFileNumber, sLine
        
        '�R�����g�s�Ƌ�s����Ȃ���Τ�̈�Ɋi�[
        If Trim(Left(sLine, 1)) <> CSV_COMMENT_CHAR And sLine <> "" Then
            
            sLineSplit = Split(sLine, ",")
            
'            If UBound(sLineSplit) = 3 Then                           ' EG20 V5.6.0.1�폜
            If UBound(sLineSplit) = 4 Then                            ' EG20 V5.6.0.1�ǉ�
            
                ReDim Preserve structDispInfo(iCnt)
                
                structDispInfo(iCnt).sTitle = sLineSplit(0)
                structDispInfo(iCnt).sFilePath = sLineSplit(1)
                structDispInfo(iCnt).iType = sLineSplit(2)
                structDispInfo(iCnt).iIdu = sLineSplit(3)
                structDispInfo(iCnt).iMaker = sLineSplit(4)           ' EG20 V5.6.0.1�ǉ�
                
                iCnt = iCnt + 1
                
            End If
            
        End If
    Loop
    
    '�t�@�C�����N���[�Y����B
    sErrEventName = LOG_ERR_FILE_CLOSE      '�t�@�C���N���[�Y�ُ�
    Close #iFileNumber
    
    iMax = iCnt - 1
    
    '�\����\�t�@�C���̃G���[�g���b�v�i�G���[�������Ă������͑����j
    On Error Resume Next
    
    'IDU�k�ރ`�F�b�N
    Call psIDUCheck
    
    For iCnt = 0 To iMax
        
        '�k�ދ@�\�t���O���Ȃ��A�܂��͏k�ޒ��ł͂Ȃ��Ƃ��̂ݏ������s��
        If structDispInfo(iCnt).iIdu = 0 Or pbIDUSts = 0 Then
        
            Select Case structDispInfo(iCnt).iType
                Case 1
                    Call psListVersion_Type1(structDispInfo(iCnt))
                Case 2
                    Call psListVersion_Type2(structDispInfo(iCnt))
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
                Case 3
                    Call psListVersion_Type3(structDispInfo(iCnt))
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
'EG20 V30.1.0.1 ADD START
                Case 4
                    Call psListVersion_Type4(structDispInfo(iCnt))
'EG20 V30.1.0.1 ADD END
                Case Else
                    '�����Ȃ�
            End Select
        End If
    Next
    
    Set fsoObj = Nothing

    Exit Sub

Err_FILE:

    '�ُ탍�O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    '���O�o�́@���t�@�C����
    Call psFileNameGet(DISP_VERFILE_FILE, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    '�t�@�C���N���[�Y
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    
    Set fsoObj = Nothing

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : psListVersion_Type1
'//  �@�\����  : ���X�g�\��
'//  �@�\�T�v  : �\���^�C�v�P�̕\�����s���i�Ď��ՁARYT�j
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-16   CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 �y�o�[�W�����\���s���Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type1(structDispInfo As DISP_FILE_INFO)
    
    Dim lLen As Long
    Dim sDateTime As String
    Dim j As Integer
    Dim uFooter As MN_FOOT              '�t�b�^���i�[�G���A
    Dim sTitle As String                '���󔒖��߂����^�C�g��
    Dim sDisp As String                 '�\���p
    Dim fsoObj As New FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim iFileNumber As Integer

    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
    Dim sErrExe As String              '�G���[���O�pINI�g���q
    Dim lngErrCode As Long             '�G���[�R�[�h
    Dim sErrEventName As String        '�G���[���N�����C�x���g��

    Dim bRet As Boolean                 ' �߂�l
    Dim szFileName As String            ' �t�@�C����
    Dim uVersion As MN_VERSION_JIKAI    ' �o�[�W�������i�[�G���A

    '�G���[�g���b�v
    On Error GoTo Err_FILE

    '�^�C�g���̉��H
    sTitle = structDispInfo.sTitle
    '�^�C�g����̃X�y�[�X�i�S�p�̉\��������̂�Format�͎g���Ȃ��j
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If

    '�t�@�C���̑��݃`�F�b�N�B�ُ펞��----�̕\��
    If fsoObj.FolderExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If

    ' �t�@�C�����X�g����t�@�C�����X�g�̍쐬
    bRet = fReadFileList(structDispInfo.sFilePath & "\" & MN_FILELIST)
    If bRet <> True Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If

    szFileName = structDispInfo.sFilePath & "\" & FileList(0)   ' �t�@�C�����X�g����o�[�W���������擾����
    If fsoObj.FileExists(szFileName) = True Then                ' �t�@�C�������݂���?
        lLen = FileLen(szFileName)                              ' �t�@�C���T�C�Y�̎擾

        iFileNumber = FreeFile                                  ' ���g�p�̃t�@�C���ԍ����擾����

        Open szFileName For Binary Access Read As #iFileNumber  ' �t�@�C���̃I�[�v��
        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter      ' �t�b�^���̎擾
        uVersion.sFileName = UCase(FileListType(0))             ' �t�@�C������啶���ɂ��ăZ�b�g
        uVersion.sMachineName = uFooter.sKisyu                  ' �@�햼�Z�b�g
        uVersion.sFooterFile = uFooter.sFileName                ' �t�@�C�����Z�b�g

        sDateTime = ""
        For j = 0 To 3
            sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
        Next
        sDateTime = sDateTime & " "
        For j = 4 To 5
            sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
        Next
        uVersion.sFileDate = sDateTime
        uVersion.sVersion = uFooter.sVersion                    ' �o�[�W�������Z�b�g
        uVersion.sComment = uFooter.sHyoji                      ' �\��������Z�b�g

        Close #iFileNumber                  '�t�@�C������܂�
    End If
    
    '�o�[�W�������i�[�G���A�̊g��
    sDisp = sTitle                                                                  ' �^�C�g��
'    sDisp = sDisp & Format(Right(FileList(0), 12), "!@@@@@@@@@@@@") & Space(11)     ' �t�@�C����   ' EG20 V6.1.0.1�폜
    sDisp = sDisp & Format(Right(FileList(0), 12), "!@@@@@@@@@@@@") & Space(7)      ' �t�@�C����    ' EG20 V6.1.0.1�ǉ�
    sDisp = sDisp & uFooter.sKisyu & Space(1)                                       ' �@�햼
    sDisp = sDisp & uFooter.sFileName & Space(2)                                    ' �t�@�C��
    sDisp = sDisp & sDateTime & Space(1)                                            ' �쐬����
    sDisp = sDisp & uFooter.sVersion                                                ' �o�[�W����
    lstKan.AddItem (sDisp)

    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 32), vbUnicode)     ' �R�����g1
    lstKan.AddItem (Space(45) & sDisp)

    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 33, 64), vbUnicode)    ' �R�����g2
    lstKan.AddItem (Space(45) & sDisp)

    Set fsoObj = Nothing
    Exit Sub
    
Err_FILE:

    '�ُ탍�O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    '���O�o�́@���t�@�C����
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)

    '�t�@�C���N���[�Y
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If

    '�ُ�p�\��
' EG20 V6.1.0.1�폜�J�n
'    lstKan.AddItem (sTitle & "------------           -------- --------  -------- ---- --")
' EG20 V6.1.0.1�폜�I��
' EG20 V6.1.0.1�ǉ��J�n
    lstKan.AddItem (sTitle & "------------       -------- --------  -------- ---- --")
' EG20 V6.1.0.1�ǉ��I��
    lstKan.AddItem (Space(45) & "--------------------------------")
    lstKan.AddItem (Space(45) & "--------------------------------")

    Set fsoObj = Nothing

End Sub

' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n�i�\�����e�ύX�j
'Private Sub psListVersion_Type1(structDispInfo As DISP_FILE_INFO)
'
'    Dim lLen As Long
'    Dim sDateTime As String
'    Dim j As Integer
'    Dim uFooter As MN_FOOT              '�t�b�^���i�[�G���A
'    Dim sTitle As String                '���󔒖��߂����^�C�g��
'    Dim sDisp As String                 '�\���p
'    Dim sDispFile As String             '��Ɨp
'    Dim sDispExe As String              '�g���q
'    Dim fsoObj As New FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
'    Dim iFileNumber As Integer
'
'    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
'    Dim sErrExe As String              '�G���[���O�pINI�g���q
'    Dim lngErrCode As Long             '�G���[�R�[�h
'    Dim sErrEventName As String        '�G���[���N�����C�x���g��
'
'    '�G���[�g���b�v
'    On Error GoTo Err_FILE
'
'    '�^�C�g���̉��H
'    sTitle = structDispInfo.sTitle
'    '�^�C�g����̃X�y�[�X�i�S�p�̉\��������̂�Format�͎g���Ȃ��j
'    If LenB(StrConv(sTitle, vbFromUnicode)) < 20 Then
'        sTitle = sTitle & Space(20 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
'    Else
'        sTitle = sTitle & Space(2)
'    End If
'
'    '�t�@�C���̑��݃`�F�b�N�B�ُ펞��----�̕\��
'    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
'        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
'        '�ُ�
'        GoTo Err_FILE
'    End If
'
'    lLen = FileLen(structDispInfo.sFilePath)              '�t�@�C���T�C�Y�̎擾
'    If lLen < Len(uFooter) Then
'        sErrEventName = LOG_ERR_FILE_LENGTH     '�t�@�C�������O�X�ُ�
'        '�ُ�
'        GoTo Err_FILE
'    End If
'
'    '���g�p�̃t�@�C���ԍ��擾
'    iFileNumber = FreeFile
'
'    '�t�@�C���̃I�[�v��
'    sErrEventName = LOG_ERR_FILE_OPEN       '�t�@�C���I�[�v���ُ�
'    Open structDispInfo.sFilePath For Binary Access Read As #iFileNumber
'
'        sErrEventName = LOG_ERR_FILE_READ       '�t�@�C���Ǎ��ُ�
'        '�t�b�^���̎擾
'        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
'
'    sErrEventName = LOG_ERR_FILE_CLOSE      '�t�@�C���N���[�Y�ُ�
'    Close #iFileNumber      '�t�@�C���̃N���[�Y
'
'    '�쐬�����̉��H
'    sDateTime = ""
'    For j = 0 To 3
'        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
'    Next
'    sDateTime = sDateTime & " "
'    For j = 4 To 5
'        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
'    Next
'
'    '�t�@�C�����̉��H
'    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             '�t�@�C���p�X����t�@�C�������擾
'    sDispFile = UCase(sDispFile & "." & sDispExe)                                 '�g���q���������啶���ɕϊ�
'
'    '�o�[�W�������i�[�G���A�̊g��
'    sDisp = sTitle                                                                  '�^�C�g��
'    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(11)       '�t�@�C����
'    sDisp = sDisp & uFooter.sKisyu & Space(1)                                       '�@�햼
'    sDisp = sDisp & uFooter.sFileName & Space(2)                                    '�t�@�C��
'    sDisp = sDisp & sDateTime & Space(1)                                            '�쐬����
'    sDisp = sDisp & uFooter.sVersion                                                '�o�[�W����
'    lstKan.AddItem (sDisp)
'
'    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 32), vbUnicode)     '�R�����g1
'    lstKan.AddItem (Space(45) & sDisp)
'
'    sDisp = StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 33, 64), vbUnicode)    '�R�����g2
'    lstKan.AddItem (Space(45) & sDisp)
'
'    Set fsoObj = Nothing
'
'    Exit Sub
'Err_FILE:
'
'    '�ُ탍�O�o��
'    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
'    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
'    '���O�o�́@���t�@�C����
'    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
'    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)
'
'    '�t�@�C���N���[�Y
'    If iFileNumber > 0 Then
'        Close #iFileNumber
'    End If
'
'    '�ُ�p�\��
'    lstKan.AddItem (sTitle & "------------           -------- --------  -------- ---- --")
'    lstKan.AddItem (Space(45) & "--------------------------------")
'    lstKan.AddItem (Space(45) & "--------------------------------")
'
'    Set fsoObj = Nothing
'
'End Sub
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2010 All Rights Reserved
'//
'//  �֐�����  : psListVersion_Type2
'//  �@�\����  : ���X�g�\��
'//  �@�\�T�v  : �\���^�C�v�Q�̕\�����s���iIDU�j
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-16   CODED   BY [TCC] S.Yamazaki
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 �y�h�b�l�o�[�W�����t�@�C�����X�g�Ή��z
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 �y�o�[�W�����\���s���Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type2(structDispInfo As DISP_FILE_INFO)

    Dim lLen As Long
    Dim sDateTime As String
    Dim j As Integer
    Dim uFooter As MN_IDU_FOOT          '�t�b�^���i�[�G���A
    Dim sTitle As String                '���󔒖��߂����^�C�g��
    Dim sDisp As String                 '�\���p
    Dim sDispFile As String             '��Ɨp�i�t�@�C���\���p�j
    Dim sDispDV As String               '��Ɨp�i�f�[�^�{�o�[�W�����j
    Dim sDispCom As String              '��Ɨp�i�R�����g�j
    Dim sDispExe As String              '�g���q
    Dim fsoObj As New FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim iFileNumber As Integer
    
    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
    Dim sErrExe As String              '�G���[���O�pINI�g���q
    Dim lngErrCode As Long             '�G���[�R�[�h
    Dim sErrEventName As String        '�G���[���N�����C�x���g��

' EG20 V5.6.0.1 �ǉ��J�n
    Dim bRet As Boolean                 ' �߂�l
    Dim szFilelist As String            ' �t�@�C�����X�g��
    Dim szFileName As String            ' �t�@�C����
' EG20 V5.6.0.1 �ǉ��I��
    
    '�G���[�g���b�v
    On Error GoTo Err_FILE
    
    '�p�X����蒼��
    structDispInfo.sFilePath = PATH_IDU_APP & "\" & structDispInfo.sFilePath

    '�^�C�g���̉��H
    sTitle = structDispInfo.sTitle
    '�^�C�g����̃X�y�[�X�i�S�p�̉\��������̂�Format�͎g���Ȃ��j
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If

' EG20 V5.6.0.1 �ǉ��J�n

    ' �t�@�C�����X�g����t�@�C�����X�g�̍쐬
    szFilelist = structDispInfo.sFilePath & "\FILELIST_" & Format(structDispInfo.iMaker, "00") & ".TXT"

    ' �t�@�C���̑��݃`�F�b�N�B�ُ펞��----�̕\��
    If fsoObj.FileExists(szFilelist) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If

    bRet = fReadFileListIDU(szFilelist)
    If bRet <> True Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If
    structDispInfo.sFilePath = structDispInfo.sFilePath & "\" & FileList(0)   ' �t�@�C�����X�g����o�[�W���������擾����

' EG20 V5.6.0.1 �ǉ��I��
    
    '--------------------------------------------
    '�t�@�C�����擾
    '--------------------------------------------
    '�t�@�C���̑��݃`�F�b�N�B�ُ펞��----�̕\��
    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If
    
    lLen = FileLen(structDispInfo.sFilePath)              '�t�@�C���T�C�Y�̎擾
    If lLen < Len(uFooter) Then
        sErrEventName = LOG_ERR_FILE_LENGTH     '�t�@�C�������O�X�ُ�
        '�ُ�
        GoTo Err_FILE
    End If
    
    '���g�p�̃t�@�C���ԍ��擾
    iFileNumber = FreeFile
    
    '�t�@�C���̃I�[�v��
    sErrEventName = LOG_ERR_FILE_OPEN       '�t�@�C���I�[�v���ُ�
    Open structDispInfo.sFilePath For Binary Access Read As #iFileNumber
    
        sErrEventName = LOG_ERR_FILE_READ       '�t�@�C���Ǎ��ُ�
        '�t�b�^���̎擾
        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
        
    sErrEventName = LOG_ERR_FILE_CLOSE      '�t�@�C���N���[�Y�ُ�
    Close #iFileNumber      '�t�@�C���̃N���[�Y
    
    '--------------------------------------------
    '�o�[�W�������\�����̕\���e�L�X�g�쐬
    '--------------------------------------------
    '�^�C�g��
    sDisp = sTitle
    
    '�t�@�C����
    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             '�t�@�C���p�X����t�@�C�������擾
    sDispFile = sDispFile & "." & sDispExe                                        '�g���q���������啶���ɕϊ�
'    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(11)      ' EG20 V6.1.0.1�폜
    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(7)        ' EG20 V6.1.0.1�ǉ�

    '���
    sDisp = sDisp & LCase(Right$("0" & Hex(uFooter.bSyubetu), 2))
    
    '���[�J��
    sDisp = sDisp & uFooter.sMakerName & Space(2)
    
    '�f�[�^���{�o�[�W����
    sDispDV = LTrim(uFooter.sDataName) & uFooter.sVersion
    If Len(Trim(uFooter.sDataName)) = 0 And Len(Trim(uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(Trim(uFooter.sVersion) & Space(20), 20) & Space(2)
    ElseIf Len(Trim(uFooter.sDataName & uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(sDispDV & Space(20), 20) & Space(2)
    Else
        sDisp = sDisp & String(20, "-") & Space(2)
    End If
    
    '�쐬����
    sDateTime = ""
    For j = 0 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    sDateTime = Format(sDateTime, "@@@@/@@/@@ @@:@@")
    sDisp = sDisp & sDateTime
    
    '���X�g�ɒǉ�
    lstKan.AddItem (sDisp)
    
    '�R�����g
    '60�����ŕۑ�����Ă���̂ŁA60�o�C�g�̗̈�ɒ����B�O��̋󔒂����B
    sDispCom = Trim(StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 60), vbUnicode))
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 1, 32), vbUnicode)     '�R�����g1
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    Else
        lstKan.AddItem (Space(45) & String(32, "-"))
    End If
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 33, 60), vbUnicode)    '�R�����g2
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    End If
    
    Set fsoObj = Nothing
    
    Exit Sub
Err_FILE:
    
    '�ُ탍�O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    '���O�o�́@���t�@�C����
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    '�t�@�C���N���[�Y
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    
    '�ُ�p�\��
' EG20 V6.1.0.1�폜�J�n
'    lstKan.AddItem (sTitle & "------------           ---  --------------------  ----/--/-- --:--")
' EG20 V6.1.0.1�폜�I��
' EG20 V6.1.0.1�ǉ��J�n
    lstKan.AddItem (sTitle & "------------       ---  --------------------  ----/--/-- --:--")
' EG20 V6.1.0.1�ǉ��I��
    lstKan.AddItem (Space(45) & "--------------------------------")
    lstKan.AddItem (Space(45) & "--------------------------------")

    Set fsoObj = Nothing

End Sub
'V1.20.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : psListVersion_Type3
'//  �@�\����  : ���X�g�\��
'//  �@�\�T�v  : �\���^�C�v�R�̕\�����s���i�����j
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V5.0.2.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.56�C���Ή��z
'//     REVISIONS :(EG20 V6.1.0.1) 2012-06-09  CODED BY  [TCC] H.Sugimoto
'//                 �y�o�[�W�����\���s���Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type3(structDispInfo As DISP_FILE_INFO)

'    Dim lLen As Long                                                 ' EG20 V5.0.2.1�폜
'    Dim sDateTime As String                                          ' EG20 V5.0.2.1�폜
'    Dim j As Integer                                                 ' EG20 V5.0.2.1�폜
'    Dim uFooter As MN_IDU_FOOT          '�t�b�^���i�[�G���A        ' EG20 V5.0.2.1�폜
    Dim sTitle As String                '���󔒖��߂����^�C�g��
    Dim sDisp As String                 '�\���p
    Dim sDispFile As String             '��Ɨp�i�t�@�C���\���p�j
'    Dim sDispCom As String              '��Ɨp�i�R�����g�j          ' EG20 V5.0.2.1�폜
    Dim sDispExe As String              '�g���q
    Dim fsoObj As New FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
'    Dim iFileNumber As Integer                                       ' EG20 V5.0.2.1�폜
    
    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
    Dim sErrExe As String              '�G���[���O�pINI�g���q
    Dim lngErrCode As Long             '�G���[�R�[�h
    Dim sErrEventName As String        '�G���[���N�����C�x���g��
    
'    Dim FsoRead As TextStream                                        ' EG20 V5.0.2.1�폜
'    Dim bFileOpen As Boolean            ' �I�[�v���t���O             ' EG20 V5.0.2.1�폜
'    Dim strBuffer As String             ' ���[�h�o�b�t�@             ' EG20 V5.0.2.1�폜
    Dim strVersion As String            ' �o�[�W����������

' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�ǉ��J�n
    Dim lSts As Long                                       '�֐��߂�l
    Dim strKansiVersion2 As String * VERSION_GATE_SIZE     '�Ď����u�S�̃o�[�W����
' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�ǉ��I��

    
    '�G���[�g���b�v
    On Error GoTo Err_FILE
    
    ' ������
'    bFileOpen = False                                                ' EG20 V5.0.2.1�폜
    
    
    '�^�C�g���̉��H
    sTitle = structDispInfo.sTitle
    '�^�C�g����̃X�y�[�X�i�S�p�̉\��������̂�Format�͎g���Ȃ��j
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If
    
    '--------------------------------------------
    '�t�@�C�����擾
    '--------------------------------------------
    '�t�@�C���̑��݃`�F�b�N�B�ُ펞��----�̕\��
    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If
    
' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�폜�J�n
'    Set FsoRead = fsoObj.OpenTextFile(structDispInfo.sFilePath, ForReading)
'    bFileOpen = True
'    ' �t�@�C������P�s���[�h
'    strBuffer = FsoRead.ReadLine
'    strVersion = Trim(strBuffer)
'
'    FsoRead.Close
'    Set FsoRead = Nothing
'    Set fsoObj = Nothing
' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�폜�I��

' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�ǉ��J�n
    Set fsoObj = Nothing
    strKansiVersion2 = ""
    strVersion = ""
    ' KansiVersion.ini���瑀���̑S�̃o�[�W�������擾���\������
    lSts = GetPrivateProfileString(KANSIVERSION_SECTION_NAME, _
                                   KANSIVERSION_KEY_NAME, _
                                   DEFAILT, _
                                   strKansiVersion2, _
                                   Len(strKansiVersion2), _
                                   structDispInfo.sFilePath)
     If lSts > 0 Then
        '�擾�����o�[�W�����ԍ���\��
        strVersion = Left$(strKansiVersion2, lSts)
    End If

' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�ǉ��I��

    '--------------------------------------------
    '�o�[�W�������\�����̕\���e�L�X�g�쐬
    '--------------------------------------------
    '�^�C�g��
    sDisp = sTitle
    
    '�t�@�C����
    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             '�t�@�C���p�X����t�@�C�������擾
    
    sDispFile = sDispFile & "." & sDispExe                                        '�g���q���������啶���ɕϊ�
    
' EG20 V5.0.2.1�y�t�@�C�����͕\�����Ȃ��z�폜�J�n
'    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(11)
' EG20 V5.0.2.1�y�t�@�C�����͕\�����Ȃ��z�폜�I��
' EG20 V5.0.2.1�y�t�@�C�����͕\�����Ȃ��z�ǉ��J�n
'    sDisp = sDisp & Space(12) & Space(11)                                       ' EG20 V6.1.0.1�폜
    sDisp = sDisp & Space(12) & Space(7)                                         ' EG20 V6.1.0.1�ǉ�
' EG20 V5.0.2.1�y�t�@�C�����͕\�����Ȃ��z�ǉ��I��

    '�o�[�W����
    If Len(strVersion) <> 0 Then
        sDisp = sDisp & Format(Left(strVersion, 11), "!@@@@@@@@@@@")
    Else
        sDisp = sDisp & "--.--.--.--"
    End If
    
    '���X�g�ɒǉ�
    lstKan.AddItem (sDisp)
    
    Exit Sub
Err_FILE:
    
    '�ُ탍�O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    '���O�o�́@���t�@�C����
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    '�ُ�p�\��
'    lstKan.AddItem (sTitle & "                       --.--.--.--")             ' EG20 V6.1.0.1�폜
    lstKan.AddItem (sTitle & "                   --.--.--.--")                  ' EG20 V6.1.0.1�ǉ�
' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�폜�J�n
'    If bFileOpen = True Then
'        FsoRead.Close
'    End If
'    Set FsoRead = Nothing
' EG20 V5.0.2.1�y����TR-No.56�C���Ή��z�폜�I��
    Set fsoObj = Nothing

End Sub
'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : psListVersion_Type4
'//  �@�\����  : ���X�g�\��
'//  �@�\�T�v  : �\���^�C�v�S�̕\�����s���i���������j
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-05-08   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psListVersion_Type4(structDispInfo As DISP_FILE_INFO)

    Dim lLen As Long
    Dim sDateTime As String
    Dim j As Integer
    Dim uFooter As MN_IDU_FOOT          '�t�b�^���i�[�G���A(�V�������D�@��IDU�̃t�b�^�t�H�[�}�b�g�Ɠ����Ȃ̂Łj
    Dim sTitle As String                '���󔒖��߂����^�C�g��
    Dim sDisp As String                 '�\���p
    Dim sDispFile As String             '��Ɨp�i�t�@�C���\���p�j
    Dim sDispDV As String               '��Ɨp�i�f�[�^�{�o�[�W�����j
    Dim sDispCom As String              '��Ɨp�i�R�����g�j
    Dim sDispExe As String              '�g���q
    Dim fsoObj As New FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim iFileNumber As Integer
    
    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
    Dim sErrExe As String              '�G���[���O�pINI�g���q
    Dim lngErrCode As Long             '�G���[�R�[�h
    Dim sErrEventName As String        '�G���[���N�����C�x���g��

    Dim bRet As Boolean                 ' �߂�l
    
    '�G���[�g���b�v
    On Error GoTo Err_FILE
    
    '�^�C�g���̉��H
    sTitle = structDispInfo.sTitle
    '�^�C�g����̃X�y�[�X�i�S�p�̉\��������̂�Format�͎g���Ȃ��j
    If LenB(StrConv(sTitle, vbFromUnicode)) < 24 Then
        sTitle = sTitle & Space(24 - LenB(StrConv(sTitle, vbFromUnicode))) & Space(2)
    Else
        sTitle = sTitle & Space(2)
    End If

    ' �t�@�C���̑��݃`�F�b�N�B�ُ펞��----�̕\��
    If fsoObj.FolderExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If

    bRet = fReadFileList(structDispInfo.sFilePath & "\" & MN_FILELIST)
    If bRet <> True Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If
    structDispInfo.sFilePath = structDispInfo.sFilePath & "\" & FileList(0)   ' �t�@�C�����X�g����o�[�W���������擾����

    '--------------------------------------------
    '�t�@�C�����擾
    '--------------------------------------------
    '�t�@�C���̑��݃`�F�b�N�B�ُ펞��----�̕\��
    If fsoObj.FileExists(structDispInfo.sFilePath) = False Then
        sErrEventName = LOG_ERR_FILE_NOTING     '�t�@�C������
        '�ُ�
        GoTo Err_FILE
    End If
    
    lLen = FileLen(structDispInfo.sFilePath)              '�t�@�C���T�C�Y�̎擾
    If lLen < Len(uFooter) Then
        sErrEventName = LOG_ERR_FILE_LENGTH     '�t�@�C�������O�X�ُ�
        '�ُ�
        GoTo Err_FILE
    End If
    
    '���g�p�̃t�@�C���ԍ��擾
    iFileNumber = FreeFile
    
    '�t�@�C���̃I�[�v��
    sErrEventName = LOG_ERR_FILE_OPEN       '�t�@�C���I�[�v���ُ�
    Open structDispInfo.sFilePath For Binary Access Read As #iFileNumber
    
        sErrEventName = LOG_ERR_FILE_READ       '�t�@�C���Ǎ��ُ�
        '�t�b�^���̎擾
        Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
        
    sErrEventName = LOG_ERR_FILE_CLOSE      '�t�@�C���N���[�Y�ُ�
    Close #iFileNumber      '�t�@�C���̃N���[�Y
    
    '--------------------------------------------
    '�o�[�W�������\�����̕\���e�L�X�g�쐬
    '--------------------------------------------
    '�^�C�g��
    sDisp = sTitle
    
    '�t�@�C����
    Call psFileNameGet(structDispInfo.sFilePath, sDispFile, sDispExe)             '�t�@�C���p�X����t�@�C�������擾
    sDispFile = sDispFile & "." & sDispExe                                        '�g���q���������啶���ɕϊ�
    sDisp = sDisp & Format(Right(sDispFile, 12), "!@@@@@@@@@@@@") & Space(7)

    '���
    sDisp = sDisp & LCase(Right$("0" & Hex(uFooter.bSyubetu), 2))
    
    '���[�J��
    sDisp = sDisp & uFooter.sMakerName & Space(2)
    
    '�f�[�^���{�o�[�W����
    uFooter.sDataName = Replace(uFooter.sDataName, vbNullChar, Space(1))
    sDispDV = LTrim(uFooter.sDataName) & uFooter.sVersion
    If Len(Trim(uFooter.sDataName)) = 0 And Len(Trim(uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(Trim(uFooter.sVersion) & Space(20), 20) & Space(2)
    ElseIf Len(Trim(uFooter.sDataName & uFooter.sVersion)) <> 0 Then
        sDisp = sDisp & Left(sDispDV & Space(20), 20) & Space(2)
    Else
        sDisp = sDisp & String(20, "-") & Space(2)
    End If
    
    '�쐬����
    sDateTime = ""
    For j = 0 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    sDateTime = Format(sDateTime, "@@@@/@@/@@ @@:@@")
    sDisp = sDisp & sDateTime
    
    '���X�g�ɒǉ�
    lstKan.AddItem (sDisp)
    
    '�R�����g
    '60�����ŕۑ�����Ă���̂ŁA60�o�C�g�̗̈�ɒ����B�O��̋󔒂����B
    sDispCom = Trim(StrConv(MidB(StrConv(uFooter.sHyoji, vbFromUnicode), 1, 60), vbUnicode))
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 1, 32), vbUnicode)     '�R�����g1
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    Else
        lstKan.AddItem (Space(45) & String(32, "-"))
    End If
    
    sDisp = StrConv(MidB(StrConv(sDispCom, vbFromUnicode), 33, 60), vbUnicode)    '�R�����g2
    If Len(Trim(sDisp)) <> 0 Then
        lstKan.AddItem (Space(45) & sDisp)
    End If
    
    Set fsoObj = Nothing
    
    Exit Sub
Err_FILE:
    
    '�ُ탍�O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, sErrEventName, lngErrCode)
    '���O�o�́@���t�@�C����
    Call psFileNameGet(structDispInfo.sFilePath, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)
    
    '�t�@�C���N���[�Y
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    
    '�ُ�p�\��
    lstKan.AddItem (sTitle & "------------       ---  --------------------  ----/--/-- --:--")
    lstKan.AddItem (Space(45) & "--------------------------------")
    lstKan.AddItem (Space(45) & "--------------------------------")

    Set fsoObj = Nothing

End Sub
'V30.1.0.1 ADD END


' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fReadFileList
'//  �@�\����  : �t�@�C�����X�g�̎擾
'//  �@�\�T�v  : �t�@�C�����X�g���A�t�@�C�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fReadFileList(sFileList As String) As Boolean
    Dim iFileNumber As Integer      '�t�@�C���ԍ�
    Dim sFileName As String         '�t�@�C����
    Dim iListCnt As Integer         '�t�@�C���i�[��

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����

    Open sFileList For Input Access Read As #iFileNumber    '�t�@�C�����X�g�̃I�[�v��
    Do While Not EOF(iFileNumber)                           '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        Line Input #iFileNumber, sFileName                  '�f�[�^��ǂݍ��݂܂��B
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                '�t�@�C���������݂���
            iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
            ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
            ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������
            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
                                                            '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
        End If
    Loop
    Close #iFileNumber      '�t�@�C������܂��B

    fReadFileList = True    '�߂�l�𐳏�Ƃ���

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    'V1.21.0.1 ADD  START
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    'V1.21.0.1 ADD  END
    fReadFileList = False   '�߂�l���G���[�Ƃ���
End Function
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��
' EG20 V5.6.0.1�y�h�b�l�o�[�W�����t�@�C�����X�g�Ή��z�ǉ��J�n
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : fReadFileListIDU
'//  �@�\����  : IDU�t�@�C�����X�g�̎擾
'//  �@�\�T�v  : �t�@�C�����X�g���A�t�@�C�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-04  CODED BY  [TCC] H.Sugimoto
'//                 �y�h�b�l�o�[�W�����t�@�C�����X�g�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fReadFileListIDU(sFileList As String) As Boolean
    Dim iFileNumber As Integer      '�t�@�C���ԍ�
    Dim sFileName As String         '�t�@�C����
    Dim iListCnt As Integer         '�t�@�C���i�[��
    Dim nIndex As Integer           ' ������

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����

    Open sFileList For Input Access Read As #iFileNumber    '�t�@�C�����X�g�̃I�[�v��
    Do While Not EOF(iFileNumber)                           '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        Line Input #iFileNumber, sFileName                  '�f�[�^��ǂݍ��݂܂��B
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                '�t�@�C���������݂���
            iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
            ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
            ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������

            nIndex = InStr(sFileName, " ")
            If nIndex = 0 Then
                ' �X�y�[�X���܂܂�Ă��Ȃ��ꍇ
                FileListType(iListCnt - 1) = UCase(Trim$(sFileName))
            Else
                ' �X�y�[�X���܂܂�Ă���ꍇ
                FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, nIndex)))
            End If
            FileList(iListCnt - 1) = FileListType(iListCnt - 1)
                                                            '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
        End If
    Loop
    Close #iFileNumber      '�t�@�C������܂��B

    fReadFileListIDU = True    '�߂�l�𐳏�Ƃ���

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    fReadFileListIDU = False   '�߂�l���G���[�Ƃ���
End Function
' EG20 V5.6.0.1�y�h�b�l�o�[�W�����t�@�C�����X�g�Ή��z�ǉ��I��

'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : fIsExistCornerType
'//  �@�\����  : �R�[�i�[�^�C�v���݃`�F�b�N
'//  �@�\�T�v  : �����R�[�i�[�����݂��邩�A�ݗ����R�[�i�[�����݂��邩�`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : byte  byCornerType�@[IN]   0:�ݗ����R�[�i�[
'//                                         1:�����R�[�i�[
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : boolean  true/false    true:�ݗ����R�[�i�[�L�� false:�����R�[�i�[�L��
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fIsExistCornerType(intCornerType As Integer)

    Dim intCount        As Integer
    Dim byFindFlg       As Byte
    
    byFindFlg = False

    '�e�R�[�i�[�̃R�[�i�[�^�C�v���擾����
    Call gsGetSettiCorner
    Call gsGetCornerType
    
    If intCornerType = CORNER_TYPE_KANSEN Then   '�����R�[�i�[�����݂��邩�m�肽���ꍇ
        For intCount = 0 To UBound(gblnCornerSet)
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN And gblnCornerSet(intCount) = True Then
                byFindFlg = True        '�����R�[�i�[����ł������OK
                Exit For
            End If
        Next intCount
    Else                                        '�ݗ����R�[�i�[�����݂��邩�m�肽���ꍇ
        For intCount = 0 To UBound(gblnCornerSet)
            If gintCornerType(intCount) = CORNER_TYPE_ZAIRAI And gblnCornerSet(intCount) = True Then
                byFindFlg = True        '�ݗ����R�[�i�[����ł������OK
                Exit For
            End If
        Next intCount
    End If
    
    fIsExistCornerType = byFindFlg
          

End Function
