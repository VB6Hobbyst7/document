VERSION 5.00
Begin VB.Form frmVerChang 
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
   Begin VB.Timer tmrMail 
      Left            =   5160
      Top             =   1080
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "���C�^��"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�o�`�r�l�n�^��"
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
      Left            =   240
      TabIndex        =   19
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "EG-R����"
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4095
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   27
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�o�[�W�����`�F�b�N�t�@�C���F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�\��2�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E���C��CPU-OS�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1380
         Width           =   2370
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E���C��CPU-Pro�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   705
         Width           =   2205
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�\���P�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E�T�uCPU-Pro�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1035
         Width           =   2175
      End
      Begin VB.Label lblVerName 
         Caption         =   "�E����CPU-Pro�F "
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2190
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   11
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3480
         TabIndex        =   10
         Top             =   705
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   9
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   8
         Top             =   1035
         Width           =   495
      End
      Begin VB.Label lblVerName 
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "����h�b�|�l"
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
      Left            =   8400
      TabIndex        =   3
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�d�f�|�q����"
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
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�m�d�f����"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   5760
      Width           =   3255
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
   Begin VB.Label lblVerName 
      Caption         =   "Z9"
      Height          =   375
      Index           =   22
      Left            =   7320
      TabIndex        =   29
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "�E���C�^���F "
      Height          =   495
      Index           =   21
      Left            =   4440
      TabIndex        =   28
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   20
      Left            =   7320
      TabIndex        =   25
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "XXXXXXXXXXXXXXXXXXXX"
      Height          =   375
      Index           =   19
      Left            =   7320
      TabIndex        =   24
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "Z9"
      Height          =   375
      Index           =   18
      Left            =   7320
      TabIndex        =   23
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label lblVerName 
      Caption         =   "�ENEG�����F"
      Height          =   375
      Index           =   15
      Left            =   4440
      TabIndex        =   22
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblVerName 
      Caption         =   "�E�h�b���ʉ^���F "
      Height          =   495
      Index           =   17
      Left            =   4440
      TabIndex        =   20
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�o�[�W�����ؑ�"
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
      TabIndex        =   5
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblVerName 
      Caption         =   "�E�h�b�|�l�F"
      Height          =   375
      Index           =   16
      Left            =   4440
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "frmVerChang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmVerChang.frm
'//  �p�b�P�[�W���F�o�[�W�����ؑ։��
'//
'//  �T�v�F�o�[�W�����ؑ։��
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//     REVISIONS :(1.0.6.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y1�s��Ή�
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
Private Const APL_INTERVAL = 390000     '�A�v���N���^�C�}�f�t�H���g�l 'V1.6.0.1 ADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����ؑ։��(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
Private Sub Form_Activate()
   On Error Resume Next
    
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����ؑ։��(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  On Error Resume Next
 
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmVerChang.Caption, False
        pfFormActive (frmVerChang.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����ؑ։��(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-30   REVISED BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

   On Error Resume Next
 
   '�u�o�[�W�����ؑ։�ʁF�\���v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_GAMEN_START, 0)

   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
      
   '���C�^���Ή��`�F�b�N
   psJikiCheck
    
   'IDU�k�ރ`�F�b�N
   psIDUCheck
    
   If pbIDUSts = 1 Then
     'IDU�Ɩ���\��
      cmdFixedExe(5).Visible = False
      cmdFixedExe(6).Visible = False
   End If
   
   '�o�[�W�����擾����
   psGetVersion

   '���[����M�p�̃^�C�}�l��ݒ肷��B
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
   
   'V1.6.0.1 ADD START
   'INI�t�@�C�����A�v���N���^�C�}�l���擾
   frmChangeVer.lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      APL_INTERVAL, HOSHU_FILE)
   '�擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
   If frmChangeVer.lngMAX_Time = 0 Then
      frmChangeVer.lngMAX_Time = APL_INTERVAL
   End If
   '�^�C�}�l�ݒ�
   frmChangeVer.tmrAplCheck.Interval = MN_MAIL_INTERVAL
   frmChangeVer.tmrAplCheck.Enabled = False
   'V1.6.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �e�t�Ώۂ̃o�[�W�����ؑ֏������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@ Index    [IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    
   On Error Resume Next
    
    Dim iRet As Integer
    Dim strWord As String
    Dim bRet As Boolean
    
    '�֑ؑΏۂ�ϐ��ɐݒ肷��B
    Change_Version = Index
    
    Select Case Index
       Case EGR_CHANGE_VER
         '�u�o�[�W�����ؑ։�ʁFEG-R�����t�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_EGRJIKAI_BUTTOM, 0)
       Case NEG_CHANGE_VER
         '�u�o�[�W�����ؑ։�ʁFNEG�����t�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_NEGJIKAI_BUTTOM, 0)
       Case ICM_CHANGE_VER
         '�u�o�[�W�����ؑ։�ʁF����IC-M�t�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_ICM_BUTTOM, 0)
       Case PASMO_CHANGE_VER
         '�u�o�[�W�����ؑ։�ʁFPASMO�^���t�����v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_PASMO_BUTTOM, 0)
       Case JIKIUNCHIN_CHANGE_VER
         '�u�o�[�W�����ؑ։�ʁF���C�^���t�����v���O�o��
          Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_KANRI_MENU_JIKIUNCHIN_BUTTOM, 0)
  End Select
    
    strWord = cmdFixedExe(Index).Caption & "�̃o�[�W������؂�ւ��܂��B" & vbCrLf & "��낵���ł����H"
    
    iRet = MsgBox(strWord, vbQuestion + vbOKCancel, "�o�[�W�����ؑ֊m�F")
    
    If iRet = vbOK Then
       Load frmChangeVer
       frmChangeVer.lblMessage(0).Caption = "�o�[�W�����ؑ֒��ł��B"
       frmChangeVer.lblMessage(1).Caption = "���΂炭���҂��������B"
       frmChangeVer.Show 1
    End If
'V1.8.0.1 ADD START
    '�o�[�W�����擾����
    psGetVersion
'V1.8.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����
'//  �@�\�T�v  : ����ʂ���������B
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
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '�u�o�[�W�����ؑ։�ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_CHANGE_GAMEN_END, 0)
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
'//     ORIGINAL :(1.4.0.1) 2009-03-17    CODED BY [TCC] S.Terao
'//     REVISIONS :(1.10.0.1) 2009-09-25   REVISED BY [TCC] T.Furuya
'//                 KK�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub psGetVersion()

  On Error Resume Next
 
  Dim sVersion  As String
  Dim sGetJikiVer As String     'V1.10.0.1 ADD
  
 'EG-R�����o�[�W�����擾
  '����CPU
  sVersion = psEGRJVersion(HANTEI_CPU)
  lblVerName(8).Caption = sVersion
  '���C��CPU
  sVersion = psEGRJVersion(MAIN_CPU)
  lblVerName(9).Caption = sVersion
 '�T�uCPU
  sVersion = psEGRJVersion(SUB_CPU)
  lblVerName(10).Caption = sVersion
 '���C��OS
  sVersion = psEGRJVersion(MAIN_OS)
  lblVerName(11).Caption = sVersion
 '�\���P
  sVersion = psEGRJVersion(YOBI1)
  lblVerName(12).Caption = sVersion
 '�\���Q
  sVersion = psEGRJVersion(YOBI2)
  lblVerName(13).Caption = sVersion
 '�o�[�W�����`�F�b�N
  sVersion = psEGRJVersion(VER_CHK)
  lblVerName(14).Caption = sVersion
  
 'NEG�����o�[�W�����擾
  sVersion = psNEGJVersion
  lblVerName(18).Caption = sVersion

 'IC-M�o�[�W�����擾
 If pbIDUSts = 1 Then
    'IDU�o�[�W������\��
    lblVerName(16).Enabled = False
    lblVerName(19).Caption = ""
 Else
    '��k�ގ��͕\������
    sVersion = psICMGetVersion
    lblVerName(19).Caption = sVersion
 End If
 
 '���ʉ^���o�[�W�����擾
 If pbIDUSts = 1 Then
    'IDU�o�[�W������\��
    lblVerName(17).Enabled = False
    lblVerName(20).Caption = ""
 Else
    '��k�ގ��͕\������
    sVersion = psICUnchinGetVersion
    lblVerName(20).Caption = sVersion
 End If
 
'V1.10.0.1 ADD START
 '���C�^���ǂݍ���
 sGetJikiVer = psJikiUnchinVersion
 lblVerName(22).Caption = CStr(sGetJikiVer)
'V1.10.0.1 ADD END

 
End Sub

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
'//     ORIGINAL  :(1.4.0.1) 2009-03-25   CODED   BY [TCC] S.Terao
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
