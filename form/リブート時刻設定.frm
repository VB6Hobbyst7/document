VERSION 5.00
Begin VB.Form frmRebootTimeSettei 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail2 
      Left            =   10440
      Top             =   7080
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "�m��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   2
      Left            =   7680
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "�m��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   1
      Left            =   9600
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
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
      Height          =   3255
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   11175
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FFFF&
         Caption         =   "��"
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
         Index           =   5
         Left            =   9360
         Style           =   1  '���̨���
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FFFF&
         Caption         =   "��"
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
         Index           =   4
         Left            =   7560
         Style           =   1  '���̨���
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FFFF&
         Caption         =   "��"
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
         Index           =   3
         Left            =   5760
         Style           =   1  '���̨���
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FF80&
         Caption         =   "��"
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
         Left            =   4080
         Style           =   1  '���̨���
         TabIndex        =   9
         Top             =   960
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FF80&
         Caption         =   "��"
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
         Left            =   2280
         Style           =   1  '���̨���
         TabIndex        =   7
         Top             =   960
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox ChkSetTaku 
         BackColor       =   &H0080FF80&
         Caption         =   "��"
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
         Index           =   0
         Left            =   480
         Style           =   1  '���̨���
         TabIndex        =   5
         Top             =   960
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   9120
         TabIndex        =   16
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   7320
         TabIndex        =   14
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5520
         TabIndex        =   12
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblCornerName 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "������������������������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2055
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   11175
      Begin VB.CommandButton cmdKakutei 
         Caption         =   "�m��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   0
         Left            =   9240
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox ChkSet 
         BackColor       =   &H0080FF80&
         Caption         =   "��"
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
         Left            =   480
         Style           =   1  '���̨���
         TabIndex        =   4
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  �V�X�e���ݒ�    ��ʂ֖߂�"
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
   Begin VB.Timer tmrMail 
      Left            =   11400
      Top             =   7080
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "���u�[�g�ݒ�"
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
      TabIndex        =   1
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmRebootTimeSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmRebootTimeSettei.frm
'//  �p�b�P�[�W���F���u�[�g�ݒ���
'//  �T�v        �F���u�[�g�ݒ���
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-15  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-264�z
'//  REVISIONS   �F(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

Private Const REBOOTSW_ON_MESSAGE = "��"    ' �t���b�Z�[�W�F�����
Private Const REBOOTSW_OFF_MESSAGE = "��"   ' �t���b�Z�[�W�F�؏��
Private Const REBOOTSW_ON_COLOR = &H80FF80  ' �t�F�F�����
Private Const REBOOTSW_OFF_COLOR = &H80FFFF ' �t�F�F�؏��
Private Const REBOOTSW_ON_VALUE = 1         ' �t��ԁF�����
Private Const REBOOTSW_OFF_VALUE = 0        ' �t��ԁF�؏��

' DA�ݒ���e
Private Const ID_KANSI_SET_RBOOT_SET = &H14  ' �Ď����u�ݒ�h�c�F���u�[�g�ݒ�
Private Const REBOOT_ON_DASTATUS = 1         ' �t��ԁF�����
Private Const REBOOT_OFF_DASTATUS = 0        ' �t��ԁF�؏��

Private Const HUTEI = 0                      ' �l�s��


'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  �֐�����     : Form_Load
'/  �@�\����     : Form_Load������
'/  �@�\�T�v     : Form_Load���������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                 EG20�t�F�[�Y�Q�Ή�
'/                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'/  REVISIONS    :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'/                 EG20�t�F�[�Y�Q�Ή��y����TR-264�z
'/  REVISIONS    :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    Dim intLoop         As Integer          ' ���[�v�J�E���^
    Dim intStatus       As Integer          ' ���u�[�g�ݒ���
    Dim strSecName(5)   As String
    Dim strDefault      As String
    Dim lngRet          As Long
    Dim strRet          As String * 32
    Const lngBufSize = 32
    
    Dim strCorner1      As String           ' ������i�[�G���A1
    Dim strCorner2      As String           ' ������i�[�G���A2
    
    On Error Resume Next
    
    '�u���u�[�g�ݒ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���[����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
        
' EG20 V6.8.0.1 ADD START
    tmrMail2.Interval = MN_MAIL_INTERVAL
    tmrMail2.Enabled = False
' EG20 V6.8.0.1 ADD END
        
    ' /////////////////////////////////////////////////////////////////////////
    ' // �����Ď��Րݒ�
    ' /////////////////////////////////////////////////////////////////////////
    
    ' /////////////////////////////////////////////////
    ' // �����Ď��Ֆt
    
    ' ���݂̐ݒ��Ԃ��擾
    intStatus = pfGetKansiArea_Sts(ID_KANSI_SET_RBOOT_SET)
    
    If intStatus = REBOOT_ON_DASTATUS Then
        ChkSet.Caption = REBOOTSW_ON_MESSAGE
        ChkSet.BackColor = REBOOTSW_ON_COLOR
        ChkSet.Value = REBOOTSW_ON_VALUE
    Else
        ChkSet.Caption = REBOOTSW_OFF_MESSAGE
        ChkSet.BackColor = REBOOTSW_OFF_COLOR
        ChkSet.Value = REBOOTSW_OFF_VALUE
    End If
    ChkSet.Visible = True
    ChkSet.Enabled = True
    
   
    ' /////////////////////////////////////////////////////////////////////////
    ' // �����ݒ�
    ' /////////////////////////////////////////////////////////////////////////
' EG20 V3.3.0.1�y����TR-264�z�폜�J�n
'    strDefault = ""
'    strSecName(0) = strAppName_station
'    strSecName(1) = strAppName_station2
'    strSecName(2) = strAppName_station3
'    strSecName(3) = strAppName_station4
'    strSecName(4) = strAppName_station5
'    strSecName(5) = strAppName_station6
' EG20 V3.3.0.1�y����TR-264�z�폜�I��
    
    ' �R�[�i���̐ݒ菈��
    Call gsGetCornerName
    
    For intLoop = 0 To UBound(strSecName)
    
        '�ݒ肠��̃R�[�i�������ɂ���
        If gudtSettiCorner(intLoop).intGokiNum > 0 Then
' EG20 V3.3.0.1�y����TR-264�z�폜�J�n
'            'Ini�t�@�C������R�[�i�[�����擾
'            lngRet = GetPrivateProfileString(strSecName(intLoop), IDU_PROFILE_KEY_NAME_CONER, _
'                                                strDefault, strRet, lngBufSize, KANSI_STATION_INI_FILE)
'            ' /////////////////////////////////////////////////
'            ' // ���x���i�R�[�i�[���̕\���j
'            lblCornerName(intLoop).Caption = Trim(strRet)
' EG20 V3.3.0.1�y����TR-264�z�폜�I��
' EG20 V3.3.0.1�y����TR-264�z�ǉ��J�n
            ' /////////////////////////////////////////////////
            ' // ���x���i�R�[�i�[���̕\���j
            strCorner1 = MidB(gstrCornerName(intLoop), 1, 12)
            strCorner2 = MidB(gstrCornerName(intLoop), 13, 24)
            lblCornerName(intLoop).Caption = strCorner1 & vbCrLf & strCorner2
' EG20 V3.3.0.1�y����TR-264�z�ǉ��I��
            lblCornerName(intLoop).Visible = True
        
            ' /////////////////////////////////////////////////
            ' // �t
            
            ' ���݂̐ݒ��Ԃ��擾
            Call pfGetJikaiSts(intStatus, intLoop + 1)
            
            If intStatus = REBOOT_ON_DASTATUS Then
                ChkSetTaku(intLoop).Caption = REBOOTSW_ON_MESSAGE
                ChkSetTaku(intLoop).BackColor = REBOOTSW_ON_COLOR
                ChkSetTaku(intLoop).Value = REBOOTSW_ON_VALUE
            Else
                ChkSetTaku(intLoop).Caption = REBOOTSW_OFF_MESSAGE
                ChkSetTaku(intLoop).BackColor = REBOOTSW_OFF_COLOR
                ChkSetTaku(intLoop).Value = REBOOTSW_OFF_VALUE
            End If
            ChkSetTaku(intLoop).Visible = True
            ChkSetTaku(intLoop).Enabled = True

        Else
            lblCornerName(intLoop).Caption = Trim(strRet)
            lblCornerName(intLoop).Visible = False
            ChkSetTaku(intLoop).Enabled = False
            ChkSetTaku(intLoop).Visible = False
            ChkSetTaku(intLoop).Value = REBOOTSW_OFF_VALUE
        End If
    
    Next intLoop
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // ���̑��R���g���[���ݒ�
    ' /////////////////////////////////////////////////////////////////////////
    
    cmdKakutei(0).Enabled = False       ' �����Ď��Ձu�m��v
    cmdKakutei(1).Enabled = False       ' �����u�m��v
    cmdKakutei(2).Enabled = False       ' �u�m��v        �F�����s�� EG20 V5.13.0.1 ADD
    cmdReturn.Enabled = True            ' �u�߂�v
    
End Sub


'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : cmdKakutei_Click
'// �@�\����    : �m��{�^������������
'// �@�\�T�v    : �m��{�^�����������ꂽ�������s��
'//
'//                   �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή�
'//               EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'// REVISIONS :(EG20 V5.13.0.1) 2012-06-07  CODED BY  [TCC] H.Sano
'/             EG20�m�F�t1�Ή�
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdKakutei_Click(Index As Integer)
    
    Dim intLoop     As Integer          ' ���[�v�J�E���^
    Dim udtSendData As ML_REBOOT_REQ    ' ���u�[�g�ݒ���v��
    Dim lngSendSize As Long             ' ���M���郁�[���T�C�Y
    Dim lngErrCode  As Long             ' �G���[�R�[�h
    Dim bRet        As Boolean          ' ���[�����M�����߂�l
    Dim intLoopMail As Integer          ' ���[�v�J�E���^2 EG20 V5.13.0.1 ADD
    
    On Error Resume Next
    
    '�u���u�[�g�ݒ��ʁF�m��t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_KAKUTEI_BUTTOM, 0)

    ' �R���g���[���ݒ�
    cmdKakutei(0).Enabled = False       ' �����Ď��Ձu�m��v    �F�����s��
    cmdKakutei(1).Enabled = False       ' �����u�m��v        �F�����s��
    cmdKakutei(2).Enabled = False       ' �u�m��v              �F�����s�� EG20 V5.13.0.1 ADD
    cmdReturn.Enabled = False           ' �u�߂�v              �F�����s��
    
    ChkSet.Enabled = False              ' �����Ď��Ձu���^�؁v  �F�����s��

    For intLoop = 0 To CONECT_CORNER_MAXINDEX   ' �����u���^�؁v      �F�����s��
        '�ݒ肠��̃R�[�i�������ɂ���
        If ChkSetTaku(intLoop).Visible = True Then
            ChkSetTaku(intLoop).Enabled = False
        End If
    Next intLoop

    '�m�F�{�^�������p�^�C�}���쓮������
    tmrMail.Interval = MN_MAIL_INTERVAL ' �{�^�������p�^�C�}���Ԑݒ�
    tmrMail.Enabled = True              ' �^�C�}�쓮

'EG20 V5.13.0.1 ADD START
For intLoopMail = 0 To 1
'EG20 V5.13.0.1 ADD END

    ' ���[���̑��M���e��ҏW����
    udtSendData.udtlHeader.dwId = ML_ID_REBOOT_REQ      ' ���[���h�c�@=�h"�ݒ���v���i���u�[�g�j"
    udtSendData.udtlHeader.dwSize = MlSize.REBOOT_REQ   ' ���[���T�C�Y=�h"�ݒ���v��"
    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' ���M���v���Z�X�h�c=�h�ێ�h
    udtSendData.udtlHeader.dwSubArea = 0                ' �⏕���@=�@0
                                                        ' �������ʂ͉��������u�m��v�t�ɑΉ�
'EG20 V5.13.0.1 MOD START
'    udtSendData.dwControl = Index                       ' �����ʁi0:�����Ď���,1:�����j
    udtSendData.dwControl = intLoopMail                  ' �����ʁi0:�����Ď���,1:�����j
'EG20 V5.13.0.1 MOD END
                                                        ' �����^�ؐݒ�͖t�ɑΉ�
    udtSendData.dwKanshi = ChkSet.Value                 ' �����Ď��Րݒ�i0:��,1:���j
    For intLoop = 0 To CONECT_CORNER_MAXINDEX                    ' �����ݒ�i0:��,1:���j
        '�ݒ肠��̃R�[�i�������ɂ���
        udtSendData.dwTaku(intLoop) = ChkSetTaku(intLoop).Value
    Next intLoop
    
    ' ���M�T�C�Y��ݒ肷��B
    lngSendSize = udtSendData.udtlHeader.dwSize
            
    ' �Ď��ՋN���`�F�b�N
    If CheckAppStart(PROC_KANRI) = 0 Then
        ' /////////////////////////////////////////////////
        ' // �Ď��Ֆ��N���F���͂Őݒ�l���X�V
        If udtSendData.dwControl = 0 Then
            ' ////////////////////////////////////////////
            ' // �����Ď���
            bRet = gspfSetKansiSts(ID_KANSI_SET_RBOOT_SET, ChkSet.Value)
        Else
            ' ////////////////////////////////////////////
            ' // �����
            For intLoop = 0 To CONECT_CORNER_MAXINDEX
                If ChkSetTaku(intLoop).Visible = True Then
                    bRet = pfSetJikaiSts(ChkSetTaku(intLoop).Value, intLoop + 1, IdGate.ID_GATE_SET_RBOOT_SET)
                End If
            Next intLoop
        End If
    Else
        ' /////////////////////////////////////////////////
        ' // �Ď��ՋN�����F���͂Őݒ�l���X�V
        
        ' �ă}�ɑ΂��āA�ݒ���v�����[���𑗐M����B
        bRet = DssSendMail(MAIL_SLOT_KANMA, lngSendSize, udtSendData.udtlHeader)
        ' ���[���𐳏�ɑ��M�������̃��O
        If bRet = False Then
            '�u�ݒ���v�����[�����M�ُ�v���O�o��
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
            Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, lngErrCode)
        Else
            '�u�ݒ���v�����[�����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, HOSHU_MENU_GAMEN_CMD, 0)
        End If
    End If

'EG20 V5.13.0.1 ADD START
Next intLoopMail
'EG20 V5.13.0.1 ADD END

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : cmdReturn_Click
'// �@�\����    : ���j���[�ɖ߂�{�^����������
'// �@�\�T�v    : ���j���[�ɖ߂�{�^�������������s��
'//
'//                  �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή�
'//               EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    On Error Resume Next

    '�u���u�[�g�ݒ��ʁF�߂�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_GAMEN_END, 0)

    '��ʂ�Unload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : Form_Activate
'// �@�\����    : Form_Activate������
'// �@�\�T�v    : Form_Activate���������s��
'//
'//                   �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή�
'//               EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'// REVISIONS   :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    tmrMail2.Enabled = True             ' EG20 V6.8.0.1 ADD
    
    pfFormActive (hwnd)

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : Form_Deactivate
'// �@�\����    : �f�B�A�N�e�B�u��
'// �@�\�T�v    : ���[����M�p�̃^�C�}��~
'//
'//                   �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή�
'//               EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'// REVISIONS   :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   
   On Error Resume Next
    
    '�^�C�}���~����B
    tmrMail.Enabled = False
    
    tmrMail2.Enabled = False             ' EG20 V6.8.0.1 ADD

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : tmrMail_Timer
'// �@�\����    : �^�C���A�E�g����
'// �@�\�T�v    : �^�C�}�^�C���A�E�g�������s��
'//
'//                   �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή�
'//               EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim intLoop     As Integer          ' ���[�v�J�E���^

    '�^�C�}���~
    tmrMail.Enabled = False

    ' �R���g���[���ݒ�
    cmdKakutei(0).Enabled = False       ' �����Ď��Ձu�m��v    �F�����s��
    cmdKakutei(1).Enabled = False       ' �����u�m��v        �F�����s��
    cmdKakutei(2).Enabled = False       ' �u�m��v        �F�����s�� EG20 V5.13.0.1 ADD
    cmdReturn.Enabled = True            ' �u�߂�v              �F�����\
    
    ChkSet.Enabled = True               ' �����Ď��Ձu���^�؁v  �F�����\

    For intLoop = 0 To CONECT_CORNER_MAXINDEX    ' �����u���^�؁v      �F�����\
        '�ݒ肠��̃R�[�i�������ɂ���
        If ChkSetTaku(intLoop).Visible = True Then
            ChkSetTaku(intLoop).Enabled = True
        End If
    Next intLoop
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : ChkSet_Click
'// �@�\����    : �����Ď��Ձu���^�؁v�t����
'// �@�\�T�v    : �����Ď��Ձu���^�؁v�t�����������s��
'//
'//                   �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή�
'//               EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub ChkSet_Click()
    
    On Error Resume Next
    
   
    '�u���u�[�g�ݒ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_KANSHI_BUTTOM, 0)

    If ChkSet.Value = REBOOTSW_ON_VALUE Then
        ' /////////////////////////////////////////////////
        ' �؁����ݒ�
        ChkSet.Caption = REBOOTSW_ON_MESSAGE
        ChkSet.BackColor = REBOOTSW_ON_COLOR
    Else
        ' /////////////////////////////////////////////////
        ' �����ؐݒ�
        ChkSet.Caption = REBOOTSW_OFF_MESSAGE
        ChkSet.BackColor = REBOOTSW_OFF_COLOR
    End If

    cmdKakutei(0).Enabled = True       ' �����Ď��Ձu�m��v
    cmdKakutei(2).Enabled = True       ' �u�m��v        �F EG20 V5.13.0.1 ADD

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : ChkSetTaku_Click
'// �@�\����    : �����u���^�؁v�t����
'// �@�\�T�v    : �����u���^�؁v�t�����������s��
'//
'//                   �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή�
'//               EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub ChkSetTaku_Click(Index As Integer)
    
    On Error Resume Next
    
   
    '�u���u�[�g�ݒ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_REBOOT_TAKU_BUTTOM, 0)

    If ChkSetTaku(Index).Value = REBOOTSW_ON_VALUE Then
        ' /////////////////////////////////////////////////
        ' �؁����ݒ�
        ChkSetTaku(Index).Caption = REBOOTSW_ON_MESSAGE
        ChkSetTaku(Index).BackColor = REBOOTSW_ON_COLOR
    Else
        ' /////////////////////////////////////////////////
        ' �����ؐݒ�
        ChkSetTaku(Index).Caption = REBOOTSW_OFF_MESSAGE
        ChkSetTaku(Index).BackColor = REBOOTSW_OFF_COLOR
    End If

    cmdKakutei(1).Enabled = True       ' �����u�m��v
    cmdKakutei(2).Enabled = True       ' �����u�m��v        �FEG20 V5.13.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGetKansiArea_Sts
'//  �@�\����  : �Ď��ݒ���(���[�h��)�B
'//  �@�\�T�v  : �Ď��ݒ��ʂ̏����������s���B
'//
'//              �^        ����          �Ӗ�
'//  ����      : Integer  intId          [IN]�G���AID
'//
'//              �^        �l       �@�@ �Ӗ�
'//  �߂�l    : Integer                 [OUT]���ݒl
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetKansiArea_Sts(intId As Integer) As Integer
    
    Dim iAreaSts     As Integer     '�Ď��ݒ��Ԓl
    Dim lSts         As Long        '�֐��߂�l
    Dim udtAreaR255  As GATE_INFO   '�Ǎ��ݗp�G���A�i255�ݒ�p�j
    Dim lngSts       As Long
    Dim lngLoop1     As Long
    Dim lngHandle    As Long
    Dim FileName     As String
    Dim lngRet       As Long
    Dim bRet         As Boolean
    Dim sSetteiFile  As String      '�t�@�C���p�X
    Dim lngAplSts    As Long        '�A�v���N��/���N������
            
    On Error Resume Next
      
    '�Ď��ՋN���L���`�F�b�N
    lngAplSts = CheckAppStart(PROC_KANRI)
    If lngAplSts = 0 Then
        '�Ď��Ֆ��N����
        '�Ď��ݒ�t�@�C�����I�[�v��
        lngHandle = CreateFile(K_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)  'V1.4.0.1�@ADD
        
        '�t�@�C���I�[�v��������ɍs��ꂽ���H
        If lngHandle = INVALID_HANDLE_VALUE Then
           '�I�[�v���ُ펞:�ُ�
           '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           pfGetKansiArea_Sts = HUTEI
           Exit Function
        End If
        
        '�Ď��ݒ�t�@�C���ǂݍ���
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           '�ǂݍ��ُ݈펞�F�ُ�
           pfGetKansiArea_Sts = HUTEI
         '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           '�n���h���̃N���[�Y
           Call CloseHandle(lngHandle)
           Exit Function
        End If
        
        '�n���h���̃N���[�Y
        Call CloseHandle(lngHandle)
        
        'ID����
        lngSts = KansiSerchId(udtAreaR255, CLng(intId))
        If lngSts >= 0 Then
           'ID���L�����ꍇ
           pfGetKansiArea_Sts = udtAreaR255.GateInfo(lngSts).bytDATA(0)
         Else
          ' �Y���h�c�����̏ꍇ:�ُ�
          '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          pfGetKansiArea_Sts = HUTEI
          Exit Function
        End If
    Else
        '�Ď��ՋN����
        Set Idinf_KansiSettei = New IdInfProc              '�Ď����u�ݒ�G���A
        '���L�G���A�I�[�v��
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei    '�Ď����u�ݒ�G���A
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
           pfGetKansiArea_Sts = HUTEI
           '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
           Exit Function
        End If
        
        '�Ď��ݒ�G���A���k�n�b�j����B
        Idinf_KansiSettei.IdLock
        If Idinf_KansiSettei.Errsts <> 0 Then
          '�f�[�^�Q�ƈُ펞:�ُ�
          pfGetKansiArea_Sts = HUTEI
          Idinf_KansiSettei.IdFree
          '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
          Exit Function
        End If
    
        '�Ď��ݒ�G���AID��ݒ�
        Idinf_KansiSettei.id = intId
        Idinf_KansiSettei.IdGet
        If Idinf_KansiSettei.Errsts <> 0 Then
          '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
          pfGetKansiArea_Sts = HUTEI
          Idinf_KansiSettei.IdFree
          '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
          Exit Function
        End If

        pfGetKansiArea_Sts = Idinf_KansiSettei.DataArea(0)   '�ݒ���e
      
        Idinf_KansiSettei.IdFree
        Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
   End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : KansiSerchId
'//  �@�\����  : �h�c��������
'//  �@�\�T�v  : �h�c�������s���B
'//
'//              �^        ����        �Ӗ�
'//  ����      : GATE_INFO udtArea255 [IN]�ϊ����f�[�^
'//�@�@�@�@�@�@�@Long�@�@�@lngId�@�@�@[IN]�G���AID
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Long�@�@�@         �@[OUT]�@0�ȏ�F����B-1�ȉ��F�G���[
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-15   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function KansiSerchId(udtArea255 As GATE_INFO, lngID As Long) As Long

    Dim lngIndex As Long                '�����p�C���f�b�N�X
    Dim lngMin As Long                  '�ŏ��C���f�b�N�X
    Dim lngMax As Long                  '�ő�C���f�b�N�X
    Dim lngChkIndex As Long             '�Y���C���f�b�N�X
    Dim lngWorkId   As Long             '�W���h�c

    On Error Resume Next
    
    '������
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '�����J�n
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             '�h�c���o��
        If lngID = lngWorkId Then                                  '�����H
            lngChkIndex = lngIndex                                  '�f�[�^���o����A�����I��
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngID < lngID) Then         '�f�[�^���\����������
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    KansiSerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGetJikaiSts
'//  �@�\����  : �����^�u�\������(�Ď��ՋN���L���Ή��Q��)
'//  �@�\�T�v  : �����^�u�̍��@�ʖt��Ԏ擾�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iJikaiSts [OUT]�\���X�e�[�^�X
'//              Integer�@iGouki  �@[IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-18   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�(�Ď��Ֆ��N�����ł��ݒ�ύX��)
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetJikaiSts(iJikaiSts As Integer, iGouki As Integer)
    Dim iAreaSts        As Integer          '�����ݒ�t�@�C����Ԓl
    Dim lSts            As Long             '�֐��߂�l
    Dim udtAreaR255     As GATE_INFO        '�Ǎ��ݗp�G���A�i255�ݒ�p�j
    Dim lngSts          As Long             '�q�b�g�G���AID
    Dim lngLoop1        As Long             '�J�E���^�[
    Dim lngHandle       As Long             '�n���h��
    Dim FileName        As String           '�t�@�C���L���`�F�b�N
    Dim lngRet          As Long             '�߂�l
    Dim bRet            As Boolean          '�ǂݍ��݌��ʖ߂�l
    Dim sSetteiFile     As String           '�t�@�C���p�X�@'V1.4.0.1�@ADD
    
    On Error Resume Next
'V1.4.0.1 DEL START
'    '�����ݒ�t�@�C���L��
'    FileName = Dir(G_SETTEI_FILE)
'    If FileName = "" Then
'       '������ΎQ�ƕs�̂��ߎQ�ƈُ�
'       iJikaiSts = GET_CONECTSTS_ERROR
'       Exit Function
'    End If
'V1.4.0.1 DEL END
'V1.4.0.1 ADD START
   '�����ݒ�t�@�C���L��
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '�����ݒ�t�@�C�����Ȃ��ꍇ
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '�����ݒ�t�@�C��������ꍇ
       sSetteiFile = G_SETTEI_FILE
    End If
'V1.4.0.1 ADD END

    '�Ď��ՋN���L���`�F�b�N
    If CheckAppStart(PROC_KANRI) = 0 Then
        
        '�����ݒ�t�@�C�����I�[�v��
'        lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 DEL
        lngHandle = CreateFile(sSetteiFile, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

        '�t�@�C���I�[�v��������ɍs��ꂽ���H
        If lngHandle = INVALID_HANDLE_VALUE Then
           '�I�[�v���ُ펞�͎Q�ƕs�̂��ߎQ�ƈُ�
'           iJikaiSts = GET_CONECTSTS_ERROR                                 ' EG20 V2.1.0.1[Mainte_03_01]�폜
           iJikaiSts = REBOOTSW_OFF_VALUE                                   ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
           Exit Function
        End If
        
        '�����ݒ�t�@�C���ǂݍ���
        For lngLoop1 = 0 To iGouki - 1
            bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        Next
        
        '�n���h���̃N���[�Y
        Call CloseHandle(lngHandle)
        
        'ID����
'        lngSts = SerchId(udtAreaR255, IdGate.JIKAI_CONECT_SETTEI)          ' EG20 V2.1.0.1[Mainte_03_01]�폜
        lngSts = SerchId(udtAreaR255, IdGate.ID_GATE_SET_RBOOT_SET)         ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
        If lngSts >= 0 Then
           'ID���L�����ꍇ
           iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         '�f�[�^�ϊ�
        Else
          ' �Y���h�c�����̏ꍇ�Q�ƈُ�
'          iJikaiSts = GET_CONECTSTS_ERROR                                  ' EG20 V2.1.0.1[Mainte_03_01]�폜
          iJikaiSts = REBOOTSW_OFF_VALUE                                    ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
          Exit Function
        End If
        
' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
'        Select Case iAreaSts
'           Case 1
'             '�ڑ�
'              iJikaiSts = CONECTSTS_ERROR
'              Exit Function
'           Case 0
'              iJikaiSts = CONECTSTS_END
'              Exit Function
'        End Select
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
        iJikaiSts = iAreaSts
        Exit Function
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
    
    Else
     
         Set Idinf_JikaiSettei = New IdInfProc              '�����ݒ�G���A
         '�����ݒ�G���A���I�[�v������B
          Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
          Idinf_JikaiSettei.IdOpen
          If Idinf_JikaiSettei.Errsts <> 0 Then
             '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
'             iJikaiSts = GET_CONECTSTS_ERROR              ' EG20 V2.1.0.1[Mainte_03_01]�폜
             iJikaiSts = REBOOTSW_OFF_VALUE                ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
             Set Idinf_JikaiSettei = Nothing               '�������u�ݒ�f�[�^�t�@�C��
             Exit Function
          End If
             
          '�����ݒ�G���A���k�n�b�j����B
          Idinf_JikaiSettei.IdLock
          If Idinf_JikaiSettei.Errsts <> 0 Then
             '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
'             iJikaiSts = GET_CONECTSTS_ERROR              ' EG20 V2.1.0.1[Mainte_03_01]�폜
             iJikaiSts = REBOOTSW_OFF_VALUE                ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
             Idinf_JikaiSettei.IdFree
             Set Idinf_JikaiSettei = Nothing               '�������u�ݒ�f�[�^�t�@�C��
             Exit Function
           End If
              
           '�G���A�̓��e��ǂݍ��ށB
'            Idinf_JikaiSettei.id = IdGate.JIKAI_CONECT_SETTEI              ' EG20 V2.1.0.1[Mainte_03_01]�폜
            Idinf_JikaiSettei.id = IdGate.ID_GATE_SET_RBOOT_SET             ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
            Idinf_JikaiSettei.GetJikai_Sts iGouki - 1
            If Idinf_JikaiSettei.Errsts <> 0 Then
               '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
'                iJikaiSts = GET_CONECTSTS_ERROR              ' EG20 V2.1.0.1[Mainte_03_01]�폜
                iJikaiSts = REBOOTSW_OFF_VALUE                ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
                Idinf_JikaiSettei.IdFree
                Set Idinf_JikaiSettei = Nothing               '�������u�ݒ�f�[�^�t�@�C��
                Exit Function
            End If
               
            '�ݒ���e���擾
             iAreaSts = Idinf_JikaiSettei.DataArea(iGouki - 1)
' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
'             Select Case iAreaSts
'                 Case 1
'                  '�ڑ�
'                   iJikaiSts = CONECTSTS_ERROR
'                   Idinf_JikaiSettei.IdFree
'                   Set Idinf_JikaiSettei = Nothing               '�������u�ݒ�f�[�^�t�@�C��
'                   Exit Function
'                 Case 0
'                   iJikaiSts = CONECTSTS_END
'                   Idinf_JikaiSettei.IdFree
'                   Set Idinf_JikaiSettei = Nothing               '�������u�ݒ�f�[�^�t�@�C��
'                   Exit Function
'             End Select
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
        iJikaiSts = iAreaSts                            ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
        Idinf_JikaiSettei.IdFree
        Set Idinf_JikaiSettei = Nothing               '�������u�ݒ�f�[�^�t�@�C��
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : SerchId
'//  �@�\����  : �h�c��������(�S�^�u��p)
'//  �@�\�T�v  : �h�c�������s���B
'//
'//              �^        ����        �Ӗ�
'//  ����      : GATE_INFO udtArea255 [IN]�ϊ����f�[�^
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Long�@�@�@         �@[OUT]�@0�ȏ�F����B-1�ȉ��F�G���[
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function SerchId(udtArea255 As GATE_INFO, lngID As Long) As Long

    Dim lngIndex As Long                '�����p�C���f�b�N�X
    Dim lngMin As Long                  '�ŏ��C���f�b�N�X
    Dim lngMax As Long                  '�ő�C���f�b�N�X
    Dim lngChkIndex As Long             '�Y���C���f�b�N�X
    Dim lngWorkId   As Long             '�W���h�c

    On Error Resume Next
    
    '������
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '�����J�n
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             '�h�c���o��
        If lngID = lngWorkId Then                                  '�����H
            lngChkIndex = lngIndex                                  '�f�[�^���o����A�����I��
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngID < lngID) Then         '�f�[�^���\����������
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    SerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : ChgData
'//  �@�\����  : �f�[�^�ϊ���������
'//  �@�\�T�v  : �f�[�^�ϊ������������s���B
'//
'//              �^        ����        �Ӗ�
'//  ����      : ID_FMT �@DataArea �@[IN]�ϊ����f�[�^
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String�@�@�@        [OUT]�@vbNullstring�ȊO�F����BvbNullString    �F�G���[
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function ChgData(DataArea As ID_FMT) As String

    Dim lngloop As Long
    Dim lngWork As Long
    Dim lngErrsts As Long

    On Error GoTo ChgDataErr
    
    lngErrsts = IdInfErr.OK
    
    Select Case DataArea.intType
    Case ID_TYPE.Flag   '���
        If (DataArea.bytDATA(0) <> 255) Then
            ChgData = str$(DataArea.bytDATA(0))
            
        Else
            ChgData = "-1"                      '�l���s��Ȃ�[�P�Z�b�g
            
        End If
            
    Case ID_TYPE.Count  '��
        lngWork = 0                              '������
        For lngloop = 3 To 0 Step -1
            lngWork = lngWork * 256 + DataArea.bytDATA(lngloop)
        Next lngloop
                        
        ChgData = str$(lngWork)
    
    Case ID_TYPE.Date_Type, ID_TYPE.time_type '���t�A����
        ChgData = StrConv(DataArea.bytDATA, vbUnicode)
        
    Case Else
        ChgData = vbNullString
        lngErrsts = IdInfErr.ID_TYPE_MISS
        Exit Function

    End Select
    
    Exit Function
    
ChgDataErr:
        ChgData = vbNullString
        lngErrsts = IdInfErr.PROC_ERR
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfSetJikaiSts
'//  �@�\����  : �����ݒ�t�@�C���X�V����
'//  �@�\�T�v  : �����ݒ�t�@�C���X�V�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iJikaiSts [IN]�ڑ��E�ؒf�^�C�v
'//              Integer�@iGouki  �@[IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-26   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfSetJikaiSts(iJikaiSts As Integer, iGouki As Integer, iJikaiID As Long) As Boolean

    Dim iAreaSts        As Integer          '�����ݒ�t�@�C����Ԓl
    Dim lSts            As Long             '�֐��߂�l
    Dim udtAreaR255     As GATE_INFO        '�Ǎ��ݗp�G���A�i255�ݒ�p�j
    Dim lngSts          As Long             '�q�b�g�G���AID
    Dim lngLoop1        As Long             '�J�E���^�[
    Dim lngHandle       As Long             '�n���h��
    Dim FileName        As String           '�t�@�C���L���`�F�b�N
    Dim lngRet          As Long             '�߂�l
    Dim bRet            As Boolean          '�ǂݍ��݌��ʖ߂�l
    Dim sSetteiFile     As String
    Dim udtAreaR255Work As GATE_INFO        '�Ǎ��ݗp�G���A�i�|�C���^�ړ��p�j
    
    On Error Resume Next
    
    '�����ݒ�t�@�C���L��
    FileName = Dir(G_SETTEI_FILE)
    If FileName = "" Then
       '�����ݒ�t�@�C�����Ȃ��ꍇ
       sSetteiFile = SHOKI_G_SETTEI_FILE
    Else
       '�����ݒ�t�@�C��������ꍇ
       sSetteiFile = G_SETTEI_FILE
    End If
        
    '�����ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�I�[�v���ُ펞�͎Q�ƕs�̂��ߍX�V�ُ�
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        pfSetJikaiSts = False
        Exit Function
    End If
        
    '�����ݒ�t�@�C���ǂݍ���
    For lngLoop1 = 0 To iGouki - 1
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           '�n���h���̃N���[�Y
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
           Call CloseHandle(lngHandle)
           pfSetJikaiSts = False
           Exit Function
        End If
    Next
    
    '�n���h���̃N���[�Y
    Call CloseHandle(lngHandle)
    
    'ID����
    lngSts = SerchId(udtAreaR255, iJikaiID)
    If lngSts >= 0 Then
       'ID���L�����ꍇ
       SetChgData udtAreaR255.GateInfo(lngSts), iJikaiSts   '�f�[�^�ݒ�
    Else
       ' �Y���h�c�����̏ꍇ�X�V�ُ�
        pfSetJikaiSts = False
       Exit Function
    End If
      
    '�����ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(sSetteiFile, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�I�[�v���ُ펞�͎Q�ƕs�̂��ߍX�V�ُ�
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        pfSetJikaiSts = False
        Exit Function
    End If
     
    '�t�@�C���|�C���^�ړ��̂��߂̓ǂݍ���
     For lngLoop1 = 1 To iGouki - 1
         bRet = ReadFile(lngHandle, udtAreaR255Work, LenB(udtAreaR255Work), lngRet, 0)
         If bRet = False Then
            '�n���h���̃N���[�Y
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_READ, 0)
            Call CloseHandle(lngHandle)
            pfSetJikaiSts = False
            Exit Function
         End If
     Next
    
    '�����ݒ�t�@�C���ɏ�������
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       '�n���h���̃N���[�Y
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       Call CloseHandle(lngHandle)
       pfSetJikaiSts = False
       Exit Function
    End If
    
    '�n���h���̃N���[�Y
     Call CloseHandle(lngHandle)

     pfSetJikaiSts = True
     
     Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CONECT_SETTEIFILE_UPDATA_OK, 0)
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : SetChgData
'//  �@�\����  : �f�[�^�ϊ���������
'//  �@�\�T�v  : �f�[�^�ϊ������������s���B
'//
'//              �^        ����        �Ӗ�
'//  ����      : ID_FMT �@DataArea �@[IN]�ϊ����f�[�^
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String�@�@�@        [OUT]�@vbNullstring�ȊO�F����BvbNullString    �F�G���[
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function SetChgData(DataArea As ID_FMT, iSts As Integer)
   
   On Error Resume Next

   DataArea.bytDATA(0) = iSts
  
End Function

'/////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'// �֐�����    : tmrMail2_Timer
'// �@�\����    : �^�C���A�E�g����
'// �@�\�T�v    : �^�C�}�^�C���A�E�g�������s��
'//
'//                   �^          ����            �Ӗ�
'// ����        :
'// �߂�l      :
'//
'// ORIGINAL    :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'// REVISIONS   :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'// ���l        :
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
'        AppActivate frmSystemSetteiMenu.Caption, False ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
'        pfFormActive (frmSystemSetteiMenu.hwnd)        ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
        AppActivate frmRebootTimeSettei.Caption, False  ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
        pfFormActive (frmRebootTimeSettei.hwnd)         ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If

End Sub
