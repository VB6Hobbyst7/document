VERSION 5.00
Begin VB.Form frmSystemDateSettei 
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   7200
      Top             =   6840
   End
   Begin Hoshu.ctlDateSetting ctlDateSetting1 
      Height          =   7000
      Left            =   720
      TabIndex        =   13
      Top             =   1000
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12356
   End
   Begin VB.Timer tmrKakunin 
      Left            =   6960
      Top             =   8400
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "�m  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9870
      Style           =   1  '���̨���
      TabIndex        =   12
      Top             =   6270
      Width           =   1725
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   9000
      Style           =   1  '���̨���
      TabIndex        =   11
      Top             =   6270
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   9000
      Style           =   1  '���̨���
      TabIndex        =   10
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   9870
      Style           =   1  '���̨���
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   10740
      Style           =   1  '���̨���
      TabIndex        =   8
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   9000
      Style           =   1  '���̨���
      TabIndex        =   7
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   9870
      Style           =   1  '���̨���
      TabIndex        =   6
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   10740
      Style           =   1  '���̨���
      TabIndex        =   5
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   9000
      Style           =   1  '���̨���
      TabIndex        =   4
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   9870
      Style           =   1  '���̨���
      TabIndex        =   3
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton cmdTenkey 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   10740
      Style           =   1  '���̨���
      TabIndex        =   2
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton cmdModoru_Menu 
      Caption         =   "     �V�X�e���ݒ�     ��ʂ֖߂�"
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
      Left            =   8760
      Style           =   1  '���̨���
      TabIndex        =   1
      Top             =   7800
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�V�X�e�����t�ݒ�"
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
      TabIndex        =   0
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmSystemDateSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////
'//   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �t�@�C����     : frmSystemDateSettei
'//  �p�b�P�[�W��   : �V�X�e�����t�ݒ���
'//  �T�v           : �V�X�e�����t�ݒ��ʂ̏������`����B
'//
'//  ORIGINAL       :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                   EG20�t�F�[�Y�Q�Ή�
'//                   EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS      :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                   2014�N�x�{�� �yEG20_KANSI05_01�z
'//  REVISIONS      :(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l           :
'//////////////////////////////////////////////////////////////////////////////

Option Explicit

Private intPos As Integer       '(0:���I���E1:�N�E2:���E3:���E4:���E5:���E6:�b)�������͍��ڂ̃J�����g�ʒu
Private gintCtrlIndex As Integer

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

Private Sub ctlDateSetting1_BtnOuka(intIndex As Integer)
    
    On Error Resume Next
    gintCtrlIndex = intIndex

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  �֐�����     : Form_Load
'/  �@�\����     : Form_Load������
'/  �@�\�T�v     : Form_Load���������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20�t�F�[�Y�Q�Ή�
'/             EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'/ REVISIONS :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next

    Dim strDateTime As String       '���ݓ����ݒ�p

    '�u�V�X�e�����t�ݒ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

' EG20 V6.8.0.1 ADD START
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
' EG20 V6.8.0.1 ADD END

    ' �R���g���[���̕ۑ��C���f�b�N�X��������
    gintCtrlIndex = -1
    
    ctlDateSetting1.psInitialize
    
    ' �����ݒ�t�R���g���[���ɏ����l��ݒ肷��B
    ctlDateSetting1.TotalArea = Format$(Now, "yyyymmddhhmmss")
    
    ' �ݒ肵�����e���R���g���[����ɕ\������B
    ctlDateSetting1.DisplaySetUp
    ' �R���g���[����\������B
    ctlDateSetting1.Enable = 0
    ctlDateSetting1.Visible = True
    
    ' �m��{�^���������s��
    cmdKakutei.Enabled = False
    
End Sub


'/////////////////////////////////////////////////////////////////////////////
'/   (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  �֐�����     : cmdTenkey_Click
'/  �@�\����     : �e���L�[����������
'/  �@�\�T�v     : �e���L�[���������ꂽ�������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20�t�F�[�Y�Q�Ή�
'/             EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdTenkey_Click(Index As Integer)

    On Error Resume Next
    
    '�u�V�X�e�����t�ݒ��ʁF�e���L�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_TENKEY_BUTTON, 0)

    
    ' ���Ƃ����Ƃ��̃{�^������������Ă��Ȃ��ꍇ�͉������Ȃ�    ' rev 02.15
    If (gintCtrlIndex = -1) Then
        Exit Sub
    End If
    
    '�m�F�{�^�����d�����������ɂ���B
    cmdKakutei.Enabled = True
    
    '�e���L�[�R���g���[�����A�p�����[�^�Ƃ��Ď󂯎�������͒l���A
    '�����ݒ�R���g���[���̌ʓ��͒l�G���A�v���p�e�B�֐ݒ肷��B
    ctlDateSetting1.InputArea = CStr(Index)
       
    '�����ݒ�{�^���R���g���[���̓��͒l�\���������\�b�h���s���B
    ctlDateSetting1.DisplayInput

End Sub
'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  �֐�����     : cmdKakutei_Click
'/  �@�\����     : �m��{�^������������
'/  �@�\�T�v     : �m��{�^�����������ꂽ�������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20�t�F�[�Y�Q�Ή�
'/             EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdKakutei_Click()
    
    Dim i As Integer
    Dim udtSendData         As ML_KYOTU_INF     ' ���ʃG���A
    Dim lngSendSize         As Long             ' ���M���郁�[���T�C�Y
    Dim lngErrCode          As Long             ' �G���[�R�[�h
    Dim bRet                As Boolean          ' ���[�����M�����߂�l
    Dim iResponse           As Integer          ' ���b�Z�[�W�{�b�N�X�߂�l
    
    Dim strDate             As String
    
    '�u�V�X�e�����t�ݒ��ʁF�m��t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_KAKUTEI_BUTTON, 0)
    
    ' ���t�`�F�b�N
    ctlDateSetting1.InputCheck
    
    ' �������ꍇ
    If ctlDateSetting1.TotalArea <> -1 Then
        '�m�F�{�^�������p�^�C�}���쓮������
        tmrKakunin.Interval = MN_MAIL_INTERVAL       '�{�^�������p�^�C�}���Ԑݒ�
        tmrKakunin.Enabled = True
        
        strDate = ctlDateSetting1.TotalArea
        ctlDateSetting1.Enable = 1
        For i = 0 To cmdTenkey.Count - 1
            cmdTenkey(i).Enabled = False
        Next
        cmdKakutei.Enabled = False
        cmdModoru_Menu.Enabled = False

        ' �����ݒ�R���g���[���̃g�[�^�����͒l�G���A�̒l�ŃV�X�e���������X�V����B
        Date = Mid(strDate, 1, 4) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2)
        Time = Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2)

        ' �Ď��Ղ����삵�Ă��Ȃ��ꍇ�̓��[�����M���s��Ȃ�
        If CheckAppStart(PROC_KANRI) <> 0 Then

            ' ���[���̑��M���e��ҏW����
            udtSendData.udtlHeader.dwId = ML_ID_DATE_SET_ORD       '���[���h�c�@=�h"�����ݒ�w��"
            udtSendData.udtlHeader.dwSize = MlSize.DATE_SET_ORD    '���[���T�C�Y=�h"�����ݒ�w��"
            udtSendData.udtlHeader.dwProid = RHOSHU_ID             '���M���v���Z�X�h�c=�h�ێ�h
            udtSendData.udtlHeader.dwSubArea = 0                   '�⏕���@=�@0

            ' ���M�T�C�Y��ݒ肷��B
            lngSendSize = udtSendData.udtlHeader.dwSize
                
            ' �ă}�ɑ΂��āA�����ݒ�w�����[���𑗐M����B
            bRet = DssSendMail(MAIL_SLOT_KANMA, lngSendSize, udtSendData.udtlHeader)
            ' ���[���𐳏�ɑ��M�������̃��O
            If bRet = False Then
                '�u��ʕ\���v�����[�����M�ُ�v���O�o��
                lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
                Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, DATESETORDER_REQ_SEND, lngErrCode)
            Else
                '�u��ʕ\���v�����[�����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, DATESETORDER_REQ_SEND, 0)
            End If
        End If

        ' �ۑ��G���A���ēx������                ' rev 02.15
        gintCtrlIndex = -1

    ' �s���ȏꍇ
    Else

        iResponse = MsgBox("���͂����l�͕s���ł��B", _
                           (vbOKOnly + vbExclamation), _
                           "���ُ͈�")
    End If

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  �֐�����     : cmdModoru_Menu_Click
'/  �@�\����     : ���j���[�ɖ߂�{�^����������
'/  �@�\�T�v     : ���j���[�ɖ߂�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/ ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/             EG20�t�F�[�Y�Q�Ή�
'/             EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Menu_Click()

    On Error Resume Next
    
    '�u�V�X�e���ݒ胁�j���[��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, SYSTEM_DATE_SETTEI_GAMEN_END, 0)
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���������j���[���(�A�N�e�B�u��)
'//  �@�\�T�v  : ��ʍĕ\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//              EG20�t�F�[�Y�Q�Ή�
'//              EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    tmrMail.Enabled = True         ' EG20 V6.8.0.1 ADD
    
    pfFormActive (hwnd)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ���������j���[���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//              EG20�t�F�[�Y�Q�Ή�
'//              EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    '�^�C�}���~����B
    tmrKakunin.Enabled = False

    tmrMail.Enabled = False         ' EG20 V6.8.0.1 ADD
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'/
'/  �֐�����     : TmrKakunin_Timer
'/  �@�\����     : �m�F�{�^�������p�^�C�}�C�x���g������
'/  �@�\�T�v     : �m�F�{�^�������p�^�C�}�C�x���g�������̏������s���B
'/                 �m�F�{�^���A���̑��{�^���̐F�������F���猳�̐F�ɖ߂�
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'/                 EG20�t�F�[�Y�Q�Ή�
'/                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'/ REVISIONS     :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrKakunin_Timer()

    Dim i As Integer
    Dim blnRet As Boolean
    Dim intCount As Integer
    
    On Error Resume Next
    
    '�m�F�{�^�������p�^�C�}���~����
    tmrKakunin.Enabled = False                   '�m�F�{�^�������p�^�C�}��~
    tmrKakunin.Interval = 0                      '�m�F�{�^�������p���ԏ�����
        
    ctlDateSetting1.Enable = 0
        
    '�e���L�[�������ɂ���
    For i = 0 To cmdTenkey.Count - 1
        cmdTenkey(i).Enabled = True
    Next
        
    '�߂�{�^���������ɂ���
    cmdModoru_Menu.Enabled = True
    
    '�m��{�^���������ɂ���
    cmdKakutei.Enabled = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z

'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
'        AppActivate frmLogMenu.Caption, False          ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
'        pfFormActive (frmLogMenu.hwnd)                 ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
        AppActivate frmSystemDateSettei.Caption, False  ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
        pfFormActive (frmSystemDateSettei.hwnd)         ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If
End Sub

