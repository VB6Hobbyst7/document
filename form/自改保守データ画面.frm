VERSION 5.00
Begin VB.Form frmGateHoshu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�����ێ�f�[�^"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
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
      Caption         =   "�ύX�O�f�[�^�ۑ�"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�����ێ�r�v�ݒ�\��(��)"
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
      Left            =   6360
      TabIndex        =   10
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�W���[�i����"
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
      Left            =   6360
      TabIndex        =   9
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�����ێ�r�v�ݒ�N���A"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�����ێ�r�v�ݒ�\��(��)"
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
      Caption         =   "�}�X�^�f�[�^���e�\��"
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
      Caption         =   "�}�X�^�f�[�^����"
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
      Caption         =   "�h�b�����e�i���X"
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
   Begin VB.Timer tmrMail 
      Left            =   3960
      Top             =   7680
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�����Ď��Ւ���"
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
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�ғ��E�����e�f�[�^���W"
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
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�f�[�^���W�E�o��"
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
Attribute VB_Name = "frmGateHoshu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmGateHoshu.frm
'//  �p�b�P�[�W���F�����ێ�f�[�^���
'//
'//  �T�v�F�����ێ�f�[�^���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�����ێ�f�[�^�N���A��ʕ\�������ǉ�
'//     REVISIONS :(1.6.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-12   CODED   BY [TCC] M.Matsumoto
'//       �y��-279,281�Ή��z
'//     REVISIONS :(EG20 V7.2.0.1) 2013-06-14  CODED   BY [TCC] T.Nakajima
'//        2013�N�x�{�� ���u�Ή�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-17  CODED   BY [TCC] T.Nakajima
'//        �k���V�����t�F�[�Y�Q�Ή�
'//         �yHKRK_Kansi07_005_01�z
'//     REVISIONS :(EG20 V32.1.0.1) 2016-06-07  CODED   BY [TCC] T.Nakajima
'//        2016�N�x�{���Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
'Private sHyoujiGoukiNo(0 To 18) As String        '�\�����@�ԍ��i�[�G���A           ' EG20 V6.9.0.1�폜
Private sHyoujiGoukiNo(0 To 31) As String         '�\�����@�ԍ��i�[�G���A           ' EG20 V6.9.0.1�ǉ�
Private Const TITLENAME_CORNER = "�R�[�i#"        ' �R�[�i��                        ' EG20 V6.9.0.1�ǉ�
Private sRonriCornerNo(0 To 31) As String         '�_���R�[�i�ԍ��i�[�G���A         ' EG20 V6.9.0.1�ǉ�
Private Const DEFAILT_HYOUJI_UMU = 1    '�u�ғ��E�����e�f�[�^���W�v�t�̃f�t�H���g�\��     'V2.7.0.1 ADD
Private iToolFlg                As Integer        ' ConfigViewer ����or�ݗ��t���O      ' EG20 V30.3.0.1 �yHKRK_Kansi07_005_01�zADD
Private Const CONFIG_VIEWER_ZAIRAI = 0            '�u�����ێ�SW�ݒ�\��(��)�v�t����    ' EG20 V30.3.0.1 �yHKRK_Kansi07_005_01�zADD
Private Const CONFIG_VIEWER_KANSEN = 1            '�u�����ێ�SW�ݒ�\��(��)�v�t����    ' EG20 V30.3.0.1 �yHKRK_Kansi07_005_01�zADD

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �����ێ�f�[�^���(�A�N�e�B�u��)
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
'//  �@�\����  : �����ێ�f�[�^���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : �����ێ�f�[�^���(���[�h��)
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
'//     REVISIONS :(2.7.0.1) 2010-12-24   CODED   BY [TCC] M.Kuroki
'//       �E�i�q���C�@�m�d�f���d�f�q�R���o�[�g�Ή�
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-12   CODED   BY [TCC] M.Matsumoto
'//       �y��-279,281�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
   Dim lSts As Long             '�֐��߂�l      'V2.7.0.1 ADD
    
    On Error Resume Next
    
    lSts = 0    '�ϐ��̏����� 'V2.7.0.1 ADD

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '�u�����ێ�f�[�^��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_GAMEN_START, 0)
    
   'V1.3.0.1 ADD START
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END

    'V2.7.0.1  ADD START
    'HOSHU.INI���A�u�ғ��E�����e�f�[�^���W�v�t�̕\���L�����擾����B
    lSts = GetPrivateProfileInt(KANSI_HOSHU_DATA_SEC, _
                                   KANSI_HOSHU_DATA_KEY, _
                                   DEFAILT_HYOUJI_UMU, _
                                   HOSHU_FILE)
    If lSts = 1 Then
        cmdFixedExe(0).Visible = True
    Else
        cmdFixedExe(0).Visible = False
    End If
    'V2.7.0.1  ADD END
    
    'EG20 V2.1.0.1 ADD START �y��-279,281�Ή��z
    '�Ď��Ֆ��N�����͈ꕔ�{�^���������s�Ƃ���
    If CheckAppStart(PROC_KANRI) = 0 Then
        cmdFixedExe(1).Enabled = False          '���ߐ؂�
        cmdFixedExe(3).Enabled = False          '�}�X�^�f�[�^����
        cmdFixedExe(4).Enabled = False          '�}�X�^�f�[�^���e�\��
        cmdFixedExe(7).Enabled = False          '�W���[�i����  EG20 V7.2.0.1
    End If
    'EG20 V2.1.0.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �e�t��������
'//  �@�\�T�v  : �e�t���̏������s���B
'//              �u�ғ��E�����e�f�[�^���W�v�u�����ێ�SW�ݒ�\���v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�����ێ�f�[�^�N���A��ʕ\�������ǉ�
'//     REVISIONS :(1.6.0.1) 2009-03-24   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-23   REVISED BY [TCC] D.Yamashita
'//                 �t�F�[�Y�R�c�����ڑΉ��@�����ێ�SW�ݒ�\�����Ɏ��w.GLT�쐬������ǉ�
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-22   REVISED BY [TCC] T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Ή��y�c��54�z
'//                �E�}�X�^�f�[�^���e�\�������ǉ�
'//                �E�����ێ�SW�ݒ�\���t�A�����ێ�r�v�ݒ�N���A�t�̕���
'//     REVISIONS :(EG20 V7.2.0.1) 2013-06-14  REVISED BY [TCC] T.Nakajima
'//                2013�N�x�{�� ���u�Ή�
'//                �E�W���[�i���󎚖t�ǉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-17  REVISED BY [TCC] T.Nakajima
'//                �k���V�����t�F�[�Y�Q�Ή�
'//                �yHKRK_Kansi07_005_01�z�����ێ�SW�ݒ�\���iConfigVirewer)���ݍ��ݑΉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
 Dim iResponse As Integer 'MsgBox�̖߂�l
 Dim lngErrCode As Long   '�G���[�R�[�h
' EG20 V2.0.1.1�y�c��54�zADD START
 Dim lRetVal As Double                    'Shell�֐��߂�l
' EG20 V2.0.1.1�y�c��54�zADD END
 
 On Error Resume Next
  
  Select Case Index
        Case 0                                 '�ғ��E�����e�f�[�^���W���
            '�u�ғ��E�����e�f�[�^���W��ʁF�ғ��E�����e�f�[�^���W�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_KADO_MENTE_BUTTOM, 0)
            Load frmSyusyu
            frmSyusyu.Show 1
        Case 1                                 '�����ێ�SW�ݒ�(�ݒ�R���t�B�O�m�F�c�[���N��)
        'EG20 V2.1.0.1 DEL START
'            '�u�ғ��E�����e�f�[�^���W��ʁF�����ێ�SW�ݒ�\���t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SWSETTEI_BUTTOM, 0)
'
'            'V1.11.0.1 ADD START
'            'GLT�t�@�C�����쐬���A���e���X�V����B
'            fMakeGLTFile
'            'V1.11.0.1 ADD END
'            '�����ێ�SW�f�[�^�t�@�C���R�s�[����
'            fGetGouki '�\�����@�擾
'            'If sSWFileCopy > 0 Then 'V1.6.0.1 DEL
'            sSWFileCopy 'V1.6.0.1 ADD
'                '�R���t�B�O�ݒ�m�F�c�[���N������
'                sToolOn
'            'V1.6.0.1 DEL START
'            'Else
'            '  '�u�����ێ�f�[�^��ʁF�����ێ�SW�f�[�^�t�@�C���R�s�[�ُ�v���O�o��
'            '  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'            '  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'            'End If
'            'V1.6.0.1 DEL END
'        'V1.4.0.1�@ADD START
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START
            '�u�ғ��E�����e�f�[�^���W��ʁF�����Ď��Ւ��ؖt�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SHIMEKIRI_BUTTOM, 0)
            Load frmShimekiriData
            frmShimekiriData.Show 1
        'EG20 V2.1.0.1 ADD END
        Case 2                                 '�����ێ�r�v�ݒ�N���A���
        'EG20 V2.1.0.1 DEL START
'            '�u�ғ��E�����e�f�[�^���W��ʁF�����ێ�r�v�ݒ�N���A�t�����v���O�o��
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SETTEICLEAR_BUTTOM, 0)
'            Load frmHoshuSwClear
'            frmHoshuSwClear.Show 1
'        'V1.4.0.1�@ADD END
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.0.1.1�y�c����54�z ADD START
            '��ʕ\���v���i��ԊĎ��@�\�ݒ�j��ID����ɑ��M����
            If (SendMessageDispInfo(ML_DT_IC_MAINTE) = False) Then
         
                iResponse = MsgBox("�h�b�����e�i���X�t�A��`�G���[�B" & _
                                   Chr(vbKeyReturn) & _
                                   "�h�b�����e�i���X��ʂ��N���ł��܂���B", _
                                   vbOKOnly, _
                                   "��ʋN���G���[")
            End If
        'EG20 V2.0.1.1�y�c����54�z ADD END
        'EG20 V2.1.0.1 ADD START
        Case 3
            '�u�ғ��E�����e�f�[�^���W��ʁF�}�X�^�f�[�^���͖t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_MST_INPUT_BUTTOM, 0)
            Load frmInputMstData
            frmInputMstData.Show 1
        'EG20 V2.1.0.1 ADD END
        'EG20 V2.0.1.1�y�c��54�zADD START
        Case 4
            '�u�}�X�^�f�[�^���e�\���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_MST_DISP_BUTTOM, 0)
            ' �}�X�^�f�[�^���e�\���c�[���N��
            lRetVal = Shell("D:\KANSI\TOOL\DataViewer\BinViewer.exe", vbNormalFocus)
        
        Case 5
            '�u�ғ��E�����e�f�[�^���W��ʁF�����ێ�SW�ݒ�\���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SWSETTEI_BUTTOM, 0)
            
            iToolFlg = CONFIG_VIEWER_ZAIRAI         '�ݗ��p��ConfigViewer���N�����ꂽ   EG20 V30.3.0.1�yHKRK_Kansi07_005_01�z ADD

            'V1.11.0.1 ADD START
            'GLT�t�@�C�����쐬���A���e���X�V����B
            fMakeGLTFile
            'V1.11.0.1 ADD END
            '�����ێ�SW�f�[�^�t�@�C���R�s�[����
            fGetGouki '�\�����@�擾
            'If sSWFileCopy > 0 Then 'V1.6.0.1 DEL
            sSWFileCopy 'V1.6.0.1 ADD
                '�R���t�B�O�ݒ�m�F�c�[���N������
                sToolOn
            'V1.6.0.1 DEL START
            'Else
            '  '�u�����ێ�f�[�^��ʁF�����ێ�SW�f�[�^�t�@�C���R�s�[�ُ�v���O�o��
            '  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            '  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
            'End If
            'V1.6.0.1 DEL END
        'V1.4.0.1�@ADD START
        Case 6
            '�u�ғ��E�����e�f�[�^���W��ʁF�����ێ�r�v�ݒ�N���A�t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SETTEICLEAR_BUTTOM, 0)
            Load frmHoshuSwClear
            frmHoshuSwClear.Show 1
        'EG20 V2.0.1.1�y�c��54�zADD START
        'EG20 V7.2.0.1 ADD START
        Case 7
            '�u�ғ��E�����e�f�[�^���W��ʁF�W���[�i���󎚖t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_JPR_PRINT_BUTTON, 0)
            Load frmJprPrint
            frmJprPrint.Show 1
        'EG20 V7.2.0.1 ADD END
        'EG20 V30.3.0.1 �yHKRK_Kansi07_005_01�zADD START
        Case 8
            '�u�ғ��E�����e�f�[�^���W��ʁF�����ێ�SW�ݒ�(��)�\���t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SWSETTEI_KAN_BUTTOM, 0)
            
            iToolFlg = CONFIG_VIEWER_KANSEN         '�����p��ConfigViewer���N�����ꂽ

            'GLT�t�@�C�����쐬���A���e���X�V����B
            fMakeGLTFile
            '�����ێ�SW�f�[�^�t�@�C���R�s�[����
            fGetGouki '�\�����@�擾
            
            sSWFileCopy 'V1.6.0.1 ADD
                '�R���t�B�O�ݒ�m�F�c�[���N������
                sToolOn
        'EG20 V30.3.0.1 �yHKRK_Kansi07_005_01�zADD END
        'EG30 V32.1.0.1 ADD START
        Case 9
            '�u�ғ��E�����e�f�[�^���W��ʁF�ύX�O�f�[�^�ۑ��t�����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_SET_BEF_BUTTON, 0)
            Load frmSetteiBefore
            frmSetteiBefore.Show 1
        'EG30 V32.1.0.1 ADD END
        
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
    
    '�u�����ێ�f�[�^��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, GATE_HOSHU_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSWFileCopy
'//  �@�\����  : �����ێ�SW�ݒ�f�[�^�t�@�C���쐬����
'//  �@�\�T�v  : �����ێ�SW�ݒ�f�[�^���A�����ێ�SW�f�[�^�t�@�C����
'//              �R�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.9.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sSWFileCopy() As Integer

     Dim iCnt As Integer                     '�J�E���^�[
     Dim sSWDataPath As String               '�����ێ�SW�f�[�^�t�@�C��
     Dim sMyPath As String                   '�����ێ�SW�ݒ�f�[�^
     
     On Error Resume Next
   
     sSWFileCopy = 0                         '�t�@�C�����ݐ�
    
    '�����ő吔�����[�v����B
    For iCnt = 1 To MAX_GATE_NO
     '�uGATE_SW##.dat�v�́u##�v��01�`16�ɕϊ�����B
     sMyPath = Replace(GATE_SW_FILE, "##", Format(iCnt, "0#"))
     '�����ێ�SW�ݒ�f�[�^�̌������s���B
     If Dir(sMyPath) <> "" Then
        '�����ێ�SW�f�[�^�t�@�C���̃p�X���쐬����B
        sSWDataPath = PATH_RMENTE_GATE_DEN_JIEKI_GOUKI
' EG20 V6.9.0.1�ǉ��J�n
        '�u�R�[�i$�v�́u$�v��1�`6�ɕϊ�����B
        sSWDataPath = Replace(sSWDataPath, "$", sRonriCornerNo(iCnt - 1))
' EG20 V6.9.0.1�ǉ��I��
        '�u##���@�v�́u##�v��01�`16�ɕϊ�����B
        sSWDataPath = Replace(sSWDataPath, "##", Format(sHyoujiGoukiNo(iCnt - 1), "0#"))
        '�t�H���_�쐬
        MkDir sSWDataPath
        sSWDataPath = sSWDataPath & TOOL_SW_File
        
        '�����ێ�SW�f�[�^�������ێ�SW�f�[�^�t�@�C���ɃR�s�[����B
        FileCopy sMyPath, sSWDataPath
        sSWFileCopy = sSWFileCopy + 1
     End If
   Next

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sToolOn
'//  �@�\����  : �����ێ�SW�ݒ�c�[���N������
'//  �@�\�T�v  : �����ێ�SW�ݒ�c�[���̋N�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-17  REVISED BY [TCC] T.Nakajima
'//                �k���V�����t�F�[�Y�Q�Ή�
'//                �yHKRK_Kansi07_005_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sToolOn() As Integer
    Dim lRetVal As Double                    'Shell�֐��߂�l
    Dim sToolName As String * MAX_PATH_SIZE  '�c�[���p�X��
    Dim lSize As Long                        '�߂�l
    
    On Error Resume Next
   
    '�ێ�@�\INI�t�@�C������A�����ێ�SW�ݒ�c�[���p�X���擾����B
    'EG20 V30.3.0.1 DEL START�yHKRK_Kansi07_005_01�z
'    lSize = GetPrivateProfileString(KANSI_HOSHU_SW_TOOL_SEC, _
'                                    KANSI_HOSHU_SW_TOOL_KEY, _
'                                    DEFAILT, sToolName, Len(sToolName), HOSHU_FILE)
    'EG20 V30.3.0.1 DEL END�yHKRK_Kansi07_005_01�z
    'EG20 V30.3.0.1 ADD START �yHKRK_Kansi07_005_01�z
    '(��)�A(��)�ǂ���̖t����������Ă��邩�H
    If iToolFlg = CONFIG_VIEWER_KANSEN Then
        ' (��)����������Ă����ConfigViewer2�̃p�X���擾
        lSize = GetPrivateProfileString(KANSI_HOSHU_SW_TOOL_SEC, _
                                        KANSI_HOSHU_SW_TOOL_KEY_KAN, _
                                        DEFAILT, sToolName, Len(sToolName), HOSHU_FILE)
    Else
        '(��)����������Ă���̂�ConfigViewer�̃p�X���擾
        lSize = GetPrivateProfileString(KANSI_HOSHU_SW_TOOL_SEC, _
                                        KANSI_HOSHU_SW_TOOL_KEY, _
                                        DEFAILT, sToolName, Len(sToolName), HOSHU_FILE)
    End If
    'EG20 V30.3.0.1 ADD END �yHKRK_Kansi07_005_01�z
    
    'INI�t�@�C���ɁA�Y���s�̒�`������ꍇ�A
    If sToolName <> "" Then
        '�����ێ�SW�ݒ�c�[�����N������B
        lRetVal = Shell(sToolName, vbNormalFocus)
    End If
 
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeGLTFile
'//  �@�\����  : ���w.GLT�t�@�C���ւ̎��������������ݏ���
'//  �@�\�T�v  : GATE.INI���Q�Ƃ��A���w.GLT�t�@�C���ցA
'//              ���@�ԍ��A�\�������AIP�A�h���X���������ށB
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
Private Function fGetGouki() As Integer
    Dim lngRet As Long          '�֐��̕Ԃ�l
    Dim iGate As Integer        '����INDEX
    Dim j As Integer            '���[�NINDEX
    Dim sGoukiNo As String      'GLT�t�@�C�����R�[�h�f�[�^(���@�ԍ��\������)
    Dim cWork As Byte           '���[�N�G���A
    Dim lngErrCode As Long      '�G���[�R�[�h
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim intGLTFileNo As Integer     '̧�ٔԍ�

    For iGate = CNT_MIN To MAX_GATE_NO - 1
      '�������D�@���擾
      sKeyName = "gate" & Format(iGate + 1, "00")
      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                     sKeyName, _
                                     DEFAILT, sGateData, Len(sGateData), _
                                     PATH_GATE_FILE)
      If iRet = 0 Then
         '�u�Ӱ�����ݽ��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
         Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
         Exit Function
      End If
        
      If Len(sGateData) <> 0 Then
         '�f�[�^�̎擾
         ReDim sFData(15)
         iFCnt = 1
            
         For iFLoop = 1 To Len(sGateData)
             If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
                iFLoop2 = iFLoop
                Do
                  iFLoop2 = iFLoop2 + 1
                  If iFLoop2 > Len(sGateData) Then
                     sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                     iFCnt = iFCnt + 1
                     If iFCnt >= 16 Then
                         Exit For
                     End If
                     
                     iFLoop = iFLoop2
                     Exit Do
                  End If
                       
                  If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
                     sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                     iFCnt = iFCnt + 1
                     If iFCnt >= 16 Then
                           Exit For
                     End If
                     
                     iFLoop = iFLoop2
                     Exit Do
                  End If
                 Loop
             End If
         Next
      End If
      
      If Len(Trim(sFData(1))) = 1 Then
         '���@�ԍ����P���Ȃ�΁A�擪�ɂO��t������B
         sGoukiNo = "0" & Trim(sFData(1))
      Else
         sGoukiNo = Trim(sFData(1))
      End If
        
      sHyoujiGoukiNo(iGate) = sGoukiNo
' EG20 V6.9.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
      sRonriCornerNo(iGate) = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.9.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��I��

    Next
    
    fGetGouki = 0
End Function

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
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmGateHoshu.Caption, False
        pfFormActive (frmGateHoshu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
'V1.11.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  �֐�����  : fMakeGLTFile
'//  �@�\����  : ���w.GLT�t�@�C���ւ̎��������������ݏ���
'//  �@�\�T�v  : GATE.INI���Q�Ƃ��A���w.GLT�t�@�C���ցA
'//              ���@�ԍ��A�\�������AIP�A�h���X���������ށB
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.11.0.1) 2009-12-23   CODED   BY [TCC] D.Yamashita
'//                 �t�F�[�Y�R�c�����ڑΉ��@�����ێ�SW�ݒ�\�����Ɏ��w.GLT�쐬������ǉ�
'//     REVISIONS :(EG20 V6.7.0.1)  2012-06-28  CODED BY  [TCC] H.Sugimoto
'//                 �y���ڃ`�F�b�N�̑Ώۂ����D�@���݂̂Ƃ���C���z
'//     REVISIONS :(EG20 V30.3.0.1)  2014-09-18  CODED BY  [TCC] T.Nakajima
'//                  �k���V�����t�F�[�Y�Q�Ή�
'//                 �yHKRK_Kansi07_005_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeGLTFile() As Integer
    Dim lngRet As Long          '�֐��̕Ԃ�l
    Dim iGate As Integer        '����INDEX
    Dim j As Integer            '���[�NINDEX
    Dim sGoukiNo As String      'GLT�t�@�C�����R�[�h�f�[�^(���@�ԍ��\������)
    Dim cWork As Byte           '���[�N�G���A
    Dim lngErrCode As Long      '�G���[�R�[�h
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim intGLTFileNo As Integer     '̧�ٔԍ�
    Dim szCorner As String      ' �R�[�i�ԍ�
    Dim szTitleName As String                       ' �^�C�g����                    ' EG20 V6.7.0.1�ǉ�
    Dim fso As New FileSystemObject                 '�t�@�C���V�X�e���I�u�W�F�N�g   ' EG20 V6.7.0.1�ǉ�

    On Error Resume Next
    MkDir PATH_RMENTE_GATE_DEN   '�����p�d�S�t�H���_���쐬����B�iGLT�t�@�C���p�j
    
    ' EG20 V30.3.0.1 ADD START �yHKRK_Kansi07_005_01�z
    ' �e�R�[�i�̃R�[�i��ʂ��擾
    gsGetCornerType
    ' EG20 V30.3.0.1 ADD END �yHKRK_Kansi07_005_01�z
    
' EG20 V6.7.0.1�ǉ��J�n
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH_RMENTE_GATE_DEN_JIEKI) = False Then
        '�R�s�[��t�H���_�쐬
        fso.CreateFolder (PATH_RMENTE_GATE_DEN_JIEKI)
    End If
    Set fso = Nothing
' EG20 V6.7.0.1�ǉ��I��
    
    'GLT�t�@�C�����J���B�t�@�C�������݂��Ȃ���ΐV�K�ɍ쐬�����B
    On Error GoTo ErrorHandlerGLTFile
    intGLTFileNo = FreeFile        ' ���g�p�̃t�@�C���ԍ����擾����B
    Open GATE_GLT_FILE_FULLPASS For Output As #intGLTFileNo     ' GLT�t�@�C�����J���B

    For iGate = CNT_MIN To MAX_GATE_NO - 1
      '�������D�@���擾
      sKeyName = "gate" & Format(iGate + 1, "00")
      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                     sKeyName, _
                                     DEFAILT, sGateData, Len(sGateData), _
                                     PATH_GATE_FILE)
      If iRet = 0 Then
         '�u�����ێ�f�[�^��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
         Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
         Exit Function
      End If
        
      If Len(sGateData) <> 0 Then
         '�f�[�^�̎擾
         ReDim sFData(15)
         iFCnt = 1
            
         For iFLoop = 1 To Len(sGateData)
             If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
                iFLoop2 = iFLoop
                Do
                  iFLoop2 = iFLoop2 + 1
                  If iFLoop2 > Len(sGateData) Then
                     sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                     iFCnt = iFCnt + 1
                     If iFCnt >= 16 Then
                         Exit For
                     End If
                     
                     iFLoop = iFLoop2
                     Exit Do
                  End If
                       
                  If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
                     sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                     iFCnt = iFCnt + 1
                     If iFCnt >= 16 Then
                           Exit For
                     End If
                     
                     iFLoop = iFLoop2
                     Exit Do
                  End If
                 Loop
             End If
         Next
      End If
      
      If Len(Trim(sFData(1))) = 1 Then
         '���@�ԍ����P���Ȃ�΁A�擪�ɂO��t������B
'         sGoukiNo = "0" & Trim(sFData(1)) & "���@"                                 ' EG20 V6.7.0.1�폜
         sGoukiNo = "0" & Trim(sFData(1))                                           ' EG20 V6.7.0.1�ǉ�
      Else
'         sGoukiNo = Trim(sFData(1)) & "���@"                                       ' EG20 V6.7.0.1�폜
         sGoukiNo = Trim(sFData(1))                                                 ' EG20 V6.7.0.1�ǉ�
      End If
        
' EG20 V6.9.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
'      szCorner = Replace(TITLENAME_CORNER, "#", Trim(sFData(GATE_IDX.IDX_RONRI_CORNER)))   ' EG20 V6.7.0.1�폜
      szCorner = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))                                    ' EG20 V6.7.0.1�ǉ�
      sRonriCornerNo(iGate) = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.9.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��I��
' EG20 V6.7.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
      ' �R�[�i�ԍ��ϊ�
      szTitleName = Replace(RMENTE_GOKITITLENAME, "$", szCorner)
      ' ���@�ԍ��ϊ�
      szTitleName = Replace(szTitleName, "##", sGoukiNo)
' EG20 V6.7.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
        
      If Trim(sFData(4)) <> "��" Then
         'Gate.ini�t�@�C���̍��@�ԍ��\�������AIP�A�h���X��GLT�t�@�C���ɏ������ށB
'         Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(5))                     ' EG20 V6.6.0.1�폜
'          Print #intGLTFileNo, szCorner & "_" & sGoukiNo & "," & Trim(sFData(5))   ' EG20 V6.6.0.1�ǉ�     ' EG20 V6.7.0.1�폜
         'EG20 V30.3.0.1 DEL START �yHKRK_Kansi07_005_01�z
         'Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))                   ' EG20 V6.7.0.1�ǉ�
         'EG20 V30.3.0.1 DEL END �yHKRK_Kansi07_005_01�z
         
         'EG20 V30.3.0.1 ADD START �yHKRK_Kansi07_005_01�z
         '���ݏ������̍��@��������_���R�[�i�̎�ʂ́H
         
         'ConfigViewer��ConfigViewr2�ǂ�����N������̂��H
         If iToolFlg = CONFIG_VIEWER_KANSEN Then
            '(��)�t����������Ă���̂Ŋ����R�[�i�̍��@�̂�GLT�t�@�C���ɍX�V����B
            If gintCornerType(CInt(szCorner) - 1) = CORNER_TYPE_KANSEN Then
                Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))
            Else
                '�����ɓ��Ă͂܂�Ȃ��ꍇ�͉���GLT�t�@�C���ɂ͓���Ȃ��B
            End If
        Else
            '(��)�t����������Ă���̂ōݗ��R�[�i�̍��@�̂�GLT�t�@�C���ɍX�V����B
            If gintCornerType(CInt(szCorner) - 1) = CORNER_TYPE_ZAIRAI Then
                Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))
            Else
                '�����ɓ��Ă͂܂�Ȃ��ꍇ�͉���GLT�t�@�C���ɂ͓���Ȃ��B
            End If
        End If
         'EG20 V30.3.0.1 ADD END �yHKRK_Kansi07_005_01�z
      End If
      
      '�\�����@�ԍ�
      If Len(Trim(sFData(1))) = 1 Then
         '���@�ԍ����P���Ȃ�΁A�擪�ɂO��t������B
         sHyoujiGoukiNo(iGate) = "0" & Trim(sFData(1))
      Else
         sHyoujiGoukiNo(iGate) = Trim(sFData(1))
      End If
    
    Next
    
    'GLT�t�@�C�������B
    Close #intGLTFileNo
    
    fMakeGLTFile = 0    '����I��
    Exit Function

ErrorHandlerGateIni:
   '�u�����ێ�f�[�^��ʁF�t�@�C���A�N�Z�X�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 1
   'GLT�t�@�C�������B
   Close #intGLTFileNo
   Exit Function
ErrorHandlerGLTFile:
   '�u�����ێ�f�[�^��ʁF�t�@�C���A�N�Z�X�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 2
   'GLT�t�@�C�������B
   Close #intGLTFileNo

End Function
'V1.11.0.1 ADD END

' EG20 V2.0.1.1�y�c����54�zADD START
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
'//  REVISIONS    :(EG20 V2.0.1.1) 2011-11-22  CODED BY  [TCC] T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�z
'//                �E�V�X�e���ݒ胁�j���[��ʂ�藬�p
'//  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'//  ���l         :�V�X�e���ݒ胁�j���[���
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

' EG20 V2.0.1.1�y�c����54�zADD END
