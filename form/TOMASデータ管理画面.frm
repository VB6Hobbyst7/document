VERSION 5.00
Begin VB.Form frmTomasDataMng 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "Form1"
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
   Begin VB.Timer tmrMail 
      Left            =   11280
      Top             =   5760
   End
   Begin VB.CommandButton cmdEject 
      Caption         =   "�}�̎�O"
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
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdDispErrInfo 
      Caption         =   "��Q�������\��"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdOutErrInfo 
      Caption         =   "��Q�������}�̏o��"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdDispVerInfo 
      Caption         =   "�o�[�W�������\��"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdOutVerInfo 
      Caption         =   "�o�[�W�������}�̏o��"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdDispKikiInfo 
      Caption         =   "�@���ԃf�[�^�\��"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdOutKikiInfo 
      Caption         =   "�@���ԃf�[�^�}�̏o��"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   2400
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
      TabIndex        =   1
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "TOMAS�f�[�^�Ǘ�"
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
      Width           =   12015
   End
End
Attribute VB_Name = "frmTomasDataMng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmTomasDataMng.frm
'//  �p�b�P�[�W���FTOMAS�f�[�^�Ǘ����
'//
'//  �T�v�F�o�[�W�����Ǘ����
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//                 �V�K�쐬�y�t�F�[�Y�R TOMAS�Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   CODED   BY [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z

'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////

Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdDispErrInfo_Click
'//  �@�\����  : �u��Q�������\���v�t����������
'//  �@�\�T�v  : ��Q�������\�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdDispErrInfo_Click()

    Dim blnRet As Boolean
    Dim lngErrCode As Long
    Dim strFileName As String
    Dim strCommand As String
    Dim lRetVal As Long
    
    On Error Resume Next
    
    '�f�[�^��ʁF��Q�������
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_ERR
    
    gblnTomasDispErr = False
    
    'TOMAS�f�[�^�\�����t�H�[�����A���[�_���E�B���h�E�ŕ\������B
    frmTomasDataDisp.Show vbModal
    
    '�G���[�̏ꍇ
    If gblnTomasDispErr = True Then
        Exit Sub
    End If
    
    '����I��
    strFileName = TOMAS_FILE_ERRINFO
    
    strCommand = MN_EXE_MEMO & PATH_WORK & strFileName      '���s�R�}���h���쐬����
    lRetVal = Shell(strCommand, vbMaximizedFocus)           '�m�[�g�p�b�h���N������
    AppActivate lRetVal, True                               '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
    SendKeys "{LEFT}", True
    
    '����I�����͏�Q�������}�̏o�͖t�������ɂ���B
    cmdOutErrInfo.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdDispKikiInfo_Click
'//  �@�\����  : �u�@���ԃf�[�^�\���v�t����������
'//  �@�\�T�v  : �@���ԃf�[�^�\�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdDispKikiInfo_Click()
    
    Dim strFileName As String
    Dim strCommand As String
    Dim lRetVal As Long
    Dim strRet As String
    Dim intRetCd As Integer
    
    On Error Resume Next
    
    '�uTOMAS�f�[�^�Ǘ���ʁF�@���ԃf�[�^�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TOMAS_DATA_KIKI_DISP, 0)
    
    '�f�[�^��ʁF�@���ԃf�[�^
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_KIKI
    
    gblnTomasDispErr = False
    
    'TOMAS�f�[�^�\�����t�H�[�����A���[�_���E�B���h�E�ŕ\������B
    frmTomasDataDisp.Show vbModal
    
    
    '����I�����͋@���ԃf�[�^�}�̏o�͖t�������ɂ���B
    If gblnTomasDispErr = False Then
        cmdOutKikiInfo.Enabled = True
        
        strCommand = MN_EXE_MEMO & PATH_WORK & TOMAS_FILE_KIKIINFO      '���s�R�}���h���쐬����
        lRetVal = Shell(strCommand, vbMaximizedFocus)           '�m�[�g�p�b�h���N������
        AppActivate lRetVal, True                               '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
        SendKeys "{LEFT}", True
    
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdDispVerInfo_Click
'//  �@�\����  : �u�o�[�W�������\���v�t����������
'//  �@�\�T�v  : �o�[�W�������\�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdDispVerInfo_Click()
    
    Dim strCommand As String
    Dim lRetVal As Long
    Dim strRet As String
    Dim intRetCd As Integer
    
    On Error Resume Next
    
    '�uTOMAS�f�[�^�Ǘ���ʁF�o�[�W�������\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TOMAS_DATA_VER_DISP, 0)
    
    '�f�[�^��ʁF�o�[�W�������
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION
    
    gblnTomasDispErr = False
    gblnRecvErr = False
    
    'TOMAS�f�[�^�\�����t�H�[�����A���[�_���E�B���h�E�ŕ\������B
    frmTomasDataDisp.Show vbModal
    
    '�o�[�W�������\��RES�𐳏��M�����ꍇ�́A�����ʒm�𑗐M����B
    If gblnRecvErr = False Then
        Call fSDATAMailSend_Commit
    End If
    
    '����I�����̓o�[�W�������}�̏o�͖t�������ɂ���B
    If gblnTomasDispErr = False Then
        
        strCommand = MN_EXE_MEMO & PATH_WORK & TOMAS_FILE_VERINFO      '���s�R�}���h���쐬����
        lRetVal = Shell(strCommand, vbMaximizedFocus)           '�m�[�g�p�b�h���N������
        AppActivate lRetVal, True                               '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
        SendKeys "{LEFT}", True
        
        cmdOutVerInfo.Enabled = True
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : fSDATAMailSend_Commit
'//  �@�\����  : �o�[�W�����擾�����ʒm���M
'//  �@�\�T�v  : �����������F���[���𑗐M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSDATAMailSend_Commit() As Boolean

    Dim udtMail As VERSION_DATA_CMT_TYPE
    Dim bRet As Boolean             '�֐��߂�l
    Dim lngErrCode As Long          '�G���[�R�[�h
    
    On Error Resume Next
 
    fSDATAMailSend_Commit = True
    
    'TOMAS�f�[�^�o�͗v���𑗐M����
    udtMail.mlHeader.dwId = ML_ID_TOMAS_VARSION_DSP_COMMIT
    udtMail.mlHeader.dwSize = MlSize.TOMAS_DATA_DSP_CMT
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    
    udtMail.dwSeqNo = gintSeqNo_Version                 '�V�[�P���X�ԍ�
    udtMail.dwDenbunSize = 8                            '�d���T�C�Y�i�Œ�j
    udtMail.byCmd(0) = &H7A
    udtMail.byCmd(1) = &H41
    udtMail.byCmd(2) = &H1
    udtMail.byCmd(3) = &H1
    udtMail.byCmd(4) = &H1
    udtMail.byCmd(5) = &H1
    udtMail.byCmd(6) = &H0
    udtMail.byCmd(7) = &H0
    
    '���[�����M
    bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMail), udtMail.mlHeader)
    
    If bRet = False Then
        '�uTOMAS�f�[�^�\����ʁFTOMAS�f�[�^�����ʒm�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, TOMAS_DATA_VER_COMMIT, lngErrCode)
        fSDATAMailSend_Commit = False
        Exit Function
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdEject_Click
'//  �@�\����  : �}�̎�O�t����������
'//  �@�\�T�v  : �}�̎�O���������s����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdEject_Click()
    
    On Error Resume Next
    
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
    
    '�}�̎�O����
    Call pfRemove(Me)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdOutVerInfo_Click
'//  �@�\����  : �u�o�[�W�������}�̏o�́v�t����������
'//  �@�\�T�v  : �o�[�W�������e�L�X�g�t�@�C����}�̂ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutVerInfo_Click()

    On Error GoTo Err_Handler
    
    '�}�̏o�͏������s��
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION
    Call sOutput
    
    Exit Sub

Err_Handler:
    
    gblnTomasDispErr = True
    frmTomasDataDisp.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdOutKikiInfo_Click
'//  �@�\����  : �u�@���ԃf�[�^�}�̏o�́v�t����������
'//  �@�\�T�v  : �o�[�W�������e�L�X�g�t�@�C����}�̂ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutKikiInfo_Click()

    On Error GoTo Err_Handler
    
    '�}�̏o�͏������s��
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_KIKI
    Call sOutput
    
    Exit Sub

Err_Handler:
    
    gblnTomasDispErr = True
    frmTomasDataDisp.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdOutErrInfo_Click
'//  �@�\����  : �u��Q�������}�̏o�́v�t����������
'//  �@�\�T�v  : �o�[�W�������e�L�X�g�t�@�C����}�̂ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-27   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutErrInfo_Click()

    On Error GoTo Err_Handler
    
    '�}�̏o�͏������s��
    gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_ERR
    Call sOutput
    
    Exit Sub

Err_Handler:
    
    gblnTomasDispErr = True
    frmTomasDataDisp.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdDispVerInfo_Click
'//  �@�\����  : �u�o�[�W�������\���v�t����������
'//  �@�\�T�v  : �o�[�W�������\�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sOutput()
    
    On Error Resume Next
    
    gstrOutPath = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    If gstrOutPath = "" Then
        Exit Sub  '�f�B���N�g�����w�肳��Ȃ���΁A�����I��
    End If
    
    gblnTomasDispErr = False
    frmTomasDataOut.Show vbModal
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    On Error Resume Next
    
    '�uTOMAS�f�[�^�\����ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TOMAS_DATA_MNG_GAMEN_END, 0)
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : TOMAS�f�[�^�Ǘ����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    On Error Resume Next
    
    '���[����M�p�^�C�}���N������
    tmrMail.Enabled = True

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : TOMAS�f�[�^�Ǘ����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : TOMAS�f�[�^�Ǘ����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���[����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '�}�̏o�͖t�͏����l�񊈐�
    cmdOutVerInfo.Enabled = False
    cmdOutKikiInfo.Enabled = False
    cmdOutErrInfo.Enabled = False
    
    '�V�[�P���X�ԍ�������
    gintSeqNo_Version = MIN_SEQ_VERSION
    gintSeqNo_KikiData = MIN_SEQ_KIKIDATA
    
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//     �T�v      : �u���[����M�p�^�C�}�v���^�C���A�b�v�������̃C�x���g�v���V�[�W��
'//     ����      : ���[����M�������s���B
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   CODED   BY [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
'        AppActivate frmTimeDataSettei.Caption, False   ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
        AppActivate frmTomasDataMng.Caption, False      ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
        pfFormActive (frmTomasDataMng.hwnd)             ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If

End Sub
