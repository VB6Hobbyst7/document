VERSION 5.00
Begin VB.Form frmTomasDataDisp 
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Timer tmrErrDisp 
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�n �j"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
End
Attribute VB_Name = "frmTomasDataDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmTomasDataMng.frm
'//  �p�b�P�[�W���FTOMAS�f�[�^�\�����
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
'//  �֐�����  : cmdOK_Click
'//  �@�\����  : �uOK�v�t����������
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
Private Sub cmdOK_Click()

    On Error Resume Next
    
    '����ʂ������B
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : TOMAS�f�[�^�\�����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
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
Private Sub Form_Activate()
    
    On Error Resume Next
    
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_TOMAS_DATA_DISP)

    If gintTomasDataDispDiv <> TOMAS_DISP_DIV.TOMAS_DATA_ERR Then
        'TOMAS�f�[�^�o�͗v�����ă}�֑��M����B
        If fSDATAMailSend = False Then
            lblMessage(0) = "�ُ�I�����܂����B"
            lblMessage(1) = ""
            cmdOK.Enabled = True
            gblnTomasDispErr = True
            gblnRecvErr = True
            Exit Sub
        End If
    End If
    
    '�������̃K�C�h��\������
    lblMessage(0) = "�������ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
    tmrErrDisp.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : TOMAS�f�[�^�\�����(�f�B�A�N�e�B�u��)
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
    
    tmrErrDisp.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : TOMAS�f�[�^�\�����(���[�h��)
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
    
    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    tmrErrDisp.Interval = MN_MAIL_INTERVAL
    tmrErrDisp.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : fSDATAMailSend
'//  �@�\����  : TOMAS�f�[�^�o�͗v�����M����
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
Private Function fSDATAMailSend() As Boolean

    Dim udtMailVer As VERSION_DATA_DISP_TYPE
    Dim udtMailKiki As KIKI_DATA_DISP_TYPE
    Dim bRet As Boolean             '�֐��߂�l
    Dim lngErrCode As Long          '�G���[�R�[�h
    Dim strLog As String
    
    On Error Resume Next
 
    fSDATAMailSend = True
    
    '�o�[�W�������\���̏ꍇ
    If gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION Then
        '�V�[�P���X�ԍ�
        If gintSeqNo_Version = MAX_SEQ_VERSION Then
            gintSeqNo_Version = MIN_SEQ_VERSION + 1
        Else
            gintSeqNo_Version = gintSeqNo_Version + 1
        End If
        
        udtMailVer.mlHeader.dwId = ML_ID_TOMAS_VARSION_DSP_REQ_CMD
        udtMailVer.mlHeader.dwSize = MlSize.VERSION_DATA_DSP_CMD
        udtMailVer.mlHeader.dwProid = RHOSHU_ID
        udtMailVer.mlHeader.dwSubArea = 0
        
        udtMailVer.dwSeqNo = gintSeqNo_Version              '�V�[�P���X�ԍ�
        udtMailVer.dwBlockCheck = 0                         '�u���b�N�ԍ��`�F�b�N�m�F
        udtMailVer.dwDenbunSize = 6                         '�d���T�C�Y�i�Œ�j
        udtMailVer.byCmd(0) = &H78                          '�R�}���h�R�[�h
        udtMailVer.byCmd(1) = &H41                          '�T�u�R�[�h
        udtMailVer.byCmd(2) = &H1                           '�R�[�iNo
        udtMailVer.byCmd(3) = &H1                           '���@No
        udtMailVer.byCmd(4) = &H1                           '�u���b�NNo
        udtMailVer.byCmd(5) = &H1                           '�ŏI�u���b�NNo
        strLog = TOMAS_DATA_VER_REQ_SEND
        
        '���[�����M
        bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMailVer), udtMailVer.mlHeader)
    
    '�@���ԃf�[�^�\���̏ꍇ
    Else
        '�V�[�P���X�ԍ�
        If gintSeqNo_KikiData = MAX_SEQ_KIKIDATA Then
            gintSeqNo_KikiData = MIN_SEQ_KIKIDATA + 1
        Else
            gintSeqNo_KikiData = gintSeqNo_KikiData + 1
        End If
        
        udtMailKiki.mlHeader.dwId = ML_ID_TOMAS_KIKIDATA_DSP_REQ_CMD
        udtMailKiki.mlHeader.dwSize = MlSize.KIKIINF_DATA_DSP_CMD
        udtMailKiki.mlHeader.dwProid = RHOSHU_ID
        udtMailKiki.mlHeader.dwSubArea = 0
    
        udtMailKiki.dwSeqNo = gintSeqNo_KikiData                '�V�[�P���X�ԍ�
        udtMailKiki.dwDenbunSize = 6                            '�d���T�C�Y�i�Œ�j
        udtMailKiki.byCmd(0) = &H79                             '�R�}���h�R�[�h
        udtMailKiki.byCmd(1) = &H41                             '�T�u�R�[�h
        udtMailKiki.byCmd(2) = &H1                              '�R�[�iNo
        udtMailKiki.byCmd(3) = &H1                              '���@No
        udtMailKiki.byCmd(4) = &H1                              '�u���b�NNo
        udtMailKiki.byCmd(5) = &H1                              '�ŏI�u���b�NNo
        strLog = TOMAS_DATA_KIKI_REQ_SEND
        
        '���[�����M
        bRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMailKiki), udtMailKiki.mlHeader)
    
    End If
    
    If bRet = False Then
        '�uTOMAS�f�[�^�\����ʁFTOMAS�f�[�^�o�͗v�����M�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, strLog, lngErrCode)
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        fSDATAMailSend = False
        Exit Function
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : fReadMailCheck
'//  �@�\����  : TOMAS�f�[�^�o�͒ʒm���[���`�F�b�N����
'//  �@�\�T�v  : ���[����M���F���[������M����B
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
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim blnRet As Boolean
    Dim lngErrCode As Long
    Dim intRetCd As Integer
    Dim strRet As String
    
    On Error Resume Next
    
    '���^�[���R�[�h�����o��
    strRet = Format(Hex(udtReadMail.lngData(3)), "00000000")
    intRetCd = CInt(Mid(strRet, 3, 2))
    
    '�������ʂ��ُ�̏ꍇ
    If intRetCd > 0 Then
        fReadMailCheck = False
        gblnRecvErr = True
        Exit Function
    End If
    Dim strMessage As String
    '�o�[�W�����擾�̏ꍇ�A�����ʒm���o��
    If gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_VERSION Then
        blnRet = dllCreateDispVerInfoFile(gintSeqNo_Version, lngErrCode)
    '�@���ԗv���̏ꍇ�A��M�f�[�^����t�@�C�����쐬����
    Else
        If sMakeDataFile(udtReadMail) = False Then
            fReadMailCheck = False
            Exit Function
        End If
        blnRet = dllCreateDispkikiStsFile(lngErrCode)
    End If
    
    '�ُ�I���̏ꍇ
    If blnRet = False Then
        fReadMailCheck = False
        Exit Function
    End If
    
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    fReadMailCheck = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sMakeDataFile
'//  �@�\����  : �@���ԃf�[�^�쐬
'//  �@�\�T�v  : ��M���[������@���ԃf�[�^���쐬����B
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
Private Function sMakeDataFile(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim lngHandle As Long
    Dim strKikiData As String
    Dim bRet As Boolean
    Dim lngRet As Long
    
    On Error Resume Next
    
    sMakeDataFile = True
    
    strKikiData = PATH_WORK & TOMAS_FILE_KIKIINFO_DAT
    
    '�@����t�@�C�����I�[�v��
    lngHandle = CreateFile(strKikiData, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           CREATE_ALWAYS, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�I�[�v���ُ펞�͎Q�ƕs�̂��ߍX�V�ُ�
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_OPEN, 0)
        sMakeDataFile = False
        Exit Function
    End If
         
    '�@����t�@�C���ɏ�������
    bRet = WriteFile(lngHandle, udtReadMail.lngData(4), udtReadMail.udtlHeader.dwSize - 32, lngRet, 0)
    If bRet = False Then
       '�n���h���̃N���[�Y
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_FILE_WRITE, 0)
       Call CloseHandle(lngHandle)
       sMakeDataFile = False
       Exit Function
    End If
    
    '�n���h���̃N���[�Y
     Call CloseHandle(lngHandle)

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : tmrErrDisp_Timer
'//  �@�\����  : ��Q���f�[�^�\������
'//  �@�\�T�v  : ��Q���f�[�^��\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2012-02-08   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrErrDisp_Timer()

    Dim blnRet As Boolean
    Dim lngErrCode As Long
    
    On Error Resume Next
    
    If gintTomasDataDispDiv = TOMAS_DISP_DIV.TOMAS_DATA_ERR Then
    
        '��Q�������\������
        blnRet = dllCreateDispErrFile(lngErrCode)
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
        
        If blnRet = False Then
            lblMessage(0) = "�ُ�I�����܂����B"
            lblMessage(1) = ""
            cmdOK.Enabled = True
            gblnTomasDispErr = True
            tmrErrDisp.Enabled = False
        Else
            Unload Me
        End If
        
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : ���[����M�p�^�C�}�A�^�C���A�b�v������
'//  �@�\�T�v  : ���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V4.1.0.1) 2011-12-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   CODED   BY [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF     '���[����M�G���A
    Dim lngLength As Long                       '��M���[���o�C�g�T�C�Y
    Dim intStatus As Integer                    '��M���[���`�F�b�N����
    Dim strLog As String
    
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
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                '�v���Z�X�̏I���������s��
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '�u�ێ��ʃA�N�e�B�u�\���v����M�����ꍇ
                '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '�\������ʁi�ێ�f�[�^���W��ʁj���A�N�e�B�u�\������B
'                AppActivate frmRenewCyu.Caption, False         ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
                AppActivate frmTomasDataDisp.Caption, False     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
                pfFormActive (frmTomasDataDisp.hwnd)            ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD

            Case ML_ID_TOMAS_VARSION_DSP_REQ_RES, ML_ID_TOMAS_KIKIDATA_DSP_REQ_RES
                '�u�o�[�W�����擾�iRES�j�v�A�u�@���ԗv���iRES�j�v����M�����ꍇ
                '�uTOMAS�f�[�^�o�͗v����M����v���O�o��
                If udtReadMail.udtlHeader.dwId = ML_ID_TOMAS_VARSION_DSP_REQ_RES Then
                    strLog = TOMAS_DATA_VER_REQ_RECV
                Else
                    strLog = TOMAS_DATA_KIKI_REQ_RECV
                End If
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, strLog, 0)
                '���e���`�F�b�N����B
                If fReadMailCheck(udtReadMail) = True Then
                    Unload Me
                Else
                    lblMessage(0) = "�ُ�I�����܂����B"
                    lblMessage(1) = ""
                    '�v���O���X�o�[����������
                    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
                    gblnTomasDispErr = True
                    cmdOK.Enabled = True
                End If
            Case Else
                 '���̑��̃��[������M�����ꍇ
                 '�u���[��ID�s���v���O�o��
                 Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
    
End Sub

