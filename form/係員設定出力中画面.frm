VERSION 5.00
Begin VB.Form frmRenewOutput 
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Timer tmrOutput 
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
      Enabled         =   0   'False
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
      TabIndex        =   1
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
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   840
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmRenewOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmRenewOutput.frm
'//  �p�b�P�[�W���F�W���ݒ�}�̏o�͒����
'//
'//  �T�v�F�W���ݒ�}�̏o�͒����
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 �y���k�t�H���_�w��Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////

Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �W���ݒ�}�̏o�͒����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    cmdOK.Enabled = False
    
    On Error Resume Next
    
    tmrMail.Enabled = True
    
'    �ۑ����̃K�C�h��\������
    lblMessage(0) = "�ݒ�l���o�͒��ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    cmdOK.Enabled = False
    DoEvents
    
    tmrOutput.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �W���ݒ�}�̏o�͒����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
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

    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �W���ݒ�}�̏o�͒����(���[�h��)
'//  �@�\�T�v  : �����������s���B
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
Private Sub Form_Load()

    Dim i As Integer '�J�E���^
    Dim intCount As Integer
    Dim intCount2 As Integer
    
    On Error Resume Next
    
    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    tmrOutput.Interval = 100
    tmrOutput.Enabled = False
    
End Sub

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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()

    On Error Resume Next
    
    '����ʂ������B
    Unload Me
    
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  '���[����M�G���A
    Dim lngLength As Long            '��M���[���o�C�g�T�C�Y
    Dim intStatus As Integer         '��M���[���`�F�b�N����

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
                AppActivate frmRenewOutput.Caption, False
                pfFormActive (frmRenewOutput.hwnd)	' EG20 V8.1.0.1 �yEG20_KANSI05_01�zADD
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
'//  �֐�����  : sOutput_Data
'//  �@�\����  : �ݒ�l�o�͏���
'//  �@�\�T�v  : �ݒ�l��ҏW���Ĕ}�̏o�͂���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-27   CODED   BY [TCC] M.Matsumoto
'//                �y����No56�Ή��z
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 �y���k�t�H���_�w��Ή��z
'//     REVISIONS :�iEG20 V30.1.0.1) 2014-04-02 CODED BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sOutput_Data()

    Dim bySyoAssort As Byte             '���O�p������
    Dim strFilePath As String           '�o�̓t�@�C���p�X
    Dim strCornerPath As String         '�ݒ�t�@�C���p�X
    Dim strStationNm As String          '�w��
    Dim strCornerNm As String           '�R�[�i��
    Dim intCount As Integer             '�J�E���^
    Dim intCount2 As Integer            '�J�E���^
    Dim intOutFile As Integer           '�o�̓t�@�C���ԍ�
    Dim intTgtFileNo As Integer         '�o�͑Ώېݒ�t�@�C���ԍ�
    Dim strTgtFileName As String        '�o�͑Ώېݒ�t�@�C��
    Dim strTargetFile() As String       '�o�͑Ώۃt�@�C��
    Dim strTargetFileKan() As String    '�o�͑Ώۃt�@�C���y�����R�[�i�����z 'EG20 V30.1.0.1 ADD
    Dim intFileNum As Integer
    Dim strDefault As String
    Dim strRet As String * 32
    Dim lngRet As Long
    Dim sLzhDirName As String
    Dim sLzhFileName As String
    Dim strCabTarget As String
    Dim lngRetZip As Long
    Dim objFileObj As FileSystemObject  '�t�@�C���V�X�e���I�u�W�F�N�g
    Const lngBufSize = 32
    Dim nIndex As Integer               ' ������                    ' EG20 V5.6.0.1�ǉ�
    
    On Error GoTo Err_Handler
    
    sLzhDirName = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    If sLzhDirName = "" Then
        Unload Me
        Exit Sub  '�f�B���N�g�����w�肳��Ȃ���΁A�����I��
    End If
    
' EG20 V5.6.0.1�y���k�t�H���_�w��Ή��z�ǉ��J�n
    ' �o�̓t�H���_�ɔ��p�X�y�[�X���܂܂�Ă���ꍇ�A���k�ňُ킪�������Ă��܂�����
    ' ���k�O�Ƀ`�F�b�N���Ĉُ��\������B
    nIndex = InStr(sLzhDirName, " ")
    If nIndex <> 0 Then
        ' �x���|�b�v�A�b�v�E�B���h�E��\������B
        Call MsgBox(CABFOLDERSELECT_ERRORMESSAGE, vbCritical, CABFOLDERSELECT_ERRORTITLE)
        Unload Me
        Exit Sub  '�f�B���N�g�����w�肳��Ȃ���΁A�����I��
    End If
' EG20 V5.6.0.1�y���k�t�H���_�w��Ή��z�ǉ��I��

    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_KAKARI_OUTPUT)
    
    Set objFileObj = New FileSystemObject
    
    '�o�͑Ώېݒ�t�@�C�����I�[�v������B
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE
    
    '�o�͑Ώېݒ�t�@�C�������݂��Ȃ��ꍇ�ُ͈�I��
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_Handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '�o�͑Ώۃt�@�C�������擾
    Input #intTgtFileNo, intFileNum
    
    '�o�͑Ώۃt�@�C�����擾
    ReDim strTargetFile(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFile)
        Input #intTgtFileNo, strTargetFile(intCount)
    Next
    
    Close #intTgtFileNo
    
    'EG20 V30.1.0.1 ADD START
    '�����R�[�i�[�ɑ΂���o�͑Ώۃt�@�C���̓��e���m�ۂ���
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE_KAN
    
    '�o�͑Ώېݒ�t�@�C�������݂��Ȃ��ꍇ�ُ͈�I��
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_Handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '�o�͑Ώۃt�@�C�������擾
    Input #intTgtFileNo, intFileNum
    
    '�o�͑Ώۃt�@�C�����擾
    ReDim strTargetFileKan(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFileKan)
        Input #intTgtFileNo, strTargetFileKan(intCount)
    Next
    
    Close #intTgtFileNo
    'EG20 V30.1.0.1 ADD END
        
    '�I���R�[�i�ɂ��ďo�͏���������
    For intCount = 0 To UBound(glngTergetCorner)
        If glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON Then
            '�R�[�i�P
            If intCount = 0 Then
                'Ini�t�@�C������w�����擾
                lngRet = GetPrivateProfileString(strAppName_station, STATIONINI_KEY_EKINAME, _
                                            strDefault, strRet, lngBufSize, KANSI_STATION_INI_FILE)
            '�R�[�i�P�ȊO
            Else
                'Ini�t�@�C������w�����擾
                lngRet = GetPrivateProfileString(strAppName_station & CStr(intCount + 1), STATIONINI_KEY_EKINAME, _
                                            strDefault, strRet, lngBufSize, KANSI_STATION_INI_FILE)
            End If
            '�o�̓t�@�C�����ҏW
            strStationNm = Trim(strRet)
            strStationNm = Replace(strStationNm, Chr(0), "")
            strStationNm = Replace(strStationNm, " ", "")           'EG20 V5.5.0.1 ADD �y����No56�Ή��z
            strCornerNm = gstrCornerName(intCount)
            strCornerNm = Replace(strCornerNm, Chr(0), "")
            strCornerNm = Replace(strCornerNm, " ", "")             'EG20 V5.5.0.1 ADD �y����No56�Ή��z
            strFilePath = strStationNm & "_" & strCornerNm & OUTPUT_LIST_FILE
            sLzhFileName = strStationNm & "_" & strCornerNm & OUTPUT_CAB_FILE
            strFilePath = sLzhDirName & strFilePath
            sLzhFileName = sLzhDirName & sLzhFileName
            
            '---- �ݒ�ꗗ�e�L�X�g�쐬 �J�n
            '�t�@�C���쐬
            If objFileObj.FileExists(strFilePath) = True Then
                objFileObj.DeleteFile (strFilePath)
            End If
            Call objFileObj.CreateTextFile(strFilePath)
            
            '�o�̓t�@�C�����I�[�v������B
            intOutFile = FreeFile
            Open strFilePath For Output As #intOutFile
    
            '�ݒu�w�E�R�[�i���o��
            Print #intOutFile, "�ݒu�w�F" & strStationNm
            Print #intOutFile, "�ݒu�R�[�i�F" & strCornerNm
            Print #intOutFile, ""
            
            'ID�ݒ�l���o��
            If gsubOutput_Id(intCount + 1, intOutFile) = False Then
                GoTo Err_Handler
            End If

            'EG20 V30.1.0.1 DEL START
            '���o��t���[�t�@�C�����o��
'            If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
'
'            '�j�Փ��t�@�C�����o��
'            If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
            'EG20 V30.1.0.1 DEL END
            
            'EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                '�����R�[�i�̏ꍇ

                '�V�����s���p�����[�^���o��
                If gsubOutput_ParaKan(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
                '�ݗ���IC����p�����[�^���o��
                If gsubOutput_ParaKan(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
                
                '�ݗ���IC�ʉߏ����p�����[�^���o��
                If gsubOutput_ParaKan(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
            Else
                '�ݗ��R�[�i�[�̏ꍇ
                '���o��t���[�t�@�C�����o��
                If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
                
                '�j�Փ��t�@�C�����o��
                If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then
                    GoTo Err_Handler
                End If
            End If
            'EG20 V30.1.0.1 ADD END
            
            Close #intOutFile
            '---- �ݒ�ꗗ�e�L�X�g�쐬 �I��
            
            '---- �ݒ�ۑ����k�t�@�C���쐬 �J�n
            '�R�[�i�ʐݒ�t�@�C���p�X
            strCornerPath = PATH_OPERATE_CORNER & CStr(intCount + 1) & PATH_OPERATE_SETTEI
            
            strCabTarget = Empty
            '�o�͑Ώۃt�@�C�����ݒ�
            ' EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                '�����R�[�i�[�̏ꍇ�͊����R�[�i�[�p�̑Ώۃt�@�C���ŏ�������
                For intCount2 = 0 To UBound(strTargetFileKan)
                    strCabTarget = strCabTarget & strCornerPath & strTargetFileKan(intCount2) & " "
                Next
            Else
                '�ݗ��R�[�i�[�̏ꍇ�͍ݗ��R�[�i�[�p�̑Ώۃt�@�C���ŏ�������
                For intCount2 = 0 To UBound(strTargetFile)
                    strCabTarget = strCabTarget & strCornerPath & strTargetFile(intCount2) & " "
                Next
            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.1.0.1 DEL START
'            For intCount2 = 0 To UBound(strTargetFile)
'                strCabTarget = strCabTarget & strCornerPath & strTargetFile(intCount2) & " "
'            Next
            'EG20 V30.1.0.1 DEL END
            
            lngRetZip = gsubCabZip(sLzhFileName, strCabTarget)
            
            If (lngRetZip <> 0) Then   '���k���ʂ�����(0)�ȊO
                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LZH_ERROR, 0)
                GoTo Err_Handler
            End If
            '---- �ݒ�ۑ����k�t�@�C���쐬 �I��
        End If
    Next intCount
    
    Set objFileObj = Nothing
    
    lblMessage(0).Caption = "����I�����܂����B"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    DoEvents
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
    Exit Sub
    
'�G���[����
Err_Handler:

    If intTgtFileNo > 0 Then
        Close #intTgtFileNo
    End If
    If intOutFile > 0 Then
        Close #intOutFile
    End If

    Set objFileObj = Nothing
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, bySyoAssort, KAKARISET_OUTPUT_ERR, 0)
    
    lblMessage(0).Caption = "�ُ�I�����܂����B"
    lblMessage(1).Caption = ""
    cmdOK.Enabled = True
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : tmrOutput_Timer
'//  �@�\����  : �o�͏������s�^�C�}
'//  �@�\�T�v  : �ݒ�l��ҏW���Ĕ}�̏o�͂���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-02   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrOutput_Timer()

    On Error Resume Next
    
    tmrOutput.Enabled = False
    Call sOutput_Data
     
End Sub
