VERSION 5.00
Begin VB.Form frmShimekiriOfflineOut 
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
   Begin VB.Timer tmrMail2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   360
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
Attribute VB_Name = "frmShimekiriOfflineOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmShimekiriOfflineOut.frm
'//  �p�b�P�[�W���F���؃f�[�^�I�t���C���o�͒����
'//
'//  �T�v�F���؃f�[�^�I�t���C���o�͒����
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.10.0.1) 2012-05-09   CODED   BY [TCC] H.Sugimoto
'//                 �y�ێ���؋@�\���P�z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////

Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���؃f�[�^�o�͒����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    On Error Resume Next
    
    ' �I�t���C���o�͒��̃K�C�h��\������
    lblMessage(0) = "���؃f�[�^���I�t���C���o�͒��ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
    tmrMail2.Enabled = True     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ���؃f�[�^�o�͒����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    On Error Resume Next
    
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
    tmrMail2.Enabled = False     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : ���؃f�[�^�o�͒����(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_SHIMEKIRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    tmrMail2.Interval = MN_MAIL_INTERVAL     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    tmrMail2.Enabled = False                 ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
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
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()

    On Error Resume Next
    
    '����ʂ������B
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
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
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 �y�@�\�������z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    On Error Resume Next
    
    ' ���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
    
' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL START
'    ' �ėp���C����M�������s��
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
'        AppActivate frmSyusyuOutPut.Caption, False
'    End If
' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL END
' EG20 V6.3.0.1�y�@�\�������z�폜�J�n
'    ' �o�̓t�@�C���쐬�������s���B
'    frmShimekiriData.gbShimekiriResult = sOutPutOfflineData
'
' EG20 V6.3.0.1�y�@�\�������z�폜�I��
' EG20 V6.3.0.1�y�@�\�������z�ǉ��J�n
    If frmShimekiriData.glShimekiriType = 1 Then
        ' �o�̓t�@�C���쐬�������s���B
        frmShimekiriData.gbShimekiriResult = sOutPutOfflineData
    Else
        frmShimekiriData.gbShimekiriResult = sReOutPutOfflineData
    End If
' EG20 V6.3.0.1�y�@�\�������z�ǉ��I��

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    If frmShimekiriData.gbShimekiriResult = True Then
        lblMessage(0) = "����I�����܂����B"
        lblMessage(1) = ""
    Else
        lblMessage(0) = "�ُ�I�����܂����B"
        lblMessage(1) = ""
    End If
    cmdOK.Enabled = True
    
End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
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
'//     ORIGINAL  :(EG20 V8.1.0.1) 2014-06-05  CODED  BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail2_Timer()

    On Error Resume Next

    ' �ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmShimekiriOfflineOut.Caption, False
        pfFormActive (frmShimekiriOfflineOut.hwnd)
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  �֐�����  : sOutPutOfflineData
'//  �@�\����  : �I�t���C���f�[�^�}�̏o�͏���
'//  �@�\�T�v  : ���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : BOOL      TRUE      ����
'//                        FALSE     �ُ�
'//
'//     ORIGINAL  :(EG20 V5.6.0.1) 2012-04-05   CODED   BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.10.0.1) 2012-05-09   CODED   BY [TCC] H.Sugimoto
'//                 �y�ێ���؋@�\���P�z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-01  CODED   BY [TCC]T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sOutPutOfflineData() As Boolean
            
    Dim nListCnt As Integer                             ' �t�@�C���i�[��
    Dim szFileName As String                            ' �t�@�C����
    Dim lResult As Long                                 ' ��������
    Dim dwCorner As Long                                ' �R�[�i
    Dim dwSequense As Long                              ' �V�[�P���X�ԍ�
    Dim szWork As String                                ' ���[�N
    Dim szNameWork As String                            ' ���[�N
    
    gsGetCornerType     '�e�R�[�i�[�̃^�C�v���擾       EG20 V30.1.0.1
            
    ' //////////////////////////////////////////////////////////////
    ' // �t�@�C���쐬����
    For nListCnt = 0 To UBound(gOfflineFileList) - 1    ' �t�@�C�����X�g��
    
        szFileName = gOfflineFileList(nListCnt)         ' �t�@�C�����̎擾
        
' EG20 V5.10.0.1 �폜�J�n
' �Ώۃt�@�C���ύX
' �i�ύX�O�j�uHOSHU_SIMEKIRI01_001.DAT�v
' �i�ύX��j�uSIMEKIRI01.DAT�v
'
'        ' �uHOSHU_SIMEKIRI01_001.DAT�v�̃R�[�i�ԍ��ƃV�[�P���X�ԍ��𒊏o
'        szNameWork = Right(szFileName, 24)
'        szWork = Mid(szNameWork, 15, 2)
'        dwCorner = CInt(szWork)
'        szWork = Mid(szNameWork, 18, 3)
'        dwSequense = CInt(szWork)
' EG20 V5.10.0.1 �폜�I��
' EG20 V5.10.0.1 �ǉ��J�n
        ' �uSIMEKIRI01.DAT�v�̃R�[�i�ԍ��𒊏o
        szNameWork = Right(szFileName, 14)
        szWork = Mid(szNameWork, 9, 2)              ' �R�[�i�ԍ�
        dwCorner = CInt(szWork)
        dwSequense = 0                              ' �V�[�P���X�ԍ�:0�Œ�
' EG20 V5.10.0.1 �ǉ��I��

        'EG20 V30.1.0.1 DEL START
'        lResult = dllCreateShimekiriFile(dwCorner, dwSequense, _
'                                frmShimekiriData.glbFilePath, _
'                                szFileName)
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        If gintCornerType(dwCorner - 1) = CORNER_TYPE_KANSEN Then
            lResult = dllCreateShimekiriFileKan(dwCorner, dwSequense, _
                                    frmShimekiriData.glbFilePath, _
                                    szFileName)
        Else
            lResult = dllCreateShimekiriFile(dwCorner, dwSequense, _
                                    frmShimekiriData.glbFilePath, _
                                    szFileName)
        End If
        'EG20 V30.1.0.1 ADD END
        If lResult = False Then
            sOutPutOfflineData = False
            Exit Function
        End If
    Next

    sOutPutOfflineData = True
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  �֐�����  : sReOutPutOfflineData
'//  �@�\����  : �I�t���C���f�[�^�}�̍ďo�͏���
'//  �@�\�T�v  : ���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : BOOL      TRUE      ����
'//                        FALSE     �ُ�
'//
'//     ORIGINAL  :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 �y�@�\�������z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sReOutPutOfflineData() As Boolean
            
    Dim objFso As New FileSystemObject                  ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim nListCnt As Integer                             ' �t�@�C���i�[��
    Dim szSrcFileName As String                         ' �o�̓t�@�C����
    Dim szDstFileName As String                         ' �ۑ���t�@�C����
    Dim FileName As String                              ' �t�@�C����
    Dim FileKaku As String                              ' �g���q
    
    On Error GoTo ErrorHandler                          ' �G���[�n���h���̓o�^
            
    ' //////////////////////////////////////////////////////////////
    ' // �t�@�C���쐬����
    For nListCnt = 0 To UBound(gOfflineFileList) - 1    ' �t�@�C�����X�g��
    
        szSrcFileName = gOfflineFileList(nListCnt)      ' �t�@�C�����̎擾
        If objFso.FileExists(szSrcFileName) = True Then
            
            ' �t�@�C�����擾
            psFileNameGet szSrcFileName, FileName, FileKaku
            
            ' �R�s�[��t�@�C�����쐬
            szDstFileName = frmShimekiriData.glbFilePath & "\" & FileName & "." & FileKaku
            
            '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
            objFso.CopyFile szSrcFileName, szDstFileName, True
        
        End If
        
    Next

    sReOutPutOfflineData = True
    Set objFso = Nothing
    Exit Function
' /////////////////////////////////////////////////////////
' // �G���[����
ErrorHandler:

    Set objFso = Nothing
    sReOutPutOfflineData = False

End Function


