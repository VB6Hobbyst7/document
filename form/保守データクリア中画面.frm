VERSION 5.00
Begin VB.Form frmHoshuClear 
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "frmHoshuClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmHoshuClear.frm
'//  �p�b�P�[�W���F�ێ�f�[�^�N���A���
'//
'//  �T�v�F�ێ�f�[�^�N���A���
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�ێ�f�[�^�N���A����ʒǉ�
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//                 �y�t�F�[�Y�Q�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �ێ�f�[�^�N���A���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
  
    '�N���A���̃K�C�h��\������
    lblMessage(0) = "�����ێ�SW�ݒ�N���A���ł��B"
    lblMessage(1) = "���΂炭���҂��������B"
    cmdOK.Enabled = False
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �ێ�f�[�^�N���A���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �ێ�f�[�^�N���A���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    On Error Resume Next
    '�u�����ێ�f�[�^�N���A����ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HODHU_SW_CLEAR_SHORI_GAMEN_START, 0)
     
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
     
    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    Dim iCnt As Integer     'V1.7.0.1 ADD

On Error Resume Next
    '����ʂ������B
    '�u�����ێ�f�[�^�N���A����ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, HODHU_SW_CLEAR_SHORI_GAMEN_END, 0)
    'V1.7.0.1 ADD START
    For iCnt = 0 To MAX_GATE_NO + 1
       '�N���A�Ώۍ��@���A�N���A��Ώۂɂď�����
       gClear_Gouki(iCnt) = CLEAR_FLAG.NOT_CLEAR
    Next
    'V1.7.0.1 ADD END
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  On Error Resume Next
    
  '�ėp���C����M�������s��
  If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
     AppActivate frmHoshuClear.Caption, False
     pfFormActive (frmHoshuClear.hwnd)
  End If
   
  '�N���A�������s���B
  psHoshuClear
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psHoshuClear
'//  �@�\����  : �����ێ�SW�ݒ�t�@�C���N���A����
'//  �@�\�T�v  : �����ێ�SW�ݒ�t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iEnd�@�@�@[IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-04   CODED   BY [TCC] M.Matsumoto
'//                 �y�t�F�[�Y�Q�Ή��z
'//     REVISIONS :(EG20 V6.9.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psHoshuClear()
    Dim iCnt As Integer
    Dim sMyPath As String 'GATE_ULL�t�H���_���A�t�@�C��
    Dim iGouki As Integer
    Dim sSWDataPath As String 'RMENTE�t�H���_�p�X
    Dim fso         As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim bRet        As Boolean '�������ʃX�e�[�^�X
    Dim nCorner     As Integer ' �R�[�i�ԍ�     ' EG20 V6.9.0.1�ǉ�

    On Error Resume Next
    
    bRet = True
    
'    For iCnt = 0 To 17         'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    For iCnt = 0 To 31          'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
       If gClear_Gouki(iCnt) = CLEAR_FLAG.TARGET_CLEAR Then
          '�uGATE_ULL�v�t�H���_�p�X���쐬
          sMyPath = Replace(GATE_SW_FILE, "##", Format(iCnt + 1, "0#"))
          '�t�@�C���̗L���`�F�b�N���s���B
          If Dir(sMyPath) <> "" Then
             Kill sMyPath
          End If

          If Dir(sMyPath) <> "" Then
             bRet = False
          End If
          
          '�uRMENTE\�{�d�S\���w\XX���@�v�t�H���_�p�X���쐬
'          iGouki = pfGetGoukiNo(iCnt + 1)              ' EG20 V6.9.0.1�폜
          iGouki = pfGetGoukiNo(iCnt + 1, nCorner)      ' EG20 V6.9.0.1�ǉ�
          If iGouki <> -1 Then
             sSWDataPath = PATH_RMENTE_GATE_DEN_JIEKI_GOUKI
' EG20 V6.9.0.1�ǉ��J�n
             '�u�R�[�i$�v�́u$�v��1�`6�ɕϊ�����B
             sSWDataPath = Replace(sSWDataPath, "$", nCorner)
' EG20 V6.9.0.1�ǉ��I��
             sSWDataPath = Replace(sSWDataPath, "##", Format(iGouki, "0#"))
             sSWDataPath = Mid(sSWDataPath, 1, Len(sSWDataPath) - 2)
             If Dir(sSWDataPath, vbDirectory) <> "" Then
                fso.DeleteFolder (sSWDataPath)
             End If
             If Dir(sSWDataPath, vbDirectory) <> "" Then
                bRet = False
             End If
          End If
        End If
     Next
     
     Set fso = Nothing
     If bRet = False Then
        sClearEnd (1)
     Else
        sClearEnd (0)
     End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGetGoukiNo
'//  �@�\����  : �\�����@�ԍ����擾����B
'//  �@�\�T�v  : GATE.INI���\�����@�ԍ����擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer  iGouki    [IN]���@�ԍ�
'//              Integer  nCorner   [OUT]�R�[�i�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.9.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Function pfGetGoukiNo(iGouki As Integer) As Integer
Private Function pfGetGoukiNo(iGouki As Integer, nCorner As Integer) As Integer

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

    On Error Resume Next

    '�������D�@���擾
    sKeyName = "gate" & Format(iGouki, "00")
    iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                   sKeyName, _
                                   DEFAILT, sGateData, Len(sGateData), _
                                   PATH_GATE_FILE)
    If iRet = 0 Then
       '�u�����ێ�SW�ݒ�N���A��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
       pfGetGoukiNo = -1
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
      
   If Trim(sFData(1)) <> "" Then
      pfGetGoukiNo = Trim(sFData(1))
   End If
' EG20 V6.9.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��J�n
   nCorner = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.9.0.1 �y���@�ԍ��ɃR�[�i�ԍ���t������Ή��z�ǉ��I��

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sClearEnd
'//  �@�\����  : �N���A�������ʕ\������
'//  �@�\�T�v  : �ێ�f�[�^�N���A���ʂ̌��ʕ�����\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iEnd�@�@�@[IN]��������
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-24   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sClearEnd(iEnd As Integer)
    Dim i As Integer       '�J�E���^
    Dim lngErrCode As Long '�G���[�R�[�h

    On Error Resume Next
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
    If iEnd = 0 Then
       '����I�����̕�����\������B
       lblMessage(0) = "����I�����܂����B"
       lblMessage(1) = ""
       '�u�ێ�f�[�^�o�͏�������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, HOSHU_SW_CLEAR_OK, 0)
    Else
       '���W���s���̕�����\������B
       lblMessage(0) = "�ُ�I�����܂����B"
       lblMessage(1) = ""
       '�u�ێ�f�[�^�o�͗��ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FCREATE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, HOSHU_SW_CLEAR_ERROR, lngErrCode)
    End If
    cmdOK.Enabled = True
End Sub
