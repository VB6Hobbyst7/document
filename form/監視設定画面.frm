VERSION 5.00
Begin VB.Form frmKansiSetteiSub 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�����[�g�����e�i���X"
   ClientHeight    =   9000
   ClientLeft      =   2160
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
   Begin Hoshu.ctlSetteiButton ctlSetteiButton1 
      Height          =   1215
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   2143
   End
   Begin VB.Timer tmrMail 
      Left            =   960
      Top             =   480
   End
   Begin VB.CommandButton cmd_Kakutei 
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
      Height          =   1095
      Left            =   7440
      TabIndex        =   2
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " �@�@���j���[ �@�@  ��ʂ֖߂�"
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
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�Ď��ݒ�"
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
      Width           =   12015
   End
End
Attribute VB_Name = "frmKansiSetteiSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmKansiSetteiSub.frm
'//  �p�b�P�[�W���F�Ď��ݒ���
'//
'//  �T�v�F�Ď��ݒ���
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�V�K�ǉ����
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l
Private Const BUTTOM_COLOR = &H8000000F '�t�F(ON��/OFF���F����F)
Private mstrFileName     As String               '�t�@�C����
Private mintMaxIndex     As Integer              'Max�C���f�b�N�X
Private Const KANSI_SETTEI = 1                   '�Ď��ݒ�
Private Const KANSI_STS = 2                      '�Ď����
Private Const HUTEI = -1                         '�l�s��
Private Const DANKI = 0                          '�g�C�^�]����

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �Ď��ݒ���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
   
   '���C����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �Ď��ݒ���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    
    On Error Resume Next
       
    '���C����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �Ď��ݒ���(���[�h��)
'//  �@�\�T�v  : �Ď��ݒ��ʂ̏����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
  
    Dim intCount            As Integer          '�J�E���^�[
    Dim strTitle            As String           '�^�C�g��
    Dim intSetteiKubun      As Integer          '�Ď��܂��͎������̐ݒ�敪
    Dim intX                As Integer          'X�ʒu
    Dim intY                As Integer          'Y�ʒu
    Dim intId               As Integer          '�h�c
    Dim iShoriNo            As Integer          '�����ԍ�
    Dim strOnMoji           As String           'ON������
    Dim strOffMoji          As String           'OFF������
    Dim iOnSts              As Integer          'ON���l
    Dim iOffSts             As Integer          'OFF���l
    Dim intBtnUmu           As Integer          '�t�\���̗L��
    Dim intFileNumber       As Integer          '�t�@�C���ԍ�
    Dim iAreaSts            As Integer          '�擾�l
    
    On Error Resume Next

    '�u�Ď��Րݒ��� �\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SETTEI_GAMEN_START, 0)

    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���g�p�̃t�@�C���ԍ����擾���܂��B
    intFileNumber = FreeFile
    
    '�ݒ���t�@�C������ݒ肷��B
    mstrFileName = HOSHU_KANSI_SETTEI_FILE
        
    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    
    '�ݒ���t�@�C�����I�[�v������B
    If Len(mstrFileName) <> 0 Then
        Open mstrFileName For Input As #intFileNumber
    End If
    
    '2���R�[�h�܂œǂݍ��݁A�ݒ�L�������擾����B
    For intCount = 0 To 1
        Input #intFileNumber, intBtnUmu, intX, intY, intId, _
                              strTitle, iShoriNo, _
                              strOnMoji, strOffMoji, intSetteiKubun

        '�ő�R���g���[������ϐ��ɐݒ肷��B
        If intCount = 1 Then
           mintMaxIndex = intBtnUmu - 1
        End If
    Next
       
    '�ݒ�L�������̂݃G���A�m�ہB
    ReDim m_Filetest(0 To mintMaxIndex)
    
    '�ێ�_��ʐ�p�t�@�C�����e�t����ǂݍ��݁A�e�R���g���[����Load����B
    For intCount = 0 To mintMaxIndex
        '�t�@�C������f�[�^��ǂށB
        '�L�������AX���W�AY���W�A�G���AID�A�^�C�g��
        '�����ԍ��AON�������AOFF�������A�ݒ�t���O���擾
        Input #intFileNumber, intBtnUmu, intX, intY, intId, _
                              strTitle, iShoriNo, _
                              strOnMoji, strOffMoji, intSetteiKubun
        
        '�ʏ�G���[���[�`���ɖ߂�
        On Error Resume Next

       '���^�؃{�^���R���g���[�����k�n�`�c����B
        If intCount > 0 Then
            Load ctlSetteiButton1(intCount)
        End If
        
        '�t�̕\�����s���ꍇ
        If intBtnUmu = 1 Then
           
           '���^�؃{�^���R���g���[���̃v���p�e�B�ɒl��ݒ肷��B
           '�t�\��
            ctlSetteiButton1(intCount).Visible = False
            
            '�ݒ�t���O
            ctlSetteiButton1(intCount).Settei_Flag = intSetteiKubun
        
            '�ΏۃG���A��茻�ݒl���擾����B
            If intSetteiKubun = KANSI_SETTEI Then
               iAreaSts = pfGetKansiArea_Sts(intId)
            End If
            
            If iAreaSts <> HUTEI Then
               '�\�����l
               ctlSetteiButton1(intCount).pSetUp = iAreaSts
               '�t�\��
               ctlSetteiButton1(intCount).Visible = True
               '�t�^�C�g���ݒ�
               ctlSetteiButton1(intCount).pButtonTitle = strTitle
               '�G���AID��ێ�
               ctlSetteiButton1(intCount).pID = intId
               'X���W�ݒ�
               ctlSetteiButton1(intCount).Top = ctlSetteiButton1(0).Top + intX
               'Y���W�ݒ�
               ctlSetteiButton1(intCount).Left = ctlSetteiButton1(0).Left + intY
               'ON�������ݒ�
               ctlSetteiButton1(intCount).On_Caption = strOnMoji
               'OFF�������ݒ�
               ctlSetteiButton1(intCount).Off_Caption = strOffMoji
               '�t�w�i�Œ�ݒ�
               ctlSetteiButton1(intCount).On_Color = BUTTOM_COLOR
               ctlSetteiButton1(intCount).Off_Color = BUTTOM_COLOR
               ctlSetteiButton1(intCount).SetteiOn_Color = BUTTOM_COLOR
               ctlSetteiButton1(intCount).SetteiOff_Color = BUTTOM_COLOR
               '�����ԍ�
               ctlSetteiButton1(intCount).SHORI_NO = iShoriNo
               '���^�؃{�^���R���g���[���̕\���������\�b�h���s���B
               '�擾���ݒl�`�F�b�N���s���B�l�s�莞�F��ʔ�\��
               ctlSetteiButton1(intCount).psDisplay
            End If
        End If
    Next
    
    '�t�@�C�����N���[�Y����B
    Close #intFileNumber
      
Exit Sub

'�G���[����
Err_LOG:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If

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
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next

    '�u�Ď��ݒ��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_SETTEI_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmd_Kakutei_Click
'//  �@�\����  : �u�m��v�t����
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmd_Kakutei_Click()
    
    On Error Resume Next

    '�u�m��t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_KENSHU_KAKUTEI_BUTTOM, 0)
    
    '��ʂ����b�N����B
    SetEnableFalse
    
    '��ʐݒ蔽�f�������s���B
    psDispSettei_Hanei

    '��ʂ̃��b�N����������B
    SetEnableTrue
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
'//  �֐�����  : psDispSettei_Hanei
'//  �@�\����  : �����t�̐ݒ�(���)�𔽉f����B
'//  �@�\�T�v  : �����t��Ԃ�l��Ώۃt�@�C���ɔ��f����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psDispSettei_Hanei()

  Dim iKansiId As Long    '�G���AID
  Dim iSetSts As Integer  '�X�V�l
  Dim iSetFlag As Integer '�ݒ�t���O
  Dim lngAplSts As Long   '�Ď��ՃA�v���N�����
  Dim iCnt As Integer     '�J�E���^�[
  Dim bRet As Boolean     '���f�����߂�l
  Dim iRet As Integer     '���b�Z�[�W�{�b�N�X�߂�l
  Dim iSettei_Flag As Boolean '�ݒ�ύX�t���O
  
  On Error Resume Next
   
  iSettei_Flag = False
  
  '�Ď��ՃA�v���N���`�F�b�N���s���B
  lngAplSts = CheckAppStart(PROC_KANRI)
  If lngAplSts <> 0 Then
       
     '�Ď��ՋN����:�����ݒ�G���A�X�V�������s��
      For iCnt = 0 To mintMaxIndex
          '�G���AID�擾
          iKansiId = ctlSetteiButton1(iCnt).pID
          '�X�V�l���擾
          iSetSts = ctlSetteiButton1(iCnt).pSetUp
          '�ݒ�t���O���擾
          iSetFlag = ctlSetteiButton1(iCnt).Settei_Flag
          
          If iSetFlag = KANSI_SETTEI Then
             bRet = Area_Updata(iKansiId, iSetSts)
          End If
          If bRet = True Then
             iSettei_Flag = True
             '���[�����M�������s���B
             psSendMail (iCnt)
          Else
             '�X�V�����ُ펞�F��������(�ُ�I��)�|�b�v�A�b�v��ʕ\��
             iRet = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f��������")
             Exit Sub
          End If
          iSettei_Flag = False
       Next
    Else
       '�Ď��Ֆ��N�����F�����ݒ�t�@�C�����l�擾
        For iCnt = 0 To mintMaxIndex
            '�G���AID�擾
            iKansiId = ctlSetteiButton1(iCnt).pID
            '�X�V�l���擾
            iSetSts = ctlSetteiButton1(iCnt).pSetUp
            '�ݒ�t���O���擾
            iSetFlag = ctlSetteiButton1(iCnt).Settei_Flag
          
            If iSetFlag = KANSI_SETTEI Then
                bRet = Settei_Updata(iKansiId, iSetSts)
            End If
         Next
           
         If bRet = False Then
            '�X�V�����ُ펞�F��������(�ُ�I��)�|�b�v�A�b�v��ʕ\��
            iRet = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f��������")
            Exit Sub
         End If
    End If
    
    iSettei_Flag = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Settei_Updata
'//  �@�\����  : �Ď��ݒ�t�@�C���X�V����
'//  �@�\�T�v  : �Ď��ݒ�t�@�C���X�V�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Long�@�@ iKansiId�@[IN]�Ď��ݒ�ID
'//              Integer�@iSetSts   [OUT]�擾�l
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]��������
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-15   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function Settei_Updata(iKansiId As Long, iSetSts As Integer) As Boolean

    Dim iAreaSts As Integer       '�Ď��ݒ��Ԓl
    Dim lSts            As Long   '�֐��߂�l
    Dim udtAreaR255 As GATE_INFO  '�Ǎ��ݗp�G���A�i255�ݒ�p�j
    Dim lngSts As Long
    Dim lngLoop1 As Long
    Dim lngHandle As Long
    Dim FileName As String
    Dim lngRet As Long
    Dim bRet As Boolean
    Dim sSetteiFile As String

    On Error Resume Next

    '�Ď��ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(K_SETTEI_FILE, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�I�[�v���ُ펞�͎Q�ƕs�̂��ߎQ�ƈُ�
       '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Settei_Updata = False
       Exit Function
    End If

   '�Ď��ݒ�t�@�C���ǂݍ���
    bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Call CloseHandle(lngHandle)
       Settei_Updata = False
       Exit Function
    End If

   '�n���h���̃N���[�Y
    Call CloseHandle(lngHandle)

    'ID����
     lngSts = KansiSerchId(udtAreaR255, iKansiId)
     If lngSts >= 0 Then
        'ID���L�����ꍇ
        udtAreaR255.GateInfo(lngSts).bytDATA(0) = iSetSts
     Else
        ' �Y���h�c�����̏ꍇ�Q�ƈُ�
        '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
        Settei_Updata = False
        Exit Function
     End If

    '�Ď��ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(K_SETTEI_FILE, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�I�[�v���ُ펞�͎Q�ƕs�̂��ߎQ�ƈُ�
        '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
    End If

    '�Ď��ݒ�t�@�C��������
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       '�n���h���̃N���[�Y
       Call CloseHandle(lngHandle)
       Exit Function
    End If

   '�n���h���̃N���[�Y
     Call CloseHandle(lngHandle)

     Settei_Updata = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Area_Updata
'//  �@�\����  : �Ď��ݒ�G���A�X�V����
'//  �@�\�T�v  : �Ď��ݒ�G���A�X�V�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Long�@�@ iKansiId�@[IN]�Ď��ݒ�ID
'//              Integer�@iSetSts   [OUT]�擾�l
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]��������
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-15   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function Area_Updata(iId As Long, iSts As Integer) As Boolean
    
    On Error Resume Next

    Set Idinf_KansiSettei = New IdInfProc              '�Ď��ݒ�G���A
    '�Ď��ݒ�G���A���I�[�v������B
    Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
    Idinf_KansiSettei.IdOpen
    If Idinf_KansiSettei.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�ُ͈��Ԃ��B
       '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
       Exit Function
    End If
             
    '�Ď��ݒ�G���A���k�n�b�j����B
    Idinf_KansiSettei.IdLock
    If Idinf_KansiSettei.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�ُ͈��Ԃ��B
       '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
       Exit Function
    End If
              
    '�G���A�̓��e��ǂݍ��ށB
    Idinf_KansiSettei.id = iId
    Idinf_KansiSettei.IdGet
    If Idinf_KansiSettei.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�ُ͈��Ԃ��B
       '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
       Exit Function
    End If
               
    '�ݒ���e���擾
    Idinf_KansiSettei.SetIDSVR iSts
    If Idinf_KansiSettei.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�ُ͈��Ԃ��B
       '�u�Ď��ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Area_Updata = False
       Idinf_KansiSettei.IdFree
       Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
       Exit Function
     End If
     
     Idinf_KansiSettei.IdFree
     Set Idinf_KansiSettei = Nothing               '�Ď����u�ݒ�f�[�^�t�@�C��
    
     Area_Updata = True
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
Private Function KansiSerchId(udtArea255 As GATE_INFO, lngId As Long) As Long

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
        If lngId = lngWorkId Then                                  '�����H
            lngChkIndex = lngIndex                                  '�f�[�^���o����A�����I��
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngId < lngId) Then         '�f�[�^���\����������
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
'//  �֐�����  : SetEnableFalse
'//  �@�\����  : ��ʃ��b�N��������
'//  �@�\�T�v  : ��ʂ̃��b�N����������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    Dim intCount As Integer '�J�E���^�[
    
    On Error Resume Next
    
    ' �{�^���̓��͕s�Ƃ���
    For intCount = 0 To mintMaxIndex
         ctlSetteiButton1(intCount).Enabled = False
    Next
    
    '�m��t�FTrue(���b�N)����B
    cmd_Kakutei.Enabled = False
    
    '���j���[��ʂ֖߂�t�FFalse(���b�N)����B
    cmdReturn.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : SetEnableTrue
'//  �@�\����  : ��ʃ��b�N��������
'//  �@�\�T�v  : ��ʂ̃��b�N����������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    Dim intCount As Integer '�J�E���^�[
    
    On Error Resume Next
    
    ' �{�^���̓��͕s�Ƃ���
    For intCount = 0 To mintMaxIndex
         ctlSetteiButton1(intCount).Enabled = True
    Next
    
    '�m��t�FTrue(���b�N����)����B
    cmd_Kakutei.Enabled = True
    
    '���j���[��ʂ֖߂�t�FTrue(���b�N����)����B
    cmdReturn.Enabled = True
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psSendMail
'//  �@�\����  : ���M���[������
'//  �@�\�T�v  : �����ԍ��ɂ�著�M���[������ʂ���B
'//
'//              �^        ����      �@�Ӗ�
'//  ����      : iCnt�@�@�@�J�E���^�[  [IN]���M�ΏۃJ�E���^�[
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub psSendMail(iCnt As Integer)
    
    On Error Resume Next
    '�����ԍ�����
    If ctlSetteiButton1(iCnt).SHORI_NO = DANKI Then
       '�g�C�^�]���[�����M
       psSendDankiMail
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psSendDankiMail
'//  �@�\����  : �g�@�^�]�ݒ�ύX�ʒm���M����
'//  �@�\�T�v  : �g�@�^�]�ݒ�ύX�ʒm���M�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub psSendDankiMail()
   Dim udtMail     As MAIL_KANSI_SET_INF    '�����ݒ�w�����[�����M�G���A
   Dim intCnt      As Integer              '�J�E���^
   Dim bRet        As Boolean
   
   On Error Resume Next

    '���ʃw�b�_�ҏW
    udtMail.mlHeader.dwId = ML_ID_KANSI_SETTEI_INF
    udtMail.mlHeader.dwSize = MlSize.KANSI_SETTEI_INF
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequestType = MailKANSI_SET_Type.ML_DT_DANKI_UNTEN
    
    '���[�����M
    bRet = DssSendMail(MAIL_SLOT_SD, MlSize.KANSI_SETTEI_INF, udtMail.mlHeader)
    If bRet = True Then
       '�u�Ď��ݒ��ʁF�Ď��ݒ�ύX�ʒm���M����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_MAIL, KANSI_SETTEI_DANKIMAIL_SEND_OK, 0)
    Else
       '�u�Ď��ݒ��ʁF�Ď��ݒ�ύX�ʒm���M�ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_MAIL, KANSI_SETTEI_DANKIMAIL_SEND_ERROR, 0)
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : �^�C���A�b�v������
'//  �@�\�T�v  : ���[����M�^�C���A�b�v���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  
    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmKansiSetteiSub.Caption, False
    End If

End Sub
