VERSION 5.00
Begin VB.Form frmRmenteMenu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�����[�g�����e�i���X"
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
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   6120
      Top             =   5640
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "���D�@"
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
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "�h�b�l"
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
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�����[�g�����e�i���X"
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
Attribute VB_Name = "frmRmenteMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmRmenteMenu.frm
'//  �p�b�P�[�W���F�����[�g�����e�i���X���j���[���
'//
'//  �T�v�F�����[�g�����e�i���X���j���[���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.37�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private sTOOLPass As String

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �����[�g�����e�i���X���j���[���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}�A�N��
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
Private Sub Form_Activate()
On Error Resume Next
     pfFormActive (hwnd)
    '�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �����[�g�����e�i���X���j���[���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}�A��~
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
Private Sub Form_Deactivate()
On Error Resume Next
    '�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �����[�g�����e�i���X���j���[���(���[�h��)
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
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next

   '�u�Ӱ�����ݽ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'IDU�k�ރ`�F�b�N
    psIDUCheck

    If pbIDUSts = 1 Then
      'IDU�Ɩ���\��
       cmdFixedExe(1).Visible = False
    End If
    'V1.3.0.1 ADD START
    '���C����M�p�̃��C����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    'V1.3.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �e�t����������
'//  �@�\�T�v  : �e�t���̉�ʂɑJ�ړ����s���B
'//              �u�����v�u����h�b�|�l�v
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@ [IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
  On Error Resume Next
  Dim lRetVal As Double     'Shell�֐��߂�l

  Select Case Index
        Case 0                                 '����
           '�u�Ӱ�����ݽ��ʁF�����t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_JIKAI_BUTTOM, 0)
           Load frmRMente
           frmRMente.Show 1
        Case 1                                '����IC-M
           '�u�Ӱ�����ݽ��ʁF����IC-M�t�����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_ICM_BUTTOM, 0)
           fMakeICMGLTFile
           psICMRMenteTool
           If sTOOLPass = "" Then
              Exit Sub
           Else
              '����IC-M�c�[���N��
            lRetVal = Shell(sTOOLPass, vbNormalFocus)
          End If
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
    
    '�u�Ӱ�����ݽ��ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_MENU_GAMEN_END, 0)
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psICMRMenteTool
'//  �@�\����  : ����IC-M�̃����[�g�����e�i���X�c�[���p�X���擾����
'//  �@�\�T�v  : ����IC-M�̃����[�g�����e�i���X�c�[���p�X���擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-01 REVISED BY [TCC] T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή��y�c����54�A�Ď�D-154�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub psICMRMenteTool()
 
    Dim sPath As String * MAX_PATH_SIZE
    Dim iRet As Integer
    
    Dim sMyPath As String               'EG20 V2.0.1.1 ADD
    
    On Error Resume Next
    
    ' HOSHU.INI��蔻��IC-M�c�[���p�X���擾����B
    iRet = GetPrivateProfileString(KANSI_HOSHU_ICM_RMENTE_SEC, _
                                    KANSI_HOSHU_ICM_RMENTE_KEY, _
                                    DEFAILT, sPath, Len(sPath), _
                                    HOSHU_FILE)

    sMyPath = Replace(sPath, Chr(0), "")
      
      If iRet = 0 Then
        sTOOLPass = ""
      Else
'        sTOOLPass = sPath              'EG20 V2.0.1.1 DEL
        sTOOLPass = sMyPath             'EG20 V2.0.1.1 ADD
      End If
      
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeICMGLTFile
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
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.37�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeICMGLTFile() As Integer
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
    Dim szIniFilePath As String     ' INI�t�@�C���p�X   ' EG20 V3.3.0.1�y����TR-No.37�z�ǉ�

    On Error Resume Next
    MkDir PATH_RMENTE_ICM_DEN   '�����p�d�S�t�H���_���쐬����B�iGLT�t�@�C���p�j
    'GLT�t�@�C�����J���B�t�@�C�������݂��Ȃ���ΐV�K�ɍ쐬�����B
    On Error GoTo ErrorHandlerGLTFile
    intGLTFileNo = FreeFile        ' ���g�p�̃t�@�C���ԍ����擾����B
    Open ICM_GLT_FILE_FULLPASS For Output As #intGLTFileNo     ' GLT�t�@�C�����J���B

    For iGate = CNT_MIN To MAX_GATE_NO - 1
' EG20 V3.3.0.1�y����TR-No.37�z�폜�J�n
'      '�������D�@���擾
'      sKeyName = "gate" & Format(iGate + 1, "00")
'      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
'                                     sKeyName, _
'                                     DEFAILT, sGateData, Len(sGateData), _
'                                     PATH_GATE_FILE)
' EG20 V3.3.0.1�y����TR-No.37�z�폜�I��
' EG20 V3.3.0.1�y����TR-No.37�z�ǉ��J�n
        ' IDU��ICM.INI������D�@�����擾
        szIniFilePath = PATH_IDU_APP & IDU_ICM_FILE
        sKeyName = "icm" & Format(iGate + 1, "00")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ICM, _
                                    sKeyName, _
                                    DEFAILT, sGateData, Len(sGateData), _
                                    szIniFilePath)
' EG20 V3.3.0.1�y����TR-No.37�z�ǉ��I��
      
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
         sGoukiNo = "0" & Trim(sFData(1)) & "���@"
      Else
         sGoukiNo = Trim(sFData(1)) & "���@"
      End If
        
' EG20 V3.3.0.1�y����TR-No.37�z�폜�J�n
'      If Trim(sFData(4)) <> "��" Then
'         'Gate.ini�t�@�C���̍��@�ԍ��\�������AIP�A�h���X��GLT�t�@�C���ɏ������ށB
'         Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(14))
'      End If
' EG20 V3.3.0.1�y����TR-No.37�z�폜�I��
' EG20 V3.3.0.1�y����TR-No.37�z�ǉ��J�n
    If Trim(sFData(5)) <> "��" Then
        'ICM.ini�t�@�C���̍��@�ԍ��\�������AIP�A�h���X��GLT�t�@�C���ɏ������ށB
        Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(7))
    End If
' EG20 V3.3.0.1�y����TR-No.37�z�ǉ��I��
              
    Next
    
    'GLT�t�@�C�������B
    Close #intGLTFileNo
    
    fMakeICMGLTFile = 0    '����I��
    Exit Function

ErrorHandlerGateIni:
   '�u�������D�@�Ӱ�����ݽ��ʁF�t�@�C���A�N�Z�X�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeICMGLTFile = 1
   'GLT�t�@�C�������B
   Close #intGLTFileNo
   Exit Function
ErrorHandlerGLTFile:
   '�u�������D�@�Ӱ�����ݽ��ʁF�t�@�C���A�N�Z�X�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeICMGLTFile = 2
   'GLT�t�@�C�������B
   Close #intGLTFileNo

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
        AppActivate frmRmenteMenu.Caption, False
        pfFormActive (frmRmenteMenu.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

