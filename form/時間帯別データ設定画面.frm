VERSION 5.00
Begin VB.Form frmTimeDataSettei 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ғ��E�����e�f�[�^���W�i�����㎩�����D�@�j"
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
   Begin VB.Timer TmrKakunin 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   8160
   End
   Begin VB.Frame Frame2 
      Caption         =   "�f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   10815
      Begin VB.CommandButton cmdDataClear 
         Caption         =   "���ԑѕʃf�[�^�N���A"
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
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ݒ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   10815
      Begin VB.CheckBox ChkSndSet 
         BackColor       =   &H0080FF80&
         Caption         =   "���M"
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
         TabIndex        =   7
         Top             =   840
         Value           =   1  '����
         Width           =   1215
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
         Height          =   855
         Left            =   9000
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "�@�@�@�@�@���M�ݒ�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   360
      Top             =   8160
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "    �V�X�e���ݒ�      ��ʂ֖߂�"
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
      Left            =   9360
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "���ԑѕʃf�[�^�ݒ�"
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
Attribute VB_Name = "frmTimeDataSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'*
'*    �錾����   : ���ԑѕʃf�[�^�ݒ�i�����㎩�����D�@�j
'*   Ӽޭ�يT�v  : ���ԑѕʃf�[�^�ݒ��ʂ̃t�H�[�����W���[��
'*
'*     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'*                 �E�t�F�[�Y�Q�Ή��yMainte_05_03�z
'*     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'*                 2014�N�x�{�� �yEG20_KANSI05_01�z
'*     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000     '���C���^�C�}�̃C���^�[�o���l

Private mintMaxIndex As Integer
Private mintID As Integer           '�G���AID
Private Type SHUSHU_STATUS
    intStatus As Integer    '�X�e�[�^�X
    strCaption As String    '�{�^������
    strColor As String      '�{�^���F
    IntValue As Integer     '�������
End Type
Private mudtBtn_Status() As SHUSHU_STATUS

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : �m��{�^�����������ꂽ���̃C�x���g�v���V�[�W��
'     ����      : ���M�ݒ�G���A���X�V����B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(EG20 V2.1.0.1) 2011-12-08   CODED   BY [TCC] M.Matsumoto
'                               �y��-221�Ή��z
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdKakutei_Click()

    Dim Idinf_KansiSettei As IdInfProc
    Dim lngErrCode As Long                  '�G���[�R�[�h
    Dim lngRet As Long
    Dim bRet As Boolean
    
    On Error Resume Next
    
    '�Ď��ՋN����
    If CheckAppStart(PROC_KANRI) <> 0 Then
        Set Idinf_KansiSettei = New IdInfProc             '�Ď����u�ݒ�G���A
        '�Q��(�����ʐM���)�G���A����ݒ�
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_SEND_SET_ERROR, lngErrCode)
           Exit Sub
        End If
    
        '�G���AID�̐ݒ�l���X�V
        Idinf_KansiSettei.IdLock
        Idinf_KansiSettei.id = mintID
        Idinf_KansiSettei.DataType = ID_TYPE.Flag
'        Call Idinf_KansiSettei.SetIDSVR(CInt(ChkSndSet.Value))     'EG20 V2.1.0.1 DEL �y��-221�Ή��z
        Call Idinf_KansiSettei.SetIDSVR(CInt(ChkSndSet.Tag))        'EG20 V2.1.0.1 ADD �y��-221�Ή��z
    
        If Idinf_KansiSettei.Errsts <> 0 Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_SEND_SET_ERROR, lngErrCode)
        Else
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_SEND_SET_OK, 0)
        End If
        Idinf_KansiSettei.IdFree
    
    '�Ď��Ֆ��N����
    Else
'        bRet = gspfSetKansiSts(mintID, CInt(ChkSndSet.Value))       'EG20 V2.1.0.1 DEL �y��-221�Ή��z
        bRet = gspfSetKansiSts(mintID, CInt(ChkSndSet.Tag))         'EG20 V2.1.0.1 ADD �y��-221�Ή��z
    End If
    
    cmdDataClear.Enabled = False
    cmdKakutei.Enabled = False
    cmdReturn.Enabled = False
    ChkSndSet.Enabled = False
    
    '�m�F�{�^�������p�^�C�}���쓮������
    tmrKakunin.Interval = 1000       '�{�^�������p�^�C�}���Ԑݒ�
    tmrKakunin.Enabled = True
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : ���ԑѕʃf�[�^�ݒ��ʂ����[�h���ꂽ���̃C�x���g�v���V�[�W��
'     ����      : ���C����M�p�̃^�C�}�l��ݒ肷��B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Load()

    Dim intFileNumber As Integer            '�t�@�C���ԍ�
    Dim strFileName As String               '�t�@�C����
    Dim strItmNum As String                 '�ݒ荀�ڐ�
    Dim strTemp As String
    Dim intCount As Integer                 '���[�v�J�E���^
    Dim intStatus As Integer                '�G���AID�l
    Dim lngErrCode As Long                  '�G���[�R�[�h
    Dim Idinf_KansiSettei As IdInfProc
    Dim intNum As Integer
    Dim lSts            As Long             '�֐��߂�l
    Dim udtAreaR255 As GATE_INFO            '�Ǎ��ݗp�G���A�i255�ݒ�p�j
    Dim lngSts As Long
    Dim lngHandle As Long
    Dim lngRet As Long
    Dim bRet As Boolean
    
    On Error Resume Next
        
    tmrKakunin.Enabled = False
 
    '�u���ԑѕʃf�[�^�ݒ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_START, 0)
   
    '���C����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���g�p�̃t�@�C���ԍ����擾���܂��B
    intFileNumber = FreeFile

    '�ݒ���t�@�C������ݒ肷��B
    strFileName = TIMEDATA_STATUS_FILE

    '�ݒ���t�@�C�����I�[�v������B
    If strFileName <> "" Then
        Open strFileName For Input As #intFileNumber
    End If

    For intCount = 0 To 2

        '�ݒ���t�@�C�����ɐݒ肳��Ă���t�ݒ�t�@�C����ǂށB
        Input #intFileNumber, strItmNum, strTemp, strTemp, strTemp

        '�ő�R���g���[������ϐ��ɐݒ肷��B
        If intCount = 1 Then
            '�G���AID
            mintID = CInt(strItmNum)
        ElseIf intCount = 2 Then
            '���ڐ�
            mintMaxIndex = CInt(strItmNum) - 1
        End If
    Next

    ReDim mudtBtn_Status(mintMaxIndex)

    For intCount = 0 To mintMaxIndex
        '�ݒ���t�@�C�����ɐݒ肳��Ă���t�ݒ�t�@�C����ǂށB
        With mudtBtn_Status(intCount)
            Input #intFileNumber, .intStatus, .strCaption, .strColor, .IntValue
        End With
    Next

    Close #intFileNumber

    strFileName = Dir(K_SETTEI_FILE)
    If strFileName = "" Then
       '�Ď��ݒ�t�@�C�����Ȃ��ꍇ
       strFileName = SHOKI_K_SETTEI_FILE
    Else
       '�Ď��ݒ�t�@�C��������ꍇ
       strFileName = K_SETTEI_FILE
    End If
    
    '�Ď��ՋN�����͊Ď����u�ݒ�G���A���瑗�M�ݒ���擾����
    If CheckAppStart(PROC_KANRI) <> 0 Then
        Set Idinf_KansiSettei = New IdInfProc             '�Ď����u�ݒ�G���A
        '�Q��(�����ʐM���)�G���A����ݒ�
        Idinf_KansiSettei.ProcMode = DATA_ID.Data_Id_KansiSettei
        Idinf_KansiSettei.IdOpen
        If Idinf_KansiSettei.Errsts <> 0 Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_START, 0)
            Exit Sub
        End If
    
        '�G���AID�̐ݒ�l���擾
        Idinf_KansiSettei.IdLock
        Idinf_KansiSettei.id = mintID
        Idinf_KansiSettei.IdGet
        intStatus = Idinf_KansiSettei.DataArea(0)
        Idinf_KansiSettei.IdFree
        
        cmdDataClear.Enabled = True
    '�Ď��Ֆ��N���̏ꍇ
    Else
    
        '�Ď��ݒ�t�@�C�����I�[�v��
        lngHandle = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
        '�t�@�C���I�[�v��������ɍs��ꂽ���H
        If lngHandle = INVALID_HANDLE_VALUE Then
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, TIMEDATA_GAMEN_START, 0)
            Exit Sub
        End If
        
        '�Ď��ݒ�t�@�C���ǂݍ���
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)

        '�n���h���̃N���[�Y
        Call CloseHandle(lngHandle)
        
        'ID����
        lngSts = KansiSerchId(udtAreaR255, CLng(mintID))
        If lngSts >= 0 Then
           'ID���L�����ꍇ
           intStatus = ChgData(udtAreaR255.GateInfo(lngSts))         '�f�[�^�ϊ�
        Else
          ' �Y���h�c�����̏ꍇ�Q�ƈُ�
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_ELSE
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, TIMEDATA_GAMEN_START, 0)
            Exit Sub
        End If
        
        cmdDataClear.Enabled = False
    End If
    
    '�擾�����l��Tag�l�ɐݒ�
    ChkSndSet.Tag = CStr(intStatus)
    
    'Tag�l�ƈ�v���镶���A�F�A������Ԃɂ���
    For intCount = 0 To UBound(mudtBtn_Status)
        If mudtBtn_Status(intCount).intStatus = CInt(ChkSndSet.Tag) Then
            ChkSndSet.Caption = mudtBtn_Status(intCount).strCaption
            ChkSndSet.BackColor = mudtBtn_Status(intCount).strColor
            ChkSndSet.Value = mudtBtn_Status(intCount).IntValue
        End If
    Next intCount
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : ���ԑѕʃf�[�^�ݒ��ʂ� �\�����ꂽ���̃C�x���g�v���V�[�W��
'     ����      : �u���[����M�p�^�C�}�v���N������B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Activate()

    On Error Resume Next
    
    '���[����M�p�^�C�}���N������
    tmrMail.Enabled = True
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : ���ԑѕʃf�[�^�ݒ��ʂ��������ꂽ���̃C�x���g�v���V�[�W��
'     ����      : �u���[����M�p�̃^�C�}�v��j������B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub Form_Deactivate()

    On Error Resume Next
    
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : �u�V�X�e���ݒ��ʂɖ߂�v�t���N���b�N���ꂽ���̃C�x���g�v���V�[�W��
'     ����      : ��ʂ���������B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdReturn_Click()

    On Error Resume Next
   '�u���ԑѕʃf�[�^�ݒ��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_END, 0)
 
    '����ʂ������B
    Unload Me
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'
'  �T�v     :�m�F�{�^�������p�^�C�}�C�x���g������
'  ����     :�m�F�{�^�������p�^�C�}�C�x���g�������̏������s���B
'            �m�F�{�^���A���̑��{�^���̐F�������F���猳�̐F�ɖ߂�
'  ���Ұ�   :
'
'    ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'    REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub tmrKakunin_Timer()
    
    On Error Resume Next

    '�m�F�{�^�������p�^�C�}���~����
    tmrKakunin.Enabled = False                   '�m�F�{�^�������p�^�C�}��~
    tmrKakunin.Interval = 0                      '�m�F�{�^�������p���ԏ�����
    
    cmdDataClear.Enabled = True
    cmdKakutei.Enabled = True
    cmdReturn.Enabled = True
    ChkSndSet.Enabled = True

End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : �u���[����M�p�^�C�}�v���^�C���A�b�v�������̃C�x���g�v���V�[�W��
'     ����      : ���[����M�������s���B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(EG20 V8.1.0.1) 2014-06-05   REVISED BY [TCC] S.Kuroda
'                 2014�N�x�{�� �yEG20_KANSI05_01�z
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub tmrMail_Timer()

    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmTimeDataSettei.Caption, False
        pfFormActive (frmTimeDataSettei.hwnd)           ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If

End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : ���M�ݒ�{�^���������̃C�x���g�v���V�[�W��
'     ����      : ���M�ݒ��؂�ւ���B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-16   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(EG20 V2.1.0.1) 2011-12-08   CODED   BY [TCC] M.Matsumoto
'                               �y��-221�Ή��z
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub ChkSndSet_Click()

    Dim intCount As Integer

'    ChkSndSet.Tag = CStr(ChkSndSet.Value)   'EG20 V2.1.0.1 DEL �y��-221�Ή��z
    
    'Tag�l�ƈ�v���镶���A�F�A������Ԃɂ���
    For intCount = 0 To UBound(mudtBtn_Status)
'        If mudtBtn_Status(intCount).intStatus = CInt(ChkSndSet.Tag) Then   'EG20 V2.1.0.1 DEL �y��-221�Ή��z
        If mudtBtn_Status(intCount).IntValue = CInt(ChkSndSet.Value) Then   'EG20 V2.1.0.1 ADD �y��-221�Ή��z
            ChkSndSet.Caption = mudtBtn_Status(intCount).strCaption
            ChkSndSet.BackColor = mudtBtn_Status(intCount).strColor
            ChkSndSet.Value = mudtBtn_Status(intCount).IntValue
            ChkSndSet.Tag = mudtBtn_Status(intCount).intStatus              'EG20 V2.1.0.1 ADD �y��-221�Ή��z
        End If
    Next intCount

End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : ���ԑѕʃf�[�^�N���A�{�^���������̃C�x���g�v���V�[�W��
'     ����      : ���M�ݒ��؂�ւ���B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-16   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub cmdDataClear_Click()

    Dim iResponse As Integer
    
    '�u���ԑѕʃf�[�^�ݒ��ʁF���ԑѕʃf�[�^�N���A�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, TIMEDATA_GAMEN_CLEAR_BUTTOM, 0)
   
    '�u���ԑѕʃf�[�^�N���A�v�|�b�v�A�b�v��\��
    iResponse = MsgBox("���ԑѕʃf�[�^���N���A���܂�����낵���ł����H", _
                        vbOKCancel, "�m�F")
    
    '�n�j�t�������ꂽ��
    If iResponse = vbOK Then
        '���ԑѕʃf�[�^�N���A���t�H�[�������[�_���E�B���h�E�ŕ\������B
        frmClearCyu.Show vbModal
    End If
    
End Sub

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


