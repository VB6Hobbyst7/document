VERSION 5.00
Begin VB.Form frmICUnkai_Type1 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�^���f�[�^�c�k�k���"
   ClientHeight    =   9000
   ClientLeft      =   1380
   ClientTop       =   1905
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   14.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
   WindowState     =   2  '�ő剻
   Begin VB.CommandButton cmdVer 
      Caption         =   "�h�b         �^���f�[�^�v��"
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
      Left            =   9600
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.ListBox lstKan 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   9135
   End
   Begin VB.Timer tmrMail 
      Left            =   9360
      Top             =   7320
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "���C         �^���f�[�^�v��"
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
      Left            =   9600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "���j���[��ʂ֖߂�"
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
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
   Begin VB.ListBox lstKan 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9135
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�^���f�[�^�c�k�k���"
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
      TabIndex        =   5
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmICUnkai_Type1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'*
'*   Ӽޭ�يT�v  : �^���f�[�^DLL��ʂ̃t�H�[�����W���[��
'*
'*      ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'*      REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'*                  �d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'*      REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Option Explicit
'���\�[�X�萔

'�I�𒆃��\�[�X��� =0�F�W�v�A=1:��
Dim iSelResource As Integer

Private Const MN_MAIL_INTERVAL = 1000 '���C���^�C�}�̃C���^�[�o���l

'���O�o�̓��b�Z�[�W
Private LogMsgStart(1) As String
Private LogMsgMidst(1) As String
Private LogMsgEnd(1) As String

Private gIndex As Integer           '�������ꂽ�t��INDEX
Private gSendMailKishu As Long      '���M���[���F�@��
Private gSendMailShubetsu As Long   '���M���[���F���
Private gSendMailShosai As Long     '���M���[���F�ڍ�

Private Enum DLL_DATA       ' DLL�Ώ�
    JIKI = 0                ' ���C�F�O
    IC                      ' �h�b�F�P
End Enum
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  �T�v     : �^���f�[�^DLL���ʕ\��
'  ����     : �^���f�[�^DLL�̌��ʂ�\�����镶�����쐬����B
'  ���Ұ�   : strMsg, I ,string, �FDLL���ʕ\������
'           :  �߂�l,O ,string, �F���X�g�{�b�N�X�\������
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Function fMakeListbox(strMsg_1 As String, strMsg_2 As String) As String
    Dim strRet As String

    strRet = vbNullString

    strRet = Format(Now, "YYYY/MM/DD   HH:MM:SS")

    strRet = strRet & Space(3) & strMsg_1 & Space(3) & strMsg_2

    fMakeListbox = strRet

End Function

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  �T�v     : �u�ێ��ʂɖ߂�v�{�^���������̃C�x���g�v���V�[�W��
'  ����     : �^���f�[�^DLL���ʕ\����ʂ����B
'  ���Ұ�   :
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              �d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub cmdReturn_Click()
   
   'V2.2.0.1 ADD START
   '�u���j���[�t�֖߂�v�t�����u�^���f�[�^�c�k�k��ʏ����v
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UNCHINDATA_DLL_GAMEN_END, 0)
   'V2.2.0.1 ADD END
   
    '�^���f�[�^DLL���ʕ\����ʂ����
    Unload Me
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'   �T�v    : �^���f�[�^�v���{�^���������̃C�x���g�v���V�[�W��
'   ����    : �^���f�[�^DLL�v�����ă}�ɑ��M����B
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              �d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub cmdVer_Click(Index As Integer)
    Dim lngMSlot_KM As Long             '�ă}�̃��[���X���b�g�n���h��
    Dim lngRet As Long                  '�߂�l
    Dim udtMail As MAIL_UNCHIN_DLL_REQ  '���M���[��
    Dim iCnt As Integer                 '�J�E���^
    Dim iDataSts As Integer             '�f�[�^�X�e�[�^�X�@'V2.2.0.1�@ADD
    
    '�������ꂽ�t�̃C���f�b�N�X�l��ۑ�
    gIndex = Index

    'DLL�J�n�����X�g�ɕ\������B
    lstKan(gIndex).AddItem fMakeListbox("����", LogMsgStart(gIndex))
    
    ' �{�^���������s�ɂ���B
    For iCnt = 0 To cmdVer.UBound
        cmdVer(iCnt).Enabled = False
    Next

    ' �ێ��ʂ֖߂�t�������s�ɂ���B
    cmdReturn.Enabled = False
'V2.2.0.1 DEL START
'    '�ă}�ւ̑��M���[���X���b�g���I�[�v������B
'    lngMSlot_KM = DssMailOpen(MAIL_SLOT_KANMA)
'    If lngMSlot_KM <> INVALID_HANDLE_VALUE Then   '�ُ�
'
'        gSendMailKishu = ML_DT_UNCHINDLL_KISHU     '���M���[���F�@��
'        gSendMailShosai = MlUnchinDllSHUBETSU.ML_DT_UNCHIN_NEW   '���M���[���F�ڍ�
'        '�f�[�^���
'        If gIndex = DLL_DATA.JIKI Then      '���C�^���c�k�k
'            gSendMailShubetsu = MlUnchinDllData.ML_DT_UNCHIN_ICHI_FUKU     '�P���p�{�������p
'        Else                                '�h�b�^���c�k�k
'            gSendMailShubetsu = MlUnchinDllData.ML_DT_UNCHIN_IC_HAN        '�h�b�^���{����v���O����
'        End If
'
'        '�ă}�ɑ΂��ĉ^��DLL�v���𑗐M����B
'        udtMail.mlHeader.dwId = ML_ID_UNTIN_REQ       '���[���h�c�F�^���f�[�^�c�k�k�v���i���X�V�R�j
'        udtMail.mlHeader.dwSize = MlSize.UNTIN_REQ    '���[���T�C�Y�F�Q�W
'        udtMail.mlHeader.dwProid = RHOSYU_ID          '���M���h�c�F�ێ�i���P�P�j
'        udtMail.mlHeader.dwSubArea = 0                '�⏕���
'        udtMail.dwKishu = gSendMailKishu                '�@��F�������D�@�i���P�j
'        udtMail.dwData = gSendMailShubetsu              '�f�[�^���
'        udtMail.dwSyosai = gSendMailShosai              '��ʏڍׁF�V�^��
'
'        lngRet = DssMailWrite(lngMSlot_KM, MlSize.UNTIN_REQ, udtMail.mlHeader)
'
'        '�ă}�ւ̑��M���[���X���b�g���N���[�Y����B
'        lngRet = DssMailClose(lngMSlot_KM)
'
'    End If
'V2.2.0.1 DEL END
'V2.2.0.1 ADD START
 psGetData_Type iDataSts
 
 '�Ď��ՋN�������[�����M
 gSendMailKishu = ML_DT_UNCHINDLL_KISHU     '���M���[���F�@��
 gSendMailShosai = MlUnchinDllSHUBETSU.ML_DT_UNCHIN_NEW   '���M���[���F�ڍ�
 '�f�[�^���
 If gIndex = DLL_DATA.JIKI Then      '���C�^���c�k�k
    gSendMailShubetsu = iDataSts                                    '�P���p�{�������p
 Else                                '�h�b�^���c�k�k
    gSendMailShubetsu = MlUnchinDllData.ML_DT_UNCHIN_IC_HAN        '�h�b�^���{����v���O����
 End If
        
 '�ă}�ɑ΂��ĉ^��DLL�v���𑗐M����B
  udtMail.mlHeader.dwId = ML_ID_UNTIN_REQ       '���[���h�c�F�^���f�[�^�c�k�k�v���i���X�V�R�j
  udtMail.mlHeader.dwSize = MlSize.UNTIN_REQ    '���[���T�C�Y�F�Q�W
  udtMail.mlHeader.dwProid = RHOSHU_ID          '���M���h�c�F�ێ�i���P�P�j
  udtMail.mlHeader.dwSubArea = 0                '�⏕���
  udtMail.dwKishu = gSendMailKishu                '�@��F�������D�@�i���P�j
  udtMail.dwData = gSendMailShubetsu              '�f�[�^���
  udtMail.dwSyosai = gSendMailShosai              '��ʏڍׁF�V�^��
    
  '���[�����M
  lngRet = DssSendMail(MAIL_SLOT_KANMA, MlSize.UNTIN_REQ, udtMail.mlHeader)
  If lngRet = False Then
     '���M�ُ�
     Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, UNCHINDATA_DLL_CMD_ERROR, 0)
     
     ' �{�^���������\�ɂ���B
     For iCnt = 0 To cmdVer.UBound
         cmdVer(iCnt).Enabled = True
     Next
            
     ' �ێ��ʂ֖߂�t�������s�ɂ���B
     cmdReturn.Enabled = True
            
     AppActivate frmICUnkai_Type1.Caption, False
  Else
    '���M����
     Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, UNCHINDATA_DLL_CMD_OK, 0)
  End If
'V2.2.0.1 ADD END
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  �T�v      : �^���f�[�^DLL���ʕ\����ʂ��A�N�e�B�u�ɂȂ������̃C�x���g�v���V�[�W��
'  ����      : ���C����M�p�̃^�C�}���N������B
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub Form_Activate()
    '���[����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  �T�v     : �^���f�[�^DLL���ʕ\����ʂ��ި��è�ނɂȂ������̲������ۼ��ެ
'  ����     : ���[����M�p�̃^�C�}���~�߂�B
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub Form_Deactivate()
    '���[����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub
'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  �T�v     : �^���f�[�^DLL���ʕ\����ʖʂ����[�h���ꂽ���̲������ۼ��ެ
'  ����     : �����������s��
'  ���Ұ�   :
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              �d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub Form_Load()

    Dim iCnt As Integer

    '���X�g�\��������
    LogMsgStart(0) = "���C�^���f�[�^�c�k�k�J�n"
    LogMsgStart(1) = "�h�b�^���f�[�^�c�k�k�J�n"

    LogMsgMidst(0) = "���C�^���f�[�^�c�k�k��"
    LogMsgMidst(1) = "�h�b�^���f�[�^�c�k�k��"

    LogMsgEnd(0) = "���C�^���f�[�^�c�k�k�I��"
    LogMsgEnd(1) = "�h�b�^���f�[�^�c�k�k�I��"
    
    'DLL���ʕ\���p�̃��X�g�{�b�N�X���N���A����B
    For iCnt = 0 To lstKan.UBound
        lstKan(iCnt).Clear
    Next

    '���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    'V2.2.0.1 ADD START
    '��ʕ\���u�^���f�[�^�c�k�k��ʕ\���v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UNCHINDATA_DLL_GAMEN_START, 0)
    'V2.2.0.1 ADD END
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'
'  �T�v     : ���[����M�p�^�C�}���^�C���A�b�v�������̃C�x���g�v���V�[�W��
'  ����     : ��M���[���̓��e�Ɋ�Â�����������B
'  ���Ұ�   :
'
'   ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'              �d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'   REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Sub tmrMail_Timer()
    Dim lngLen As Long                      '���C���T�C�Y
    Dim bRet As Boolean                     '�߂�l
    Dim uMail As MAIL_LGMINF_INF            '���C��
    Dim iCnt As Integer                     '�J�E���^

    '���[����M
    Do Until fDssMailReadMN(plMSlot_MN, uMail) <= 0
        lngLen = uMail.mlHeader.dwSize    '���C���T�C�Y�l�l��ݒ�B
       Select Case uMail.mlHeader.dwId   '���[���h�c
        '�u�v���Z�X�I���w���v����M�����ꍇ
        Case ML_ID_PROEND_ORD
            '���[����M���O�o��
            'Call dllWriteMailLog(plMSlot_LG, "�v���Z�X�I���w��", uMail, lngLen)  'V2.2.0.1 DEL
            'V2.2.0.1 ADD START
            '�u�v���Z�X�I���w����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            'V2.2.0.1 ADD END
            '�����I���������s��
            pfAbortProc

        '�u�^���f�[�^DLL�����ʒm�v����M�����ꍇ
        Case ML_ID_UNTIN_INF
            '���[����M���O�o��
            '���ʓ��e�Ɋ�Â��������s���B
            Dim i As Integer
            For i = 0 To 83
                Debug.Print "byTxtDat(" & i & ") : " & uMail.stMonFmt.byTxtDat(i)
            Next

            '��M���[���̃f�[�^�m�F���s���B�@��A�f�[�^��ʁA��ʏڍ�
            If uMail.stMonFmt.byTxtDat(0) <> gSendMailKishu Or _
                uMail.stMonFmt.byTxtDat(4) <> gSendMailShubetsu Or _
                uMail.stMonFmt.byTxtDat(8) <> gSendMailShosai Then

                '���[���ُ��M���O�o��
                'Call dllErrorMailLog(plMSlot_LG, uMail, lngLen) 'V2.2.0.1 DEL
                'V2.2.0.1 ADD START
                '�u���[���ُ��M�v���O�o��
                Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_RECV_ERROR, 0)
                'V2.2.0.1 ADD END
                '�������Ȃ��ŏI��
                Exit Sub
            End If

            'V2.2.0.1 ADD START
            '�u�^���f�[�^�c�k�k�����ʒm�v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, UNCHINDATA_DLL_END_REQ, 0)
            'V2.2.0.1 ADD END

            Select Case uMail.stMonFmt.byTxtDat(12)
                Case 0
                    '�e�X�g����I����\������B
                    lstKan(gIndex).AddItem fMakeListbox("����", LogMsgEnd(gIndex))
                Case 1
                    '�e�X�g�ُ�I����\������B
                    lstKan(gIndex).AddItem fMakeListbox("�ُ�", LogMsgEnd(gIndex))
                Case 2
                    '�e�X�g���s�s�\��\������B
                    lstKan(gIndex).AddItem fMakeListbox("�ُ�", LogMsgMidst(gIndex))
            End Select

            ' �{�^���������\�ɂ���B
            For iCnt = 0 To cmdVer.UBound
                cmdVer(iCnt).Enabled = True
            Next
            
            ' �ێ��ʂ֖߂�t�������s�ɂ���B
            cmdReturn.Enabled = True
            
            '�ܕԂ��e�X�g��ʂ��A�N�e�B�u�ɂ���B
            'AppActivate frmICUnkai.Caption, False 'V2.2.0.1 DEL
            AppActivate frmICUnkai_Type1.Caption, False 'V2.2.0.1 ADD

        '�ێ��ʃA�N�e�B�u�\���̏ꍇ
        'Case ML_ID_HOSYU_ACTIVE_REQ 'V2.2.0.1 DEL
        Case ML_ID_HOSHU_ACTIVE_REQ 'V2.2.0.1 ADD
            'V2.2.0.1 DEL START
            '���[����M���O�o��
            'Call dllWriteMailLog(plMSlot_LG, "�ێ��ʃA�N�e�B�u�\��", uMail, lngLen)
            'V2.2.0.1 DEL END
            'V2.2.0.1 ADD STRT
            '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
            'V2.2.0.1 ADD END
            '�ܕԂ��e�X�g��ʂ��A�N�e�B�u�ɂ���B
            'AppActivate frmICUnkai.Caption, False 'V2.2.0.1 END
            AppActivate frmICUnkai_Type1.Caption, False 'V2.2.0.1 ADD

        '(���[���h�c�s���j
        Case Else
            'V2.2.0.1 DEL START
            '���[���ُ��M���O�o��
            'Call dllErrorMailLog(plMSlot_LG, uMail, lngLen)
            'V2.2.0.1 DEL END
            'V2.2.0.1 ADD START
            '�u���[��ID�s���v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
            'V2.2.0.1 ADD END
        End Select
    Loop
End Sub

'*****************************************************************************
'*    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'*
'* �@�@�\�F���[���X���b�g��ǂݍ��ށi�ێ�v���Z�X�E���j�^��M��p�j
'*   �����@�@�@�@      (I/O) ����
'*   lngMailHamdle      I    ���[���X���b�g�n���h��
'*   vReadBuf           O  ���[�����e���i�[���邽�߂̃G���A���w���|�C���^
'*   �߂�l
'*   Err                  O  0�ȏ�              �Ǎ��݃T�C�Y
'*                           0                 ��M�f�[�^����
'*                           -1                �G���[
'*
'*  ORIGINAL  :(ED7.0.0.1) 2006-05-10   CODED   BY [TCC] Y.Takezawa
'*   REVISIONS :(V2.2.0.1)  2010-09-13   REVISED BY [TCC] S.Terao
'*              �d�f�q���g���@�m�d�f���d�f�q�R���o�[�g�Ή�
'*  REVISIONS :(xx0.0.0.0) 0000-00-00   REVISED BY [   ]
'*****************************************************************************
Private Function fDssMailReadMN(lngMailHamdle As Long, udtReadBuf As MAIL_LGMINF_INF) As Long
    Dim lngBool As Long             ' ��������
    Dim lngNextMsg As Long          ' ���̃��b�Z�[�W�̃T�C�Y
    Dim lngMsg As Long              ' ���b�Z�[�W��
    Dim lngMailRcvLength As Long    ' �Ǎ��݃T�C�Y
    Dim lngTraceSize As Long        ' ���O�ɂƂ郁�[���T�C�Y
    Dim lngErrCode As Long  '�G���[�R�[�h
    
      On Error GoTo MailReadError

    lngMsg = 0
    lngBool = GetMailslotInfo(lngMailHamdle, 0, lngNextMsg, lngMsg, 0)

    ' ���[���ɏ�񂪂���΁A���[����M
    If (lngNextMsg = -1) Then
        fDssMailReadMN = 0
        Exit Function
    End If
   
    ' ���[���T�C�Y����M�G���A���傫�����
    If lngNextMsg > LenB(udtReadBuf) Then
       '�ُ탁�[����M���O���o�͂���B
       'Call dllErrorMailLog(plMSlot_LG, udtReadBuf, lngNextMsg) 'V2.2.0.1 DEL
      'V2.2.0.1 ADD START
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MRECEIVE
      Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_RECV_SIZE_ERROR, lngErrCode)
      'V2.2.0.1 ADD END
      'MsgBox "��M�ł��Ȃ��T�C�Y�̃��[�������M���ꂽ���߁A�ێ��ʃv���Z�X���ُ�I�����܂��B" 'V2.2.0.1 DEL
        pfAbortProc
    End If

    On Error Resume Next
    ' ���[����M����
    lngBool = ReadFile(lngMailHamdle, udtReadBuf, lngNextMsg, lngMailRcvLength, 0)
    fDssMailReadMN = lngMailRcvLength

    Exit Function

MailReadError:
    fDssMailReadMN = -1 'INVALID_HANDLE_VALUE
End Function
