VERSION 5.00
Begin VB.Form frmUtilityUSR 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "���[�e�B���e�B�N��"
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
      Interval        =   3000
      Left            =   5640
      Top             =   8520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����"
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
      Index           =   1
      Left            =   6840
      TabIndex        =   12
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�O���"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   6600
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Left            =   6360
      TabIndex        =   9
      Top             =   6600
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      TabIndex        =   7
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      TabIndex        =   4
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Top             =   840
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
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "���[�e�B���e�B�N��"
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
      TabIndex        =   13
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmUtilityUSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmUtilityUSR.frm
'//  �p�b�P�[�W���F���[�e�B���e�B�N��(��ʃ����e�i���X)���
'//
'//  �T�v�F���[�e�B���e�B�N��(��ʃ����e�i���X)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const iHoshuAplMax = 19           '�o�^�ő匏��
Private sFixedExePass(0 To 31) As String  '�Œ�N���t�ɑΉ���������̧���߽���i��޴ر���܂ށj
Private sFixedExeName(0 To 31) As String  '�Œ�N���t�ɑΉ������t���́i��޴ر���܂ށj
Private iHyoujiCnt As Integer             '�\���J�E���^�[
Private iGamenSts As Integer              '���ݕ\����ʐ�

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���[�e�B���e�B�N���i��ʃ����e�i���X)���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
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
    '���[����M�p�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ���[�e�B���e�B�N���i��ʃ����e�i���X)���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
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
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : ���[�e�B���e�B�N���i��ʃ����e�i���X)���(���[�h��)
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
    Dim i As Integer    '�J�E���^�[
    
On Error Resume Next
   
   '�uհè�è��ʁF�\���v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '������
    iHyoujiCnt = 0    '�\���J�E���^�[
    iGamenSts = 0     '���ݕ\����ʐ�
    Command1(0).Visible = False  '�u�O��ʁv�t��\��
    Command1(1).Visible = False  '�u����ʁv�t��\��
    
    For i = 0 To 31
        '�\�����G���A������
        sFixedExeName(i) = ""
    Next
    For i = 0 To 31
        '�c�[���p�X�G���A������
        sFixedExePass(i) = ""
    Next
     
    'V1.3.0.1 ADD START
    '���[����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    '1.3.0.1 ADD END
    
    '�Œ�t�\����������
    sFixedExeDisplay
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �Œ�N���t����������
'//  �@�\�T�v  : �Y���A�v���̋N�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      :Integer�@�@Index    [IN]�N���A�v���t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    Dim lRetVal As Double         'Shell�֐��߂�l
    Dim iResponse As Integer      'MsgBox�߂�l
    Dim iSetupAplIndex As Integer '�N���A�v���C���f�b�N�X
    
On Error GoTo ERROR_MSG
   '�uհè�è��ʁF�N���t�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_KIDOU_BUTTOM, 0)
  
    '��ʐݒ�C���f�b�N�X��0�`9�Ȃ̂ŁA�t�C���f�b�N�X�l���Z�o���A
    '�N���A�v���̃p�X�ŋN������B
    '�N���A�v���C���f�b�N�X=(���݉�ʐ�-1���)�~1��ʍő�t���{�����C���f�b�N�X(0�`9)
    '��F2��ʖڂ̉����t�C���f�b�N�X3���������ꂽ�ꍇ�A�N���A�v���p�X�C���f�b�N�X��13
    '13=(2-1)��10�{3
    iSetupAplIndex = (iGamenSts - 1) * 10 + Index
    
    '�Y���{�^���̃A�v���P�[�V�������N������B
    lRetVal = Shell(sFixedExePass(iSetupAplIndex), vbNormalFocus)
    '�uհè�è��ʁF�c�[���N������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_API, UTILITY_GAMEN_TOOL_OK, 0)
 
    Exit Sub
    
ERROR_MSG:
'===�A�v���N���G���[�̏ꍇ�A
    '�uհè�è��ʁF�c�[���N���ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_API, UTILITY_GAMEN_TOOL_ERROR, 0)
    '�u�N�����s�v�|�b�v�A�b�v��ʂ�\������B
    iResponse = MsgBox(cmdFixedExe(Index).Caption & "�t�A��`�G���[�B" & _
                Chr(vbKeyReturn) & _
                sFixedExePass(iSetupAplIndex) & "���N���ł��܂���B", _
                vbYes, _
               "�Œ�N���A�v�����s�G���[")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Command1_Click
'//  �@�\����  : �u����ʁv�u�O��ʁv�t����������
'//  �@�\�T�v  : �u����ʁv�u�O��ʁv�t�����ɂ��A�Ώۉ�ʂ�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      :Integer�@�@Index    [IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Command1_Click(Index As Integer)
  Dim i As Integer          'INI̧�كL�[�J�E���^�FDSPi ���N���tINDEX
  Dim iMax As Integer       '�Œ�N���tINDEX�ő�l

On Error Resume Next

 Select Case Index
  Case 0
   '�uհè�è��ʁF�O��ʖt�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_BACK_BUTTOM, 0)
   If iGamenSts = 2 Then
    '���ݕ\����ʐ��F2��ʖځB�u�O��ʁv�t���������ꂽ�B
    '�\���J�n�_��0�A���\����ʐ���1��ʖڂ̂��߁A���ݕ\����ʐ���1��ݒ肷��B
    iHyoujiCnt = 0
    iGamenSts = 1
   Else
    '���ݕ\����ʐ��F1��ʖځB�u�O��ʁv�t���������ꂽ�B
    '���\����ʐ���2��ʖڂ̂��߁A���ݕ\����ʐ���2��ݒ肷��B
    iGamenSts = 2
   End If
    
    '�S�Ă̌Œ�N���t�ɂ��āA�ȉ������{����B
    iMax = cmdFixedExe.UBound     '�Œ�N���tINDEX�̍ő�l�𓾂�B
    For i = 0 To iMax
     '�Œ�N���t���\���ɂ���B
      cmdFixedExe(i).Visible = False
        '�N���A�v���p�X���ƁA�\���t���̂̒�`�`�F�b�N���s���B
        If sFixedExePass(iHyoujiCnt) <> "" And sFixedExeName(iHyoujiCnt) <> "" Then
          '��`�L��̏ꍇ�̂݁A�L���v�V�����ɋN���t�\��������������݁A�N���t��\������B
          cmdFixedExe(i).Visible = True
          cmdFixedExe(i).Caption = sFixedExeName(iHyoujiCnt)
        End If
         '�\���J�E���^�A�b�v����B
         iHyoujiCnt = iHyoujiCnt + 1
    Next i
  
  Case 1
    '�uհè�è��ʁF����ʖt�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_NEXT_BUTTOM, 0)
  
    If iGamenSts = 2 Then
     '���ݕ\����ʐ��F2��ʖځB�u����ʁv�t���������ꂽ�B
     '�\���J�n�_��0�A���\����ʐ���1��ʖڂ̂��߁A���ݕ\����ʐ���1��ݒ肷��B
     iGamenSts = 1
      iHyoujiCnt = 0
    Else
     '���ݕ\����ʐ��F1��ʖځB�u����ʁv�t���������ꂽ�B
     '���\����ʐ����J�E���g�A�b�v����B
     iGamenSts = iGamenSts + 1
    End If
    
    '�S�Ă̌Œ�N���t�ɂ��āA�ȉ������{����B
    iMax = cmdFixedExe.UBound     '�Œ�N���tINDEX�̍ő�l�𓾂�B
    For i = 0 To iMax
     '�Œ�N���t���\���ɂ���B
       cmdFixedExe(i).Visible = False
        '�N���A�v���p�X���ƁA�\���t���̂̒�`�`�F�b�N���s���B
        If sFixedExePass(iHyoujiCnt) <> "" And sFixedExeName(iHyoujiCnt) <> "" Then
           '��`�L��̏ꍇ�̂݁A�L���v�V�����ɋN���t�\��������������݁A�N���t��\������B
           cmdFixedExe(i).Visible = True
           cmdFixedExe(i).Caption = sFixedExeName(iHyoujiCnt)
        End If
          '�\���J�E���^�A�b�v����B
          iHyoujiCnt = iHyoujiCnt + 1
    Next i
   
   Case Else
    '��������
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
   '�uհè�è��ʁF�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_END, 0)
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFixedExeDisplay
'//  �@�\����  : �Œ�A�v���N���t�����\������
'//  �@�\�T�v  : �����\���������s���B
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
Private Sub sFixedExeDisplay()
Dim i As Integer                     'INI̧�كL�[�J�E���^�FDSPi ���N���tINDEX
Dim iMax As Integer                  '�Œ�N���tINDEX�ő�l
Dim sLine As String * MAX_PATH_SIZE  '�P�s���̕�����B�i������hDSPi=�h�������j
Dim lSize As Long                    '�P�s�����޲Đ��B�i������hDSPi=�h�������j
Dim iK As Integer                    '�J���}�L�q�ʒu
Dim iKensuFlag As Integer            '�����t���O

On Error Resume Next

'��ʐ����P��ʖڂƂ���B
iGamenSts = 1
iMax = cmdFixedExe.UBound     '�Œ�N���tINDEX�̍ő�l�𓾂�B

 For i = 0 To iHoshuAplMax
   '�A�v���N�������lINI�t�@�C������A�P�s���̕�����iDSPi=�������j��Ǎ��ށB
    lSize = GetPrivateProfileString(PROFILE_SECTION_NAME_FIXED_EXE, _
                                    PROFILE_KEY_NAME_FIXED_EXE & CStr(i), _
                                    DEFAILT, sLine, Len(sLine), HOSHUAPL_FILE)
    iK = InStr(sLine, ",")        '�t�@�C�����i�t���p�X�j�̋�ؕ����ʒu�𓾂�B
    'INI�t�@�C���ɁA�Y���s�̒�`������ꍇ�A
    If lSize > 0 And iK <> 0 Then
     '�t�@�C�����Ɩt���̂�����o���A�ۑ����Ă����B
      sFixedExePass(i) = Trim$(Left$(sLine, iK - 1))
      sFixedExeName(i) = Trim$(Mid$(sLine, iK + 1, lSize - iK))
    End If
Next i

'�S�Ă̌Œ�N���t�ɂ��āA�ȉ������{����B
 For i = 0 To iMax
   '�Œ�N���t���\���ɂ���B
    cmdFixedExe(i).Visible = False
    '�N���A�v���p�X���ƁA�\���t���̂̒�`�`�F�b�N���s���B
    If sFixedExePass(i) <> "" And sFixedExeName(i) <> "" Then
       '��`�L��̏ꍇ�A�L���v�V�����ɋN���t�\��������������݁A�N���t��\������B
       cmdFixedExe(i).Visible = True
       cmdFixedExe(i).Caption = sFixedExeName(i)
    End If
    '�\���J�E���^�A�b�v����B
    iHyoujiCnt = iHyoujiCnt + 1
  Next i

For i = 0 To iHoshuAplMax
   If sFixedExePass(i) <> "" And sFixedExeName(i) <> "" And i > 9 Then
      Command1(0).Visible = True
      Command1(1).Visible = True
   End If
Next i
 
End Sub

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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmUtilityUSR.Caption, False
        pfFormActive (frmUtilityUSR.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
