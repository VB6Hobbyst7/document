VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLduVer 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "                                                               �k�c���[�e�B���e�B�o�[�W�����Ǘ�"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAplTimer 
      Left            =   8640
      Top             =   4080
   End
   Begin VB.Timer tmrLogTimer 
      Left            =   8760
      Top             =   3360
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " �}�� �� ���[�N �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   18
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyWork_Jikko 
      Caption         =   " ���[�N �� ���s �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   17
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   �� �� ���s     �R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   16
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   8760
      Top             =   5040
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�@�o�[�W�����Ǘ��@��ʂ֖߂�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      TabIndex        =   15
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "�\���X�V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   14
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdOutPut 
      Caption         =   " �o�[�W������� �}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   13
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "�}�̎�O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   12
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox txtDummy 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   12480
      Width           =   2535
   End
   Begin VB.Frame frmFolder 
      Height          =   1815
      Left            =   9360
      TabIndex        =   6
      Top             =   480
      Width           =   2055
      Begin VB.CheckBox chkFolder 
         Caption         =   "O ��"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "N ���s"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "W ���[�N"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.ListBox LstFile 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   240
      MultiSelect     =   2  '�g��
      TabIndex        =   1
      Top             =   2520
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H0000C000&
      Caption         =   "LDU�A�v���P�[�V�����o�[�W�����Ǘ�"
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
      TabIndex        =   11
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label lblZenVer 
      Caption         =   "�S�̃o�[�W����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label lblVer 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�o�[�W�����ԍ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblTime 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�X�V���t"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lblFolder 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "̫���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblFile 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "�t�@�C����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "frmLduVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmLduVer.frm
'//  �p�b�P�[�W���F�o�[�W�����Ǘ�(LDU)���
'//
'//  �T�v�F�o�[�W�����Ǘ�(LDU)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �ELD���[�e�B���e�B���A�o�[�W�����Ǘ�(LDU)��ʗ��p�B
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.129�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.123�z
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.273�C���Ή��z
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.22�C���Ή��z
'//     REVISIONS :(EG20 V5.2.0.1) 2012-03-10  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.59�C���Ή��z
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y2,3�����Ή�
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-23  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����t�F�[�Y�Q�Ή��i�}�̎�O���G���[�Ή��j
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'�S�̃o�[�W�������ۑ��Ǘ�
Private sMainVer As String

'�t�H���_��ʕ�
Public mlngChkFolderType        As Long
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l  'V1.3.0.1 ADD

' EG20 V2.1.0.1[Mainte_03_01] �ǉ��J�n
Private Const HEADERTITLE_WRK = "LDU�A�v���P�[�V�����o�[�W�����i���[�N�j�F"
Private Const HEADERTITLE_NOW = "�@�@�@�@�@�@�@�@�@�@�@�@�@�@ �i���s�j�@�F"
Private Const HEADERTITLE_OLD = "�@�@�@�@�@�@�@�@�@�@�@�@�@�@ �i���j�@�@�F"
Private Const HEADERVERSION_NON = "--.--.--.--"
' EG20 V2.1.0.1[Mainte_03_01] �ǉ��I��

' EG20 V3.3.0.1�y����TR-No.123�z �ǉ��J�n
Private Const APL_INTERVAL = 390000         ' �A�v���N���^�C�}�f�t�H���g�l
Private Const LOG_INTERVAL = 30000          ' ���O�N���^�C�}�f�t�H���g�l(30�b)
Dim lngAplMAX_Time As Long                  ' INI�擾�ݒ�l�i�`�o�k�j
Dim lngLogMAX_Time As Long                  ' INI�擾�ݒ�l�i���O�j
Dim lngtime        As Long                  ' ���݃^�C�}�l
Dim lngChangeKind  As Long                  ' �o�[�W�����ؑ֎��
' EG20 V3.3.0.1�y����TR-No.123�z �ǉ��I��

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdCopyBaitai_Work_Click
'//  �@�\����    �F�}�́����[�N�R�s�[
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()

    '�u�}�́����[�N�R�s�[�v�{�^���̏ꍇ�B
    '�u�o�[�W�����Ǘ���ʁF�}�́����[�N�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_USB_COPY_WRK_BUTTOM, 0)
    sCmdBtnEnabled False                        ' ��ʑ���s��
    '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����
    Call sFDInstall
    sCmdBtnEnabled True                         ' ��ʑ����
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sFDInstall
'//  �@�\����  : �}�̃C���X�g�[������
'//  �@�\�T�v  : �C���X�g�[���}�̃t�@�C�����A���[�N�t�H���_�ɃR�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.273�C���Ή��z
'//  REVISIONS   �F(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y2,3�����Ή�
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F (EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                  �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//  REVISIONS   �F (EG20 V30.3.0.1)2014-10-23  CODED BY  [TCC] T.Nakajima
'//                  �k���V�����t�F�[�Y�Q�Ή��i�}�̎�O���G���[�Ή��j
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F���D�@�o�[�W�����Ǘ���ʂ�sFDInstall���p
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall()
    Dim MyName As String            '�t�@�C���t���p�X��
    Dim iResponse As Integer        'MsgBox�{�^���R�[�h
    Dim sInputPass As String        '�C���X�g�[�����f�B���N�g����(STD)or�t�@�C����(LZH)
    Dim lngErrCode As Long          '�G���[�R�[�h
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                    '�t�@�C���I�u�W�F�N�g
    
    Dim lngProcId As Long                ' �v���Z�XID
    Dim hProc As Variant                 ' �v���Z�X�n���h��
    Dim FileName As String               ' ���o�t�@�C����                ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim FileKaku As String               ' �g���q                        ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim ExecCommand As String            ' ���s������                    ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim CurrentDirectory As String       ' �J�����g�f�B���N�g��          ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim ExecDirectory As String          ' ���s�t�@�C���f�B���N�g��      ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    
    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^

    '���k�t�@�C���w��̎�:
    '�擾�t�@�C������������
    CommonDialog1.FileName = ""
    '�����f�B���N�g����ݒ�
    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    '�t�H���_�I����ʃf�t�H���g�p�X�P�����݂��邩
        '���݂��邽�߁A�f�t�H���g�p�X�P�iH:�j��ݒ�
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
    Else
        '���݂��Ȃ����߁A�f�t�H���g�p�X�Q�iC:�j��ݒ�
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
    End If
    '�g���q��ݒ�
    CommonDialog1.Filter = "���s�t�@�C���i*.EXE�j|*.exe|" & _
                            "�o�b�`�t�@�C���i*.BAT�j|*.bat|" & _
                            "�X�N���v�g�t�@�C���i*.VBS�j|*.vbs|"
    '�t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    '�I�������t�@�C�������擾
    sInputPass = CommonDialog1.FileName
    If sInputPass = "" Then '�t�@�C�����I��
        Set objFso = Nothing
        Set objFi = Nothing
        Exit Sub    '�t�@�C�����I������Ȃ���Ώ������f
    End If
        
    '�u���[�N�R�s�[�m�F�v�|�b�v�A�b�v��ʕ\��
    iResponse = MsgBox("�I�����ꂽ�C���X�g�[�����ނ̓��e���k�c�t�A�v���P�[�V������" _
                       & Chr(vbKeyReturn) _
                       & "�ؑ֗̈�ɓW�J���܂��B��낵���ł����H", _
                       (vbOKCancel + vbExclamation), _
                       "�}�́����[�N�@�R�s�[")
        
    If iResponse = vbCancel Then
    '[������] �{�^����I��:�������Ȃ��B
        'V1.20.0.1 ADD START
        Set objFso = Nothing
        Set objFi = Nothing
        'V1.20.0.1 ADD END
        Exit Sub
    End If
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
'    lngProcId = Shell(sInputPass, vbNormalFocus)       ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�폜
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��J�n
    ' �J�����g�f�B���N�g���擾
    CurrentDirectory = CurDir$()
    Call psFolderPathGet(sInputPass, ExecDirectory)
    Call ChDir(ExecDirectory)
    ' �t�@�C�����O�擾
    psFileNameGet sInputPass, FileName, FileKaku
    If UCase(FileKaku) = "VBS" Then
        ExecCommand = "wscript.exe " & sInputPass
    Else
        ExecCommand = sInputPass
    End If
    lngProcId = Shell(ExecCommand, vbNormalFocus)
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��I��
    
    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, lngProcId)   ' �v���Z�X�n���h�����擾���܂��B
    If hProc > 0 Then                                           ' �v���Z�X�n���h�����擾�ł����ꍇ
        Call dllWaitForSingleObject(hProc)                      ' �v���Z�X���V�O�i����ԂɂȂ�܂ő҂��܂��B
        CloseHandle hProc                                       ' �v���Z�X�n���h����������܂��B
    End If

    'EG20 V30.0.3.1 ADD START
    'ChDir�ł�CommonDialog�̏ꍇ�AH�h���C�u���I�����ꂽ�܂ܕύX���ꂸ�A�}�̎�O�����ł��Ȃ��Ȃ邽�߁AChDrive�ɕύX
    ChDrive "C"
    'EG20 V30.0.3.1 ADD END
    Call ChDir(CurrentDirectory)                        ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    
    Set objFso = Nothing
    Set objFi = Nothing
    
    '�Ď��Ղ̃o�[�W�����ԍ���\������
    psVersionDisp
    
' EG20 V5.8.0.1�폜�J�n
'    ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
    ' �^����ԍX�V
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_LDU, BOOTINFO_UNKAI_ARI)
    Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMELDU, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1�ǉ��I��
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD START
    ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
    Call funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_CLEAR_LDU)
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD END

' EG20 V5.8.0.1 ADD START
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (PATH_LDU_APLNEW)
' EG20 V5.8.0.1 ADD END
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    Exit Sub    '�������I������

ErrorHandler:   ' �G���[�����B
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing

' EG20 V5.8.0.1 ADD START
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (PATH_LDU_APLNEW)
' EG20 V5.8.0.1 ADD END
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    MsgBox "�C���X�g�[���}�̂���̃R�s�[�G���[���������܂����B" _
            & Chr(vbKeyReturn) & "�G���[�R�[�h��" _
            & str$(Err.Number), _
            vbOKOnly + vbExclamation, _
            "�}�́����[�N�@�R�s�["
    
    '�u�����ް�ޮ݁F�}�́�ܰ���߰�����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_USB_COPY_WRK_ERROR, lngErrCode)
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdCopyOld_Jikko_Click
'//  �@�\����    �F�������s�R�s�[
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.123�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.22�C���Ή��z
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.372�C���Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    
    Dim udtSendData As ML_KANEND_REQ_CMD  ' ���ʃG���A
    Dim lngSendSize As Long               ' ���M���郁�[���T�C�Y
    Dim lngErrCode  As Long               ' �G���[�R�[�h
    Dim bRet        As Boolean            ' ���[�����M�����߂�l
    Dim iResponse   As Integer            ' MsgBox�{�^���R�[�h
    Dim iAplChk     As Integer            ' �A�v���N���`�F�b�N�߂�l    'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
    
    On Error Resume Next
    
    '�u�o�[�W�����Ǘ���ʁF�������s�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OLD_COPY_NOW_BUTTOM, 0)

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�u���v�t�H���_�̓��e���A" _
            & Chr(vbKeyReturn) & "�u���s�v�t�H���_�ɖ߂����Ƃɂ��A" _
            & Chr(vbKeyReturn) & "�k�c�t�̈ꐢ��O�o�[�W��������s�o�[�W�����Ƃ��܂��" _
            & Chr(vbKeyReturn) & "��낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "�������s�@�R�s�[")
    If iResponse = vbCancel Then
        Exit Sub
    End If

'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��J�n
    ' ���o�[�W�����t�H���_�ɑ�\�o�[�W�����t�@�C�������݂��Ȃ��ꍇ�ُ͈�Ƃ���B
    ' ���o�[�W�����EEW4500JR�ELDU�E
    bRet = dllCheckAplVersion(4, PATH_LDU_APP, 1)
    If bRet = False Then
        MsgBox "�ُ�I�����܂����B", vbCritical, "�������s�@�R�s�["
        Exit Sub
    End If
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��I��
        
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD START
    ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONDOWN_LDU)
    If bRet = False Then
        MsgBox "�ُ�I�����܂����B", vbCritical, "�������s�@�R�s�["
        Exit Sub
    End If

    ' �I���m�F
    iResponse = MsgBox("���s�R�s�[��K�p���邽�߂ɓ����Ď��Ղ�" & Chr(vbKeyReturn) _
                        & "�ċN�����܂����H", _
                        vbOKCancel + vbExclamation, _
                        "�������s�@�R�s�[")
    If iResponse = vbCancel Then
        Exit Sub
    End If
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD END
        
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL START
''EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��J�n
'    ' �����Ď��Ղ��N�����̏ꍇ�Ƀ��b�Z�[�W�{�b�N�X��\������B
'    iAplChk = CheckAppStart(PROC_KANRI)
'    If iAplChk <> 0 Then
''EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��I��
'        '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
'        iResponse = MsgBox("�����Ď��դ�h�c�t��k�c�t�A�v���P�[�V������" _
'                & Chr(vbKeyReturn) & "�I�����܂��B��낵���ł����H", _
'               vbOKCancel + vbExclamation, _
'               "�I���m�F")
'
'        If iResponse = vbCancel Then
'            Exit Sub
'        End If
'    End If  'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL END

' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
' AplVersionChangeProc�Ƀ��W���[����
'    ' ���[���̑��M���e��ҏW����
'    udtSendData.udtlHeader.dwId = ML_ID_KANEND_REQ      ' ���[���h�c�@=�h"�Ď����u�I���v��"
'    udtSendData.udtlHeader.dwSize = MlSize.KANEND_REQ   ' ���[���T�C�Y=�h"�Ď����u�I���v��"
'    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' ���M���v���Z�X�h�c=�h�ێ�h
'    udtSendData.udtlHeader.dwSubArea = 0                ' �⏕���@=�@0
'
'    udtSendData.dwStartProc = ML_DT_VERSIONDOWN         ' �N���v���Z�X��� = �o�[�W�����_�E��
'    ' ���M�T�C�Y��ݒ肷��B
'    lngSendSize = udtSendData.udtlHeader.dwSize
'
'    ' �ă}�ɑ΂��āA�ݒ���v�����[���𑗐M����B
'    bRet = DssSendMail(MAIL_SLOT_KANRI, lngSendSize, udtSendData.udtlHeader)
'    ' ���[���𐳏�ɑ��M�������̃��O
'    If bRet = False Then
'        '�u�ݒ���v�����[�����M�ُ�v���O�o��
'        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, lngErrCode)
'    Else
'        '�u�ݒ���v�����[�����M����v���O�o��
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, 0)
'    End If
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
' EG20 V3.3.0.1�y����TR-No.123�z�폜�J�n
'' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
'    ' �A�v���P�[�V�����o�[�W�����ؑ֎��s����
'    If (AplVersionChangeProc(ML_DT_VERSIONDOWN_LDU) = False) Then
'        ' // �ێ���I������B
'        Call psEndHoshuProc
'        '�ێ�v���Z�X�I��
'        End
'    End If
'' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
' EG20 V3.3.0.1�y����TR-No.123�z�폜�I��
' EG20 V3.3.0.1�y����TR-No.123�z�ǉ��J�n

    sCmdBtnEnabled False                            ' ��ʑ���s��
    ' �����Ď��ՂփA�v���I���v���̑��M
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
               vbOKOnly + vbExclamation, _
               "�k�c�t�o�[�W�����Ǘ�"
        sCmdBtnEnabled True                         ' ��ʑ����
    Else

        lngtime = MN_MAIL_INTERVAL                  ' ���݃^�C�}�l������
        tmrAplTimer.Enabled = True                  ' ���݃^�C�}�N��
    
        lngChangeKind = ML_DT_VERSIONDOWN_LDU       ' �ؑ֎�ʂ�ݒ�
    End If
' EG20 V3.3.0.1�y����TR-No.123�z�ǉ��I��

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdCopyWork_Jikko_Click
'//  �@�\����    �F���[�N�����s�R�s�[
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.123�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.22�C���Ή��z
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.372�C���Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y�A�v���ؑ։��P�Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
    
    Dim udtSendData As ML_KANEND_REQ_CMD  ' ���ʃG���A
    Dim lngSendSize As Long               ' ���M���郁�[���T�C�Y
    Dim lngErrCode  As Long               ' �G���[�R�[�h
    Dim bRet        As Boolean            ' ���[�����M�����߂�l
    Dim iResponse   As Integer            ' MsgBox�{�^���R�[�h
    Dim iAplChk     As Integer            ' �A�v���N���`�F�b�N�߂�l    'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
    
    On Error Resume Next
    
    '�u�o�[�W�����Ǘ���ʁF���[�N�����s�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_COPY_NOW_BUTTOM, 0)


    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�u���[�N�v�t�H���_�̓��e���A" _
            & Chr(vbKeyReturn) & "�u���s�v�t�H���_�ɓo�^���邱�Ƃɂ��A" _
            & Chr(vbKeyReturn) & " �k�c�t�̍ŐV�o�[�W�������A���s�o�[�W�����Ƃ��܂��B" _
            & Chr(vbKeyReturn) & "��낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "���[�N�����s �R�s�[")
    If iResponse = vbCancel Then
        Exit Sub
    End If
        
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��J�n
    ' ���[�N�o�[�W�����t�H���_�ɑ�\�o�[�W�����t�@�C�������݂��Ȃ��ꍇ�ُ͈�Ƃ���B
    ' ���[�N�o�[�W�����EEW4500JR�ELDU�E
    bRet = dllCheckAplVersion(1, PATH_LDU_APP, 1)
    If bRet = False Then
        MsgBox "�ُ�I�����܂����B", vbCritical, "���[�N�����s �R�s�["
        Exit Sub
    End If
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��I��

' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD START
    ' �ؑ֎��s�R�s�[�c�[���p�����[�^�X�V����
    bRet = funcUpdateCopyExecParam(KanendReq_ProcType.ML_DT_VERSIONUP_LDU)
    If bRet = False Then
        MsgBox "�ُ�I�����܂����B", vbCritical, "���[�N�����s �R�s�["
        Exit Sub
    End If

    ' �I���m�F
    iResponse = MsgBox("���s�R�s�[��K�p���邽�߂ɓ����Ď��Ղ�" & Chr(vbKeyReturn) _
                        & "�ċN�����܂����H", _
                        vbOKCancel + vbExclamation, _
                        "���[�N�����s �R�s�[")
    If iResponse = vbCancel Then
        Exit Sub
    End If
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zADD END
        
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL START
'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��J�n
'    ' �����Ď��Ղ��N�����̏ꍇ�Ƀ��b�Z�[�W�{�b�N�X��\������B
'    iAplChk = CheckAppStart(PROC_KANRI)
'    If iAplChk <> 0 Then
''EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ��I��
'        '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
'        iResponse = MsgBox("�����Ď��դ�h�c�t��k�c�t�A�v���P�[�V������" _
'                & Chr(vbKeyReturn) & "�I�����܂��B��낵���ł����H", _
'               vbOKCancel + vbExclamation, _
'               "�I���m�F")
'
'        If iResponse = vbCancel Then
'            Exit Sub
'        End If
'    End If  'EG20 V3.6.0.1�y03����TR-No.22�C���Ή��z�ǉ�
' EG20 V6.9.0.1�y�ʎY�Ή��F�A�v���ؑ։��P�Ή��zDEL END

' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
' AplVersionChangeProc�Ƀ��W���[����
'    ' ���[���̑��M���e��ҏW����
'    udtSendData.udtlHeader.dwId = ML_ID_KANEND_REQ      ' ���[���h�c�@=�h"�Ď����u�I���v��"
'    udtSendData.udtlHeader.dwSize = MlSize.KANEND_REQ   ' ���[���T�C�Y=�h"�Ď����u�I���v��"
'    udtSendData.udtlHeader.dwProid = RHOSHU_ID          ' ���M���v���Z�X�h�c=�h�ێ�h
'    udtSendData.udtlHeader.dwSubArea = 0                ' �⏕���@=�@0
'
'    udtSendData.dwStartProc = ML_DT_VERSIONUP           ' �N���v���Z�X��� = �o�[�W�����A�b�v
'    ' ���M�T�C�Y��ݒ肷��B
'    lngSendSize = udtSendData.udtlHeader.dwSize
'
'    ' �ă}�ɑ΂��āA�ݒ���v�����[���𑗐M����B
'    bRet = DssSendMail(MAIL_SLOT_KANRI, lngSendSize, udtSendData.udtlHeader)
'    ' ���[���𐳏�ɑ��M�������̃��O
'    If bRet = False Then
'        '�u�ݒ���v�����[�����M�ُ�v���O�o��
'        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'        Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, lngErrCode)
'    Else
'        '�u�ݒ���v�����[�����M����v���O�o��
'        Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSHISYSTEM_INSTALL_CMD_SEND, 0)
'    End If
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
' EG20 V3.3.0.1�y����TR-No.123�z�폜�J�n
'' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
'    ' �A�v���P�[�V�����o�[�W�����ؑ֎��s����
'    If (AplVersionChangeProc(ML_DT_VERSIONUP_LDU) = False) Then
'        ' // �ێ���I������B
'        Call psEndHoshuProc
'        '�ێ�v���Z�X�I��
'        End
'    End If
'' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��
' EG20 V3.3.0.1�y����TR-No.123�z�폜�I��
' EG20 V3.3.0.1�y����TR-No.123�z�ǉ��J�n

    sCmdBtnEnabled False                            ' ��ʑ���s��
    ' �����Ď��ՂփA�v���I���v���̑��M
    bRet = pubFuncAplEndRequest()
    If bRet = False Then
        MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
               vbOKOnly + vbExclamation, _
               "�k�c�t�o�[�W�����Ǘ�"
        sCmdBtnEnabled True                         ' ��ʑ����
    Else

        lngtime = MN_MAIL_INTERVAL                  ' ���݃^�C�}�l������
        tmrAplTimer.Enabled = True                  ' ���݃^�C�}�N��
    
        lngChangeKind = ML_DT_VERSIONUP_LDU         ' �ؑ֎�ʂ�ݒ�
    End If
' EG20 V3.3.0.1�y����TR-No.123�z�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : CmdRemove_Click
'//  �@�\����  : �u�}�̎�O�v�t����������
'//  �@�\�T�v  : �}�̂̎��O�����s���B
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
Private Sub cmdRemove_Click()
   On Error Resume Next
   
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����Ǘ�(LDU)���(�A�N�e�B�u��)
'//  �@�\�T�v  : ��ʍőO�ʕ\��
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
Private Sub Form_Activate()
    
    pfFormActive (hwnd)
    'V1.3.0.1 ADD START
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
    'V1.3.0.1 ADD END

End Sub

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����Ǘ�(LDU)���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
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
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �o�[�W�����Ǘ�(LDU)���(���[�h��)
'//  �@�\�T�v  : �����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.123�z
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    On Error Resume Next
    
    '�uLDհè�è�ް�ޮ݊Ǘ���ʁF�\���v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_VERASION_KANRI_GAMEN_START, 0)

    gStrCurrentForm = sFormName_LDUVer

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   
   '������
    LstFile.Clear
    lblZenVer.Caption = ""
    mlngChkFolderType = 0

    '�t�H���_�I�𕔁F�I��L��
    chkFolder(0).Value = 1
    chkFolder(1).Value = 1
    chkFolder(2).Value = 1
        
    mlngChkFolderType = 7       ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
        
    '�o�[�W�������o�͏���
    Call psVersionDisp
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   
   'V1.3.0.1 ADD START
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END

' EG20 V3.3.0.1�y����TR-No.123�z �ǉ��J�n
    ' INI�t�@�C�����A�v���N���^�C�}�l���擾
    lngAplMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                       APL_INTERVAL, HOSHU_FILE)
    ' �擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
    If lngAplMAX_Time = 0 Then
       lngAplMAX_Time = APL_INTERVAL
    End If

    ' �^�C�}�l�ݒ�
    tmrAplTimer.Interval = MN_MAIL_INTERVAL
    tmrAplTimer.Enabled = False

    ' INI�t�@�C����胍�O�N���^�C�}�l���擾
    lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
    ' �擾�l��0�̏ꍇ�A�f�t�H���g�l��ݒ�
    If lngLogMAX_Time = 0 Then
       lngLogMAX_Time = LOG_INTERVAL
    End If

    ' �^�C�}�l�ݒ�
    tmrLogTimer.Interval = MN_MAIL_INTERVAL
    tmrLogTimer.Enabled = False

    ' �ؑ֎�ʂ�������
    lngChangeKind = 0
' EG20 V3.3.0.1�y����TR-No.123�z �ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : chkFolder_Click
'//  �@�\����  : �u�t�H���_�`�F�b�N�v�`�F�b�N��������
'//  �@�\�T�v  : �t�H���_�I�𕔃`�F�b�N���s���B
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
Private Sub chkFolder_Click(Index As Integer)
    
    Dim ValueCnt                As Integer

    '��ނɂ���đ����l��ύX����
    ValueCnt = 0
    '���[�N
    If Index = 0 Then
        ValueCnt = 1
    '���s
    ElseIf Index = 1 Then
        ValueCnt = 2
    '��
    ElseIf Index = 2 Then
        ValueCnt = 4
    End If

    '�`�F�b�N���͂����ꂽ��
    If chkFolder(Index).Value = 0 Then
        mlngChkFolderType = mlngChkFolderType - ValueCnt
    '�`�F�b�N���ꂽ��
    ElseIf chkFolder(Index).Value = 1 Then
        mlngChkFolderType = mlngChkFolderType + ValueCnt
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdRefresh_Click
'//  �@�\����  : �u�\���X�V�v�t����������
'//  �@�\�T�v  : �ŐV�̃o�[�W��������\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdRefresh_Click()
    Dim i As Integer        '�J�E���^�[
    Dim bFlag As Boolean    '�\���t�H���_�I���`�F�b�N(TRUE�F�`�F�b�N�L�BFALSE�F�`�F�b�N��)
  
    On Error Resume Next
  
  '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�\���X�V�t�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
   
   '�\���t�H���_�I���`�F�b�N���`�F�b�N�����ŏ���������B
   bFlag = False
   '�\���t�H���_�I���`�F�b�N�L�����`�F�b�N����B
   For i = 0 To 2
     If chkFolder(i).Value = CHECKBOX_ON Then
       '�P�ł��`�F�b�N�L��̏ꍇ�A�\���t�H���_�I���`�F�b�N���A�`�F�b�N�L�ɂ���B
        bFlag = True
         Exit For
      End If
   Next
   
   '�\���t�H���_�I���̃`�F�b�N���Ȃ��ꍇ�́A�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
   If bFlag = False Then
      MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
              vbOKOnly + vbExclamation, _
               "�k�c�t�o�[�W�����Ǘ�"
       Exit Sub
    End If

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    '�o�[�W�������o�͏���
    Call psVersionDisp

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdOutPut_Click
'//  �@�\����  : �u�o�[�W�������}�̏o�́v�t����������
'//  �@�\�T�v  : ��ʏ�ɕ\�����ꂽ�o�[�W���������A�}�̂ɏo�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.129�z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()

'*******************************
'VB�G���[����
On Error GoTo Error_cmdOutPut_Click
'*******************************

    Dim strVerFile  As String               'LD���[�e�B���e�B�t�@�C���p�X
    Dim strCopySaki As String               '�o�̓t�@�C���p�X
    Dim strWriteDir As String               '�o�͐�t�H���_
    Dim fso         As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim lngErrCode  As Long                 '�G���[�R�[�h
  
    Dim strStationName       As String          ' �w����                ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim szCornerName         As String          ' �R�[�i����            ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim nNullIndex           As Integer         ' ���������[�N          ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
    Dim strRecord            As String          ' ���[�N
    Dim strFileName         As String           ' �t�@�C����
    Dim bRet                As Boolean          ' �߂�l

    Set fso = CreateObject("Scripting.FileSystemObject")

  
   '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͖t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OUTPUT, 0)

' EG20 V3.3.0.1 �y����TR-No.129�z�ǉ��J�n
    ' ���X�g�ɂP�����f�[�^���Ȃ��ꍇ�ُ͈�I��
    If LstFile.ListCount = 0 Then
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Set fso = Nothing           ' EG20 V3.3.0.1�ǉ�
        Exit Sub
    End If
' EG20 V3.3.0.1 �y����TR-No.129�z�ǉ��I��

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)

    '�w��t�H���_�Ȃ�
    If Len(strWriteDir) = 0 Then
        Set fso = Nothing
        Exit Sub
    End If

    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��


' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    strStationName = gsGetStationEkiName
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ �ێ��p�֐�:IDU�o�[�W�����t�@�C���i���\�p�j�쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, LDUVERLIST_REPORTFILE, PATH_LDU_APP, _
                                    VERLISTKIND_REPORT, 1)

    ' �o�[�W�����t�@�C������
    If bRet Then
        '�u�o�[�W�������t�@�C���쐬����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    ' �o�[�W�����t�@�C�����s
    Else
        '�u�o�[�W�������t�@�C���쐬�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Set fso = Nothing           ' EG20 V3.3.0.1�ǉ�
       Exit Sub
    End If
    
    '�t�@�C���̗L���m�F
    If fso.FileExists(LDUVERLIST_REPORTFILE) = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Set fso = Nothing           ' EG20 V3.3.0.1�ǉ�
        Exit Sub
    End If
    strFileName = Dir(LDUVERLIST_REPORTFILE)
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��

' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
'    'LD���[�e�B���e�B�o�[�W�����t�@�C��
'    strVerFile = PATH_LDU_APP & PATH_LDU_WORK & LDU_VER_FILE
'
'    '�t�@�C���̗L���m�F
'    If fso.FileExists(strVerFile) = False Then
'        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
'        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
'        Exit Sub
'    End If
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�J�n
'    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
''    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")                       'V1.12.0.1 DEL
'    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.12.0.1 ADD
'
'    '�w��t�H���_�Ȃ�
'    If Len(strWriteDir) = 0 Then
'        Set fso = Nothing           ' EG20 V3.3.0.1�ǉ�
'        Exit Sub
'    End If
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�I��

' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    '�R�s�[��t�H���_�̗L���m�F
    If fso.FolderExists(strWriteDir) = False Then
        '�R�s�[��t�H���_�쐬
        fso.CreateFolder (strWriteDir)
    End If
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��

' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
'    '�R�s�[��t�H���_�p�X�쐬(�w��t�H���_��LDU_VER)
'    strWriteDir = strWriteDir & "\" & LDU_VER
'
'    '�R�s�[��t�H���_�̗L���m�F
'    If fso.FolderExists(strWriteDir) = False Then
'
'        '�R�s�[��t�H���_�쐬
'        fso.CreateFolder (strWriteDir)
'
'    End If
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��

    '�R�s�[��t�@�C�����쐬
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    '�R�s�[��t�@�C�����쐬
    strCopySaki = strWriteDir & "\" & strStationName & "_" & strFileName

    '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
    fso.CopyFile LDUVERLIST_REPORTFILE, strCopySaki, True
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��

    '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
'    fso.CopyFile strVerFile, strCopySaki, True
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    MsgBox "����I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
   '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͏�������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_INFO_OUTPUT_OK, 0)

    Set fso = Nothing

    Exit Sub

'*******************************
'VB�G���[����
Error_cmdOutPut_Click:
   '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͏����ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OUTPUT_ERROR, lngErrCode)
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
    Set fso = Nothing

'*******************************

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdCancel_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����������
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
Private Sub cmdCancel_Click()
   On Error Resume Next
   
   '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_VERASION_KANRI_GAMEN_END, 0)
   frmVersion.ZOrder
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psVersionDisp
'//  �@�\����  : �o�[�W�������\������
'//  �@�\�T�v  : �o�[�W�������\�����̕\���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Public Sub psVersionDisp()

    Dim strFilePath     As String   '�o�[�W�����t�@�C���p�X
    Dim bRet            As Boolean  '�߂�l
    Dim intFileNo       As Integer  '�t�@�C���ԍ�
    Dim strWork         As String   '��ƃG���A
    Dim strVerData      As String   '�S�̃o�[�W����
    Dim intCnt          As Integer  '�J�E���^�[
    Dim lngErrCode      As Long     '�G���[�R�[�h

'*******************************
'VB�G���[����
On Error GoTo Error_psVersionDisp
'*******************************

    '�}�̏o�͖t�����s��
    cmdOutPut.Enabled = False

    '���X�g������
    LstFile.Clear

    '�S�̃o�[�W����������
' EG20 V2.1.0.1[Mainte_03_01]�폜�J�n
'    lblZenVer.Caption = "�S�̃o�[�W�����i���[�N�j:--.--.--.--" & vbCrLf & _
'                        "�@�@�@�@�@�@�@�i���s�j�@:--.--.--.--" & vbCrLf & _
'                        "�@�@�@�@�@�@�@�i���j    :--.--.--.--"
' EG20 V2.1.0.1[Mainte_03_01]�폜�I��
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��J�n
    lblZenVer.Caption = HEADERTITLE_WRK & HEADERVERSION_NON & vbCrLf & _
                        HEADERTITLE_NOW & HEADERVERSION_NON & vbCrLf & _
                        HEADERTITLE_OLD & HEADERVERSION_NON
' EG20 V2.1.0.1[Mainte_03_01]�ǉ��I��

    '��ƃG���A������
    strWork = ""

    '�S�̃o�[�W����������
    strVerData = ""

    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���p�X�쐬
    strFilePath = PATH_LDU_APP & PATH_LDU_WORK & LDU_VER_FILE
    
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ ����DA:LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
'    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, strFilePath, PATH_LDU_APP)                       ' EG20 V2.1.0.1[Mainte_03_01]�폜
    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, strFilePath, PATH_LDU_APP, VERLISTKIND_DISP, 1)   ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�

    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���쐬����
    If bRet Then
       '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���쐬���s
    Else
       '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
       Exit Sub
    End If

    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���̗L���m�F
    If Len(Trim(Dir(strFilePath))) = 0 Then
        Exit Sub
    End If

    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���̃t�@�C���ԍ����擾����B
    intFileNo = FreeFile

    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���I�[�v��
    Open strFilePath For Input As #intFileNo


        '���[�N
        Line Input #intFileNo, strWork

        If (Trim(strWork) = "") Then
'            strVerData = "�S�̃o�[�W�����i���[�N�j�F--.--.--.--" & vbCrLf                              ' EG20 V2.1.0.1[Mainte_03_01]�폜
            strVerData = HEADERTITLE_WRK & HEADERVERSION_NON & vbCrLf                                   ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
        Else
            '�S�̃o�[�W����������쐬
'            strVerData = strVerData & strWork & vbCrLf                                                 ' EG20 V2.1.0.1[Mainte_03_01]�폜
            strVerData = strWork & vbCrLf                                                               ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
        End If

        '���s
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "�@�@�@�@�@�@�@�i���s�j�@�F--.--.--.--" & vbCrLf                 ' EG20 V2.1.0.1[Mainte_03_01]�폜
            strVerData = strVerData & HEADERTITLE_NOW & HEADERVERSION_NON & vbCrLf                      ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
        Else
            '�S�̃o�[�W����������쐬
            strVerData = strVerData & strWork & vbCrLf
        End If

        '��
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "�@�@�@�@�@�@�@�i���j    �F--.--.--.--" & vbCrLf                 ' EG20 V2.1.0.1[Mainte_03_01]�폜
            strVerData = strVerData & HEADERTITLE_OLD & HEADERVERSION_NON & vbCrLf                      ' EG20 V2.1.0.1[Mainte_03_01]�ǉ�
        Else
            '�S�̃o�[�W����������쐬
            strVerData = strVerData & strWork & vbCrLf
        End If

        '�S�̃o�[�W�����o��
        lblZenVer.Caption = strVerData

        strWork = ""

        '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1�폜
        Do While Not EOF(intFileNo)                         ' EG20 V3.3.0.1�ǉ�

            Line Input #intFileNo, strWork

            '���s�R�[�h�݂͓̂ǂ݂Ƃ΂�
            If Trim(strWork) <> "" Then

                '���X�g�ɏo��
                LstFile.AddItem (strWork)

            End If
        Loop

    '�t�@�C���N���[�Y
    Close #intFileNo

    '�}�̏o�͖t������
    cmdOutPut.Enabled = True

    Exit Sub

'*******************************
'VB�G���[����
Error_psVersionDisp:
   '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'    �t�@�C���N���[�Y
    Close #intFileNo
'*******************************
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
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��i�Ď��Ճo�[�W�����A�b�v�Ή��j
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
'    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then          ' EG20 V3.0.0.2�폜
    If pfVersionDispMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then   ' EG20 V3.0.0.2�ǉ�
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmLduVer.Caption, False
        pfFormActive (frmLduVer.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'/
'/  �֐�����     : sCmdBtnEnabled
'/  �@�\����     : �R�}���h�{�^�������E�s����
'/  �@�\�T�v     : �R�}���h�{�^���������Ɋ�ĉ����E�s�������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'//  ORIGINAL    :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                EG20�t�F�[�Y�Q�Ή�
'//                EG20�����Ď���USDM�Ή��ԍ��yMainte_02_06�z
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)
    Dim iLoopCnt    As Integer

    '�t�H���_�I�𕔁F�I��L��
    chkFolder(0).Enabled = blnFlg
    chkFolder(1).Enabled = blnFlg
    chkFolder(2).Enabled = blnFlg

    cmdRefresh.Enabled = blnFlg                     ' �\���X�V
    cmdCopyBaitai_Work.Enabled = blnFlg             ' �}�́����[�N�R�s�[
    cmdCopyWork_Jikko.Enabled = blnFlg              ' ���[�N�����s�R�s�[
    cmdCopyOld_Jikko.Enabled = blnFlg               ' �������s�R�s�[
    cmdOutPut.Enabled = blnFlg                      ' �o�[�W�������}�̏o��
    cmdRemove.Enabled = blnFlg                      ' �}�̎�O
    cmdCancel.Enabled = blnFlg                      ' �o�[�W�����Ǘ���ʂ֖߂�

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : tmrAplTimer_Timer
'//  �@�\����  : �A�v���N���`�F�b�N�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v  : �^�C���A�b�v���ɃA�v���N����Ԃ��`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  : (EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή��y����TR-No.123�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrAplTimer_Timer()

    Dim bIDURet As Boolean
    Dim bLDURet As Boolean

    On Error Resume Next

    If CheckAppStart(PROC_KANRI) <> 0 Then
        If lngtime >= lngAplMAX_Time Then
            tmrAplTimer.Enabled = False
            '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
            '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)

' EG20 V5.2.0.1�폜�J�n
'            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
'                    vbOKOnly + vbExclamation, _
'                    "LD���[�e�B���e�B�o�[�W�����Ǘ�"
' EG20 V5.2.0.1�폜�I��
' EG20 V5.2.0.1�ǉ��J�n
            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
                    vbOKOnly + vbExclamation, _
                    "�k�c�t�o�[�W�����Ǘ�"
' EG20 V5.2.0.1�ǉ��I��
            sCmdBtnEnabled True                         ' ��ʑ����
        Else
            '�^�C�}���蒼��
            tmrAplTimer.Interval = MN_MAIL_INTERVAL
            lngtime = lngtime + MN_MAIL_INTERVAL
        End If
    Else
        tmrAplTimer.Enabled = False
        If CheckAppStart(PROCESS_IDU_LOG) <> 0 Then
            bIDURet = EndIDULog 'IDU���O�N������IDU���O�ɑ΂��ă��O�I���v��CMD���M
        Else
            bIDURet = True
        End If

        If CheckAppStart(PROCESS_LDU_LOG) <> 0 Then
            bLDURet = EndLDULog  'LDU���O�N������LDU���O�ɑ΂��ă��O�I���v��CMD���M
        Else
            bLDURet = True
        End If

        If bIDURet = True And bLDURet = True Then
            lngtime = MN_MAIL_INTERVAL
            tmrLogTimer.Enabled = True
        Else
            '�Ǘ��AIDU���O�ALDU���O���I�����Ă��Ȃ���΁A�I�������ُ�
            '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
' EG20 V5.2.0.1�폜�J�n
'            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
'                    vbOKOnly + vbExclamation, _
'                    "LD���[�e�B���e�B�o�[�W�����Ǘ�"
' EG20 V5.2.0.1�폜�I��
' EG20 V5.2.0.1�ǉ��J�n
            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
                    vbOKOnly + vbExclamation, _
                    "�k�c�t�o�[�W�����Ǘ�"
' EG20 V5.2.0.1�ǉ��I��
            sCmdBtnEnabled True                         ' ��ʑ����
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : tmrLogTimer_Timer
'//  �@�\����  : ���O�N���`�F�b�N�^�C�}����
'//  �@�\�T�v  : ���O�N���`�F�b�N�^�C�}�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  : (EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//               EG20�t�F�[�Y�Q�Ή��y����TR-No.123�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrLogTimer_Timer()

    On Error Resume Next

    If CheckAppStart(PROCESS_IDU_LOG) <> 0 _
        Or CheckAppStart(PROCESS_LDU_LOG) <> 0 Then

        If lngtime >= lngLogMAX_Time Then
            '���O�N���`�F�b�N�^�C�}���~����B
            tmrLogTimer.Enabled = False
            '�u�A�v���N���E�I����ʁF�A�v���I�������ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, APL_END_ERROR, 0)
' EG20 V5.2.0.1�폜�J�n
'            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
'                    vbOKOnly + vbExclamation, _
'                    "LD���[�e�B���e�B�o�[�W�����Ǘ�"
' EG20 V5.2.0.1�폜�I��
' EG20 V5.2.0.1�ǉ��J�n
            MsgBox "�A�v���P�[�V�����̏I���������Ɉُ킪�������܂����B", _
                    vbOKOnly + vbExclamation, _
                    "�k�c�t�o�[�W�����Ǘ�"
' EG20 V5.2.0.1�ǉ��I��
            sCmdBtnEnabled True                         ' ��ʑ����
        Else
            '�^�C�}���蒼��
            tmrLogTimer.Interval = MN_MAIL_INTERVAL
            lngtime = lngtime + MN_MAIL_INTERVAL
        End If
    Else
        tmrLogTimer.Enabled = False
        '�u�A�v���N���E�I����ʁF�A�v���I����������v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, APL_END_OK, 0)

        '�ؑփc�[���N��
        Call AplVersionChangeProc(lngChangeKind)

        '�I������
        psEndHoshuProc
        '�ێ�v���Z�X�I��
        End
    End If
End Sub

